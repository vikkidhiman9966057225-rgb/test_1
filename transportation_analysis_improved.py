# -*- coding: utf-8 -*-
"""
Transportation Analysis Script - Improved Version
Analyzes cart consumption vs scheduled loads and sends email reports

@author: dhimv
@version: 2.0
"""

import pandas as pd
import math
import datetime
import logging
import sys
import os
from math import floor
from typing import List, Dict, Optional

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('transportation_analysis.log'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

class TransportationAnalyzer:
    """Main class for transportation analysis"""
    
    def __init__(self, carts_file: str, fmc_file: str, current_serial: int = 45898):
        """
        Initialize the analyzer
        
        Args:
            carts_file: Path to carts consumption Excel file
            fmc_file: Path to FMC data Excel file
            current_serial: Current Excel serial date
        """
        self.carts_file = carts_file
        self.fmc_file = fmc_file
        self.current_serial = current_serial
        self.final_df = None
        self.fmc_df = None
        
    def serial_to_date(self, serial: int) -> datetime.datetime:
        """
        Convert Excel serial date to actual date
        
        Args:
            serial: Excel serial date
            
        Returns:
            datetime object
        """
        base = datetime.datetime(1900, 1, 1)
        delta = datetime.timedelta(days=int(serial) - 2)  # Adjust for Excel 1900 leap bug
        return base + delta
    
    def load_carts_data(self) -> bool:
        """
        Load carts consumption data from Excel file
        
        Returns:
            True if successful, False otherwise
        """
        try:
            logger.info(f"Loading carts data from: {self.carts_file}")
            self.final_df = pd.read_excel(self.carts_file, sheet_name='Final_Sheet')
            logger.info(f"Successfully loaded {len(self.final_df)} rows from carts data")
            return True
        except Exception as e:
            logger.error(f"Error reading Carts_Consumption.xlsx: {e}")
            return False
    
    def load_fmc_data(self) -> bool:
        """
        Load FMC data from Excel file
        
        Returns:
            True if successful, False otherwise
        """
        try:
            logger.info(f"Loading FMC data from: {self.fmc_file}")
            self.fmc_df = pd.read_excel(self.fmc_file, sheet_name='fmc-data-source(5)')
            
            # Ensure Planned Dock Arrival is numeric
            self.fmc_df['Planned Dock Arrival'] = pd.to_numeric(
                self.fmc_df['Planned Dock Arrival'], errors='coerce'
            )
            
            # Convert to date
            self.fmc_df['Date'] = self.fmc_df['Planned Dock Arrival'].apply(
                lambda x: floor(x) if not pd.isna(x) else None
            )
            
            logger.info(f"Successfully loaded {len(self.fmc_df)} rows from FMC data")
            return True
        except Exception as e:
            logger.error(f"Error reading fmc-data-source(5).xlsx: {e}")
            return False
    
    def get_date_range(self) -> tuple:
        """
        Get the date range for analysis
        
        Returns:
            Tuple of (date_serials, date_strings)
        """
        # Dates: columns 3 to 9: 45894 to 45900
        date_serials = [45894 + i for i in range(7)]
        date_strs = [self.serial_to_date(d).strftime('%m/%d/%Y') for d in date_serials]
        
        # Filter for current and future dates
        future_dates = [d for d in date_serials if d >= self.current_serial]
        logger.info(f"Analyzing {len(future_dates)} future dates from {len(date_serials)} total dates")
        
        return date_serials, date_strs
    
    def count_scheduled_trucks(self, site: str, site_type: str, serial_date: int) -> int:
        """
        Count scheduled trucks for a specific site and date
        
        Args:
            site: Site name
            site_type: Type of site (FC, SC, DS)
            serial_date: Excel serial date
            
        Returns:
            Number of scheduled trucks
        """
        if site_type in ['FC', 'SC']:
            # Count DROPOFF to site on that date
            count = len(self.fmc_df[
                (self.fmc_df['Stop'] == site) &
                (self.fmc_df['Action Type'] == 'DROPOFF') &
                (self.fmc_df['Date'] == serial_date)
            ])
        elif site_type == 'DS':
            # Count PICKUP from site on that date
            count = len(self.fmc_df[
                (self.fmc_df['Stop'] == site) &
                (self.fmc_df['Action Type'] == 'PICKUP') &
                (self.fmc_df['Date'] == serial_date)
            ])
        else:
            count = 0
            
        return count
    
    def analyze_data(self) -> pd.DataFrame:
        """
        Perform the main analysis
        
        Returns:
            DataFrame with analysis results
        """
        if self.final_df is None or self.fmc_df is None:
            logger.error("Data not loaded. Please load data first.")
            return pd.DataFrame()
        
        date_serials, date_strs = self.get_date_range()
        table_data = []
        
        logger.info("Starting data analysis...")
        
        for idx, row in self.final_df.iterrows():
            region = row['Region']
            site_type = row['Type']
            site = row['Site']
            
            # Only process FC, SC, DS site types
            if site_type not in ['FC', 'SC', 'DS']:
                continue
            
            for i, serial_date in enumerate(date_serials):
                if serial_date < self.current_serial:
                    continue
                
                # Get needed trucks
                needed80_col = 10 + i  # K:Q are 10 to 16
                needed93_col = 17 + i  # R:X are 17 to 23
                needed80 = row.iloc[needed80_col] if not pd.isna(row.iloc[needed80_col]) else 0
                needed93 = row.iloc[needed93_col] if not pd.isna(row.iloc[needed93_col]) else 0
                
                # Count scheduled trucks
                count_have = self.count_scheduled_trucks(site, site_type, serial_date)
                
                # Calculate difference
                difference = round(needed80 - count_have, 2)
                
                date_str = date_strs[i]
                
                table_data.append({
                    'Region': region,
                    'Type': site_type,
                    'Site': site,
                    'Date': date_str,
                    'Needed TFR 80': round(needed80, 2),
                    'Needed TFR 93': round(needed93, 2),
                    'Scheduled Trucks': count_have,
                    'Difference': difference
                })
        
        logger.info(f"Analysis complete. Generated {len(table_data)} data points")
        return pd.DataFrame(table_data)
    
    def generate_html_report(self, table_df: pd.DataFrame) -> str:
        """
        Generate HTML report from analysis results
        
        Args:
            table_df: DataFrame with analysis results
            
        Returns:
            HTML string
        """
        # Generate HTML table with styling
        html_table = table_df.to_html(
            index=False, 
            border=1, 
            classes='table table-striped',
            formatters={
                'Needed TFR 80': lambda x: f"{x:.2f}",
                'Needed TFR 93': lambda x: f"{x:.2f}",
                'Difference': lambda x: f"{x:.2f}"
            }
        )
        
        # Add CSS for better presentation
        html_body = f"""
        <html>
        <head>
        <style>
        .table {{
            border-collapse: collapse;
            width: 100%;
            font-family: Arial, sans-serif;
            margin: 20px 0;
        }}
        .table th, .table td {{
            border: 1px solid #ddd;
            padding: 12px;
            text-align: left;
        }}
        .table th {{
            background-color: #4CAF50;
            color: white;
            font-weight: bold;
        }}
        .table tr:nth-child(even) {{
            background-color: #f9f9f9;
        }}
        .table tr:hover {{
            background-color: #f5f5f5;
        }}
        .positive {{
            color: green;
            font-weight: bold;
        }}
        .negative {{
            color: red;
            font-weight: bold;
        }}
        .summary {{
            background-color: #e7f3ff;
            padding: 15px;
            border-radius: 5px;
            margin: 20px 0;
        }}
        </style>
        </head>
        <body>
        <div class="summary">
            <h2>Transportation Analysis Report</h2>
            <p><strong>Generated:</strong> {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
            <p><strong>Total Records:</strong> {len(table_df)}</p>
            <p><strong>Analysis Period:</strong> {table_df['Date'].min()} to {table_df['Date'].max()}</p>
        </div>
        
        <p>Hi Team,</p>
        <p>Please review the scheduled loads as per consumption. The table below shows the comparison between needed and scheduled trucks for each site.</p>
        
        {html_table}
        
        <div class="summary">
            <h3>Summary</h3>
            <p><strong>Sites with Shortage (Negative Difference):</strong> {len(table_df[table_df['Difference'] < 0])}</p>
            <p><strong>Sites with Surplus (Positive Difference):</strong> {len(table_df[table_df['Difference'] > 0])}</p>
            <p><strong>Sites with Perfect Match:</strong> {len(table_df[table_df['Difference'] == 0])}</p>
        </div>
        
        <p>Best Regards,<br>Vipin Dhiman</p>
        </body>
        </html>
        """
        
        return html_body
    
    def send_email(self, html_body: str, to_email: str = "dhimv@amazon.com", 
                   cc_email: str = "test@amazon.com") -> bool:
        """
        Send email with analysis results
        
        Args:
            html_body: HTML content of the email
            to_email: Recipient email address
            cc_email: CC email address
            
        Returns:
            True if successful, False otherwise
        """
        try:
            # Try to import win32com for Windows
            import win32com.client
            outlook = win32com.client.Dispatch("Outlook.Application")
            
            mail = outlook.CreateItem(0)  # Create a new email
            mail.To = to_email
            mail.CC = cc_email
            mail.Subject = f"Transportation Analysis Report - {datetime.datetime.now().strftime('%Y-%m-%d')}"
            mail.HTMLBody = html_body
            mail.Send()
            
            logger.info("Email sent successfully!")
            return True
            
        except ImportError:
            logger.warning("win32com not available. Email functionality disabled.")
            logger.info("HTML report generated. Please send manually.")
            return False
        except Exception as e:
            logger.error(f"Error sending email: {e}")
            return False
    
    def run_analysis(self) -> bool:
        """
        Run the complete analysis workflow
        
        Returns:
            True if successful, False otherwise
        """
        try:
            # Load data
            if not self.load_carts_data():
                return False
            if not self.load_fmc_data():
                return False
            
            # Perform analysis
            table_df = self.analyze_data()
            if table_df.empty:
                logger.error("No data to analyze")
                return False
            
            # Generate report
            html_body = self.generate_html_report(table_df)
            
            # Save report to file
            with open('transportation_report.html', 'w', encoding='utf-8') as f:
                f.write(html_body)
            logger.info("Report saved to transportation_report.html")
            
            # Send email
            self.send_email(html_body)
            
            return True
            
        except Exception as e:
            logger.error(f"Error in analysis workflow: {e}")
            return False


def main():
    """Main function"""
    # File paths (local)
    carts_file = r'\\ant\fc\Dept\Transportation\LTN2\ATS\Reverse_Logistics\maseerat\TRAX\Overbias & Underbias_Trigger\Carts_Consumption.xlsx'
    fmc_file = r'\\ant\fc\Dept\Transportation\LTN2\ATS\Reverse_Logistics\maseerat\TRAX\Overbias & Underbias_Trigger\fmc-data-source(5).xlsx'
    
    # Current serial date (August 28, 2025)
    current_serial = 45898
    
    # Create analyzer and run analysis
    analyzer = TransportationAnalyzer(carts_file, fmc_file, current_serial)
    
    if analyzer.run_analysis():
        logger.info("Analysis completed successfully!")
    else:
        logger.error("Analysis failed!")
        sys.exit(1)


if __name__ == "__main__":
    main()