# -*- coding: utf-8 -*-
"""
Created on Thu Aug 28 15:18:08 2025

@author: dhimv
"""

# -*- coding: utf-8 -*-
"""
Created on Thu Aug 28 17:51:00 2025
@author: dhimv
"""
import pandas as pd
import math
import datetime
import win32com.client
from math import floor
# File paths (local)
carts_file = r'\\ant\fc\Dept\Transportation\LTN2\ATS\Reverse_Logistics\maseerat\TRAX\Overbias & Underbias_Trigger\Carts_Consumption.xlsx'
fmc_file = r'\\ant\fc\Dept\Transportation\LTN2\ATS\Reverse_Logistics\maseerat\TRAX\Overbias & Underbias_Trigger\fmc-data-source(5).xlsx'
# Current serial date (August 28, 2025)
current_serial = 45898
# Function to convert Excel serial date to actual date
def serial_to_date(serial):
    base = datetime.datetime(1900, 1, 1)
    delta = datetime.timedelta(days=int(serial) - 2) # Adjust for Excel 1900 leap bug
    return base + delta
# Read Final_Sheet
try:
    final_df = pd.read_excel(carts_file, sheet_name='Final_Sheet')
except Exception as e:
    print(f"Error reading Carts_Consumption.xlsx: {e}")
    exit()
# The consumption columns are D:J (index 3:9), needed80 K:Q (10:16), needed93 R:X (17:23)
# Dates: columns 3 to 9: 45894 to 45900
date_serials = [45894 + i for i in range(7)]
date_strs = [serial_to_date(d).strftime('%m/%d/%Y') for d in date_serials]
# Filter for current and future dates (August 28, 2025 onward)
future_dates = [d for d in date_serials if d >= current_serial]
future_indices = [date_serials.index(d) for d in future_dates]
# Read FMC data
try:
    fmc_df = pd.read_excel(fmc_file, sheet_name='fmc-data-source(5)')
except Exception as e:
    print(f"Error reading fmc-data-source(5).xlsx: {e}")
    exit()
# Ensure Planned Dock Arrival is float
fmc_df['Planned Dock Arrival'] = pd.to_numeric(fmc_df['Planned Dock Arrival'], errors='coerce')
# Group to count per site per date per action
fmc_df['Date'] = fmc_df['Planned Dock Arrival'].apply(lambda x: floor(x) if not pd.isna(x) else None)
# Prepare table data
table_data = []
for idx, row in final_df.iterrows():
    region = row['Region']
    site_type = row['Type']
    site = row['Site']
   
    # Only process FC, SC, DS site types
    if site_type not in ['FC', 'SC', 'DS']:
        continue
   
    for i, serial_date in enumerate(date_serials):
        if serial_date < current_serial:
            continue
       
        # Get needed trucks
        needed80_col = 10 + i # K:Q are 10 to 16
        needed93_col = 17 + i # R:X are 17 to 23
        needed80 = row.iloc[needed80_col] if not pd.isna(row.iloc[needed80_col]) else 0
        needed93 = row.iloc[needed93_col] if not pd.isna(row.iloc[needed93_col]) else 0
       
        # Count 'have' trucks
        if site_type in ['FC', 'SC']:
            # Count DROPOFF to site on that date
            count_have = len(fmc_df[(fmc_df['Stop'] == site) &
                                    (fmc_df['Action Type'] == 'DROPOFF') &
                                    (fmc_df['Date'] == serial_date)])
        elif site_type == 'DS':
            # Count PICKUP from site on that date
            count_have = len(fmc_df[(fmc_df['Stop'] == site) &
                                    (fmc_df['Action Type'] == 'PICKUP') &
                                    (fmc_df['Date'] == serial_date)])
       
        difference = round(needed80 - count_have, 2) # Difference for TFR 80, rounded to 2 decimals
       
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
# Create DataFrame for table
table_df = pd.DataFrame(table_data)
# Generate HTML table with basic styling
html_table = table_df.to_html(index=False, border=1, classes='table table-striped',
                             formatters={
                                 'Needed TFR 80': lambda x: f"{x:.2f}",
                                 'Needed TFR 93': lambda x: f"{x:.2f}",
                                 'Difference': lambda x: f"{x:.2f}"
                             })
# Add CSS for better presentation
html_body = f"""
<html>
<head>
<style>
.table {{
    border-collapse: collapse;
    width: 100%;
    font-family: Arial, sans-serif;
}}
.table th, .table td {{
    border: 1px solid #ddd;
    padding: 8px;
    text-align: left;
}}
.table th {{
    background-color: #f2f2f2;
    font-weight: bold;
}}
.table tr:nth-child(even) {{
    background-color: #f9f9f9;
}}
</style>
</head>
<body>
<p>Hi Team,</p>
<p>Please review the scheduled loads as per consumption.</p>
{html_table}
<p>Best Regards,</p>
<p>Vipin Dhiman</p>
</body>
</html>
"""
# Initialize Outlook
try:
    outlook = win32com.client.Dispatch("Outlook.Application")
except Exception as e:
    print(f"Error initializing Outlook: {e}")
    exit()
# Create and send email
try:
    mail = outlook.CreateItem(0) # Create a new email
    mail.To = "dhimv@amazon.com"
    mail.CC = "test@amazon.com"
    mail.Subject = "Review Scheduled Loads as per Consumption"
    mail.HTMLBody = html_body
    mail.Send()
    print("Email sent successfully!")
except Exception as e:
    print(f"Error sending email: {e}")