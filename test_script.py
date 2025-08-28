# -*- coding: utf-8 -*-
"""
Test script for Transportation Analysis
Demonstrates usage and provides sample data for testing
"""

import pandas as pd
import datetime
from transportation_analysis_improved import TransportationAnalyzer

def create_sample_data():
    """Create sample data for testing"""
    
    # Sample carts data
    carts_data = {
        'Region': ['NA', 'NA', 'EU', 'EU'],
        'Type': ['FC', 'SC', 'DS', 'FC'],
        'Site': ['FC001', 'SC001', 'DS001', 'FC002'],
        'Consumption_1': [100, 50, 75, 120],
        'Consumption_2': [110, 55, 80, 125],
        'Consumption_3': [95, 45, 70, 115],
        'Needed80_1': [10, 5, 8, 12],
        'Needed80_2': [11, 6, 9, 13],
        'Needed80_3': [10, 5, 7, 12],
        'Needed93_1': [8, 4, 6, 10],
        'Needed93_2': [9, 5, 7, 11],
        'Needed93_3': [8, 4, 6, 10]
    }
    
    # Sample FMC data
    fmc_data = {
        'Stop': ['FC001', 'FC001', 'SC001', 'DS001', 'FC002', 'FC002'],
        'Action Type': ['DROPOFF', 'DROPOFF', 'DROPOFF', 'PICKUP', 'DROPOFF', 'DROPOFF'],
        'Planned Dock Arrival': [45898, 45899, 45898, 45898, 45898, 45899]
    }
    
    return pd.DataFrame(carts_data), pd.DataFrame(fmc_data)

def test_analyzer():
    """Test the TransportationAnalyzer with sample data"""
    
    print("Creating sample data...")
    carts_df, fmc_df = create_sample_data()
    
    # Save sample data to Excel files
    carts_df.to_excel('sample_carts.xlsx', sheet_name='Final_Sheet', index=False)
    fmc_df.to_excel('sample_fmc.xlsx', sheet_name='fmc-data-source(5)', index=False)
    
    print("Sample data saved to Excel files")
    
    # Create analyzer with sample files
    analyzer = TransportationAnalyzer('sample_carts.xlsx', 'sample_fmc.xlsx', 45898)
    
    # Test individual methods
    print("\nTesting individual methods...")
    
    # Test data loading
    if analyzer.load_carts_data():
        print("✓ Carts data loaded successfully")
    else:
        print("✗ Failed to load carts data")
    
    if analyzer.load_fmc_data():
        print("✓ FMC data loaded successfully")
    else:
        print("✗ Failed to load FMC data")
    
    # Test analysis
    if analyzer.final_df is not None and analyzer.fmc_df is not None:
        table_df = analyzer.analyze_data()
        print(f"✓ Analysis completed. Generated {len(table_df)} records")
        
        # Display sample results
        print("\nSample analysis results:")
        print(table_df.head())
        
        # Test HTML generation
        html_body = analyzer.generate_html_report(table_df)
        print("✓ HTML report generated")
        
        # Save HTML report
        with open('test_report.html', 'w', encoding='utf-8') as f:
            f.write(html_body)
        print("✓ Test report saved to test_report.html")
        
    else:
        print("✗ Cannot run analysis - data not loaded")
    
    print("\nTest completed!")

def test_date_conversion():
    """Test the date conversion functionality"""
    
    print("Testing date conversion...")
    
    analyzer = TransportationAnalyzer('', '', 45898)
    
    # Test some known dates
    test_dates = [
        (45898, "2025-08-28"),
        (45899, "2025-08-29"),
        (45900, "2025-08-30")
    ]
    
    for serial, expected in test_dates:
        converted = analyzer.serial_to_date(serial)
        result = converted.strftime('%Y-%m-%d')
        if result == expected:
            print(f"✓ {serial} -> {result}")
        else:
            print(f"✗ {serial} -> {result} (expected {expected})")

if __name__ == "__main__":
    print("Transportation Analysis Test Script")
    print("=" * 40)
    
    # Test date conversion
    test_date_conversion()
    
    print("\n" + "=" * 40)
    
    # Test full analyzer
    test_analyzer()