#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Driver Weekly Data Analysis Program - Version 3
Focus on IL region drivers with pinyin matching and fuel cost analysis
"""

import pandas as pd
import sys
import os
import argparse
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# Set environment for proper Unicode handling
def setup_console_encoding():
    """Setup console encoding for Windows compatibility"""
    try:
        if sys.platform.startswith('win'):
            os.environ['PYTHONIOENCODING'] = 'utf-8'
            # For exe compatibility, use safe encoding
            import locale
            locale.setlocale(locale.LC_ALL, 'C.UTF-8')
    except Exception:
        # Fallback if encoding setup fails
        pass

setup_console_encoding()

def read_driver_data(file_path):
    """
    Read driver weekly data file (dwd_929.xlsx or 921-925dwd.xlsx)
    Expected columns: Driver Name, Hours, Tasks, Mileage, Date(?)
    """
    try:
        df = pd.read_excel(file_path)

        # Assign column names based on analysis
        df.columns = ['Driver_Name', 'Work_Hours', 'Tasks', 'Mileage', 'Date_Column']

        # Group by driver to get totals
        driver_summary = df.groupby('Driver_Name').agg({
            'Work_Hours': 'sum',
            'Tasks': 'sum',
            'Mileage': 'sum'
        }).reset_index()
        return driver_summary

    except Exception as e:
        print(f"Error reading driver file: {e}")
        return None

def read_transaction_data(file_path):
    """
    Read fuel transaction data file and filter for IL region
    """
    try:
        df = pd.read_excel(file_path)
        # Filter for IL region only
        if 'SiteState' in df.columns:
            il_df = df[df['SiteState'] == 'IL'].copy()
            print(f"Transaction data: {len(il_df)} IL transactions, {il_df['Driver'].nunique()} drivers")
        else:
            print("Warning: No SiteState column, using all transactions")
            il_df = df.copy()

        # Convert date format
        if 'TransDate' in il_df.columns:
            il_df['TransDate'] = pd.to_datetime(il_df['TransDate'], format='%m/%d/%y')

        return il_df

    except Exception as e:
        print(f"Error reading transaction file: {e}")
        return None

def create_chinese_to_pinyin_mapping():
    """
    Create a mapping between Chinese names and their pinyin equivalents
    Based on the IL drivers found in the transaction data
    Note: Chinese names are in format "姓名" (Last First), English names are in format "FIRST LAST"
    """
    # Mapping based on actual IL drivers found in transaction data
    name_mapping = {
        '王洋': 'YANG WANG',        # 王(Wang) 洋(Yang)
        '朱海山': 'HAISHAN ZHU',    # 朱(Zhu) 海山(Haishan)
        '韩境烨': 'JINGYE HAN',     # 韩(Han) 境烨(Jingye)
        '安博': 'BO AN',            # 安(An) 博(Bo) -> BO AN (corrected)
        '邢旭': 'XU XING',
        '李健': 'JIAN LI',          # 李(Li) 健(Jian)
        '刘增榕': 'ZENGRONG LIU',   # 刘(Liu) 增榕(Zengrong)
        '陈磊': 'LEI CHEN',         # 陈(Chen) 磊(Lei)
        '吴波': 'BO WU',            # 吴(Wu) 波(Bo)
    }
    return name_mapping

def match_driver_names(driver_name, il_transaction_data):
    """
    Match driver names between Chinese names and IL region transaction drivers
    """
    if il_transaction_data is None or len(il_transaction_data) == 0:
        return None

    il_drivers = il_transaction_data['Driver'].unique()

    # Use predefined mapping for Chinese to Pinyin
    name_mapping = create_chinese_to_pinyin_mapping()
    if driver_name in name_mapping:
        mapped_name = name_mapping[driver_name]
        if mapped_name in il_drivers:
            return mapped_name

    # Exact match (unlikely but check anyway)
    if driver_name in il_drivers:
        return driver_name

    return None

def calculate_fuel_statistics(driver_transactions):
    """
    Calculate detailed fuel statistics for a driver
    """
    if driver_transactions is None or len(driver_transactions) == 0:
        return 0, 0.0, 0.0

    fuel_count = len(driver_transactions)
    total_cost = driver_transactions['TotalAmount'].sum() if 'TotalAmount' in driver_transactions.columns else 0.0
    total_quantity = driver_transactions['Quantity'].sum() if 'Quantity' in driver_transactions.columns else 0.0

    return fuel_count, round(total_cost, 2), round(total_quantity, 2)


def calculate_performance_metrics(total_hours, total_tasks, total_mileage):
    """
    Calculate basic efficiency metrics and balance score
    """
    # Calculate task and mileage efficiency
    task_efficiency = round(total_tasks / total_hours, 2) if total_hours > 0 else 0
    mileage_efficiency = round(total_mileage / total_hours, 2) if total_hours > 0 else 0

    # Calculate efficiency balance score (geometric mean)
    import math
    efficiency_balance = round(math.sqrt(task_efficiency * mileage_efficiency), 2) if task_efficiency > 0 and mileage_efficiency > 0 else 0

    # Calculate new pickup-specific metrics
    avg_pickup_distance = round(total_mileage / total_tasks, 2) if total_tasks > 0 else 0  # miles per pickup
    avg_time_per_task = round(total_hours / total_tasks, 2) if total_tasks > 0 else 0      # hours per pickup
    workload_index = round((total_tasks * total_mileage) / total_hours, 2) if total_hours > 0 else 0  # comprehensive workload

    # Calculate daily workload density (assuming 6 working days per week)
    working_days = 6
    daily_workload_density = round((total_tasks + total_mileage / 10) / working_days, 2) if working_days > 0 else 0

    return {
        'Task_Efficiency': task_efficiency,        # tasks/hour
        'Mileage_Efficiency': mileage_efficiency,  # miles/hour
        'Efficiency_Balance': efficiency_balance,  # sqrt(task_eff * mile_eff)
        'Avg_Pickup_Distance': avg_pickup_distance,  # miles/task
        'Avg_Time_Per_Task': avg_time_per_task,      # hours/task
        'Workload_Index': workload_index,            # (tasks * miles) / hour
        'Daily_Workload_Density': daily_workload_density,  # (tasks + miles/10) / 6 days
    }

def analyze_driver_work(driver_data, transaction_data):
    """
    Analyze driver work data combining both files with performance metrics
    """
    if driver_data is None:
        return None

    results = []

    print(f"\nAnalyzing {len(driver_data)} drivers:")

    # Process each driver from the main driver file
    for idx, row in driver_data.iterrows():
        driver_name = row['Driver_Name']
        total_hours = row['Work_Hours']
        total_tasks = row['Tasks']
        total_mileage = row['Mileage']

        # Try to match with IL transaction data
        matched_driver = match_driver_names(driver_name, transaction_data)
        fuel_count = 0
        fuel_total_cost = 0.0
        fuel_quantity = 0.0

        if matched_driver and transaction_data is not None:
            driver_transactions = transaction_data[transaction_data['Driver'] == matched_driver]
            fuel_count, fuel_total_cost, fuel_quantity = calculate_fuel_statistics(driver_transactions)
            print(f"  Driver {idx+1} -> {matched_driver}: {fuel_count} trans, ${fuel_total_cost}")
        else:
            print(f"  Driver {idx+1}: no match")

        # Calculate performance metrics
        metrics = calculate_performance_metrics(total_hours, total_tasks, total_mileage)

        result = {
            'Driver': driver_name,
            'Total_Hours': round(total_hours, 2),
            'Total_Tasks': total_tasks,
            'Total_Mileage': total_mileage,
            'Fuel_Count': fuel_count,
            'Fuel_Total_Cost': fuel_total_cost,
            'Fuel_Quantity': fuel_quantity,
            'Matched_Transaction_Driver': matched_driver if matched_driver else 'No Match'
        }

        # Add all performance metrics
        result.update(metrics)

        results.append(result)

    return pd.DataFrame(results)

def generate_summary_report(analysis_results):
    """
    Generate summary report
    """
    if analysis_results is None or len(analysis_results) == 0:
        print("No analysis results available for report generation")
        return

    print("\n" + "="*70)
    print("Driver Weekly Work Analysis Report (IL Region Focus)")
    print("="*70)

    # Overall statistics
    total_drivers = len(analysis_results)
    total_hours = analysis_results['Total_Hours'].sum()
    total_tasks = analysis_results['Total_Tasks'].sum()
    total_mileage = analysis_results['Total_Mileage'].sum()
    total_fuel_transactions = analysis_results['Fuel_Count'].sum()
    total_fuel_cost = analysis_results['Fuel_Total_Cost'].sum()
    total_fuel_quantity = analysis_results['Fuel_Quantity'].sum()
    drivers_with_fuel = len(analysis_results[analysis_results['Fuel_Count'] > 0])

    print(f"\n[Overall Statistics]")
    print(f"Total Drivers: {total_drivers}")
    print(f"Total Work Hours: {total_hours:.2f}")
    print(f"Total Tasks: {total_tasks}")
    print(f"Total Mileage: {total_mileage}")
    print(f"Total Fuel Transactions: {total_fuel_transactions}")
    print(f"Total Fuel Cost: ${total_fuel_cost:.2f}")
    print(f"Total Fuel Quantity: {total_fuel_quantity:.2f} gallons")
    print(f"Drivers with IL Fuel Records: {drivers_with_fuel}/{total_drivers}")
    print(f"Average Hours per Driver: {total_hours/total_drivers:.2f}")
    print(f"Average Fuel Cost per Driver: ${total_fuel_cost/total_drivers:.2f}")

    print(f"\n[Individual Driver Summary with Performance Metrics]")
    for idx, row in analysis_results.iterrows():
        print(f"\nDriver: {row['Driver']}")
        print(f"  Basic Data:")
        print(f"    Work Hours: {row['Total_Hours']} | Tasks: {row['Total_Tasks']} | Mileage: {row['Total_Mileage']}")
        print(f"    Fuel: {row['Fuel_Count']} transactions, ${row['Fuel_Total_Cost']}")

        print(f"  Efficiency Metrics:")
        print(f"    Task Efficiency: {row['Task_Efficiency']} tasks/hour")
        print(f"    Mileage Efficiency: {row['Mileage_Efficiency']} miles/hour")

        print(f"  Cost Metrics:")
        print(f"    Mileage Cost: ${row['Mileage_Cost']}/mile")
        print(f"    Task Cost: ${row['Task_Cost']}/task")
        print(f"    Time Cost: ${row['Time_Cost']}/hour")

        print(f"  Vehicle Fit Indicators:")
        print(f"    Avg Task Distance: {row['Avg_Task_Mileage']} miles/task")

        if row['Matched_Transaction_Driver'] != 'No Match':
            print(f"  Matched with: {row['Matched_Transaction_Driver']}")

def save_results_to_excel(analysis_results, filename="921-925dwa.xlsx"):
    """
    Save results to Excel file with comprehensive metrics
    """
    if analysis_results is not None and len(analysis_results) > 0:
        # Create the final output with basic data and efficiency metrics
        output_columns = [
            'Driver', 'Total_Hours', 'Total_Tasks', 'Total_Mileage',
            'Fuel_Count', 'Fuel_Total_Cost',
            'Task_Efficiency', 'Mileage_Efficiency', 'Efficiency_Balance',
            'Avg_Pickup_Distance', 'Avg_Time_Per_Task', 'Workload_Index',
            'Daily_Workload_Density'
        ]

        output_df = analysis_results[output_columns].copy()

        # Rename columns to Chinese with calculation formulas
        output_df.columns = [
            '司机',
            '总时长(小时)',
            '总任务(个)',
            '总里程(英里)',
            '加油次数',
            '加油总耗费($)',
            '任务效率(任务/小时)',
            '里程效率(英里/小时)',
            '效率平衡分数(√(任务效率×里程效率))',
            '平均揽收间距(英里/任务)',
            '每任务时长(小时/任务)',
            '工作负荷指数(任务×里程/小时)',
            '综合工作密度((任务+里程/10)/6天)'
        ]

        # Save to Excel
        output_df.to_excel(filename, index=False, engine='openpyxl')
        print(f"Results saved to: {filename}")

    else:
        print("No results to save")

def extract_date_prefix(filename):
    """
    Extract date prefix from driver data filename
    Supports both old format ('921-925dwd.xlsx') and new format ('dwd_929.xlsx')
    Example: 'dwd_929.xlsx' -> '929'
    Example: '921-925dwd.xlsx' -> '921-925'
    """
    import re

    # Extract base filename without extension
    base_name = os.path.splitext(os.path.basename(filename))[0]

    # New format: 'dwd_929' or 'transaction_929'
    pattern_new = r'(?:dwd|transaction)_(\d{3,6})'
    match = re.search(pattern_new, base_name, re.IGNORECASE)
    if match:
        return match.group(1)

    # Old format: '921-925dwd' (date range at the beginning)
    pattern_old = r'^(\d{3,4}-\d{3,4})'
    match = re.match(pattern_old, base_name)
    if match:
        return match.group(1)

    # Fallback: if base name ends with 'dwd', try to extract before it
    if base_name.endswith('dwd'):
        return base_name[:-3]

    # Last fallback: use first part before any non-digit characters
    match = re.match(r'^([^a-zA-Z]+)', base_name)
    if match:
        return match.group(1).rstrip('-_')

    return '929'  # Default fallback

def parse_arguments():
    """
    Parse command line arguments
    """
    parser = argparse.ArgumentParser(
        description='Driver Weekly Data Analysis (DWA) Program',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  process_dwa.exe -f dwd_929.xlsx transaction_929.xlsx
  process_dwa.exe -f 921-925dwd.xlsx Transactions.xlsx
  process_dwa.exe -f dwd_929.xlsx transaction_929.xlsx -o analysis_929.xlsx
        """
    )

    parser.add_argument('-f', '--files', nargs=2, required=True,
                        metavar=('DRIVER_FILE', 'TRANSACTION_FILE'),
                        help='Driver data file and Transaction data file (Excel format)')

    parser.add_argument('-o', '--output', default=None,
                        help='Output file name (default: auto-generated from driver file date prefix)')

    parser.add_argument('-v', '--version', action='version', version='DWA 3.0')

    return parser.parse_args()

def main():
    """
    Main function with command line support
    """
    # Parse command line arguments
    args = parse_arguments()

    # File paths from command line
    driver_file = args.files[0]
    transaction_file = args.files[1]

    # Generate output filename if not provided
    if args.output is None:
        date_prefix = extract_date_prefix(driver_file)
        output_file = f"dwa_{date_prefix}.xlsx"
    else:
        output_file = args.output

    print("=== Driver Weekly Data Analysis ===")
    print(f"Input: {args.files[0]}, {args.files[1]}")
    print(f"Output: {output_file}")

    print("\nReading data files...")
    driver_data = read_driver_data(driver_file)
    transaction_data = read_transaction_data(transaction_file)

    # Check if files were read successfully
    if driver_data is None:
        print(f"Error: Cannot read driver file '{driver_file}'")
        sys.exit(1)

    if transaction_data is None:
        print(f"Error: Cannot read transaction file '{transaction_file}'")
        sys.exit(1)

    print("\nAnalyzing data...")
    analysis_results = analyze_driver_work(driver_data, transaction_data)

    # Generate report and save results
    if analysis_results is not None:
        save_results_to_excel(analysis_results, output_file)
        matched_drivers = len(analysis_results[analysis_results['Fuel_Count'] > 0])
        print(f"\nComplete! {len(analysis_results)} drivers, {matched_drivers} with fuel records")
    else:
        print("Error: Analysis failed")
        sys.exit(1)

if __name__ == "__main__":
    main()