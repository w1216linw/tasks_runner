"""
Weekly Order Processing Template
用途: 处理一周的TEMU和YY订单数据
使用方法: 修改下方的文件名和日期，然后运行脚本
"""

import pandas as pd
import sys
import pgeocode
from datetime import timedelta, date

# ============================================================
# 配置区域 - 请修改这里的文件名和日期
# ============================================================
DATE_RANGE = '260216-260222'                                  # 输出文件日期范围 (YYMMDD-YYMMDD)
TEMU_FILES = ['TEMU/TEMU_260216-260219.xlsx', 'TEMU/TEMU_260220-260222.xlsx']  # TEMU文件列表
YY_FILE = f'YY/YY_{DATE_RANGE}.csv'                                # YY文件名

# Set UTF-8 encoding for output
sys.stdout.reconfigure(encoding='utf-8')

# State name to state code mapping
STATE_NAME_TO_CODE = {
    'Alabama': 'AL', 'Alaska': 'AK', 'Arizona': 'AZ', 'Arkansas': 'AR',
    'California': 'CA', 'Colorado': 'CO', 'Connecticut': 'CT', 'Delaware': 'DE',
    'Florida': 'FL', 'Georgia': 'GA', 'Hawaii': 'HI', 'Idaho': 'ID',
    'Illinois': 'IL', 'Indiana': 'IN', 'Iowa': 'IA', 'Kansas': 'KS',
    'Kentucky': 'KY', 'Louisiana': 'LA', 'Maine': 'ME', 'Maryland': 'MD',
    'Massachusetts': 'MA', 'Michigan': 'MI', 'Minnesota': 'MN', 'Mississippi': 'MS',
    'Missouri': 'MO', 'Montana': 'MT', 'Nebraska': 'NE', 'Nevada': 'NV',
    'New Hampshire': 'NH', 'New Jersey': 'NJ', 'New Mexico': 'NM', 'New York': 'NY',
    'North Carolina': 'NC', 'North Dakota': 'ND', 'Ohio': 'OH', 'Oklahoma': 'OK',
    'Oregon': 'OR', 'Pennsylvania': 'PA', 'Rhode Island': 'RI', 'South Carolina': 'SC',
    'South Dakota': 'SD', 'Tennessee': 'TN', 'Texas': 'TX', 'Utah': 'UT',
    'Vermont': 'VT', 'Virginia': 'VA', 'Washington': 'WA', 'West Virginia': 'WV',
    'Wisconsin': 'WI', 'Wyoming': 'WY', 'District of Columbia': 'DC',
    'PUERTO RICO': 'PR', 'Puerto Rico': 'PR'
}

def get_state_info_from_zip(zip_code, nomi):
    """Convert ZIP code to state code and name using pgeocode"""
    if pd.isna(zip_code):
        return ('Unknown', 'Unknown')

    zip_str = str(int(zip_code)).zfill(5)

    # Handle Puerto Rico manually (pgeocode doesn't include PR)
    if zip_str.startswith(('006', '007', '008', '009')):
        return ('PR', 'Puerto Rico')

    location = nomi.query_postal_code(zip_str)

    if pd.notna(location.state_code):
        state_code = location.state_code
        state_name = location.state_name if pd.notna(location.state_name) else state_code
        return (state_code, state_name)
    else:
        return ('Unknown', 'Unknown')

def get_business_date(timestamp):
    """
    Get business date based on 8am cutoff (time window).

    Business date definition: Orders from previous day 8:00 AM to current day 8:00 AM
    are counted as current day's orders.

    Example: To get 11/23 orders
      Time window: 11/22 8:00 AM - 11/23 8:00 AM
      - Orders at/after 11/22 8:00 AM -> counted as 11/23
      - Orders before 11/23 8:00 AM -> counted as 11/23
      - Orders at/after 11/23 8:00 AM -> counted as 11/24

    Logic:
      - If order time < current date 8:00 AM -> return current date
      - If order time >= current date 8:00 AM -> return next date
    """
    if timestamp.hour < 8:
        return timestamp.date()
    else:
        return timestamp.date() + timedelta(days=1)

print("="*80)
print(f"Processing Week {DATE_RANGE} Data")
print("="*80)

# Initialize pgeocode
print("\nInitializing pgeocode for US...")
nomi = pgeocode.Nominatim('us')

# ============================================================
# Process TEMU files (one or more files for this week)
# ============================================================
df_temu_list = []
tracking_numbers = set()

for temu_file in TEMU_FILES:
    print(f"\n--- Processing TEMU: {temu_file} ---")

    df_temu = pd.read_excel(temu_file)
    print(f"  Total TEMU orders: {len(df_temu):,}")

    # Convert date column to datetime
    df_temu['Date'] = pd.to_datetime(df_temu['下单时间'], format='%m/%d/%Y %H:%M:%S')
    df_temu['Date_Only'] = df_temu['Date'].apply(get_business_date)

    # Get unique ZIP codes for efficiency
    unique_zips = df_temu['邮编'].unique()
    print(f"  Processing {len(unique_zips):,} unique ZIP codes...")

    # Create mapping dictionaries
    zip_to_state_code = {}
    zip_to_state_name = {}

    for i, zip_code in enumerate(unique_zips):
        if (i + 1) % 500 == 0:
            print(f"    Processed {i + 1:,} / {len(unique_zips):,} ZIP codes...")
        state_code, state_name = get_state_info_from_zip(zip_code, nomi)
        zip_to_state_code[zip_code] = state_code
        zip_to_state_name[zip_code] = state_name

    # Apply the mapping
    df_temu['State_Code'] = df_temu['邮编'].map(zip_to_state_code)
    df_temu['State_Name'] = df_temu['邮编'].map(zip_to_state_name)

    # Get tracking numbers for deduplication
    # TEMU uses column "运单号" (tracking number) for deduplication
    file_tracking_numbers = set(df_temu['运单号'].dropna())
    tracking_numbers.update(file_tracking_numbers)
    print(f"  TEMU tracking numbers in this file: {len(file_tracking_numbers):,}")

    # Get date range
    temu_min = df_temu['Date_Only'].min()
    temu_max = df_temu['Date_Only'].max()
    print(f"  TEMU date range: {temu_min} to {temu_max}")

    df_temu_list.append(df_temu)

# Combine all TEMU data
df_temu_combined = pd.concat(df_temu_list, ignore_index=True)
print(f"\n  Combined TEMU orders: {len(df_temu_combined):,}")
print(f"  Total unique TEMU tracking numbers: {len(tracking_numbers):,}")

# Count orders by date and state
temu_counts = df_temu_combined.groupby(['Date_Only', 'State_Code', 'State_Name']).size().reset_index(name='Order_Count')

# ============================================================
# Process YY file
# ============================================================
print(f"\n--- Processing YY: {YY_FILE} ---")

df_yy = pd.read_csv(YY_FILE, encoding='utf-16', sep='\t')
print(f"  Total YY orders: {len(df_yy):,}")

# Check column names
print(f"  YY columns: {list(df_yy.columns)}")

# Convert date column BEFORE deduplication to get original counts
df_yy['Date'] = pd.to_datetime(df_yy['预约揽收时间'])
df_yy['Date_Only'] = df_yy['Date'].dt.date

# Save original YY counts by date (before deduplication)
yy_original_counts = df_yy.groupby('Date_Only').size().to_dict()
print(f"  YY original counts by date: {yy_original_counts}")

# Remove orders that exist in TEMU (deduplication)
# YY uses column "单据号" (tracking number) which corresponds to TEMU's "运单号"
# Remove any YY orders where "单据号" exists in TEMU's "运单号" set
df_yy = df_yy[~df_yy['单据号'].isin(tracking_numbers)]
print(f"  After removing TEMU duplicates: {len(df_yy):,} orders")

# Filter weekdays only (Monday=0, Sunday=6)
df_yy['Weekday'] = df_yy['Date'].dt.dayofweek
df_yy = df_yy[df_yy['Weekday'] < 5]  # Keep Monday-Friday (0-4)
print(f"  After weekday filter: {len(df_yy):,} orders")

# Map state names to state codes (case-insensitive)
state_name_lower_to_code = {k.lower(): v for k, v in STATE_NAME_TO_CODE.items()}
df_yy['State_Code'] = df_yy['目的地行政洲'].str.strip().str.lower().map(state_name_lower_to_code)
df_yy['State_Name'] = df_yy['目的地行政洲'].str.strip()

# Check for unmapped states
unmapped = df_yy[df_yy['State_Code'].isna()]
if len(unmapped) > 0:
    print(f"  WARNING: {len(unmapped)} orders have unmapped state names:")
    for state_name in unmapped['目的地行政洲'].unique():
        print(f"    - {state_name}")
    df_yy.loc[df_yy['State_Code'].isna(), 'State_Code'] = 'Unknown'
    df_yy.loc[df_yy['State_Code'] == 'Unknown', 'State_Name'] = 'Unknown'

# Get date range
if len(df_yy) > 0:
    yy_min = df_yy['Date_Only'].min()
    yy_max = df_yy['Date_Only'].max()
    print(f"  YY date range: {yy_min} to {yy_max}")

# Count orders by date and state
yy_counts = df_yy.groupby(['Date_Only', 'State_Code', 'State_Name']).size().reset_index(name='Order_Count')

# ============================================================
# Combine data
# ============================================================
combined_counts = pd.concat([temu_counts, yy_counts], ignore_index=True)
combined_counts = combined_counts.groupby(['Date_Only', 'State_Code', 'State_Name'])['Order_Count'].sum().reset_index()

# Create pivot table
pivot_table = combined_counts.pivot_table(
    index=['State_Code', 'State_Name'],
    columns='Date_Only',
    values='Order_Count',
    fill_value=0
).astype(int)

# Add Total column
pivot_table['Total'] = pivot_table.sum(axis=1)
pivot_table = pivot_table.sort_values('Total', ascending=False)
pivot_table = pivot_table.reset_index()

# Reorder columns
date_columns = [col for col in pivot_table.columns if isinstance(col, date)]
date_columns_sorted = sorted(date_columns)
column_order = ['State_Code', 'State_Name'] + date_columns_sorted + ['Total']
pivot_table = pivot_table[column_order]

# Output filename
output_file = f'orders_weekly_{DATE_RANGE}.csv'
pivot_table.to_csv(output_file, index=False, encoding='utf-8-sig')

print(f"\n{'='*80}")
print(f"✓ SUCCESS! Saved to: {output_file}")
print(f"{'='*80}")
print(f"\nSummary:")
print(f"  - States/Territories: {len(pivot_table)}")
print(f"  - Date columns: {len(date_columns_sorted)}")
print(f"  - Total TEMU orders: {temu_counts['Order_Count'].sum():,}")
print(f"  - Total YY orders (after dedup): {yy_counts['Order_Count'].sum():,}")
print(f"  - Combined total: {pivot_table['Total'].sum():,}")

# ============================================================
# Generate Daily Statistics File
# ============================================================
print(f"\n{'='*80}")
print(f"Generating Daily Statistics...")
print(f"{'='*80}")

# Generate daily breakdown
daily_stats_list = []
for date_col in date_columns_sorted:
    # Get TEMU orders for this date
    temu_date = df_temu_combined[df_temu_combined['Date_Only'] == date_col]
    temu_date_count = len(temu_date)

    # Get YY orders for this date (before dedup) - from yy_original_counts
    yy_date_original_count = yy_original_counts.get(date_col, 0)

    # Get YY orders for this date (after dedup) - from yy_counts
    yy_date_dedup = yy_counts[yy_counts['Date_Only'] == date_col]
    yy_date_dedup_count = yy_date_dedup['Order_Count'].sum() if len(yy_date_dedup) > 0 else 0

    # Combined total
    combined_total = temu_date_count + yy_date_dedup_count

    daily_stats_list.append({
        'Date': date_col,
        'TEMU_Orders': temu_date_count,
        'YY_Original': yy_date_original_count,
        'YY_Dedup': yy_date_dedup_count,
        'Combined_Total': combined_total
    })

daily_stats_df = pd.DataFrame(daily_stats_list)
daily_stats_file = f'daily_statistics_{DATE_RANGE}.csv'
daily_stats_df.to_csv(daily_stats_file, index=False, encoding='utf-8-sig')

print(f"\n✓ Daily statistics saved to: {daily_stats_file}")
print(f"\nDaily Statistics:")
print(daily_stats_df.to_string(index=False))
