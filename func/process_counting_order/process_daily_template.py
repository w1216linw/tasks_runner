"""
Daily Order Processing Template
用途: 处理单个日期的TEMU和YY订单数据
使用方法: 修改下方的文件名和日期，然后运行脚本
"""

import pandas as pd
import sys
import pgeocode
from datetime import timedelta

# ============================================================
# 配置区域 - 请修改这里的文件名和日期
# ============================================================
DATE_STRING = '260223'         # 输出文件日期字符串 (YYMMDD)
TEMU_FILE = f'TEMU/TEMU_{DATE_STRING}.xlsx'  # TEMU文件名
YY_FILE = f'YY/YY_{DATE_STRING}.csv'        # YY文件名

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
print(f"Processing Daily Data for {DATE_STRING}")
print("="*80)

# Initialize pgeocode
print("\nInitializing pgeocode for US...")
nomi = pgeocode.Nominatim('us')

# ============================================================
# Process TEMU file
# ============================================================
print(f"\n--- Processing TEMU: {TEMU_FILE} ---")

df_temu = pd.read_excel(TEMU_FILE)
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
tracking_numbers = set(df_temu['运单号'].dropna())
print(f"  TEMU tracking numbers: {len(tracking_numbers):,}")

# Get date range
temu_min = df_temu['Date_Only'].min()
temu_max = df_temu['Date_Only'].max()
print(f"  TEMU date range: {temu_min} to {temu_max}")

# Count orders by state
temu_counts = df_temu.groupby(['State_Code', 'State_Name']).size().reset_index(name='TEMU_Orders')

# ============================================================
# Process YY file
# ============================================================
print(f"\n--- Processing YY: {YY_FILE} ---")

df_yy = pd.read_csv(YY_FILE, encoding='utf-16', sep='\t')
print(f"  Total YY orders: {len(df_yy):,}")

# Check column names
print(f"  YY columns: {list(df_yy.columns)}")

# Convert date column BEFORE deduplication (needed for statistics)
df_yy['Date'] = pd.to_datetime(df_yy['预约揽收时间'])
df_yy['Date_Only'] = df_yy['Date'].dt.date

# Save YY before deduplication for statistics
df_yy_before_dedup = df_yy.copy()

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

# Count orders by state
yy_counts = df_yy.groupby(['State_Code', 'State_Name']).size().reset_index(name='YY_Orders')

# ============================================================
# Combine data
# ============================================================
# Merge TEMU and YY counts
combined = pd.merge(temu_counts, yy_counts, on=['State_Code', 'State_Name'], how='outer').fillna(0)
combined['TEMU_Orders'] = combined['TEMU_Orders'].astype(int)
combined['YY_Orders'] = combined['YY_Orders'].astype(int)
combined['Total'] = combined['TEMU_Orders'] + combined['YY_Orders']

# Sort by total descending
combined = combined.sort_values('Total', ascending=False)

# Output filename
output_file = f'orders_daily_{DATE_STRING}.csv'
combined.to_csv(output_file, index=False, encoding='utf-8-sig')

print(f"\n{'='*80}")
print(f"✓ SUCCESS! Saved to: {output_file}")
print(f"{'='*80}")
print(f"\nSummary:")
print(f"  - States/Territories: {len(combined)}")
print(f"  - Total TEMU orders: {combined['TEMU_Orders'].sum():,}")
print(f"  - Total YY orders (after dedup): {combined['YY_Orders'].sum():,}")
print(f"  - Combined total: {combined['Total'].sum():,}")

# Show top 10 states
print(f"\nTop 10 States by Order Volume:")
print(combined[['State_Code', 'State_Name', 'TEMU_Orders', 'YY_Orders', 'Total']].head(10).to_string(index=False))

# ============================================================
# Generate Daily Statistics File
# ============================================================
print(f"\n{'='*80}")
print(f"Generating Daily Statistics...")
print(f"{'='*80}")

# Count TEMU orders by date
temu_by_date = df_temu.groupby('Date_Only').size().reset_index(name='TEMU_Orders')

# Count YY orders BEFORE deduplication by date
# Filter weekdays for YY before dedup (same logic as after dedup)
df_yy_before_dedup['Weekday'] = df_yy_before_dedup['Date'].dt.dayofweek
df_yy_before_dedup_weekday = df_yy_before_dedup[df_yy_before_dedup['Weekday'] < 5]
yy_before_dedup_by_date = df_yy_before_dedup_weekday.groupby('Date_Only').size().reset_index(name='YY_Orders_Before_Dedup')

# Count YY orders AFTER deduplication by date
yy_after_dedup_by_date = df_yy.groupby('Date_Only').size().reset_index(name='YY_Orders_After_Dedup')

# Merge all statistics
stats_df = temu_by_date.merge(yy_before_dedup_by_date, on='Date_Only', how='outer')
stats_df = stats_df.merge(yy_after_dedup_by_date, on='Date_Only', how='outer')

# Fill NaN with 0 and convert to int
stats_df = stats_df.fillna(0)
stats_df['TEMU_Orders'] = stats_df['TEMU_Orders'].astype(int)
stats_df['YY_Orders_Before_Dedup'] = stats_df['YY_Orders_Before_Dedup'].astype(int)
stats_df['YY_Orders_After_Dedup'] = stats_df['YY_Orders_After_Dedup'].astype(int)

# Calculate combined total
stats_df['Combined_Total'] = stats_df['TEMU_Orders'] + stats_df['YY_Orders_After_Dedup']

# Sort by date
stats_df = stats_df.sort_values('Date_Only')

# Rename column for clarity
stats_df = stats_df.rename(columns={'Date_Only': 'Date'})

daily_stats_file = f'daily_statistics_{DATE_STRING}.csv'
stats_df.to_csv(daily_stats_file, index=False, encoding='utf-8-sig')

print(f"\n✓ Daily statistics saved to: {daily_stats_file}")
print(f"\nDaily Statistics:")
print(stats_df.to_string(index=False))
