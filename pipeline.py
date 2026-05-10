import pandas as pd
import os
import glob

# Load one file first to see structure
file_path = r"C:\Users\zaidr\OneDrive\Desktop\Projects\fca-complaints\data\firm-level-complaints-data-2025-h1.xlsx"

# FCA files often have multiple sheets
xl = pd.ExcelFile(file_path)
print(f"Sheet names: {xl.sheet_names}")

# Load first sheet
df = pd.read_excel(file_path, sheet_name=xl.sheet_names[0])
print(f"\nShape: {df.shape}")
print(f"\nColumns: {df.columns.tolist()}")
print(f"\nFirst few rows:")
print(df.head(10))
print(f"\nData types:")
print(df.dtypes)

# Check upheld sheet
df_upheld = pd.read_excel(file_path, sheet_name='Percentage upheld')
print(f"Shape: {df_upheld.shape}")
print(f"Columns: {df_upheld.columns.tolist()}")
print(df_upheld.head(10))

# Check all files
all_files = glob.glob(r"C:\Users\zaidr\OneDrive\Desktop\Projects\fca-complaints\data\*.xlsx")

for file in sorted(all_files):
    xl = pd.ExcelFile(file)
    filename = os.path.basename(file)
    print(f"{filename}: {xl.sheet_names}")
    
  
data_path = r"C:\Users\zaidr\OneDrive\Desktop\Projects\fca-complaints\data"

# Sheet name mapping for each file
sheet_mapping = {
    'firm-level-complaints-data-2022-h1.xlsx': {
        'opened': 'opened',
        'upheld': '% upheld '
    },
    'firm-level-complaints-data-2022-h2.xlsx': {
        'opened': 'opened',
        'upheld': 'upheld '
    },
    'firm-level-complaints-data-2023-h1.xlsx': {
        'opened': 'Opened',
        'upheld': 'Upheld'
    },
    'firm-level-complaints-data-2023-h2.xlsx': {
        'opened': 'Opened',
        'upheld': 'Upheld'
    },
    'firm-level-complaints-data-2024-h1.xlsx': {
        'opened': 'Opened',
        'upheld': 'Upheld'
    },
    'firm-level-complaints-data-2024-h2.xlsx': {
        'opened': 'Opened',
        'upheld': 'Upheld'
    },
    'firm-level-complaints-data-2025-h1.xlsx': {
        'opened': 'Opened',
        'upheld': 'Percentage upheld'
    }
}

# Product categories
product_cols = [
    'Banking and credit cards',
    'Decumulation & pensions',
    'Home finance',
    'Insurance & pure protection',
    'Investments'
]

all_opened = []
all_upheld = []

for filename, sheets in sheet_mapping.items():
    filepath = os.path.join(data_path, filename)
    
    try:
        # Load opened sheet
        df_open = pd.read_excel(filepath, sheet_name=sheets['opened'])
        df_open['period'] = filename.replace('firm-level-complaints-data-', '').replace('.xlsx', '')
        df_open['type'] = 'opened'
        all_opened.append(df_open)
        
        # Load upheld sheet
        df_uph = pd.read_excel(filepath, sheet_name=sheets['upheld'])
        df_uph['period'] = filename.replace('firm-level-complaints-data-', '').replace('.xlsx', '')
        df_uph['type'] = 'upheld'
        all_upheld.append(df_uph)
        
        print(f"✅ Loaded {filename}")
        
    except Exception as e:
        print(f"❌ Error loading {filename}: {e}")

# Combine
master_opened = pd.concat(all_opened, ignore_index=True)
master_upheld = pd.concat(all_upheld, ignore_index=True)

print(f"\nOpened complaints shape: {master_opened.shape}")
print(f"Upheld data shape: {master_upheld.shape}")
print(f"\nPeriods in opened: {master_opened['period'].unique()}")
print(f"\nColumns in opened: {master_opened.columns.tolist()}")

# Standardise column names for opened complaints
def standardise_columns(df):
    # Merge duplicate columns
    if 'Firm Group' in df.columns and 'Group' in df.columns:
        df['Group'] = df['Group'].fillna(df['Firm Group'])
    if 'Joint Reporting' in df.columns and 'Joint Report' in df.columns:
        df['Joint Reporting'] = df['Joint Reporting'].fillna(df['Joint Report'])
    
    # Keep only standard columns
    keep_cols = [
        'Firm Name', 'Group', 'Joint Reporting', 'Reporting period',
        'Banking and credit cards', 'Decumulation & pensions',
        'Home finance', 'Insurance & pure protection', 'Investments',
        'Grand Total', 'period'
    ]
    
    # Only keep columns that exist
    keep_cols = [c for c in keep_cols if c in df.columns]
    return df[keep_cols]

# Apply to both datasets
master_opened = standardise_columns(master_opened)
master_upheld = standardise_columns(master_upheld)

# Rename Grand Total
master_opened = master_opened.rename(columns={'Grand Total': 'total_complaints'})
master_upheld = master_upheld.rename(columns={'Grand Total': 'total_upheld_rate'})

# Clean firm names
master_opened['Firm Name'] = master_opened['Firm Name'].str.strip()
master_upheld['Firm Name'] = master_upheld['Firm Name'].str.strip()

# Remove rows with no firm name
master_opened = master_opened.dropna(subset=['Firm Name'])
master_upheld = master_upheld.dropna(subset=['Firm Name'])

# Add year and half
master_opened['year'] = master_opened['period'].str[:4].astype(int)
master_opened['half'] = master_opened['period'].str[-2:]
master_upheld['year'] = master_upheld['period'].str[:4].astype(int)
master_upheld['half'] = master_upheld['period'].str[-2:]

# Convert upheld rates to percentage
product_cols = [
    'Banking and credit cards', 'Decumulation & pensions',
    'Home finance', 'Insurance & pure protection', 'Investments'
]
for col in product_cols:
    if col in master_upheld.columns:
        master_upheld[col] = pd.to_numeric(master_upheld[col], errors='coerce')

# Save cleaned files
master_opened.to_csv(r"C:\Users\zaidr\OneDrive\Desktop\Projects\fca-complaints\data\opened_clean.csv", index=False)
master_upheld.to_csv(r"C:\Users\zaidr\OneDrive\Desktop\Projects\fca-complaints\data\upheld_clean.csv", index=False)

# Sense check
print(f"Opened clean shape: {master_opened.shape}")
print(f"Upheld clean shape: {master_upheld.shape}")
print(f"\nSample opened data:")
print(master_opened.head())
print(f"\nTotal complaints range: {master_opened['total_complaints'].min()} to {master_opened['total_complaints'].max()}")
print(f"\nUnique firms: {master_opened['Firm Name'].nunique()}")

# Fix total_complaints — calculate from product columns where Grand Total missing
product_cols = [
    'Banking and credit cards', 'Decumulation & pensions',
    'Home finance', 'Insurance & pure protection', 'Investments'
]

for col in product_cols:
    master_opened[col] = pd.to_numeric(master_opened[col], errors='coerce')

master_opened['total_complaints'] = master_opened['total_complaints'].fillna(
    master_opened[product_cols].sum(axis=1)
)

# Drop duplicate columns
master_opened = master_opened.drop(columns=['Firm Group', 'Joint Reporting', 'Joint Report'], errors='ignore')
master_upheld = master_upheld.drop(columns=['Firm Group', 'Joint Reporting', 'Joint Report'], errors='ignore')

# Remove rows with zero total complaints
master_opened = master_opened[master_opened['total_complaints'] > 0]

# Save again
master_opened.to_csv(r"C:\Users\zaidr\OneDrive\Desktop\Projects\fca-complaints\data\opened_clean.csv", index=False)
master_upheld.to_csv(r"C:\Users\zaidr\OneDrive\Desktop\Projects\fca-complaints\data\upheld_clean.csv", index=False)

print(f"Final opened shape: {master_opened.shape}")
print(f"Final upheld shape: {master_upheld.shape}")
print(f"\nMissing total_complaints: {master_opened['total_complaints'].isnull().sum()}")
print(f"Unique firms: {master_opened['Firm Name'].nunique()}")
print(f"\nTop 10 firms by total complaints:")
print(master_opened.groupby('Firm Name')['total_complaints'].sum().sort_values(ascending=False).head(10))

# ---- Sector Classification ----

# Rule-based sector assignment using firm names and dominant product
def classify_sector(row):
    name = str(row['Firm Name']).lower()
    
    # Energy firms
    energy_keywords = ['british gas', 'eon', 'edf', 'npower', 'scottish power', 
                       'sse', 'octopus', 'bulb', 'utilita', 'ovo energy']
    if any(k in name for k in energy_keywords):
        return 'Energy & Utilities'
    
    # Banking keywords
    bank_keywords = ['bank', 'building society', 'credit union', 'monzo', 
                     'starling', 'revolut', 'metro', 'virgin money']
    if any(k in name for k in bank_keywords):
        return 'Banking'
    
    # Mortgage/Lending keywords
    mortgage_keywords = ['mortgage', 'finance', 'lending', 'loan', 'credit']
    if any(k in name for k in mortgage_keywords):
        return 'Mortgage & Lending'
    
    # Insurance keywords
    insurance_keywords = ['insurance', 'insure', 'assurance', 'life', 
                          'protection', 'bupa', 'axa', 'aviva', 'zurich',
                          'allianz', 'admiral', 'direct line', 'hastings']
    if any(k in name for k in insurance_keywords):
        return 'Insurance'
    
    # Use dominant product column as fallback
    products = {
        'Banking and credit cards': 'Banking',
        'Insurance & pure protection': 'Insurance',
        'Home finance': 'Mortgage & Lending',
        'Decumulation & pensions': 'Insurance',
        'Investments': 'Investments'
    }
    
    # Find which product has highest value
    product_cols = list(products.keys())
    for col in product_cols:
        if col in row.index:
            row[col] = pd.to_numeric(row[col], errors='coerce')
    
    max_col = None
    max_val = 0
    for col in product_cols:
        if col in row.index and pd.notna(row[col]) and row[col] > max_val:
            max_val = row[col]
            max_col = col
    
    if max_col:
        return products[max_col]
    
    return 'Other'

# Apply sector classification
master_opened['sector'] = master_opened.apply(classify_sector, axis=1)
master_upheld['sector'] = master_upheld.apply(classify_sector, axis=1)

print("Sector distribution:")
print(master_opened.groupby('sector')['total_complaints'].sum().sort_values(ascending=False))
print(f"\nFirms per sector:")
print(master_opened.groupby('sector')['Firm Name'].nunique().sort_values(ascending=False))

master_opened.to_csv(r"C:\Users\zaidr\OneDrive\Desktop\Projects\fca-complaints\data\opened_clean.csv", index=False)
master_upheld.to_csv(r"C:\Users\zaidr\OneDrive\Desktop\Projects\fca-complaints\data\upheld_clean.csv", index=False)

import pandas as pd

upheld = pd.read_csv(r"C:\Users\zaidr\OneDrive\Desktop\Projects\fca-complaints\data\upheld_clean.csv")

product_cols = [
    'Banking and credit cards', 'Decumulation & pensions',
    'Home finance', 'Insurance & pure protection', 'Investments'
]

for col in product_cols:
    upheld[col] = pd.to_numeric(upheld[col], errors='coerce')

# Calculate average upheld rate across product columns
upheld['avg_upheld_rate'] = upheld[product_cols].mean(axis=1)

# Save back
upheld.to_csv(r"C:\Users\zaidr\OneDrive\Desktop\Projects\fca-complaints\data\upheld_clean.csv", index=False)

print("Done!")
print(upheld[['Firm Name', 'period', 'avg_upheld_rate']].head(10))

import pandas as pd

upheld = pd.read_csv(r"C:\Users\zaidr\OneDrive\Desktop\Projects\fca-complaints\data\upheld_clean.csv")

product_cols = [
    'Banking and credit cards', 'Decumulation & pensions',
    'Home finance', 'Insurance & pure protection', 'Investments'
]

for col in product_cols:
    upheld[col] = pd.to_numeric(upheld[col], errors='coerce')

# Fix avg_upheld_rate — multiply by 100 if values are decimal
if upheld['avg_upheld_rate'].mean() < 1:
    upheld['avg_upheld_rate'] = upheld['avg_upheld_rate'] * 100
    print("Fixed — multiplied by 100")
else:
    print("Values already in percentage format")

print(f"Sample values: {upheld['avg_upheld_rate'].head()}")

upheld.to_csv(r"C:\Users\zaidr\OneDrive\Desktop\Projects\fca-complaints\data\upheld_clean.csv", index=False)
print("Saved!")

# Load both files
opened = pd.read_csv(r"C:\Users\zaidr\OneDrive\Desktop\Projects\fca-complaints\data\opened_clean.csv")
upheld = pd.read_csv(r"C:\Users\zaidr\OneDrive\Desktop\Projects\fca-complaints\data\upheld_clean.csv")

# Add sort order column
period_order = {
    '2022-h1': 1, '2022-h2': 2,
    '2023-h1': 3, '2023-h2': 4,
    '2024-h1': 5, '2024-h2': 6,
    '2025-h1': 7
}

opened['period_sort'] = opened['period'].map(period_order)
upheld['period_sort'] = upheld['period'].map(period_order)

# Save
opened.to_csv(r"C:\Users\zaidr\OneDrive\Desktop\Projects\fca-complaints\data\opened_clean.csv", index=False)
upheld.to_csv(r"C:\Users\zaidr\OneDrive\Desktop\Projects\fca-complaints\data\upheld_clean.csv", index=False)

print("Done!")
print(opened[['period', 'period_sort']].drop_duplicates().sort_values('period_sort'))

import pandas as pd

upheld = pd.read_csv(r"C:\Users\zaidr\OneDrive\Desktop\Projects\fca-complaints\data\upheld_clean.csv")

print("Columns:", upheld.columns.tolist())
print("\nSample:")
print(upheld[['period', 'period_sort']].drop_duplicates().sort_values('period_sort'))

# Create proper date column
period_to_date = {
    '2022-h1': '2022-01-01',
    '2022-h2': '2022-07-01',
    '2023-h1': '2023-01-01',
    '2023-h2': '2023-07-01',
    '2024-h1': '2024-01-01',
    '2024-h2': '2024-07-01',
    '2025-h1': '2025-01-01'
}

upheld['period_date'] = upheld['period'].map(period_to_date)
upheld.to_csv(r"C:\Users\zaidr\OneDrive\Desktop\Projects\fca-complaints\data\upheld_clean.csv", index=False)
print("Done!")
print(upheld[['period', 'period_date']].drop_duplicates())


# Calculate industry average per period
period_avg = upheld.groupby('period')['avg_upheld_rate'].mean().reset_index()
period_avg.columns = ['period', 'industry_avg']

upheld = upheld.merge(period_avg, on='period', how='left')

upheld['upheld_deviation'] = upheld['avg_upheld_rate'] - upheld['industry_avg']

upheld.to_csv(r"C:\Users\zaidr\OneDrive\Desktop\Projects\fca-complaints\data\upheld_clean.csv", index=False)

print("Done!")
print(upheld[['sector', 'period', 'avg_upheld_rate', 'industry_avg', 'upheld_deviation']].head(10))

first = upheld[upheld['period_sort'] == 1][['Firm Name', 'avg_upheld_rate']].rename(columns={'avg_upheld_rate': 'first_rate'})
last = upheld[upheld['period_sort'] == 7][['Firm Name', 'avg_upheld_rate']].rename(columns={'avg_upheld_rate': 'last_rate'})

merged = first.merge(last, on='Firm Name')
merged['firm_yoy_change'] = merged['last_rate'] - merged['first_rate']

print(f"Average change: {merged['firm_yoy_change'].mean():.2f}%")

upheld = upheld.merge(merged[['Firm Name', 'firm_yoy_change']], on='Firm Name', how='left')
upheld.to_csv(r"C:\Users\zaidr\OneDrive\Desktop\Projects\fca-complaints\data\upheld_clean.csv", index=False)
print("Saved!")


# Get one row per firm with their yoy change
firm_changes = upheld.drop_duplicates(subset=['Firm Name'])[['Firm Name', 'firm_yoy_change', 'sector']].dropna()

# Create buckets
def bucket(x):
    if x < -10:
        return '1. < -10%'
    elif x < -5:
        return '2. -10 to -5%'
    elif x < 0:
        return '3. -5 to 0%'
    elif x < 5:
        return '4. 0 to 5%'
    elif x < 10:
        return '5. 5 to 10%'
    else:
        return '6. > 10%'

firm_changes['change_bucket'] = firm_changes['firm_yoy_change'].apply(bucket)

# Save
firm_changes.to_csv(r"C:\Users\zaidr\OneDrive\Desktop\Projects\fca-complaints\data\firm_changes.csv", index=False)

print("Done!")
print(firm_changes['change_bucket'].value_counts().sort_index())
print(f"\nTotal firms: {len(firm_changes)}")

# Count firms per bucket
bucket_counts = firm_changes.groupby('change_bucket')['Firm Name'].count().reset_index()
bucket_counts.columns = ['change_bucket', 'firm_count']

print(bucket_counts)

bucket_counts.to_csv(r"C:\Users\zaidr\OneDrive\Desktop\Projects\fca-complaints\data\bucket_counts.csv", index=False)
print("Saved!")


import pandas as pd

upheld = pd.read_csv(r"C:\Users\zaidr\OneDrive\Desktop\Projects\fca-complaints\data\upheld_clean.csv")

# Count periods per firm
period_counts = upheld.groupby('Firm Name')['period'].nunique().reset_index()
period_counts.columns = ['Firm Name', 'period_count']

# Only firms with 4+ periods
stable_firms = period_counts[period_counts['period_count'] >= 4]['Firm Name']

# Recalculate volatility for stable firms only
upheld_stable = upheld[upheld['Firm Name'].isin(stable_firms)]

volatility = upheld_stable.groupby('Firm Name')['avg_upheld_rate'].apply(
    lambda x: x.max() - x.min()
).reset_index()
volatility.columns = ['Firm Name', 'volatility_score']

# Merge back
upheld = upheld.drop(columns=['volatility_score'], errors='ignore')
upheld = upheld.merge(volatility, on='Firm Name', how='left')

upheld.to_csv(r"C:\Users\zaidr\OneDrive\Desktop\Projects\fca-complaints\data\upheld_clean.csv", index=False)

print("Done!")
print(f"\nTop 10 most volatile firms:")
print(volatility.sort_values('volatility_score', ascending=False).head(10))

# Set volatility to null for firms with < 4 periods
upheld.loc[~upheld['Firm Name'].isin(stable_firms), 'volatility_score'] = None

upheld.to_csv(r"C:\Users\zaidr\OneDrive\Desktop\Projects\fca-complaints\data\upheld_clean.csv", index=False)
print("Done!")

import pandas as pd

bucket_counts = pd.read_csv(r"C:\Users\zaidr\OneDrive\Desktop\Projects\fca-complaints\data\bucket_counts.csv")

# Clean bucket names without numbers
bucket_rename = {
    '1. < -10%': '< -10%',
    '2. -10 to -5%': '-10 to -5%',
    '3. -5 to 0%': '-5 to 0%',
    '4. 0 to 5%': '0 to 5%',
    '5. 5 to 10%': '5 to 10%',
    '6. > 10%': '> 10%'
}

bucket_counts['change_bucket'] = bucket_counts['change_bucket'].map(bucket_rename)

# Keep sort order
bucket_counts['sort_order'] = [1, 2, 3, 4, 5, 6]

bucket_counts.to_csv(r"C:\Users\zaidr\OneDrive\Desktop\Projects\fca-complaints\data\bucket_counts.csv", index=False)
print(bucket_counts)

bucket_rename = {
    '1. < -10%': 'A. < -10%',
    '2. -10 to -5%': 'B. -10 to -5%',
    '3. -5 to 0%': 'C. -5 to 0%',
    '4. 0 to 5%': 'D. 0 to 5%',
    '5. 5 to 10%': 'E. 5 to 10%',
    '6. > 10%': 'F. > 10%'
}