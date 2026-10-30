import pandas as pd

# 1. Load Dataset
print("Loading dataset for time series analysis...")
df = pd.read_csv('../online_retail_II.csv', encoding='ISO-8859-1')

# 2. Data Preparation for Time Series
# Convert InvoiceDate to datetime objects
df['InvoiceDate'] = pd.to_datetime(df['InvoiceDate'])

# Create a Total Sales column using the correct column names ('Price' and 'Customer ID')
df['TotalSales'] = df['Quantity'] * df['Price']

# Filter out cancellations and missing Customer IDs
df = df.dropna(subset=['Customer ID'])
df['Invoice'] = df['Invoice'].astype(str)
df = df[~df['Invoice'].str.startswith('C')]
df = df[(df['Quantity'] > 0) & (df['Price'] > 0)]

# Set InvoiceDate as index for time series resampling
df.set_index('InvoiceDate', inplace=True)

# 3. Resample for Daily and Monthly Revenue
print("\n--- Calculating Daily and Monthly Revenue ---")
daily_revenue = df['TotalSales'].resample('D').sum()
monthly_revenue = df['TotalSales'].resample('ME').sum()

print("Daily Revenue (Head):")
print(daily_revenue.head())

print("\nMonthly Revenue:")
print(monthly_revenue)

# 4. Calculate Rolling Averages (e.g., 7-day rolling average of daily revenue)
daily_rolling_avg = daily_revenue.rolling(window=7).mean()
print("\n--- 7-Day Rolling Average of Daily Revenue (Head) ---")
print(daily_rolling_avg.head(10))

# Export summaries
daily_revenue.to_csv('daily_revenue.csv')
monthly_revenue.to_csv('monthly_revenue.csv')
print("\nTime series summaries exported successfully!")