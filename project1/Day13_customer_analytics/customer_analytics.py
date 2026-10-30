import pandas as pd
import numpy as np
import datetime as dt

print("--- STARTING DAY 13: CUSTOMER ANALYTICS (RFM & COHORTS) ---")

# 1. Load Cleaned Data from Day 12
df = pd.read_csv('../Day12_datacleaning/cleaned_online_retail.csv', parse_dates=['InvoiceDate'])

# Ensure Customer ID is integer/string without decimals
df['Customer ID'] = df['Customer ID'].astype(int)

# Calculate Total Price for each transaction
df['TotalPrice'] = df['Quantity'] * df['Price']

# -------------------------------------------------------------
# PART 1: RFM SEGMENTATION
# -------------------------------------------------------------
print("\nComputing RFM scores...")

# Set snapshot date as 1 day after the max invoice date in the dataset
snapshot_date = df['InvoiceDate'].max() + dt.timedelta(days=1)

# Group by Customer ID to calculate Recency, Frequency, Monetary
rfm = df.groupby('Customer ID').agg({
    'InvoiceDate': lambda x: (snapshot_date - x.max()).days,
    'Invoice': 'nunique',
    'TotalPrice': 'sum'
}).reset_index()

# Rename columns
rfm.columns = ['CustomerID', 'Recency', 'Frequency', 'Monetary']

# Calculate RFM scores using quintiles (1 to 4)
# Note: Lower recency is better, higher frequency/monetary are better
rfm['R_Score'] = pd.qcut(rfm['Recency'], 4, labels=[4, 3, 2, 1])
rfm['F_Score'] = pd.qcut(rfm['Frequency'].rank(method='first'), 4, labels=[1, 2, 3, 4])
rfm['M_Score'] = pd.qcut(rfm['Monetary'], 4, labels=[1, 2, 3, 4])

# Combine RFM Score
rfm['RFM_Score'] = rfm['R_Score'].astype(str) + rfm['F_Score'].astype(str) + rfm['M_Score'].astype(str)

# Export RFM results
rfm.to_csv('customer_rfm_segments.csv', index=False)
print("Saved: customer_rfm_segments.csv")


# -------------------------------------------------------------
# PART 2: COHORT RETENTION ANALYSIS
# -------------------------------------------------------------
print("\nBuilding cohort retention matrix...")

# Function to get month period
def get_month(x):
    return dt.datetime(x.year, x.month, 1)

df['InvoiceMonth'] = df['InvoiceDate'].apply(get_month)

# Find the first purchase month (Cohort Month) for each customer
df['CohortMonth'] = df.groupby('Customer ID')['InvoiceMonth'].transform('min')

# Calculate integer index for months between transaction and cohort
def get_month_int(df, column):
    year_diff = df['InvoiceDate'].dt.year - df['CohortMonth'].dt.year
    month_diff = df['InvoiceDate'].dt.month - df['CohortMonth'].dt.month
    return year_diff * 12 + month_diff

df['CohortIndex'] = get_month_int(df, 'InvoiceDate') + 1

# Group by CohortMonth and CohortIndex to count unique active customers
cohort_data = df.groupby(['CohortMonth', 'CohortIndex'])['Customer ID'].nunique().reset_index()
cohort_matrix = cohort_data.pivot(index='CohortMonth', columns='CohortIndex', values='Customer ID')

# Calculate retention matrix (percentages)
cohort_sizes = cohort_matrix.iloc[:, 0]
retention_matrix = cohort_matrix.divide(cohort_sizes, axis=0)

# Export cohort retention matrix
retention_matrix.to_csv('cohort_retention_matrix.csv')
print("Saved: cohort_retention_matrix.csv")

print("\n--- DAY 13 CUSTOMER ANALYTICS COMPLETE ---")