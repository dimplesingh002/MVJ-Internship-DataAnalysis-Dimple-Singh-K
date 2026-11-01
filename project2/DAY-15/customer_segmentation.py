import pandas as pd
import numpy as np
import os

# Set paths
base_dir = r"C:\Users\Dimple Singh K\OneDrive\Desktop\MVJ-Internship-DataAnalysis-Dimple-Singh-K"
cleaned_data_path = os.path.join(base_dir, "project1", "Day12_datacleaning", "cleaned_online_retail.csv")
project2_dir = os.path.join(base_dir, "project2")
os.makedirs(project2_dir, exist_ok=True)

print("Loading cleaned dataset for Project 2...")
df = pd.read_csv(cleaned_data_path)
df['InvoiceDate'] = pd.to_datetime(df['InvoiceDate'])
if 'TotalPrice' not in df.columns:
    df['TotalPrice'] = df['Quantity'] * df['Price']

# 1. RFM SEGMENTATION ANALYSIS
print("Calculations: RFM Segmentation...")
snapshot_date = df['InvoiceDate'].max() + pd.Timedelta(days=1)

# Aggregate customer-level metrics
rfm = df.dropna(subset=['Customer ID']).groupby('Customer ID').agg({
    'InvoiceDate': lambda x: (snapshot_date - x.max()).days, # Recency
    'Invoice': 'nunique',                                   # Frequency
    'TotalPrice': 'sum'                                     # Monetary
}).reset_index()

rfm.columns = ['CustomerID', 'Recency', 'Frequency', 'Monetary']

# Assign RFM Quartile Scores (1 to 4, where 4 is best)
rfm['R_Score'] = pd.qcut(rfm['Recency'], 4, labels=[4, 3, 2, 1]).astype(int)
# Handle duplicate bins in frequency/monetary if any
rfm['F_Score'] = pd.qcut(rfm['Frequency'].rank(method='first'), 4, labels=[1, 2, 3, 4]).astype(int)
rfm['M_Score'] = pd.qcut(rfm['Monetary'].rank(method='first'), 4, labels=[1, 2, 3, 4]).astype(int)

rfm['RFM_Score'] = rfm['R_Score'].astype(str) + rfm['F_Score'].astype(str) + rfm['M_Score'].astype(str)

# Segment mapping function
def segment_customers(row):
    if row['R_Score'] >= 4 and row['F_Score'] >= 4:
        return 'Champions'
    elif row['R_Score'] >= 3 and row['F_Score'] >= 3:
        return 'Loyal Customers'
    elif row['R_Score'] >= 3 and row['F_Score'] <= 2:
        return 'Recent / New Buyers'
    elif row['R_Score'] <= 2 and row['F_Score'] >= 3:
        return 'At Risk'
    else:
        return 'Hibernating / Lost'

rfm['Segment'] = rfm.apply(segment_customers, axis=1)

# Save RFM results
rfm_output_path = os.path.join(project2_dir, "customer_rfm_segments.csv")
rfm.to_csv(rfm_output_path, index=False)
print(f"RFM Segmentation saved to: {rfm_output_path}")

# 2. COHORT RETENTION ANALYSIS
print("Calculations: Cohort Retention Matrix...")
df['InvoiceMonth'] = df['InvoiceDate'].dt.to_period('M')
df['CohortMonth'] = df.groupby('Customer ID')['InvoiceDate'].transform('min').dt.to_period('M')

def get_month_index(row):
    return (row['InvoiceMonth'].year - row['CohortMonth'].year) * 12 + (row['InvoiceMonth'].month - row['CohortMonth'].month)

df['CohortIndex'] = df.apply(get_month_index, axis=1)

cohort_data = df.dropna(subset=['Customer ID']).groupby(['CohortMonth', 'CohortIndex'])['Customer ID'].nunique().reset_index()
cohort_pivot = cohort_data.pivot(index='CohortMonth', columns='CohortIndex', values='Customer ID')

# Calculate retention percentage
cohort_size = cohort_pivot.iloc[:, 0]
retention_matrix = cohort_pivot.divide(cohort_size, axis=0)

# Save cohort matrix
cohort_output_path = os.path.join(project2_dir, "cohort_retention_matrix.csv")
retention_matrix.to_csv(cohort_output_path)
print(f"Cohort Retention Matrix saved to: {cohort_output_path}")
print("Project 2 data preprocessing and segmentation complete!")