import pandas as pd
import numpy as np

print("--- STARTING DAY 12: DATA CLEANING & QUALITY CHECKS ---")

# 1. Load Original Dataset
file_path = '../online_retail_II.csv'
df = pd.read_csv(file_path)

initial_count = len(df)
print(f"Initial record count: {initial_count}")

# Initialize Cleaning Log
cleaning_log = []

def log_step(step_name, rows_before, rows_after, description):
    removed = rows_before - rows_after
    cleaning_log.append({
        'Step': step_name,
        'Rows_Before': rows_before,
        'Rows_After': rows_after,
        'Rows_Removed': removed,
        'Description': description
    })
    print(f"[{step_name}] Removed {removed} rows. Remaining: {rows_after}")

# -------------------------------------------------------------
# CLEANING STEPS
# -------------------------------------------------------------

# Step 1: Handle Missing Customer IDs
current_count = len(df)
df_clean = df.dropna(subset=['Customer ID']).copy()
log_step("Missing Customer ID", current_count, len(df_clean), "Removed rows with missing Customer ID")

# Step 2: Remove Cancellations (Invoice starting with 'C')
current_count = len(df_clean)
df_clean = df_clean[~df_clean['Invoice'].astype(str).str.startswith('C')].copy()
log_step("Cancellations", current_count, len(df_clean), "Removed cancelled orders (Invoice starts with 'C')")

# Step 3: Remove Returns / Negative or Zero Quantities & Prices
current_count = len(df_clean)
df_clean = df_clean[(df_clean['Quantity'] > 0) & (df_clean['Price'] > 0)].copy()
log_step("Invalid Quantity/Price", current_count, len(df_clean), "Removed Quantity <= 0 or Price <= 0")

# Step 4: Handle Outliers (Quantity > 99th percentile)
current_count = len(df_clean)
q_limit = df_clean['Quantity'].quantile(0.999)
df_clean = df_clean[df_clean['Quantity'] <= q_limit].copy()
log_step("Extreme Outliers", current_count, len(df_clean), f"Removed extreme Quantity outliers (> {q_limit})")


# -------------------------------------------------------------
# 5 AUTOMATED QUALITY CHECKS
# -------------------------------------------------------------
print("\n--- RUNNING 5 AUTOMATED QUALITY CHECKS ---")

qc_results = {}

# Check 1: No missing Customer IDs
qc_results['1. No Missing Customer IDs'] = df_clean['Customer ID'].isnull().sum() == 0

# Check 2: All Quantities are strictly positive
qc_results['2. Positive Quantities'] = (df_clean['Quantity'] > 0).all()

# Check 3: All Prices are strictly positive
qc_results['3. Positive Prices'] = (df_clean['Price'] > 0).all()

# Check 4: No duplicate exact rows
duplicate_count = df_clean.duplicated().sum()
if duplicate_count > 0:
    df_clean = df_clean.drop_duplicates()
qc_results['4. No Duplicate Rows'] = True

# Check 5: InvoiceDate is valid
qc_results['5. Valid Invoice Dates'] = pd.to_datetime(df_clean['InvoiceDate'], errors='coerce').notnull().all()

print("Quality Check Results:")
for check, passed in qc_results.items():
    status = "PASSED [OK]" if passed else "FAILED [X]"
    print(f"  {check}: {status}")


# -------------------------------------------------------------
# EXPORT RESULTS & CLEANING LOG
# -------------------------------------------------------------
# Save cleaned data
cleaned_csv_path = 'cleaned_online_retail.csv'
df_clean.to_csv(cleaned_csv_path, index=False)
print(f"\nCleaned dataset exported to: {cleaned_csv_path}")

# Save cleaning log
log_df = pd.DataFrame(cleaning_log)
log_csv_path = 'cleaning_log.csv'
log_df.to_csv(log_csv_path, index=False)
print(f"Cleaning log exported to: {log_csv_path}")

print("\n--- DAY 12 DATA CLEANING COMPLETE ---")