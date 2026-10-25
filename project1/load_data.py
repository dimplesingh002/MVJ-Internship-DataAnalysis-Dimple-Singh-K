import sqlite3
import pandas as pd

print("Loading dataset into memory...")

# Read the CSV dataset
df = pd.read_csv('online_retail_II.csv')

print(f"Dataset loaded with {len(df):,} rows.")

# Connect to a local SQLite database (this creates 'online_retail.db')
conn = sqlite3.connect('online_retail.db')

# Write the dataframe into an SQLite table named 'online_retail'
df.to_sql('online_retail', conn, index=False, if_exists='replace')

print("Data successfully loaded into local SQLite database 'online_retail.db'!")

# Verify by running a quick test query
query = "SELECT COUNT(*) AS total_rows FROM online_retail;"
result = pd.read_sql(query, conn)
print(result)

# Close connection
conn.close()