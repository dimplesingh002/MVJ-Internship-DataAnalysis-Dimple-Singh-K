<<<<<<< HEAD
import sqlite3
import pandas as pd

=======
<<<<<<< HEAD

=======
import sqlite3
import pandas as pd

>>>>>>> 3ee80a8 (Day 10: Final commit consolidating pandas KPI script)
# 1. Connect to SQLite database and load data into a DataFrame
conn = sqlite3.connect('../online_retail.db')  # Adjust path if db is in root
query = "SELECT Invoice, StockCode, Description, Quantity, InvoiceDate, Price, `Customer ID` FROM online_retail"
df = pd.read_sql(query, conn)
conn.close()

print("Data loaded successfully! Shape:", df.shape)

# 2. Clean data (remove missing descriptions and cancellations)
df_clean = df.dropna(subset=['Description']).copy()
df_clean = df_clean[~df_clean['Invoice'].astype(str).str.startswith('C')]

# 3. Calculate Total Revenue
df_clean['Revenue'] = df_clean['Quantity'] * df_clean['Price']
total_revenue = df_clean['Revenue'].sum()
print(f"Total Revenue: ₹{total_revenue:,.2f}")

# 4. Calculate Average Order Value (AOV)
total_orders = df_clean['Invoice'].nunique()
aov = total_revenue / total_orders
print(f"Average Order Value (AOV): ₹{aov:,.2f}")

# 5. Top 5 Products by Revenue
top_products = df_clean.groupby('Description')['Revenue'].sum().reset_index()
top_products = top_products.sort_values(by='Revenue', ascending=False).head(5)
print("\nTop 5 Products by Revenue:")
<<<<<<< HEAD
print(top_products)
=======
print(top_products)
>>>>>>> 003d770 (Day 10: Added pandas script to replicate Project 1 KPIs)
>>>>>>> 3ee80a8 (Day 10: Final commit consolidating pandas KPI script)
