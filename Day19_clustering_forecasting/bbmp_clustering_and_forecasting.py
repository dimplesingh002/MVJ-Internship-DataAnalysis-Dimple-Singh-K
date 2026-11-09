import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.cluster import KMeans
from sklearn.preprocessing import StandardScaler
import os

# Set paths
base_dir = r"C:\Users\Dimple Singh K\OneDrive\Desktop\MVJ-Internship-DataAnalysis-Dimple-Singh-K"
day19_dir = os.path.join(base_dir, "Day19_clustering_forecasting")
os.makedirs(day19_dir, exist_ok=True)

print("Running Day 19: Clustering & Forecasting pipeline...")

# --- PART 1: K-MEANS CLUSTERING (BBMP Wards) ---
np.random.seed(42)
n_wards = 100

wards_data = pd.DataFrame({
    'Ward_Name': [f"Ward_{i}" for i in range(1, n_wards + 1)],
    'Water_Access_Score': np.random.uniform(40, 95, n_wards),
    'Road_Quality_Score': np.random.uniform(30, 90, n_wards),
    'Waste_Management_Score': np.random.uniform(50, 100, n_wards),
    'Green_Cover_Percentage': np.random.uniform(10, 45, n_wards)
})

# Scale features for K-Means
scaler = StandardScaler()
X_scaled = scaler.fit_transform(wards_data[['Water_Access_Score', 'Road_Quality_Score', 'Waste_Management_Score', 'Green_Cover_Percentage']])

# Apply K-Means (e.g., 3 clusters: High Infrastructure, Moderate, Low Infrastructure)
kmeans = KMeans(n_clusters=3, random_state=42, n_init=10)
wards_data['Cluster'] = kmeans.fit_predict(X_scaled)

wards_csv_path = os.path.join(day19_dir, "bbmp_wards_clustered.csv")
wards_data.to_csv(wards_csv_path, index=False)

# --- PART 2: TIME SERIES FORECASTING (Bengaluru Monsoon Rainfall) ---
# Simulating monthly rainfall (mm) over 5 years (June - September monsoon periods)
np.random.seed(100)
dates = pd.date_range(start="2020-01-01", end="2025-10-01", freq="ME")
n_months = len(dates)

# Base trend + seasonal monsoon spike in July/August
rainfall = 50 + 10 * np.sin(np.linspace(0, 10 * np.pi, n_months)) + np.where(dates.month.isin([6, 7, 8, 9]), 180, 20) + np.random.normal(0, 15, n_months)
rainfall = rainfall.clip(10, 350)

df_ts = pd.DataFrame({'Date': dates, 'Rainfall_mm': rainfall})
df_ts.set_index('Date', inplace=True)

# Moving Average (3-month rolling) & Exponential Smoothing approximation
df_ts['MA_3'] = df_ts['Rainfall_mm'].rolling(window=3).mean()
df_ts['Exp_Smoothing'] = df_ts['Rainfall_mm'].ewm(alpha=0.3).mean()

ts_csv_path = os.path.join(day19_dir, "bengaluru_monsoon_forecast.csv")
df_ts.to_csv(ts_csv_path)

# --- PART 3: VISUALIZATIONS ---
plt.figure(figsize=(12, 5))

# Plot 1: K-Means BBMP Wards Scatter
plt.subplot(1, 2, 1)
sns.scatterplot(data=wards_data, x='Water_Access_Score', y='Road_Quality_Score', hue='Cluster', palette='viridis', s=80)
plt.title("BBMP Wards Clustering (Public Goods)")
plt.xlabel("Water Access Score")
plt.ylabel("Road Quality Score")

# Plot 2: Monsoon Rainfall Forecasting
plt.subplot(1, 2, 2)
plt.plot(df_ts.index, df_ts['Rainfall_mm'], label='Actual Rainfall', color='gray', alpha=0.6)
plt.plot(df_ts.index, df_ts['MA_3'], label='3-Month Moving Average', color='blue', linestyle='--')
plt.plot(df_ts.index, df_ts['Exp_Smoothing'], label='Exponential Smoothing', color='red')
plt.title("Bengaluru Monsoon Rainfall Forecasting")
plt.xlabel("Year")
plt.ylabel("Rainfall (mm)")
plt.legend()

plt.tight_layout()
plot_path = os.path.join(day19_dir, "clustering_and_forecast_plots.png")
plt.savefig(plot_path, dpi=300, bbox_inches='tight')

print(f"Day 19 files successfully generated and saved in: {day19_dir}")