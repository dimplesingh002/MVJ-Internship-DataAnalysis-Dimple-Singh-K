import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.model_selection import train_test_split
from sklearn.linear_model import LinearRegression
from sklearn.metrics import mean_squared_error, r2_score
import os

# Set paths
base_dir = r"C:\Users\Dimple Singh K\OneDrive\Desktop\MVJ-Internship-DataAnalysis-Dimple-Singh-K"
day17_dir = os.path.join(base_dir, "Day17_regression")
os.makedirs(day17_dir, exist_ok=True)

print("Generating/Loading Bengaluru Housing Dataset for Regression...")

# Simulating a realistic Bengaluru housing dataset based on typical metrics (Area, Bedrooms, Bathrooms, Age -> Price in Lakhs INR)
np.random.seed(42)
n_samples = 500

area = np.random.normal(1200, 350, n_samples).clip(500, 3500) # Square feet
bhk = np.random.choice([1, 2, 3, 4], size=n_samples, p=[0.1, 0.4, 0.4, 0.1])
bathrooms = bhk + np.random.choice([0, 1], size=n_samples, p=[0.7, 0.3])
age = np.random.randint(0, 15, n_samples)

# Price formula with some noise (Price in Lakhs ₹)
price = (area * 0.055) + (bhk * 15.0) + (bathrooms * 10.0) - (age * 1.2) + np.random.normal(0, 8, n_samples)
price = price.clip(20, 350)

df_house = pd.DataFrame({
    'Area_SqFt': area,
    'BHK': bhk,
    'Bathrooms': bathrooms,
    'Property_Age': age,
    'Price_Lakhs': price
})

# Save dataset snapshot
df_house.to_csv(os.path.join(day17_dir, "bengaluru_housing_data.csv"), index=False)

# 1. Define Features and Target
X = df_house[['Area_SqFt', 'BHK', 'Bathrooms', 'Property_Age']]
y = df_house['Price_Lakhs']

# 2. Train/Test Split (80% train, 20% test)
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

# 3. Train Multiple Linear Regression Model
model = LinearRegression()
model.fit(X_train, y_train)

# 4. Predictions & Evaluation
y_pred = model.predict(X_test)
rmse = np.sqrt(mean_squared_error(y_test, y_pred))
r2 = r2_score(y_test, y_pred)

print(f"\n--- Model Evaluation ---")
print(f"R-squared (R2): {r2:.4f}")
print(f"Root Mean Squared Error (RMSE): ₹{rmse:.2f} Lakhs")

print("\n--- Model Coefficients ---")
for feature, coef in zip(X.columns, model.coef_):
    print(f"{feature}: {coef:.4f}")
print(f"Intercept: {model.intercept_:.4f}")

# 5. Visualizing Actual vs Predicted (Parity Plot)
plt.figure(figsize=(8, 6))
sns.scatterplot(x=y_test, y=y_pred, alpha=0.7, color="teal")
plt.plot([y.min(), y.max()], [y.min(), y.max()], '--r', linewidth=2)
plt.title("Bengaluru House Prices: Actual vs Predicted (Linear Regression)")
plt.xlabel("Actual Price (Lakhs ₹)")
plt.ylabel("Predicted Price (Lakhs ₹)")
plt.grid(True, linestyle="--", alpha=0.5)

plot_path = os.path.join(day17_dir, "actual_vs_predicted_regression.png")
plt.savefig(plot_path, dpi=300, bbox_inches='tight')
print(f"\nParity plot saved successfully at: {plot_path}")