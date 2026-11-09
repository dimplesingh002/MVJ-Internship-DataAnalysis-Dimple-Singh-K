import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.model_selection import train_test_split
from sklearn.linear_model import LogisticRegression
from sklearn.tree import DecisionTreeClassifier
from sklearn.metrics import classification_report, confusion_matrix, accuracy_score
import os

# Set paths
base_dir = r"C:\Users\Dimple Singh K\OneDrive\Desktop\MVJ-Internship-DataAnalysis-Dimple-Singh-K"
day18_dir = os.path.join(base_dir, "Day18_classification")
os.makedirs(day18_dir, exist_ok=True)

print("Simulating/Loading Bank Customer Churn Dataset for Classification...")

# Simulating bank churn data (Credit Score, Age, Tenure, Balance, NumOfProducts, HasCrCard, IsActiveMember, EstimatedSalary -> Exited/Churned)
np.random.seed(42)
n_samples = 1000

credit_score = np.random.normal(650, 80, n_samples).clip(350, 850)
age = np.random.normal(38, 10, n_samples).clip(18, 70)
tenure = np.random.randint(0, 11, n_samples)
balance = np.random.exponential(50000, n_samples)
num_products = np.random.choice([1, 2, 3, 4], size=n_samples, p=[0.5, 0.4, 0.08, 0.02])
has_cr_card = np.random.choice([0, 1], size=n_samples, p=[0.3, 0.7])
is_active = np.random.choice([0, 1], size=n_samples, p=[0.4, 0.6])
salary = np.random.uniform(30000, 150000, n_samples)

# Churn logic probability based on factors
churn_prob = 1 / (1 + np.exp(-(-2 + 0.03*(age - 40) - 0.002*(credit_score - 650) - 0.8*is_active + 0.000005*balance)))
churn = np.random.binomial(1, churn_prob)

df_churn = pd.DataFrame({
    'CreditScore': credit_score,
    'Age': age,
    'Tenure': tenure,
    'Balance': balance,
    'NumOfProducts': num_products,
    'HasCrCard': has_cr_card,
    'IsActiveMember': is_active,
    'EstimatedSalary': salary,
    'Churn': churn
})

# Save dataset snapshot
df_churn.to_csv(os.path.join(day18_dir, "bank_customer_churn.csv"), index=False)

# 1. Define Features and Target
X = df_churn.drop('Churn', axis=1)
y = df_churn['Churn']

# 2. Train/Test Split (80% train, 20% test)
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

# 3. Train Logistic Regression
lr_model = LogisticRegression(max_iter=1000, random_state=42)
lr_model.fit(X_train, y_train)
y_pred_lr = lr_model.predict(X_test)

# 4. Train Decision Tree Classifier
dt_model = DecisionTreeClassifier(max_depth=5, random_state=42)
dt_model.fit(X_train, y_train)
y_pred_dt = dt_model.predict(X_test)

print("\n--- Logistic Regression Performance ---")
print(f"Accuracy: {accuracy_score(y_test, y_pred_lr):.4f}")
print(classification_report(y_test, y_pred_lr))

print("\n--- Decision Tree Classifier Performance ---")
print(f"Accuracy: {accuracy_score(y_test, y_pred_dt):.4f}")
print(classification_report(y_test, y_pred_dt))

# 5. Visualizing Confusion Matrices
fig, axes = plt.subplots(1, 2, figsize=(12, 5))

sns.heatmap(confusion_matrix(y_test, y_pred_lr), annot=True, fmt='d', cmap='Blues', ax=axes[0])
axes[0].set_title('Logistic Regression Confusion Matrix')
axes[0].set_xlabel('Predicted')
axes[0].set_ylabel('Actual')

sns.heatmap(confusion_matrix(y_test, y_pred_dt), annot=True, fmt='d', cmap='Purples', ax=axes[1])
axes[1].set_title('Decision Tree Confusion Matrix')
axes[1].set_xlabel('Predicted')
axes[1].set_ylabel('Actual')

plt.tight_layout()
cm_plot_path = os.path.join(day18_dir, "confusion_matrices.png")
plt.savefig(cm_plot_path, dpi=300, bbox_inches='tight')
print(f"\nConfusion matrices saved successfully at: {cm_plot_path}")