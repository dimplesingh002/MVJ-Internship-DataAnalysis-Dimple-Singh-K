import pandas as pd
from scipy.stats import chi2_contingency

# Sample A/B Testing Data (e.g., Checkout Page Redesign)
# Variant A (Control): 10,000 visitors, 1,200 conversions
# Variant B (Treatment): 10,200 visitors, 1,350 conversions

data = [
    [1200, 10000 - 1200],  # Variant A: [Conversions, Non-Conversions]
    [1350, 10200 - 1350]   # Variant B: [Conversions, Non-Conversions]
]

# Perform Chi-Square Test of Independence
chi2, p_value, dof, ex = chi2_contingency(data)

conv_rate_a = 1200 / 10000
conv_rate_b = 1350 / 10200
lift = (conv_rate_b - conv_rate_a) / conv_rate_a * 100

print(f"Variant A Conversion Rate: {conv_rate_a:.2%}")
print(f"Variant B Conversion Rate: {conv_rate_b:.2%}")
print(f"Observed Lift: {lift:.2f}%")
print(f"Chi2 Statistic: {chi2:.4f}")
print(f"P-value: {p_value:.5f}")

if p_value < 0.05:
    print("Result: Statistically significant difference! Reject the null hypothesis.")
else:
    print("Result: No statistically significant difference. Fail to reject null hypothesis.")