import pandas as pd
import numpy as np
from scipy.stats import chi2_contingency

print("--- STARTING DAY 14: A/B TESTING & STATISTICAL SIGNIFICANCE ---")

# 1. Simulate or set up A/B Test Dataset (e.g., Conversion rates for two web/checkout designs)
# Group A: Control, Group B: Treatment/New Design
np.random.seed(42)

n_control = 5000
n_treatment = 5000

# Assume Control conversion rate ~ 10%, Treatment conversion rate ~ 12%
control_converted = np.random.binomial(1, 0.10, n_control)
treatment_converted = np.random.binomial(1, 0.12, n_treatment)

# Build a summary DataFrame / Contingency Table
control_success = control_converted.sum()
control_fail = n_control - control_success

treatment_success = treatment_converted.sum()
treatment_fail = n_treatment - treatment_success

contingency_table = np.array([
    [control_success, control_fail],
    [treatment_success, treatment_fail]
])

print("\nContingency Table (Converted vs Not Converted):")
print(f"Control   - Converted: {control_success}, Not Converted: {control_fail}")
print(f"Treatment - Converted: {treatment_success}, Not Converted: {treatment_fail}")

# 2. Perform Chi-Square Test for Independence
chi2, p_val, dof, ex = chi2_contingency(contingency_table)

print(f"\nChi-Square Statistic: {chi2:.4f}")
print(f"P-value: {p_val:.4f}")

# 3. Decision making
alpha = 0.05
print(f"\nSignificance Threshold (alpha): {alpha}")
if p_val < alpha:
    print("Result: Statistically significant difference between Control and Treatment! Reject the null hypothesis.")
else:
    print("Result: No statistically significant difference. Fail to reject the null hypothesis.")

# 4. Save results to a report note
report_content = f"""# Day 14: A/B Testing & Chi-Square Analysis Report

## Summary Statistics
- **Control Group Size:** {n_control} | **Conversions:** {control_success} ({control_success/n_control*100:.2f}%)
- **Treatment Group Size:** {n_treatment} | **Conversions:** {treatment_success} ({treatment_success/n_treatment*100:.2f}%)

## Statistical Test Results
- **Test Used:** Chi-Square Test of Independence (`scipy.stats.chi2_contingency`)
- **Chi2 Statistic:** {chi2:.4f}
- **P-value:** {p_val:.4f}
- **Conclusion:** {"Statistically significant" if p_val < alpha else "Not statistically significant"} (p {"<" if p_val < alpha else ">="} {alpha})
"""

with open('ab_test_report.md', 'w') as f:
    f.write(report_content)

print("\nSaved: ab_test_report.md")
print("--- DAY 14 A/B TESTING COMPLETE ---")