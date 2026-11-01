# Day 14: A/B Testing & Funnel Analysis Report

## 1. Objective
Evaluate whether the new checkout page redesign (Variant B) results in a statistically significant improvement in conversion rate compared to the standard layout (Variant A).

## 2. Methodology
* **Statistical Test:** Chi-Square Test of Independence (`scipy.stats.chi2_contingency`).
* **Significance Level ($\alpha$):** 0.05.
* **Sample Size:** 
  * Variant A (Control): 10,000 users
  * Variant B (Treatment): 10,200 users

## 3. Results Summary
* **Variant A Conversion Rate:** 12.00%
* **Variant B Conversion Rate:** 13.24%
* **Relative Lift:** +10.33%
* **P-value:** < 0.01

## 4. Business Recommendation
Since the p-value is well below the 0.05 threshold, the increase in conversion is statistically significant and not due to random chance. We recommend rolling out Variant B to 100% of users.