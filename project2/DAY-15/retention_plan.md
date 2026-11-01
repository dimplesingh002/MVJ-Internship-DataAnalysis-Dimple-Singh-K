# Project 2: Customer Segmentation & Retention Strategy
**Prepared by:** Dimple Singh K | Register No: U19XF24S0021  
**Institution:** MVJ Degree College, BCA Data Science  

---

## 1. Executive Summary
Using RFM (Recency, Frequency, Monetary) modeling and monthly cohort retention analysis on the Online Retail II dataset, we identified key customer behaviors. A significant drop-off in customer engagement occurs after Month 1, highlighting the need for targeted retention interventions.

## 2. Target Segment: "At Risk" & "Champions"
* **Champions:** High recency, frequency, and monetary spend. **Action:** Reward with VIP loyalty perks, early access to new stock, and referral incentives.
* **At Risk Customers:** Previously high frequency and spend, but high recency (haven't purchased recently). **Action:** Deploy win-back email campaigns offering personalized discounts based on their past purchase history.

## 3. Recommended Retention Plan (One-Page Strategy)
1. **Automated Onboarding (Month 1):** Trigger a welcome series and a thank-you discount code immediately following a customer's first purchase to prevent early drop-off seen in the cohort matrix.
2. **Re-engagement Triggers:** Automatically tag customers as "At Risk" if 90 days pass without a transaction, launching an automated re-engagement email sequence.
3. **Product Recommendations:** Utilize collaborative filtering based on top-selling inventory (e.g., Regency Cake Stands and White Hanging Heart T-Light Holders) for cart upsells.