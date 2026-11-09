# Day 18: Predictive Analytics 2 – Classification Report
**Prepared by:** Dimple Singh K | Register No: U19XF24S0021  
**Institution:** MVJ Degree College, BCA Data Science  

---

## 1. Objective
Analyze bank customer churn data using classification algorithms (**Logistic Regression** and **Decision Trees**) to predict whether a customer will leave the bank, focusing on evaluation metrics like accuracy, precision, recall, and confusion matrices.

## 2. Model Evaluation Highlights
* **Logistic Regression:** Provides a linear baseline probability estimation. Useful for interpreting direct coefficient impacts (e.g., how active membership reduces churn).
* **Decision Tree:** Captures non-linear feature splits (e.g., combinations of customer age, balance, and product counts).

## 3. Business Importance of Precision vs. Recall
* **Recall (Sensitivity):** Crucial for identifying as many actual churners as possible so the bank can deploy proactive retention offers.
* **Precision:** Ensures retention offers aren't unnecessarily wasted on loyal customers who weren't going to leave.

## 4. Deliverables Committed
* `bank_customer_churn.csv`
* `bank_churn_classification.py`
* `confusion_matrices.png`
* `classification_report_notes.md`