# Day 20: Power BI Data Modeling & Star Schema
**Prepared by:** Dimple Singh K | Register No: U19XF24S0021  
**Institution:** MVJ Degree College, BCA Data Science  

---

## 1. Data Model Architecture (Star Schema)
* **Fact_Sales:** Contains transaction IDs, Quantity, UnitPrice, CustomerID, and StockCode.
* **Dim_Customer:** Contains CustomerID, Region, and Segment attributes.
* **Dim_Product:** Contains StockCode, Description, and Category.
* **Dim_Calendar:** Generated date table covering all transaction date ranges for time intelligence.

## 2. Relationship Mapping
* Established 1-to-Many ($1: \infty$) relationships from Dimension tables (`Dim_Customer`, `Dim_Product`, `Dim_Calendar`) to the central Fact table (`Fact_Sales`).