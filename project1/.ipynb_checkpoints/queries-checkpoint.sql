-- ==========================================================
-- PROJECT 1: ONLINE RETAIL DATA ANALYSIS (DAY 5 SQL BASICS)
-- ==========================================================

-- 1. Top 5 Countries by Total Transactions
SELECT Country, COUNT(Invoice) AS Total_Transactions
FROM online_retail
WHERE Quantity > 0 AND Price > 0
GROUP BY Country
ORDER BY Total_Transactions DESC
LIMIT 5;

-- 2. Top 5 Best-Selling Products by Total Revenue
SELECT Description, SUM(Quantity * Price) AS Total_Revenue
FROM online_retail
WHERE Quantity > 0 AND Price > 0
GROUP BY Description
ORDER BY Total_Revenue DESC
LIMIT 5;

-- 3. Monthly Revenue Trend
SELECT strftime('%Y-%m', InvoiceDate) AS Month, 
       SUM(Quantity * Price) AS Monthly_Revenue
FROM online_retail
WHERE Quantity > 0 AND Price > 0
GROUP BY Month
ORDER BY Month ASC;