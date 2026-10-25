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

-- ============================================================
-- Day 6: Query 1 - Daily Revenue & Running Total
-- ============================================================
WITH DailySales AS (
    SELECT 
        SUBSTR(InvoiceDate, 1, 10) AS SaleDate,
        SUM(Quantity * Price) AS DailyRevenue
    FROM online_retail
    WHERE Quantity > 0 AND Price > 0
    GROUP BY SaleDate
)
SELECT 
    SaleDate,
    ROUND(DailyRevenue, 2) AS DailyRevenue,
    ROUND(SUM(DailyRevenue) OVER (ORDER BY SaleDate), 2) AS RunningTotalRevenue
FROM DailySales
ORDER BY SaleDate;


-- ============================================================
-- Day 6: Query 2 - Customer Spending Rank using Window Functions
-- ============================================================
WITH CustomerSpending AS (
    SELECT 
        "Customer ID" AS CustomerID,
        Country,
        SUM(Quantity * Price) AS TotalSpent,
        COUNT(DISTINCT Invoice) AS TotalOrders
    FROM online_retail
    WHERE "Customer ID" IS NOT NULL AND Quantity > 0 AND Price > 0
    GROUP BY "Customer ID", Country
)
SELECT 
    CustomerID,
    Country,
    ROUND(TotalSpent, 2) AS TotalSpent,
    TotalOrders,
    RANK() OVER (ORDER BY TotalSpent DESC) AS SpendingRank
FROM CustomerSpending
LIMIT 20;