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


-- Project 1: Online Retail KPI Analysis Queries
-- Database: online_retail.db
-- Day 8:Queries
-- 1. Total Revenue
SELECT 
    ROUND(SUM(Quantity * Price), 2) AS Total_Revenue
FROM online_retail;

-- 2. Average Order Value (AOV)
SELECT 
    ROUND(SUM(Quantity * Price) / COUNT(DISTINCT Invoice), 2) AS Average_Order_Value
FROM online_retail;

-- 3. Top 10 Products by Revenue
SELECT 
    Description,
    SUM(Quantity) AS Total_Quantity_Sold,
    ROUND(SUM(Quantity * Price), 2) AS Total_Revenue
FROM online_retail
WHERE Description IS NOT NULL
GROUP BY Description
ORDER BY Total_Revenue DESC
LIMIT 10;

-- 4. Month-over-Month (MoM) Revenue Growth Trends
WITH MonthlyRevenue AS (
    SELECT 
        STRFTIME('%Y-%m', InvoiceDate) AS Month,
        ROUND(SUM(Quantity * Price), 2) AS Monthly_Revenue
    FROM online_retail
    WHERE InvoiceDate IS NOT NULL
    GROUP BY Month
)
SELECT 
    Month,
    Monthly_Revenue,
    LAG(Monthly_Revenue, 1) OVER (ORDER BY Month) AS Previous_Month_Revenue,
    ROUND(
        ((Monthly_Revenue - LAG(Monthly_Revenue, 1) OVER (ORDER BY Month)) / 
        LAG(Monthly_Revenue, 1) OVER (ORDER BY Month)) * 100, 2
    ) AS MoM_Growth_Percentage
FROM MonthlyRevenue;

-- 5. Repeat Customer Analysis
SELECT 
    `Customer ID`,
    COUNT(DISTINCT Invoice) AS Total_Orders,
    ROUND(SUM(Quantity * Price), 2) AS Total_Customer_Spend
FROM online_retail
WHERE `Customer ID` IS NOT NULL
GROUP BY `Customer ID`
HAVING Total_Orders > 1
ORDER BY Total_Customer_Spend DESC;