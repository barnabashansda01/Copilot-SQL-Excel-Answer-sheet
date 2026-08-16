/* ============================================================
   Question 2
   Prompts given to GitHub Copilot in SSMS:
   a) "Find total sales per city"
   b) "Top 5 cities by revenue"
   ============================================================ */

-- a) Total sales per city
SELECT
    City,
    SUM(PurchaseAmount) AS TotalSales,
    COUNT(*)            AS NumberOfOrders,
    AVG(PurchaseAmount) AS AvgOrderValue
FROM SalesData
GROUP BY City
ORDER BY TotalSales DESC;
GO

-- b) Top 5 cities by revenue
SELECT TOP 5
    City,
    SUM(PurchaseAmount) AS TotalRevenue
FROM SalesData
GROUP BY City
ORDER BY TotalRevenue DESC;
GO
