/* ============================================================
   Question 3
   Prompt given to GitHub Copilot in SSMS:
   "Find customers with purchases above average"
   ============================================================ */

SELECT
    CustomerID,
    Name,
    City,
    PurchaseAmount,
    (SELECT AVG(PurchaseAmount) FROM SalesData) AS OverallAverage
FROM SalesData
WHERE PurchaseAmount > (SELECT AVG(PurchaseAmount) FROM SalesData)
ORDER BY PurchaseAmount DESC;
GO

-- Optional: count how many customers are above average
SELECT
    COUNT(*) AS CustomersAboveAverage,
    (SELECT COUNT(*) FROM SalesData) AS TotalCustomers,
    CAST(COUNT(*) * 100.0 / (SELECT COUNT(*) FROM SalesData) AS DECIMAL(5,2)) AS PercentAboveAverage
FROM SalesData
WHERE PurchaseAmount > (SELECT AVG(PurchaseAmount) FROM SalesData);
GO
