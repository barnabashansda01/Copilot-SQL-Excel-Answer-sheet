/* ============================================================
   Question 1
   Prompt given to GitHub Copilot in SSMS:
   "Create a table SalesData with columns CustomerID, Name, Age,
   City, PurchaseAmount, PurchaseDate. Then insert 10,000 rows
   of random data."
   ============================================================ */

-- Step 1: Create the table
CREATE TABLE SalesData (
    CustomerID    INT IDENTITY(1,1) PRIMARY KEY,
    Name          VARCHAR(50),
    Age           INT,
    City          VARCHAR(50),
    PurchaseAmount DECIMAL(10,2),
    PurchaseDate  DATE
);
GO

-- Step 2: Copilot-suggested approach to generate 10,000 random rows.
-- A numbers/tally CTE is used instead of a slow row-by-row loop
-- (this is the pattern Copilot proposes when asked to "insert
-- 10,000 rows of random data" for performance reasons).

;WITH Numbers AS (
    SELECT TOP (10000)
           ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) AS n
    FROM sys.all_objects a
    CROSS JOIN sys.all_objects b
),
Names AS (
    SELECT * FROM (VALUES
        ('Aarav'),('Vivaan'),('Aditya'),('Vihaan'),('Arjun'),
        ('Sai'),('Reyansh'),('Ayaan'),('Krishna'),('Ishaan'),
        ('Ananya'),('Diya'),('Saanvi'),('Aadhya'),('Kiara'),
        ('Myra'),('Anika'),('Navya'),('Riya'),('Meera')
    ) AS T(Name)
),
Cities AS (
    SELECT * FROM (VALUES
        ('Mumbai'),('Delhi'),('Bengaluru'),('Hyderabad'),('Chennai'),
        ('Kolkata'),('Pune'),('Ahmedabad'),('Jaipur'),('Lucknow')
    ) AS T(City)
)
INSERT INTO SalesData (Name, Age, City, PurchaseAmount, PurchaseDate)
SELECT
    (SELECT Name FROM (SELECT Name, ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) rn FROM Names) x
        WHERE rn = (ABS(CHECKSUM(NEWID())) % 20) + 1),
    18 + (ABS(CHECKSUM(NEWID())) % 50),                         -- Age 18-67
    (SELECT City FROM (SELECT City, ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) rn FROM Cities) y
        WHERE rn = (ABS(CHECKSUM(NEWID())) % 10) + 1),
    CAST(50 + (ABS(CHECKSUM(NEWID())) % 49950) AS DECIMAL(10,2)), -- Amount 50-50000
    DATEADD(DAY, -(ABS(CHECKSUM(NEWID())) % 1095), GETDATE())     -- last 3 years
FROM Numbers;
GO

-- Step 3: Verify
SELECT COUNT(*) AS TotalRowsInserted FROM SalesData;
SELECT TOP 10 * FROM SalesData;
