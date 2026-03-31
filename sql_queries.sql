1- Category Analysis

SELECT 
    Category,
    SUM(Sales) AS Revenue,
    SUM(Profit) AS Profit,
    ROUND(SUM(Profit) * 1.0 / NULLIF(SUM(Sales), 0), 4) AS Margin
FROM orders_raw
GROUP BY Category;

2- Segment Analysis

SELECT 
    Segment,
    SUM(Sales) AS Revenue,
    SUM(Profit) AS Profit,
    ROUND(SUM(Profit) * 1.0 / NULLIF(SUM(Sales), 0), 4) AS Margin
FROM orders_raw
GROUP BY Segment;


3- Discount Grouping
 
SELECT 
    CASE 
        WHEN Discount = 0 THEN '0%'
        WHEN Discount <= 0.10 THEN '0–10%'
        WHEN Discount <= 0.20 THEN '10–20%'
        WHEN Discount <= 0.30 THEN '20–30%'
        ELSE '30%+'
    END AS Discount_Group,
    SUM(Sales) AS Revenue,
    SUM(Profit) AS Profit
FROM orders_raw
GROUP BY Discount_Group;

4- Time Analysis

SELECT 
    strftime('%Y', "Order Date") AS Year,
    SUM(Sales) AS Revenue,
    SUM(Profit) AS Profit
FROM orders_raw
GROUP BY Year
ORDER BY Year;
