
SELECT Salary  
FROM 
    (SELECT Salary 
     FROM Employee 
     WHERE ManagerId IS NOT NULL AND (SELECT count(*) FROM Employee WHERE ManagerId IS NOT NULL)  >=2 
     ORDER BY Salary DESC 
     LIMIT 3) AS Comp 
ORDER BY Salary 
LIMIT 1;