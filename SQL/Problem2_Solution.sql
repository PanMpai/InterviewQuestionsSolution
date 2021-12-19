

SELECT employee.Name
FROM Employee AS employee
JOIN Employee AS manager ON manager.Id = employee.ManagerId 
WHERE employee.Salary > manager.Salary