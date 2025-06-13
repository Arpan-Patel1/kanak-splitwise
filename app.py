IF OBJECT_ID('Employee', 'U') IS NULL
CREATE TABLE Employee (
    ID INT PRIMARY KEY,
    Name NVARCHAR(255) NOT NULL,
    Department NVARCHAR(255) NOT NULL,
    Salary FLOAT NOT NULL
);
