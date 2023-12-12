CREATE DATABASE WarehouseManagementDB;
USE WarehouseManagementDB;

CREATE TABLE Equipment (
    EquipmentID INT PRIMARY KEY IDENTITY(1,1),
    Name NVARCHAR(100) NOT NULL,
    Category NVARCHAR(50),
    PurchaseDate DATE,
    Price INT,
    Quantity INT,
    Location NVARCHAR(50)
);

CREATE TABLE EquipmentMovement (
    MovementID INT PRIMARY KEY IDENTITY(1,1),
    EquipmentID INT,
    MovementDate DATETIME,
    MovementType NVARCHAR(10) CHECK (MovementType IN ('IN', 'OUT')),
    Quantity INT,
    FOREIGN KEY (EquipmentID) REFERENCES Equipment(EquipmentID)
);

CREATE TABLE Supplier (
    SupplierID INT PRIMARY KEY IDENTITY(1,1),
    Name NVARCHAR(100) NOT NULL,
    ContactPerson NVARCHAR(50),
    Phone NVARCHAR(20),
    Email NVARCHAR(50)
);

CREATE TABLE EquipmentSupplier (
    EquipmentID INT,
    SupplierID INT,
    PRIMARY KEY (EquipmentID, SupplierID),
    FOREIGN KEY (EquipmentID) REFERENCES Equipment(EquipmentID),
    FOREIGN KEY (SupplierID) REFERENCES Supplier(SupplierID)
);

CREATE TABLE Registration (
	UserID INT PRIMARY KEY IDENTITY(1,1),
	UserLogin VARCHAR(50),
	UserPassword VARCHAR(50),
	IsAdmin bit
);

INSERT INTO Equipment (Name, Category, PurchaseDate, Price, Quantity, Location)
VALUES 
    ('Laptop', 'Electronics', '2023-01-15', 1200, 5, 'Office A'),
    ('Projector', 'Electronics', '2023-02-20', 1500, 2, 'Conference Room'),
    ('Printer', 'Office Supplies', '2023-03-10', 500, 3, 'Printer Room');

INSERT INTO EquipmentMovement (EquipmentID, MovementDate, MovementType, Quantity)
VALUES 
    (1, '2023-01-01', 'IN', 2),
    (2, '2023-01-01', 'OUT', 1),
    (3, '2023-01-01', 'IN', 5);

INSERT INTO Supplier (Name, ContactPerson, Phone, Email)
VALUES 
    ('Tech Supplies Inc.', 'John Doe', '123-456-7890', 'john.doe@techsupplies.com'),
    ('Office Solutions Co.', 'Jane Smith', '987-654-3210', 'jane.smith@officesolutionsco.com');

INSERT INTO EquipmentSupplier (EquipmentID, SupplierID)
VALUES 
    (1, 1),
    (2, 2),
    (3, 1),
    (3, 2);

INSERT INTO Registration (UserLogin, UserPassword, IsAdmin)
VALUES
	('admin', 'admin', 1),
	('user', 'user', 0);

SELECT * FROM Equipment;
SELECT * FROM EquipmentMovement;
SELECT * FROM Supplier;
SELECT * FROM EquipmentSupplier;
SELECT * FROM Registration;

DROP DATABASE WarehouseManagementDB;