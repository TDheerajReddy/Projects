CREATE DATABASE BloodBankDB;
USE BloodBankDB;
SHOW TABLES;

CREATE TABLE Donor(Donor_ID INT PRIMARY KEY, Name VARCHAR(50), Date_Of_Birth DATE, Gender VARCHAR(6), Address VARCHAR(70));
desc donor;

CREATE TABLE Blood_Bank_Manager(Manager_ID INT PRIMARY KEY, Name VARCHAR(50), Phone_Number VARCHAR(13));
desc Blood_Bank_Manager;

CREATE table Blood_Bank(Blood_Bank_Id INT PRIMARY KEY, Address VARCHAR(70));
desc Blood_Bank;

CREATE TABLE Blood_Unit(Blood_Code INT UNIQUE, Blood_Amount INT, Donate_Date DATE, Blood_Type VARCHAR(3));
desc Blood_Unit;

CREATE TABLE Hospital(Hospital_Id INT PRIMARY KEY, Name VARCHAR(50), Address VARCHAR(70), Email VARCHAR(50));
desc Hospital;

CREATE TABLE Donor_Blood_Bank (
    Donor_ID INT,
    Blood_Bank_ID INT,
    PRIMARY KEY (Donor_ID, Blood_Bank_ID),
    FOREIGN KEY (Donor_ID) REFERENCES Donor(Donor_ID),
    FOREIGN KEY (Blood_Bank_ID) REFERENCES Blood_Bank(Blood_Bank_ID)
);
desc Donor_Blood_Bank;

ALTER TABLE Blood_Bank ADD COLUMN Manager_Id INT UNIQUE, 
ADD CONSTRAINT fk_manager FOREIGN KEY (Manager_Id) REFERENCES Blood_Bank_Manager(Manager_Id);

ALTER TABLE Blood_Unit ADD COLUMN Donor_ID INT NOT NULL,
ADD CONSTRAINT fk_donor FOREIGN KEY (Donor_ID) REFERENCES Donor(Donor_ID);

ALTER TABLE Hospital ADD COLUMN Blood_Bank_Id INT NOT NULL,
ADD CONSTRAINT fk_blood_bank_id FOREIGN KEY (Blood_Bank_Id) REFERENCES Blood_Bank(Blood_Bank_Id);

SELECT * FROM Donor;
SELECT * FROM Blood_Unit;
SELECT * FROM Blood_Bank_Manager;
SELECT * FROM Blood_Bank;
SELECT * FROM Donor_Blood_Bank;
SELECT * FROM Hospital;

-- TRUNCATE Hospital; -- 

SELECT d.Donor_Id, b.Blood_Code, d.Name, b.Blood_Type, b.Blood_Amount, b.Donate_Date
FROM Donor d 
JOIN Blood_Unit b ON d.Donor_Id = b.Donor_Id;


SELECT b.Blood_Bank_Id, b.Address, d.Donor_Id, d.Name FROM Blood_Bank b
JOIN Donor_Blood_Bank db ON b.Blood_Bank_Id = db.Blood_Bank_Id
JOIN Donor d ON db.Donor_Id = d.Donor_Id;		-- Where a Donor is donating his/her blood to which Blood Bank ?

SELECT b.Blood_Bank_Id, b.Address
FROM Blood_Bank b
LEFT JOIN Hospital h ON b.Blood_Bank_Id = h.Blood_Bank_Id
WHERE h.Blood_Bank_Id IS NULL;		-- Not assigned to any hospital
-- OR
SELECT b.Blood_Bank_Id, b.Address
FROM Blood_Bank b
WHERE b.Blood_Bank_Id NOT IN 
(SELECT Blood_Bank_Id FROM Hospital);

/*EXPLAIN SELECT b.Blood_Bank_Id, b.Address
FROM Blood_Bank b
LEFT JOIN Hospital h ON b.Blood_Bank_Id = h.Blood_Bank_Id
WHERE h.Blood_Bank_Id IS NULL; */


-- Find Top 3 Donors with the Highest Total Blood Donated
SELECT d.Donor_ID, d.Name, SUM(bu.Blood_Amount) AS Total_Amount
FROM Donor d
JOIN Blood_Unit bu ON d.Donor_ID = bu.Donor_ID
GROUP BY d.Donor_ID, d.Name
ORDER BY Total_Amount DESC
LIMIT 3;


-- Monthly Donation Summary
SELECT 
    DATE_FORMAT(Donate_Date, '%Y-%m') AS Month,
    COUNT(*) AS Total_Donations,
    SUM(Blood_Amount) AS Total_Blood_Collected
FROM Blood_Unit
GROUP BY Month
ORDER BY Month DESC;


-- Hospitals with Total Blood Collected via Associated Donors
SELECT h.Hospital_Id, h.Name AS Hospital_Name,
	SUM(bu.Blood_Amount) AS Total_Blood_Received
FROM Hospital h
JOIN Blood_Bank bb ON h.Blood_Bank_Id = bb.Blood_Bank_Id
JOIN Donor_Blood_Bank dbb ON bb.Blood_Bank_Id = dbb.Blood_Bank_Id
JOIN Blood_Unit bu ON dbb.Donor_ID = bu.Donor_ID
GROUP BY h.Hospital_Id, h.Name;


-- Blood Type Shortage Detection (Threshold < 150 Units)
SELECT Blood_Type, SUM(Blood_Amount) AS Total_Available
FROM Blood_Unit
GROUP BY Blood_Type
HAVING Total_Available < 150;


-- Donors Who Haven’t Donated in the Last 6 Months
SELECT d.Donor_ID, d.Name, MAX(bu.Donate_Date) AS Last_Donation
FROM Donor d
LEFT JOIN Blood_Unit bu ON d.Donor_ID = bu.Donor_ID
GROUP BY d.Donor_ID, d.Name
HAVING Last_Donation IS NOT NULL OR Last_Donation < CURDATE() - INTERVAL 6 MONTH;


-- Blood Type Distribution Per Blood Bank
SELECT bb.Blood_Bank_Id, bb.Address, bu.Blood_Type,
    SUM(bu.Blood_Amount) AS Total_Blood
FROM Blood_Bank bb
JOIN Donor_Blood_Bank dbb ON bb.Blood_Bank_Id = dbb.Blood_Bank_Id
JOIN Blood_Unit bu ON dbb.Donor_ID = bu.Donor_ID
GROUP BY bb.Blood_Bank_Id, bu.Blood_Type;


-- List of Blood Banks Not Linked to Any Donor
SELECT b.Blood_Bank_Id, b.Address
FROM Blood_Bank b
LEFT JOIN Donor_Blood_Bank db ON b.Blood_Bank_Id = db.Blood_Bank_Id
WHERE db.Donor_ID IS NULL;

