CREATE DATABASE cbis_icphDB;

CREATE TABLE Departments(
		id  INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
		department VARCHAR(255) NOT NULL
);

CREATE TABLE Employee_Types(
		id INT IDENTITY(1,1)  NOT NULL PRIMARY KEY,
		empType VARCHAR(255) NOT NULL
);

CREATE TABLE Employees(
		id INT IDENTITY(20170000,1)  NOT NULL PRIMARY KEY, 
		firstName VARCHAR(255) NOT NULL,
		middleName VARCHAR(255) NOT NULL,
		lastName VARCHAR(255) NOT NULL,
		gender CHAR(7) NOT NULL,
		address VARCHAR(255) NOT NULL,
		birthDate DATE NOT NULL,
		empTypeId INT NOT NULL,
		departmentId INT NOT NULL,
		FOREIGN KEY(empTypeId) REFERENCES Employee_Types(id) ON DELETE CASCADE ON UPDATE CASCADE, 
		FOREIGN KEY(departmentId) REFERENCES Departments(id) ON DELETE CASCADE ON UPDATE CASCADE
);

CREATE TABLE Users(
		id INT IDENTITY(20161000,1)  NOT NULL PRIMARY KEY,
		username VARCHAR(255),
		password VARCHAR(255),
		employeeId INT NOT NULL,
		dateCreated VARCHAR(255),
		dateUpdated VARCHAR(255),
		FOREIGN KEY(employeeId) REFERENCES Employees(id) ON DELETE CASCADE ON UPDATE CASCADE
);

CREATE TABLE Users_session_status(
		id INT IDENTITY(1,1)  NOT NULL PRIMARY KEY,
		userId INT NOT NULL,
		timeLogin VARCHAR(50),
		FOREIGN KEY(userId) REFERENCES Users(id) ON DELETE CASCADE ON UPDATE CASCADE
);

CREATE TABLE Users_activity_logs(
		id INT IDENTITY(1,1)  NOT NULL PRIMARY KEY,
		userId INT NOT NULL,
		dateLog VARCHAR(50),
		timeLogin VARCHAR(50),
		timeLogout VARCHAR(50),
		FOREIGN KEY(userId) REFERENCES Users(id)
);

CREATE TABLE AdminUsers(
		id INT IDENTITY(20162000,1) NOT NULL PRIMARY KEY,
		username VARCHAR(255) NOT NULL,
		password VARCHAR(255) NOT NULL,
		createdAt DATETIME DEFAULT CURRENT_TIMESTAMP,
		lastupdate VARCHAR(255),
		permission VARCHAR(5) NOT NULL
);


CREATE TABLE AdminSessionsStatus(
		id INT IDENTITY(1,1)  NOT NULL PRIMARY KEY,
		adminId INT NOT NULL,
		timeLogin VARCHAR(50),
		FOREIGN KEY(adminId) REFERENCES AdminUsers(id) ON DELETE CASCADE ON UPDATE CASCADE
);

CREATE TABLE AdminActivityLog(
		id INT IDENTITY(1,1)  NOT NULL PRIMARY KEY,
		adminId INT NOT NULL,
		dateLog VARCHAR(50),
		timeLogin VARCHAR(50),
		timeLogout VARCHAR(50),
		FOREIGN KEY(adminId) REFERENCES AdminUsers(id)
);

CREATE TABLE MedicineClassification(
		id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
		classification VARCHAR(255)
);

CREATE TABLE MedicineCategory(
		id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
		category VARCHAR(255)
);

CREATE TABLE Medicines(
		id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
		brandName VARCHAR(255),
		genericName VARCHAR(255),
		unitCost DECIMAL(19,4),
		categoryId INT NOT NULL,
		medClassificationId INT NOT NULL,
		FOREIGN KEY(categoryId) REFERENCES MedicineCategory(id),
		FOREIGN KEY(medClassificationId) REFERENCES MedicineClassification(id),
);

CREATE TABLE Suppliers(
		id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
		companyName VARCHAR(255),
		name VARCHAR(255),
		phoneNum VARCHAR(255),
		email VARCHAR(255),
		address VARCHAR(255),
);

CREATE TABLE MedicinesStockBatch(
		id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
		lotNumber VARCHAR(50),
		batchDate DATETIME DEFAULT CURRENT_TIMESTAMP,
		supplierId INT NOT NULL,
		FOREIGN KEY(supplierId) REFERENCES Suppliers(id)
);

CREATE TABLE MedicinesStockBatchItems(
		id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
		medicineId INT NOT NULL,
		quantity INT NOT NULL,
		currentStocks INT NOT NULL,
		expirationDate DATE,
		batchId INT NOT NULL,
		FOREIGN KEY(medicineId) REFERENCES Medicines(id),
		FOREIGN KEY(batchId) REFERENCES MedicinesStockBatch(id)
);


CREATE TABLE PharmacyStocks(
		id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
		quantity INT NOT NULL,
		currentStocksQuantity INT NOT NULL,
		medicineID INT NOT NULL,
		batchID INT NOT NULL,
		FOREIGN KEY(medicineID) REFERENCES Medicines(id),
		FOREIGN KEY(batchID) REFERENCES MedicinesStockBatch(id)
);

CREATE TABLE PharmacyCashierTransactions(
		id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
		transID VARCHAR(255),
		amount DECIMAL(19,4),
		transDate DATETIME DEFAULT CURRENT_TIMESTAMP,
		transStatus INT,
);

CREATE TABLE Temporary_PharmacyStockOutTransactions(
		id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
		pharCashierTransId INT NOT NULL,
		transDate DATETIME DEFAULT CURRENT_TIMESTAMP,
		stockOutQuantity INT NOT NULL,
		stockOutPrice DECIMAL(19,4),
		whopay VARCHAR(20),
		patientType CHAR(5),
		medicineID INT NOT NULL,
		userID INT NOT NULL
);

CREATE TABLE PharmacyStockOutTransactions(
		id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
		transDate DATETIME DEFAULT CURRENT_TIMESTAMP,
		stockOutQuantity INT NOT NULL,
		stockOutPrice DECIMAL(19,4),
		whopay VARCHAR(20),
		patientType VARCHAR(5),
		medicineID INT NOT NULL,
		userID INT NOT NULL,
		pharCashierTransID INT NOT NULL,
		FOREIGN KEY(medicineID) REFERENCES Medicines(id),
		FOREIGN KEY(userID) REFERENCES Users(id)
);

CREATE TABLE PharmacyRequests(
		id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
		requestDate DATETIME DEFAULT CURRENT_TIMESTAMP,
		requestStatus VARCHAR(30),
		userID INT NOT NULL,
		FOREIGN KEY(userID) REFERENCES Users(id)
);


CREATE TABLE PharmacyRequestsItems(
		id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
		medicineID INT NOT NULL,
		medQuantity INT NOT NULL,
		pharmaRequestID INT NOT NULL,
		FOREIGN KEY(pharmaRequestID) REFERENCES PharmacyRequests(id)
);


CREATE TABLE Room_types(
		rm_type_ID INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
		rm_type_name VARCHAR(255),
		rm_price DECIMAL(19,4)
);

CREATE TABLE Rooms(
		rm_ID INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
		rm_No VARCHAR(255),
		rm_type_ID INT NOT NULL,
		FOREIGN KEY(rm_type_ID) REFERENCES Room_types(rm_type_ID)
);

CREATE TABLE patient(
		pt_ID INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
		pt_Fname VARCHAR(255),
		pt_Mname VARCHAR(255),
		pt_LName VARCHAR(255),
);

CREATE TABLE patient_admission(
		pt_admit_Id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
		pt_status CHAR(20),
		pt_ID INT NOT NULL,
		pt_adm_startDate DATETIME DEFAULT CURRENT_TIMESTAMP,
		pt_adm_EndDate DATETIME,
		rm_ID INT NOT NULL,
		pt_id_price DECIMAL(19,4),
		pt_admit_price DECIMAL(19,4),
		FOREIGN KEY(pt_ID) REFERENCES patient(pt_ID),
		FOREIGN KEY(rm_ID) REFERENCES Rooms(rm_ID),
);


CREATE TABLE patient_medicine(
		pt_med_ID INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
		med_ID INT NOT NULL,
		quantity INT NOT NULL,
		totalprice DECIMAL(19,4),
		pt_adm_ID INT NOT NULL,
		FOREIGN KEY(pt_adm_ID) REFERENCES patient_admission(pt_admit_Id)
);

CREATE TABLE doctor_department(
		dr_dept_ID INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
		dr_dept_Name VARCHAR(255)
);


CREATE TABLE doctor(
		dr_ID INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
		dr_Name VARCHAR(255),
		dr_dept_ID INT NOT NULL,
		FOREIGN KEY(dr_dept_ID) REFERENCES doctor_department(dr_dept_ID)
);

CREATE TABLE patient_doctor(
		pt_dr_ID INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
		dr_ID INT NOT NULL,
		doctor_fee DECIMAL(19,4),
		pt_adm_ID INT NOT NULL,
		FOREIGN KEY(dr_ID) REFERENCES doctor(dr_ID),
		FOREIGN KEY(pt_adm_ID) REFERENCES patient_admission(pt_admit_Id)
);


CREATE TABLE services(
		srvc_ID INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
		srvc_Name VARCHAR(255),
		srvc_Price Decimal(19,4)
);

CREATE TABLE patient_laboratory_services(
		pt_lab_ID INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
		pt_adm_ID INT NOT NULL,
		srvc_ID INT NOT NULL,
		FOREIGN KEY(pt_adm_ID) REFERENCES patient_admission(pt_admit_Id),
		FOREIGN KEY(srvc_ID) REFERENCES services(srvc_ID)
);

CREATE TABLE Supplies(
		sup_id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
		sup_name VARCHAR(255),
		sup_time DATETIME DEFAULT CURRENT_TIMESTAMP,
		sup_price DECIMAL(19,4),
		sup_qty INT NOT NULL,
		sup_total DECIMAL(19,4),
		pt_adm_ID INT NOT NULL,
		FOREIGN KEY(pt_adm_ID) REFERENCES patient_admission(pt_admit_Id)
);

CREATE TABLE laboratory_transaction_out_patient(
		trns_ID INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
		Name VARCHAR(255),
		Gender CHAR(10),
		srvcID INT NOT NULL,
		status INT NOT NULL,
		date DATETIME DEFAULT CURRENT_TIMESTAMP,
		FOREIGN KEY(srvcID) REFERENCES services(srvc_ID)
);

CREATE TABLE CashierAuditTrail(
		cash_id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
		cash_transac_no VARCHAR(255),
		cash_transac_date DATETIME DEFAULT CURRENT_TIMESTAMP,
		cash_dept VARCHAR(255),
		cash_total DECIMAL(19,4),
		cash_actor_id INT NOT NULL,
		FOREIGN KEY(cash_actor_id) REFERENCES Users(id)
);

CREATE TABLE ERPharmacyStocks(
		id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
		quantity INT NOT NULL,
		currentStocksQuantity INT NOT NULL,
		medicineID INT NOT NULL,
		batchID INT NOT NULL,
		FOREIGN KEY(medicineID) REFERENCES Medicines(id),
		FOREIGN KEY(batchID) REFERENCES MedicinesStockBatch(id)
);


CREATE TABLE ERPharmacySales(
		id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
		transactionNO VARCHAR(255),
		totalPurchaseAmount DECIMAL(19,4),
		transDate DATETIME DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE ERPharmacyStockOutTransactions(
		id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
		transDate DATETIME DEFAULT CURRENT_TIMESTAMP,
		stockOutQuantity INT NOT NULL,
		stockOutPrice DECIMAL(19,4),
		transactionNO VARCHAR(255),
		whopay VARCHAR(20),
		patientType VARCHAR(5),
		medicineID INT NOT NULL,
		userID INT NOT NULL,
		FOREIGN KEY(medicineID) REFERENCES Medicines(id),
		FOREIGN KEY(userID) REFERENCES Users(id)
);

CREATE TABLE ERPharmacyRequests(
		id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
		requestDate DATETIME DEFAULT CURRENT_TIMESTAMP,
		requestStatus VARCHAR(30),
		userID INT NOT NULL,
		FOREIGN KEY(userID) REFERENCES Users(id)
);

CREATE TABLE ERPharmacyRequestsItems(
		id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
		medicineID INT NOT NULL,
		medQuantity INT NOT NULL,
		erpharmaRequestID INT NOT NULL,
		FOREIGN KEY(erpharmaRequestID) REFERENCES ERPharmacyRequests(id)
);
