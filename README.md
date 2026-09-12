# CBIS — Centralized Billing and Information System (Sub-Systems)

**Immaculate Concepcion Polyclinic and Hospital (ICPH)**

A suite of four Windows desktop applications that together power the hospital's centralized billing and pharmacy information system. Built with **VB.NET / WinForms (.NET Framework 4.0)** on **Microsoft SQL Server**, the system covers the complete pharmacy and billing lifecycle — from Excel-based medicine intake, to point-of-sale dispensing, to centralized administration, multi-department audit trails, and self-service sales reporting.

---

## Executive Summary

CBIS connects every pharmacy and billing touchpoint in the hospital into one SQL Server database (`cbis_icphDB`). Rather than four isolated programs, each subsystem is a dedicated, purpose-built module of a single integrated platform:

- **InputStocks** — the stock intake engine that turns drug suppliers' Excel spreadsheets into live, validated pharmacy inventory.
- **Pharmacy Point-of-Sale** — the main hospital pharmacy's cashier front-end for selling and dispensing medicine to in-patients and out-patients.
- **ER Pharmacy POS** — a second, Emergency-Room-specific pharmacy front-end with its own stock, requisition flow, and cash/change handling.
- **Admin Dashboard** — the central back-office console for HR, user accounts, inventory oversight, room management, consolidated billing views, and daily/weekly/monthly/yearly reporting.

### What this project demonstrates

- **Full stack desktop engineering** — layered architecture separating UI (Forms) from business logic (Controller classes) over a shared ADO.NET data layer.
- **Relational database design** — a 35+ table schema spanning authentication, HR, inventory, admissions, billing, and cashier auditing, with referential integrity and cascading deletes.
- **Security engineering** — SHA-512 password hashing, permission-tiered admin rights, live online-user session tracking, and login/logout audit logging.
- **Workflow design for real operations** — staged sales pipelines (Unpaid → Paid → Claimed), in-patient charge-to-bill flows, returns and post-payment corrections, and branch-to-central stock requisitions.
- **Office automation** — full Excel COM integration: bulk data ingestion on one side and one-click report generation (`.xlsx`) on the other.
- **Performance-conscious UI** — heavy grids and report queries run on dedicated background threads so the interface stays responsive.

---

## System Architecture

All four applications are independent executables that talk to the same SQL Server database and are deployed to different departments of the hospital.

### Technology Stack

| Layer | Technology |
|---|---|
| Language | VB.NET |
| UI Framework | Windows Forms (WinForms) |
| Runtime | .NET Framework 4.0 (Client Profile) |
| Database | Microsoft SQL Server (`cbis_icphDB`) |
| Data Access | ADO.NET (`System.Data.SqlClient`) — `SqlCommand` / `SqlDataReader` |
| Reporting Export | Microsoft Excel Interop (`.xlsx` generation) |
| Bulk Data Ingestion | Microsoft Excel COM Automation |
| Password Security | SHA-512 (hex-digest hashing) |
| Concurrency | Background `System.Threading.Thread` workers for bulk imports and report generation |
| Platform | Windows, x86 |

### Architecture Pattern

Every subsystem follows the same clean, consistent pattern:

```
WinForms Form (UI)  →  Controller Class (business logic)  →  SqlControl (data access base)  →  SQL Server
```

- **`SqlControl.vb`** is the shared base class holding the SQL Server connection and a generic `ExecuteCommand()` helper for all inserts/updates/deletes.
- **`Security.vb`** provides SHA-512 hashing and safe hash comparison shared by every login flow.
- **`Authentication.vb`** handles credential verification, session creation/teardown, and activity-log recording.
- Controllers expose methods that execute `SqlCommand` objects, iterate a `SqlDataReader`, and feed rows back to the form through delegate callbacks.
- Report and import screens offload work to background threads to keep the UI live.

---

## Sub-System 1: AdminDashboard

**Role:** Central administration and back-office console. The main "Admin Panel" every authorized administrator lands on, giving complete control over staff, accounts, inventory, rooms, cashier operations, and hospital-wide reporting.

### Feature Breakdown

**Security & Authentication**

- SHA-512-hashed administrator credentials (`AdminUsers`), verified on every login.
- Session tracking: every admin login is recorded in `AdminSessionsStatus` (marking the admin online) and appended to `AdminActivityLog`.
- Graceful logout: sessions are torn down and the logout time is stamped on the latest activity-log entry.
- **Permission-tiered access** (`wrx` / `rw` / `r` semantics): write operations such as adding or updating employees are denied with a "Permission denied" dialog unless the admin holds a write-level permission.
- Full **Login / Logout / Change Username / Change Password** management, where changing credentials requires re-entering the current password (re-validated against the stored hash).

**Employee Management**

- Complete employee directory CRUD: add, update, remove, refresh, and search (by employee id) over the `Employees` table.
- Fields cover full name, gender, address, birth date, employee type, and department, with required-field validation.
- Dedicated **Departments** manager showing each department with a **live employee headcount**.
- Dedicated **Employee Types** manager showing each position/type with the **number of employees assigned**.
- Dual-purpose Add/Update form that switches context and pre-fills from the selected grid row.

**User Account Management**

- Create system user accounts and bind them to employee records (parent-employee must exist; duplicate usernames are rejected).
- Update usernames/passwords with SHA-512 hashing and a `dateUpdated` stamp; remove users with confirmation.
- Per-account intelligence panel showing the linked employee's **full name and department**, live **Online/Offline status**, and **last login time**.
- Show / hide password convenience toggle on the sign-up form.

**Pharmacy & Inventory Oversight**

- **Pharmacy Stocks** browser: browse medicine inventory by batch (lot), category, and classification, with live as-you-type search across product, category, and classification names and a combined multi-axis filter.
- **ER Pharmacy Stocks** browser: the same inventory experience against the Emergency Room pharmacy's independent stock (products are only sellable from lots with positive stock).
- **Stock Room** browser: the central warehouse view tied to incoming supplier batches, including batch date, a **View Supplier** lookup, per-batch stock details, and combined filtering.
- Every inventory view shows lot number, brand/generic name, category, classification, unit cost, **current stock**, and **expiration date** — a full lot-and-expiry traceability picture.
- **Rooms & Room Types**: add/update/remove hospital rooms and room types (with per-type pricing).

**Billing & Cashier Audit Trails**

- **Cashier-Pharmacy Audit Trail**: daily and weekly/monthly/yearly reports over `PharmacyCashierTransactions` with paid/claimed status mapping and a running sales total.
- **Pharmacy Sales Transaction / Audit Trail Report**: itemized stock-out history joined with medicine names and cashier usernames, filterable by **All / In-Patient / Out-Patient**, plus **Daily/WMY** and **Daily/WMY Summary** (per-medicine aggregated) views.
- **ER-Pharmacy Sales Report**: the same reporting suite for the Emergency Room pharmacy's independent sales.
- **Cashier Audit Trail**: department-level transaction records with the responsible cashier's username and cumulative totals.
- **Laboratory Out-Patient Audit Trail**: laboratory service transactions with service name/price and paid/claimed status.
- **Admission-Billing to Cashier Liquidation**: for discharged in-patients, automatically aggregates the **entire itemized hospital bill** — room charge (rate × stay days), medicines, professional/pro-fee, laboratory services, and supplies — into a per-patient grand total.

### Dedicated Data Scope

Auth/session (`AdminUsers`, `AdminSessionsStatus`, `AdminActivityLog`), HR (`Employees`, `Departments`, `Employee_Types`, `Users`, `Users_session_status`, `Users_activity_logs`), inventory (`PharmacyStocks`, `ERPharmacyStocks`, `Medicines*`, `Suppliers`, `MedicinesStockBatch*`), admissions/rooms (`patient*`, `Rooms`, `Room_types`, `patient_doctor`, `patient_laboratory_services`, `services`, `Supplies`), and cashier/audit (`PharmacyCashierTransactions`, `PharmacyStockOutTransactions`, `ERPharmacyStockOutTransactions`, `CashierAuditTrail`, `laboratory_transaction_out_patient`).

---

## Sub-System 2: Pharmacy Point-of-Sale

**Role:** The main hospital pharmacy's cashier front-end. Search for medicine, build a cart, price and validate, then complete the sale either as cash (out-patient) or as a charge to an in-patient's hospital bill.

### Feature Breakdown

**Login & Cashier Session**

- SHA-512 password hashing with credential verification against the `Users` table.
- Every login/logout is timestamped into the activity log; a live session table tracks online users.
- The logged-in cashier's username is displayed on the main screen header.

**Product Search**

- **Live/predictive search by brand or generic name** — the product grid refreshes on every keystroke.
- Each result shows lot number, generic/trade name, category, classification, unit cost, **current available stock**, and expiration date — with only in-stock lots surfaced.

**Cart & Pricing Controls**

- Keyboard-validated numeric fields (selling price allows a single decimal; quantity is digits-only).
- **Profit-floor guard** — selling price cannot be below the medicine's unit cost.
- **Stock-aware quantity validation** — quantities must be positive and never exceed available stock ("Insufficient Stock" warnings).
- Re-add same medicine to replace its cart line; decrease quantity; remove; clear; and a **live running total**.

**Dual Patient-Type Dispensing Workflow**

- **Out-Patient (OUT-P) — cashier pipeline:**
  1. Transaction number auto-built as fiscal year + transaction id (e.g., `2016…`).
  2. Sale staged as **UNPAID** in a temporary table with items, price, payor (`who pays`), and patient type.
  3. Cashier receives the transaction; patient pays.
  4. A single **Done** action commits the sale, **decrements the correct lot's stock**, marks the transaction **CLAIMED**, and prevents double-dispensing.
- **In-Patient (IN-P) — charge-to-bill:**
  - Cashier validates the patient exists and is actively admitted, and auto-fills the patient's name and room from the hospital system.
  - Medicines are written directly to `patient_medicine` **against the patient's hospital admission bill**, stock is decremented, and the sale is recorded with patient type `IN`.
  - Deep integration with the hospital information system rather than a cash transaction at the pharmacy counter.

**Returns & Corrections**

- The **Return** flow pulls an *unpaid* transaction back into the POS cart (identity verified), lets the cashier correct items/quantities/prices, then re-prices the cashier record — a real-world "cancellation/correction" workflow that blocks edits once paid/claimed.

**Inventory & Replenishment**

- **Stocks browser**: load all, load by batch, search product/category/classification (live), and a combined cross-filter across batch + classification + category.
- **Low-stock alerting**: rows are painted red as soon as current stock falls to **25% or less of the original batch quantity** — an at-a-glance reorder signal.
- **Stock ordering**: search the medicine catalog, build an order list, and submit a formal replenishment **request to the main pharmacy** (`PharmacyRequests`), with the latest request status shown on screen.

**Reporting & Export**

- Daily and Weekly/Monthly/Yearly sales reports with **All / In-Patient / Out-Patient** filters.
- **Daily/WMY Summary reports** aggregate per-medicine (summed quantities and totals) for fast-moving-item analysis.
- Running grand total displayed live; queries execute on background threads so the UI never freezes.
- **One-click Excel export** of any report to `.xlsx` in the user's Documents folder.

**Receipt Engine**

- A built-in thermal receipt layout engine (targeted at an EPSON TM-T88V class printer) with the hospital's letterhead, department, transaction number, item lines, and total — wired and ready for activation.

### Dedicated Data Scope

`PharmacyStocks`, `PharmacyStockOutTransactions`, `Temporary_PharmacyStockOutTransactions`, `PharmacyCashierTransactions`, `PharmacyRequests`, `PharmacyRequestsItems`, `Medicines*`, `MedicineCategory`, `MedicineClassification`, `MedicinesStockBatch*`, `patient`, `patient_admission`, `Rooms`, `patient_medicine`, `Users*`.

---

## Sub-System 3: ER Pharmacy POS

**Role:** The Emergency Room's dedicated pharmacy front-end with its own independent stock, its own cash/change workflow, and a requisition link to the main pharmacy.

### Feature Breakdown

**Login & Session Security (shared CBIS baseline)**

- SHA-512 login verification, session creation, online tracking, and full login/logout audit logging.

**Live Product Search**

- As-you-type search by brand or generic name against the ER's own stock; only lots with positive stock are shown, each with lot number, names, category, classification, unit cost, stock level, and expiration date.

**Cart & Validation**

- Selling-price floor guard (price ≥ unit cost), stock availability checks, positive-and-available quantity validation, item decrease/remove/clear, and a live purchase total.
- Cash amount field accepts numeric input with a single decimal point.

**Out-Patient Cash POS**

- Transaction numbering prefixed with the fiscal year (e.g., `20261234`).
- Cash tendered vs. total calculation with **change computation**.
- Sale header written to `ERPharmacySales`; each item written to `ERPharmacyStockOutTransactions` (with `who-pays` tracking); the correct batch's stock is decremented.
- **Dedicated Change dialog** displays the transaction number and computed change, then resets the screen for the next customer.
- **Thermal receipt printing** with the hospital letterhead, cashier identity, itemized lines, and grand total (suppressed print dialog for fast turnaround).

**In-Patient Charge-to-Bill**

- Validates the patient and their active admission, displays the patient's full name and room from the hospital system, and posts every ER medicine to `patient_medicine` **against the admission bill** — no cash changes hands at the point of care.
- Also records an `IN`-type stock-out entry and decrements ER stock.

**Stock Requisition (Emergency Room → Main Pharmacy)**

- Search the medicine catalog, stage an order list, and submit a formal requisition (`ERPharmacyRequests`) with status `new requests` for the main pharmacy to fulfill — a satellite-warehouse replenishment workflow.
- The form shows the status of the user's latest request on screen.

**Stock Browser**

- Full batch/category/classification browsing with live searches and a combined cross-filter.
- Same **low-stock 25% reorder-point alert** (rows turn red) used across the platform.

**Reporting & Export**

- Daily and Weekly/Monthly/Yearly itemized reports plus **per-medicine summary reports**, filtered by In-Patient / Out-Patient / All.
- Running totals and **one-click Excel export** to `.xlsx`.

### Dedicated Data Scope

`ERPharmacyStocks`, `ERPharmacySales`, `ERPharmacyStockOutTransactions`, `ERPharmacyRequests`, `ERPharmacyRequestsItems`, plus the shared catalog (`Medicines*`, `MedicinesStockBatch*`) and patient/admission (`patient*`, `Rooms`, `patient_medicine`), `Users*`.

---

## Sub-System 4: InputStocks

**Role:** The pharmacy **stock intake engine**. It bulk-loads the entire medicine catalog and supplier stock batches into the central database by reading drug distributors' Microsoft Excel workbooks — the entry point of the whole inventory chain.

### Feature Breakdown

**Bulk Excel Ingestion (four data types)**

- **Medicine Classifications** — imports classifications, skipping any that already exist.
- **Medicine Categories** — imports categories, with the same duplicate protection.
- **Medicines** — imports brand name, generic name, unit cost, category, and classification. **Foreign-key safe:** a medicine is only inserted when its category and classification already exist in the database, and identical brand+generic duplicates are skipped automatically.
- **Suppliers** — imports company name, contact name, phone, email, and address.

**Smart Batch Management**

- **Automatic batch numbering** in `YYYY-n` format (e.g., `2017-5`) prefilled as the next lot number for the selected supplier.
- Live supplier search and selection; one click creates the supplier-linked stock batch.

**Medicine-per-Batch Stock Entry**

- Live as-you-type search for batches (by lot number) and medicines (by brand/generic).
- Stage each batch item with **quantity, current on-hand stock, and expiration date** into an editable staging grid (Add / Remove before committing).
- **One-click commit** of the entire staged batch into `MedicinesStockBatchItems` — quantities and expiry dates then flow into the real-time stock used by both pharmacy POS subsystems.

**Performance & Usability**

- Every bulk import runs on its own **background thread**; imported rows paint into the grids live, row-by-row, with auto-scroll to the newest entry — the UI never freezes, even for large workbooks.
- Duplicate/sanity checks, apostrophe sanitization on imported names, read-only display grids, and graceful Excel-workbook save/close/quit on exit (no orphan processes).

### Dedicated Data Scope

`MedicineClassification`, `MedicineCategory`, `Medicines`, `Suppliers`, `MedicinesStockBatch`, `MedicinesStockBatchItems`.

**Typical workflow:** Import Classifications → Import Categories → Import Medicines → Import Suppliers → Create Batch → Attach Medicines to Batch (qty/current stocks/expiry) → Insert Batch Items.

---

## Shared Infrastructure

| Concern | Implementation |
|---|---|
| Password security | SHA-512 hashing (`Security.vb`) on every subsystem's login and password-change flow |
| Session & audit | Online session tables plus login/logout activity logs for admins and cashiers |
| Inventory model | Lot/batch + expiration-date tracking across master, main-pharmacy, ER, and stock-room stock |
| Reporting | Daily / Weekly / Monthly / Yearly itemized and per-medicine summary views with patient-type filters |
| Excel export | Interop-generated `.xlsx` reports saved to the user's Documents folder |
| Responsiveness | Background-thread data loading for imports, report queries, and heavy grids |

---

## Database Overview

The central `cbis_icphDB` schema groups into five domains:

| Domain | Tables |
|---|---|
| Authentication & HR | `AdminUsers`, `AdminSessionsStatus`, `AdminActivityLog`, `Users`, `Users_session_status`, `Users_activity_logs`, `Employees`, `Departments`, `Employee_Types` |
| Pharmacy catalog & inventory | `MedicineClassification`, `MedicineCategory`, `Medicines`, `Suppliers`, `MedicinesStockBatch`, `MedicinesStockBatchItems`, `PharmacyStocks`, `ERPharmacyStocks` |
| Sales & cashier | `PharmacyCashierTransactions`, `Temporary_PharmacyStockOutTransactions`, `PharmacyStockOutTransactions`, `ERPharmacySales`, `ERPharmacyStockOutTransactions`, `CashierAuditTrail`, `laboratory_transaction_out_patient` |
| Requisitions | `PharmacyRequests`, `PharmacyRequestsItems`, `ERPharmacyRequests`, `ERPharmacyRequestsItems` |
| Admissions & billing | `patient`, `patient_admission`, `Rooms`, `Room_types`, `patient_medicine`, `doctor`, `doctor_department`, `patient_doctor`, `services`, `patient_laboratory_services`, `Supplies` |

---

## Getting Started

- Each subsystem is a Visual Studio solution (.NET Framework 4.0, x86) requiring Microsoft SQL Server with the `cbis_icphDB` schema and Office Excel installed for import/export features.
- The schema DDL and sample queries are provided in this repository (`DB-Tables-Queries.sql.txt`, `MessyCBISQuery.sql.txt`).