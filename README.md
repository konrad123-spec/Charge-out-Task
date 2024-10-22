# Chargeout Task IT - SAP Data Extraction and Validation

This project involves a robust macro that connects directly to SAP to automate data extraction, processing, and validation for Chargeout-related tasks. The macro significantly reduces manual effort by automating multiple steps in the data retrieval process for different divisions and Product Groups (PGs).

## Macro Overview

The primary macro extracts essential data for various divisions, focusing on:

- **Chargeout COS**
- **Chargeout GA**
- **Chargeout Sales**
- **Data Specific to PGs**

The macro operates within the Profit and Loss (P&L) report, extracting and validating the required data automatically.

## Key Features

### 1. **Data Retrieval and Manipulation**
- The macro identifies relevant **Division/PG nodes** (SAP coordinates) based on data stored in the "Macro" tab of the Master File.
- It searches for specified values in the **Chargeouts** section of the P&L report.
- If the values exist, the macro:
  - Accesses these accounts in SAP.
  - Creates an export of the data.
  - Identifies and removes any values that reset to zero.
  - Extracts values with their associated currency and pastes them into the **"datacharge"** tab, organized by division, PG, and chargeout.

### 2. **Counterparty Identification**
- A second macro focuses on identifying the **abbacli** (counterparty) for each transaction pasted in the "datacharge" sheet.
- In some cases, SAP does not provide the counterparty information, which requires manual verification.

### 3. **Sum Verification**
After completing data extraction and counterparty identification, two smaller macros perform the following operations:

1. **Division-wise Sum Calculation**:
   - The macro calculates the sums from the **"datacharge"** sheet, aggregating data by division.
   - These sums are then placed in a summary table within the **"Macro"** sheet.

2. **CFIN System Sum Calculation**:
   - The second macro calculates sums based on the **CFIN system** data.
   - A third table compares the differences between these two sets of sums, and conditional formatting highlights any discrepancies.

## Process Flow

1. **Data Extraction**: Extracts relevant chargeout data from SAP based on Division/PG coordinates.
2. **Counterparty Identification**: Attempts to locate and assign abbacli (counterparties) to each transaction.
3. **Sum Comparison**: Compares sums from both "datacharge" and CFIN systems, highlighting discrepancies.

---
