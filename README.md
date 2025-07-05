# 1. Introduction  
E-commerce in the Philippines has grown rapidly over the past decade, led by major platforms like Shopee, Lazada, and TikTok Shop. In 2024, the government introduced a mandatory tax on e-commerce businesses to boost revenue and regulate the fast-expanding digital market. This move signals a shift toward greater formalization and accountability in the online retail sector.  
As part of this shift, my work focuses on automating tax input processes using Excel VBA macros, helping businesses streamline compliance and reduce manual errors in tax reporting.

# 2. Description of Project  
This project is an Excel-based VBA macro tool designed for e-commerce sellers to automate their tax reporting workflows. It handles the consolidation of sales data from Shopee, Lazada, and TikTok Shop, automates sales invoice counting, matches digital sales with physical paper invoice numbers, and computes taxes on a daily, weekly, or monthly basis, depending on the business’s reporting needs. The tool helps streamline compliance and reduce manual input in high-volume transaction environments.

## Business Rules / Tax Compliance  
- Orders with individual amounts below PHP 500 within the same day are combined into a single invoice for tax reporting.  
- Orders with amounts exceeding PHP 500 require separate invoices in compliance with tax regulations.  
- This ensures proper documentation and alignment with government guidelines on invoice issuance and tax filing.

# 3. Objective  
The main objective of this project is to provide a simple yet powerful solution for automating tax-related tasks specific to e-commerce operations. It aims to:  
- Improve accuracy in tax reporting  
- Consolidate multi-platform sales into one system  
- Minimize manual encoding sales data  
- Align digital records with official paper invoice sequences  
- Support timely and structured tax filing across various reporting periods

# 4. Scope and Limitation  

## Scope  
- Import raw sales reports from Shopee, Lazada, and TikTok Shop  
- Match and tally invoice numbers with physical sales invoices and book records  
- Consolidate sales orders regardless of reporting interval (daily, weekly, or monthly)

## Limitations  
- The system does not integrate directly with e-commerce platforms or the BIR e-filing system  
- Manual input is still required to assign official invoice numbers from printed physical invoice books  
- Cannot be used to mix or process physical store invoices alongside e-commerce invoices

# 5. Methodology  
- **Data Extraction**  
  Sales reports are downloaded from each e-commerce platform — Shopee, Lazada, and TikTok Shop — and saved into a designated folder. Within this folder, subfolders are created for each platform to organize the raw data files.  
- **Data Preparation**  
  Using Power Query in Excel, the raw sales data files are imported and cleaned. This includes steps such as removing unnecessary columns, correcting data formats, filtering out irrelevant records, and consolidating multiple files into a unified dataset.  
- **Automation with VBA Macros**  
  After data preparation, the VBA macros process the cleaned dataset to consolidate orders based on business rules, such as grouping orders below PHP 500 into a single invoice and assigning separate invoices for orders exceeding PHP 500, while automatically assigning invoice numbers for both consolidated and separate invoices.

# 6. Snippet  

# 7. How to Use  
**Step 1: Download Sales Reports**  
- Log in to Shopee, Lazada, and TikTok Shop seller centers.  
- Download your sales reports for the relevant period (daily, weekly, or monthly).  
- Save each report in its corresponding platform subfolder within the main project folder.

**Step 2: Open the Excel Workbook**  
- Launch the Excel file containing the VBA macros.  
- Make sure macros are enabled when prompted.

**Step 3: Refresh**  
- Click refresh all button on Data tab.

**Step 4: Run the VBA Macros**  
- Click the “update” button to run the macros.  
- The tool will:  
  - Consolidate orders based on business rules (group orders below PHP 500 into single invoices and separate those above PHP 500).  
  - Automatically assign invoice numbers for consolidated and separate invoices.

**Step 5: Save Your Work**  
- Save the workbook with updated sales invoice records for your documentation or submission.

# 8. Summary of the VBA Macro  
1. **Worksheet Setup**  
The macro begins by identifying and linking the worksheets for Lazada, Shopee, TikTok, and the summary sheet (CashReceipt). It clears any old data in the summary to prepare for new entries.

2. **Data Extraction and Consolidation**  
Each platform’s sales report is processed differently based on its layout:  
- Lazada:  
  Uses a Do While loop to scan through blocks of rows that represent each transaction. It checks for a specific label to identify transaction start points and calculates totals from related rows.  
- Shopee & TikTok:  
  Use For loops to process each row one by one. The macro extracts transaction details like date, price, charges, and net income.  
All processed records are added to the summary sheet (CashReceipt), along with the source platform.

3. **Sorting by Date**  
After consolidation, all data is sorted chronologically by transaction date. This prepares the dataset for correct invoice grouping in the next step.

4. **Assigning Invoice Numbers**  
The macro groups’ transactions by date and assigns Sales Invoice (SI) numbers:  
- Transactions over PHP 500 receive individual SI numbers.  
- Transactions PHP 500 and below (per order) are grouped under a single SI number for that day.  
This ensures compliance with invoicing rules and allows easy tracking.

5. **Sorting by Invoice Number**  
To organize the data further, the summary is sorted again this time by SI number, so grouped orders are visually together.

6. **Formatting**  
The final part of the macro sets up the structure to apply alternating row colors and cell borders for readability. This part appears to be in progress.

### Techniques Used:  
- Do While loops for dynamic row processing  
- For loops for row-by-row extraction  
- Conditional logic for tax rules  
- Excel sort functions for organization  
- Structured grouping and numbering of invoices
  
---

## 🔐 Security Note
To protect the macro logic:
- The Excel file is VBA password-protected

---

## 📂 Files Included
- `TAX RECORD.xlsm` — Excel file with Power Query and VBA macros
- `README.md` — Project documentation

---

## 📞 Contact  
For questions or requests:  
**[Engr. Alfreime Alloye Lazarte]**  
📧 lazartealfreimealloye.com  
📍 Philippines

---
