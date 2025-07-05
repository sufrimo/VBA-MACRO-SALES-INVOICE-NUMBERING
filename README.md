# 📊 E-Commerce Tax Automation Tool (Excel VBA)

## 1. Introduction  
E-commerce in the Philippines has grown rapidly over the past decade, led by major platforms like Shopee, Lazada, and TikTok Shop. In 2024, the government introduced a mandatory tax on e-commerce businesses to boost revenue and regulate the fast-expanding digital market. This project supports that transition by helping sellers streamline their compliance processes.

This Excel-based tool uses **VBA macros** to automate the consolidation of sales data and generation of sales invoices, reducing manual effort and minimizing errors in tax reporting.

---

## 2. Description of Project  
This project is an Excel macro automation tool tailored for Philippine e-commerce sellers. It consolidates raw sales data from Shopee, Lazada, and TikTok Shop, matches these with official invoice numbers, and helps automate tax calculations.

### 🔒 Business Rules / Tax Compliance
- Orders **below PHP 500** (per order, per day) are grouped into a **single invoice**.
- Orders **exceeding PHP 500** require **separate invoices**.
- Aligns with BIR tax rules and ensures proper invoice issuance and record-keeping.

---

## 3. Objective  
The main goal of this tool is to simplify tax-related workflows for online sellers by:
- Improving accuracy in tax reporting  
- Consolidating multi-platform sales  
- Reducing manual data entry  
- Matching with official printed invoice numbers  
- Supporting daily, weekly, or monthly reporting

---

## 4. Scope and Limitation

### ✅ Scope:
- Import sales reports from **Shopee, Lazada, and TikTok Shop**
- Match digital sales with **printed invoice books**
- Consolidate orders across different timeframes

### ⚠️ Limitations:
- No direct integration with e-commerce APIs or BIR e-filing
- Manual input still needed for assigning physical invoice numbers
- Not designed for brick-and-mortar store invoices

---

## 5. Methodology

### 📥 Data Extraction
- Sellers download CSV or Excel sales reports from their platform dashboards
- Files are saved in a structured folder with subfolders by platform

### 🧹 Data Preparation
- Power Query is used to clean and combine data:
  - Remove unnecessary columns
  - Normalize formats
  - Filter out non-sales data

### 🤖 Automation with VBA
- VBA macros consolidate orders based on business rules
- Orders below PHP 500 are grouped under a single invoice
- Orders above PHP 500 are assigned individual invoices
- Invoice numbers are automatically assigned

---

## 6. How to Use

### Step 1: Download Sales Reports
- Get sales files from Shopee, Lazada, and TikTok seller dashboards
- Save them into platform-specific subfolders

### Step 2: Open the Excel File
- Launch the workbook and enable macros

### Step 3: Refresh Data
- Click **Refresh All** from the Excel **Data** tab

### Step 4: Run Macros
- Click the **"Update"** button (or run the macro manually)
- The macro will:
  - Consolidate and group transactions
  - Assign sales invoice numbers automatically

### Step 5: Save Your Work
- Save the updated workbook for tax documentation or BIR submission

---

## 7. Summary of VBA Logic

### 🧠 Core Logic
1. **Setup** – Assign worksheet references and clear old data
2. **Import** – Extract and process each platform’s report differently
3. **Sort** – First by date, then by invoice number
4. **Invoice Assignment** – Apply business rules for grouping and numbering
5. **Format Output** – (In progress) Apply borders and alternating colors for readability

### 🔧 Techniques Used:
- `Do While` loops for dynamic multi-row logic (Lazada)
- `For` loops for sequential data (Shopee and TikTok)
- Conditional logic for tax grouping
- Excel's built-in sorting methods
- Automated invoice generation and grouping logic

---

## 7. Snippet


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
📧 lazartealfreimealloye@yahoo.com.
📍 Philippines

---

