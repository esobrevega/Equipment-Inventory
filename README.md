# Equipment-Inventory
# STMC RT Equipment Inventory System

## Overview
The **STMC RT Equipment Inventory System** is an Excel-based tracking system enhanced with **VBA macros** to facilitate equipment logging, monitoring, and searching. The system helps users efficiently manage medical equipment data, search for records, and generate reports.

## File Structure
### **Excel Sheets**
- **STMC EQUIPMENT INPUT** – Allows users to input new equipment data.
- **EQUIPMENT MONITORING DATA** – Stores and tracks recorded equipment details.
- **STMC EQUIPMENT SEARCH** – Provides search functionality for equipment.
- **LOCATION** – Maintains location data for different equipment.
- **EQUIPMENT INPUT** – Alternative data entry sheet.
- **FORMULA** – Contains key formulas for calculations.
- **FORM SEARCH** – Supports form-based searching.
- **DROPDOWN** – Holds dropdown values for easier selection.
- **RAW DATA** – Stores unprocessed equipment data.

### **VBA Modules**
- **Module1** – Handles form logging and data transfer.
- **Module2** – Manages equipment log recording and clearing updates.
- **Module3** – Implements advanced search functions.
- **Module4** – Contains a test macro for debugging.
- **Module5** – Includes a deletion function for test purposes.
- **Module6** – Implements formula applications and additional search features.

## Key Features
### **1️⃣ Equipment Logging (Module1 & Module2)**
- Transfers form data to the **RAW DATA** and **EQUIPMENT MONITORING DATA** sheets.
- Sorts data by **ID Number** and ensures records are stored correctly.
- Clears input fields after successful data entry.

### **2️⃣ Advanced Search Functions (Module3)**
- Search by:
  - Reference Number
  - Date Logged
  - Serial Number
  - Equipment Type
  - Brand
  - Location
  - PMM Date & Status
  - Description
  - Maintenance History
- Uses **Excel FILTER functions** to dynamically display relevant results.

### **3️⃣ Equipment Maintenance Tracking (Module2 & Module6)**
- Uses **XLOOKUP** to retrieve equipment details based on Reference Number.
- Tracks PMM (Preventive Maintenance Management) dates and statuses.
- Displays maintenance history for selected equipment.

### **4️⃣ PDF Report Generation (Module6 & Sheets 4, 6, 7)**
- Generates **monthly reports** as PDFs.
- Exports search results into formatted **PDF reports**.
- Saves files to a predefined directory.

## VBA Code Analysis
### **Module1: Logging Forms**
- **`LogForm` & `LogForm2`**
  - Copies inputted form data to the **RAW DATA** sheet.
  - Pastes data as **values only** to avoid formula errors.
  - Clears input fields after saving.
  - Ensures data is **sorted properly** in the main database.

### **Module2: Equipment Log Recording & Clearing Updates**
- **`RecordLog`**
  - Transfers new equipment records to **EQUIPMENT MONITORING DATA**.
  - Sorts entries by **Reference Number**.
  - Clears form fields after entry.
  - Displays a **success message** upon completion.

- **`CLEAR_UPDATE`**
  - Uses **XLOOKUP** to autofill data based on Reference Number.
  - Pulls information like **Serial Number, Equipment Type, Brand, Location, PMM Status**.

### **Module3: Advanced Search Functions**
- **`SearchRef`**, **`SearchAll`**, and other search macros:
  - Use `FILTER` functions to dynamically find data.
  - Allow searching by **Reference Number, Date Logged, Serial Number, Brand, Location, and more**.
  - Update the search interface dynamically.

### **Module4 & Module5: Testing & Debugging**
- Contain test macros for deleting rows and debugging search queries.
- Used for **development purposes** and improving accuracy.

### **Module6: PDF & Report Generation**
- **`toPDF` & `PDFMonthlyReport`**
  - Convert selected data into a **PDF report**.
  - Use **ExportAsFixedFormat** to create professional reports.
  - Automatically name and save reports to a directory.

## How to Use the System
1. **Adding Equipment:**
   - Enter details in **STMC EQUIPMENT INPUT**.
   - Click the **Log Equipment** button.
   - Data will be stored and sorted automatically.

2. **Searching for Equipment:**
   - Go to **STMC EQUIPMENT SEARCH**.
   - Use search filters (Reference Number, Brand, Location, etc.).
   - Results will update dynamically.

3. **Generating Reports:**
   - Click the **Generate PDF Report** button.
   - Select the report type (Monthly, Current Equipment, etc.).
   - The file will be saved automatically.

4. **Updating or Deleting Equipment:**
   - Enter the **Reference ID** in **STMC EQUIPMENT INPUT**.
   - Click **Update Equipment** to modify existing records.
   - Click **Delete Equipment** to remove an entry.

## License
This project is licensed under the **MIT License**. See the [LICENSE](LICENSE) file for details.

