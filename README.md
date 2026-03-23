# McrJb: Automated Insurance Analytics Engine

This project is a VBA-based automation framework designed for the insurance sector. It acts as a bridge between SQL Server databases and Excel, automating the ETL (Extract, Transform, Load) process to generate professional reports containing Pivot Tables, 3D Charts, and financial summaries.

## 🚀 Key Features

*   **Object-Oriented VBA**: Uses a custom class (`FormlessClass`) to encapsulate database logic, making the main execution scripts clean and readable.
*   **Dynamic SQL Injection**: Reads raw SQL queries from external text files and injects parameters (Dates, Policy IDs, Company Codes) at runtime.
*   **Automated Visualization**: Programmatically creates:
    *   Pivot Tables (e.g., Loss Ratios, Authorizations).
    *   3D Pie Charts (e.g., Call Center utilization).
*   **Professional Formatting**: Automatically applies corporate color schemes (RGB 72, 61, 139), fonts (Arial Bold), and number formatting.

## 📂 Project Structure

### Core Files
*   **`FormlessClass.cls`**: The engine of the project. It handles ADODB connections, recordset processing, and Excel DOM manipulation.
*   **`MainDisaster.bas`**: The execution script that iterates through a list of companies to generate "Disaster" (Claims/Siniestralidad) reports.
*   **`Querys/` Directory**: Contains `.txt` files with SQL templates and connection strings.

### Required Directory Layout
For the tool to work, your folder must look like this:
```text
/ProjectFolder
    ├── AutomationTool.xlsm  (Your Excel File)
    ├── disasters/           (Empty folder for report output)
    └── Querys/              (Folder containing text files)
        ├── 127Settings.txt  (Connection string)
        ├── CloudSettings.txt (Connection string)
        ├── DisasterFormless.txt
        ├── AuthorizationsFormless.txt
        └── CallCenterFormless.txt
```

## 🛠️ Installation & Setup (How to add to Excel)

Follow these steps to set up this project inside a fresh Excel workbook.

### 1. Prepare the Environment
1.  Open **Microsoft Excel**.
2.  Save the file as an **Excel Macro-Enabled Workbook (.xlsm)**.
3.  Enable the **Developer Tab** (File > Options > Customize Ribbon > Check "Developer").

### 2. Import the VBA Code
1.  Press `Alt + F11` to open the **Visual Basic Editor (VBE)**.
2.  **Add the Class Module**:
    *   Go to `Insert` > `Class Module`.
    *   In the Properties window (press `F4` if not visible), rename the class from `Class1` to `FormlessClass` (or `Formless` depending on how it's called in Main). *Note: The provided Main script refers to it as `Formless`, but the file is `FormlessClass.cls`. Ensure the Name property matches `Formless`.*
    *   Paste the code from `FormlessClass.cls`.
3.  **Add the Main Module**:
    *   Go to `Insert` > `Module`.
    *   Rename it to `MainDisaster`.
    *   Paste the code from `MainDisaster.bas`.

### 3. Add External References
This project uses SQL database connections, so you must enable the ADO library.
1.  In the VBE, go to `Tools` > `References`.
2.  Scroll down and check **Microsoft ActiveX Data Objects 6.1 Library** (or 2.8 if 6.1 is unavailable).
3.  Click **OK**.

### 4. Configure the Input Sheet
The `Main()` subroutine reads configuration from **"Hoja2"**. Set up your sheet headers as follows:

| Cell | Header (Concept) | Description |
| :--- | :--- | :--- |
| **A** | `CodEmpresa` | Database ID for the Company |
| **B** | `CodPymeColectivo` | Sub-group ID (0 if none) |
| **C** | `StartDate` | Report Start Date (YYYY/MM/DD) |
| **D** | `EndDate` | Report End Date (YYYY/MM/DD) |
| **E** | `ReportName` | Name of the output file |
| **F** | `CodAfiliado` | Affiliate ID (Optional) |
| **G** | `Policy` | Policy Number (Optional) |

*Populate row 2 onwards with the data you want to process.*

### 5. Setup Connection Strings
Create text files in the `Querys/` folder named `127Settings.txt` (or whatever name is passed to `.SetConnection`) containing your ADODB connection string.
*Example content:*
```text
Provider=SQLOLEDB;Data Source=YOUR_SERVER_IP;Initial Catalog=YOUR_DB;User ID=YOUR_USER;Password=YOUR_PASSWORD;
```

## 💻 Usage

1.  Ensure your `Querys` folder contains the necessary SQL text files.
2.  Ensure your `disasters` folder exists in the same path as the Excel file.
3.  Fill out "Hoja2" with the parameters.
4.  Run the macro:
    *   Press `Alt + F8` in Excel.
    *   Select `Main`.
    *   Click **Run**.

## 📈 Implementation Details

### The Formless Engine
The code uses a fluent-interface style pattern. Instead of passing 10 arguments to a function, properties are set sequentially:

```vba
Dim F As Formless
Set F = New Formless

F.SetName = "Company X"
F.SetStartDate = #1/1/2023#
F.SetDisasterQuery = "DisasterQueryFile" ' Loads SQL from txt
F.SetRecordSet = F.Query()             ' Executes SQL
F.RecordToSheet("Sheet1") = 1          ' Dumps data
```

### Reporting Logic
*   **Disaster Module**: Analyzes claims (`Siniestralidad`) and authorizations.
*   **Census Module**: Handles population demographics.
*   **Call Center**: visualizes call reasons using exploded 3D pie charts.