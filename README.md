# 🚀 Key Features

    Object-Oriented Design: Uses a Formless Class to encapsulate properties and methods, making the codebase reusable and easy to maintain.

    Dynamic SQL Integration: Powered by ADODB to execute stored queries and pull live data directly into Excel.

    Automated Analytics: Automatically generates 3D Pie Charts, Pivot Tables, and statistical summaries (CountIf/Sum functions) upon data retrieval.

    Multi-Report Support:

        Disaster Module: Processes claim history, policy details, and affiliate codes.

        Population Module: Handles census data and demographic reporting.

    Smart Formatting: Includes automated cell styling, header coloring (Arial/Bold), and number formatting (Percentages).

# 📂 Project Structure

    FormlessClass.cls: The core engine. Contains logic for database connectivity, recordset handling, and Excel sheet manipulation.

    MainDisaster.bas: Execution script for processing insurance disaster reports based on a list of company IDs.

    MainPopulation.bas: Execution script for generating population and census summaries.

# 🛠️ Implementation Details
    
    How to Use

    Configure Connection: Update the .SetConnetion property in the Main modules (e.g., "CloudSettings" or "127Settings") to point to your SQL environment.

    Input Data: Ensure "Hoja2" contains the parameters (Company Code, Policy, Dates, etc.).

    Run: Execute the Main() or Population() subroutines.

    Core Class Snippet

    The engine handles data ingestion through a streamlined interface:

# 📈 Automated Reporting Output

    The engine doesn't just pull data; it visualizes it. It includes a SummaryCallCenter function that:

    Calculates distribution percentages.

    Applies xl3DPie chart types.

    Explodes chart segments for emphasis on key metrics.