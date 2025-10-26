Press Production Sheet Analyst

This is a Streamlit web application designed to analyze Excel-based "PRESS PROD SHEET" files from an aluminum extrusion facility.

It automatically parses the complex, multi-row header format, applies a sophisticated, press-specific ruleset to identify underperforming dies, and generates a downloadable "Flagged Dies Report".

Features

Smart Sheet Detection: Automatically finds the "PRESS PROD SHEET" and the "Sheet1" (or "Mapping") sheet, ignoring case and spacing.

Complex Header Parsing: Correctly reads data from sheets with multi-row headers and merged cells.

Header Data Extraction: Pulls key single-value data (Date, Press, Operator, Supervisor) and displays them in metric cards.

Advanced Rule Engine:

Maps "DIE NAME" (e.g., "R.T 75x12x1.7 mm") to a standardized "Die Family" (e.g., "Rectangular Tube").

Applies different performance rules based on the "Press" value (P1 or P2).

Flags dies for up to three metrics: PROD/HOUR, RECOVERY %, and Speed(mm).

Provides clear, combined reasons for flagging (e.g., "Production rate below 220 Prod/hour and Recovery below 80%").

Department Mapping: Uses the "Sheet1" / "Mapping" tab to map production "Remarks" to their corresponding department (e.g., "Tool Room", "Production Department").

Report Generation: Generates a clean, filterable data table of only the flagged dies, including all relevant context.

Excel Export: Allows the user to download the final "Flagged Dies Report" as a .xlsx file for sharing and further analysis.

How to Run

Clone the repository:

git clone <your-repository-url>
cd <your-repository-directory>


Install dependencies:
It's recommended to use a virtual environment.

python -m venv venv
source venv/bin/activate  # On Windows, use `venv\Scripts\activate`
pip install -r requirements.txt


Run the app:

streamlit run press_prod_viewer.py


Use the app:
Open the local URL (e.g., http://localhost:8501) in your web browser and upload your Excel production file.
