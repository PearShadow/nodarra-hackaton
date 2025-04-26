Link to UPWork task: https://www.upwork.com/jobs/~021915973742727480014?referrer_url_path=%2Fnx%2Fsearch%2Fjobs%2Fdetails%2F~021915973742727480014

ID: 021915973742727480014

📄 Breakdown of the Script

Inputs Tab
Assumptions are dynamically placed into the Inputs tab.

Assumptions include:
-Capacity
-CapEx
-PPA Price
-OpEx Rate
-Loan Amount
-Loan Term
-Interest Rate

Financial Model Tab
Automatically calculates:
-Revenue
-OpEx
-EBITDA
-Debt Service
-Cash Flows
-Uses dynamic Excel formulas for all fields.
-Debt Service is calculated based on the loan amount, interest rate, and loan term.

Summary Tab
-Calculates IRR, DSCR, and NPV for the first 5 years.
-These calculations use built-in Excel formulas (you can adjust the formula ranges if needed).

Formatting
-Headers are bolded for clarity.
-Currency format is applied to all monetary values.
-Borders are added to give the Excel model a professional look.

Inputs are visually distinguished from outputs using different cell styles.

🚀 How to Test the Script

pip install openpyxl

python generate_financial_model.py
