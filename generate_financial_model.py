import openpyxl
from openpyxl.styles import Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter

# Create workbook and add tabs
wb = openpyxl.Workbook()

# Add the 'Inputs' tab
inputs_sheet = wb.active
inputs_sheet.title = "Inputs"

# Add assumptions in the 'Inputs' tab
inputs_data = [
    ("Capacity (MW)", 100),  # Example: 100 MW capacity
    ("CapEx ($)", 5000000),  # Example: $5M CapEx
    ("PPA Price ($/MWh)", 50),  # Example: $50 PPA Price
    ("OpEx Rate ($/MW/year)", 20000)  # Example: $20K OpEx Rate
]

for row_num, (label, value) in enumerate(inputs_data, 1):
    inputs_sheet[f"A{row_num}"] = label
    inputs_sheet[f"B{row_num}"] = value
    inputs_sheet[f"B{row_num}"].number_format = '#,##0.00'  # Currency format

# Add the 'Financial Model' tab
financial_model_sheet = wb.create_sheet("Financial Model")

# Headers for the financial model
headers = ["Year", "Revenue ($)", "OpEx ($)", "EBITDA ($)", "Debt Service ($)", "Cash Flows ($)"]

for col_num, header in enumerate(headers, 1):
    financial_model_sheet.cell(row=1, column=col_num, value=header)
    financial_model_sheet.cell(row=1, column=col_num).font = Font(bold=True)

# Calculate financial data
capacity = inputs_sheet["B1"].value  # MW
ppa_price = inputs_sheet["B3"].value  # $/MWh
opex_rate = inputs_sheet["B4"].value  # $/MW/year

for year in range(1, 11):  # 10 years of data
    financial_model_sheet[f"A{year+1}"] = year
    financial_model_sheet[f"B{year+1}"] = f"= {capacity} * 8760 * {ppa_price}"  # Revenue calculation
    financial_model_sheet[f"C{year+1}"] = f"= {capacity} * 1000 * {opex_rate}"  # OpEx calculation
    financial_model_sheet[f"D{year+1}"] = f"= B{year+1} - C{year+1}"  # EBITDA calculation
    financial_model_sheet[f"E{year+1}"] = "Fixed"  # Placeholder for Debt Service (to be defined)
    financial_model_sheet[f"F{year+1}"] = f"= D{year+1} - E{year+1}"  # Cash Flow calculation

# Add the 'Summary' tab
summary_sheet = wb.create_sheet("Summary")

# IRR, DSCR, and NPV headers
summary_headers = ["Year", "IRR", "DSCR", "NPV"]
for col_num, header in enumerate(summary_headers, 1):
    summary_sheet.cell(row=1, column=col_num, value=header)
    summary_sheet.cell(row=1, column=col_num).font = Font(bold=True)

# Calculate IRR, DSCR, and NPV (example formulas)
summary_sheet["A2"] = "Year 5"
summary_sheet["B2"] = "=IRR(Financial_Model!F2:F6)"  # Example: IRR based on Cash Flows for years 1-5
summary_sheet["C2"] = "=DSCR(Financial_Model!F2:F6, Financial_Model!E2:E6)"  # Example: Debt Service Coverage Ratio
summary_sheet["D2"] = "=NPV(0.1, Financial_Model!F2:F6)"  # Example: NPV at 10% discount rate

# Apply formatting: Currency, Borders, Alignment
currency_format = '#,##0.00'
thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

# Format 'Inputs' tab
for row in inputs_sheet.iter_rows(min_row=1, max_row=len(inputs_data), min_col=1, max_col=2):
    for cell in row:
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center", vertical="center")
        if cell.column == 2:  # Apply currency format to values
            cell.number_format = currency_format

# Format 'Financial Model' tab
for row in financial_model_sheet.iter_rows(min_row=1, max_row=11, min_col=1, max_col=6):
    for cell in row:
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center", vertical="center")
        if cell.column in [2, 3, 4, 5, 6]:  # Currency columns
            cell.number_format = currency_format

# Format 'Summary' tab
for row in summary_sheet.iter_rows(min_row=1, max_row=2, min_col=1, max_col=4):
    for cell in row:
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center", vertical="center")
        if cell.column in [2, 3, 4]:  # Currency columns
            cell.number_format = currency_format

# Save the workbook
wb.save("financial_model.xlsx")

print("Excel model generated successfully!")
