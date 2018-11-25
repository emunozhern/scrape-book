from openpyxl import load_workbook

new_row_data = [
    ['odhgos', 'e/p', 'dromologio', 'ora'],
    ['odigosou', 'dromou', 'dromologio', 'ora']]

wb = load_workbook("Sample.xlsx")
# Select First Worksheet
ws = wb.worksheets[0]

# Append 2 new Rows - Columns A - D
for row_data in new_row_data:
    # Append Row Values
    ws.append(row_data)

wb.save("Sample.xlsx")