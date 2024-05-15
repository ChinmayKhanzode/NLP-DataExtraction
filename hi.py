from openpyxl import Workbook

# Create a workbook and select the active worksheet
wb = Workbook()
ws = wb.active

# Data to be added
data = [['John', 28, 'USA'], ['Anna', 24, 'Sweden'], ['Peter', 22, 'Germany']]

# Add data row by row
for row in data:
    ws.append(row)

wb.save('output.xlsx')
