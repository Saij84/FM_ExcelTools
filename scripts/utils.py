from openpyxl import load_workbook

filePath = r"C:\tools\FM_ExcelTools\testData\SampleData.xlsx"

wb = load_workbook(filePath)
salesOrderSheet = wb["SalesOrders"]

print(salesOrderSheet["B2"].value)
