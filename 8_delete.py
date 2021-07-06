from openpyxl import load_workbook
wb = load_workbook("sample.xlsx")
ws = wb.active

# ws.delete_rows(8)
# ws.delete_rows(8, 3) #8번부터 3명의 학생 삭제

ws.delete_cols(2,2)

wb.save("sample_delete_row.xlsx")