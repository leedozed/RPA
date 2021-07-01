from openpyxl import load_workbook
wb = load_workbook("sample.xlsx")
ws = wb.active

# ws.insert_rows(8) #8번쨰 줄이 삽입됨
# ws.insert_rows(8,5) #8부터 5 줄쨰 줄이 삽입됨

ws.insert_cols(2)
ws.insert_cols(2,3)


# wb.save("sample_insert_rows.xlsx")
wb.save("sample_insert_cols.xlsx")
