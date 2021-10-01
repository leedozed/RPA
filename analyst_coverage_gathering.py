from openpyxl import load_workbook
wb = load_workbook("These2_Company Coverage_210928__15.xlsx")

ws = wb["1"]

# for x in range(11485, 18426):
for x in range(11485, 2):
    for y in range(1,3):
        print(ws.cell(row=x, column=y).value, end=" ")
    print()

# wb = Workbook() # 새 워크북 생성
# ws = wb.active #현재 활성화된 sheet 가져옴
# ws.title = "Coverage" #Sheet의 이름을 변경
# wb.save("th2_analyst_coverage.xlsx")
# wb.close()


# new_ws = wb["NewSheet"] #Dict 형태로 Sheet 접근이 가능함
# print(wb.sheetnames) #모든 시트 이름 확인

# #sheet 복사
# new_ws["A1"] = "Test"
# target = wb.copy_worksheet(new_ws)
# target.title = "Copied Sheet"

# ws["A1"] = 1 #1값을 입력함

# print(ws["A1"])
# print(ws["A1"].value) #A1셀의 값을 출력함

# ws.cell(row = 1, column = 1).value #A1셀
# print(ws.cell(row = 1, column = 1).value ) #ws["A1"].value

# from random import *

# #반목문으로 랜덤 숫자 채우기

# index = 1
# for x in range(1, 11):
#     for y in range(1,11):
#         # ws.cell(row = x, column = y, value = randint(0,100))
#         ws.cell(row = x, column = y, value = index)
#         index += 1

