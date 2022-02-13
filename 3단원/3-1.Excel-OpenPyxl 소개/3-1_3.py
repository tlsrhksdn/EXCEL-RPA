
# import openpyxl as op

# wb=op.load_workbook(r"test.xlsx")
# ws=wb.active

# #방법 1: Sheet의 Cell 속성 사용하기
# data1=ws.cell(row=1,column=2).value

# #방법 2: 엑셀 인덱스(Range) 사용하기
# data2=ws["B1"].value

# print("cell(1,2) : ",data1)
# print('Range("B1"): ',data2)

# #영역으로 된 부분의 data 출력
# rng=ws["A1:B1"]  #A1:B1 범위 저장
# print("Range(a1:b1) : ", rng)   #출력값은 튜플

# #이중 for문 이용해 위치정보 출력
# rng=ws["A1:C3"]  #A1:C3 범위 저장
# for rng_data in rng:
#     for cell_data in rng_data:
        
#     print(cell_data.value)

# #Cell에 데이터 쓰기

# import openpyxl as op

# wb=op.load_workbook(r"test.xlsx")  #Workbook 객체 생성
# ws=wb["무"]  #WorkSheet 객체 생성

# #"B1" Cell에 입력하기
# ws.cell(row=1,column=2).value = "입력테스트1"

# #"C1" Cell에 입력하기
# ws["C1"].value="입력테스트2"

# wb.save("result.xlsx")   

# #임의의 숫자 리스트를 작성하고 일렬로 세워 입력하는 코드
# datalist=[2,4,8,16,32,64,128,256]

# i=1

# for data in datalist:
#     ws.cell(row=i,column=1).value=data  #A열에 행을 바꾸면서 입력
#     i=i+1 #행값 증가
    
# wb.save("result.xlsx")

#Cell data 삭제하기

# 1.cell.value 값을 공백으로 설정하기

# import openpyxl as op

# wb=op.load_workbook("test data.xlsx")

# ws=wb['test']

# rng=ws["A1:C3"]

# for row_data in rng:
#     for data in row_data:
#         if (data.value%2)==0:
#             data.value=""
            
# wb.save("delete_result.xlsx")

#2.delete_rows, delete_cols 사용하기

#1)delete_rows 사용

# import openpyxl as op

# wb=op.load_workbook(r"delete_test.xlsx")
# ws=wb.active

# ws.delete_rows(1,2)

# wb.save("delete_result.xlsx")

#2)delete_cols 사용

# import openpyxl as op

# wb=op.load_workbook(r"delete_test.xlsx") #경로 헷갈림방지 위해 파일명 앞에 r을 붙인다
# ws=wb.active

# ws.delete_cols(2,1)

# wb.save("delete_result.xlsx")

#시트를 통째로 삭제하고 다시 생성하기

# import openpyxl as op

# wb=op.load_workbook(r"delete_test.xlsx")
# ws=wb['test']

# wb.remove(ws)
# wb.create_sheet('test')

# wb.save("delete_result.xlsx")
