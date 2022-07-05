# import openpyxl as op

# wb=op.Workbook()

# print(wb.sheetnames)


# #.sheetnames 함수: 해당 Workbook의 시트명을 리스트로 출력해주는 함수

# #시트 생성/삭제하기

# #.create_sheet 함수를 사용해 Workbook 객체에 시트를 새로 생성한다

# import openpyxl as op

# wb=op.Workbook()

# ws=wb.create_sheet("연습")
# print(wb.sheetnames)

# #첫 번째 위치에 생성
# ws=wb.create_sheet("첫번째",0)
# print("첫번째 : ",wb.sheetnames)

# #마지막에서 두번째 위치에 생성
# ws=wb.create_sheet("뒤에서 두번째",-1)
# print("뒤에서 두번째 : ",wb.sheetnames)

# #두번째 위치에 생성
# ws=wb.create_sheet("두번째",1)
# print("두번째 : ",wb.sheetnames)

# #세번째 위치에 생성
# ws=wb.create_sheet('세번째',2)
# print("세번째 : ",wb.sheetnames)

# #create_sheet 내부의 숫자 입력값에 따라 시트의 위치가 다르게 생성되는 것을 확인할 수 있다.

# #시트삭제

# #Workbook 객체의 remove 함수를 사용해 시트를 삭제한다

# #원본 시트 리스트 출력
# print("원본 : ",wb.sheetnames)

# #시트 삭제: 삭제시 괄호 안의 내용은 Worksheet 객체이다
# # ws=wb['First']
# # wb.remove(ws)

# #삭제 후 시트 리스트 출력
# print("삭제 후: ",wb.sheetnames)

# #시트 이름 변경하기

# #WorkSheet 객체의 title 속성을 사용해 시트 이름을 변경한다

# #'첫번째'라는 이름을 가진 시트를 'First'로 변경
# ws1=wb["첫번째"]
# ws1.title="First"
# print(wb.sheetnames)

# #'두번째'라는 이름을 가진 시트를 'Second'로 변경
# ws2=wb["두번째"]
# ws2.title="Second"
# print(wb.sheetnames)

# #시트 이동/복사

# #.move_sheet 함수와 copy_worksheet 함수를 사용해 시트 위치를 이동하고 복사한다

# #시트 이동

# #move_sheet는 현재 위치에서 상대적인 위치 값을 입력하여 시트를 이동시킨다

# #원본 출력
# print("원본 : ",wb.sheetnames)

# #Worksheet 객체 설정(시트명: Second Sheet)
# ws=wb['Second sheet']

# #'Second sheet'를 기준점에서 앞으로 5칸 이동
# wb.move_sheet(ws,-5)

# #위치 이동 후 출력
# print("이동 후 : ",wb.sheetnames)

# #시트 복사

# #copy_worksheet를 사용해 시트를 복사할 수 있다

# #'Second' 시트를 Worksheet 객체로 설정
# sht=wb["Second"]
# #'Second' 시트를 복사하고 복사한 시트를 ws_copy 객체로 설정
# ws_copy=wb.copy_worksheet(sht)

# #출력해보기
# print("복사한 시트명 : ",ws_copy)
# print("시트리스트 : ",wb.sheetnames)

# #시트 탭 속성 설정

# #WorkSheet 객체의 sheet_properties 속성을 활용하면 탭의 색을 변경할 수 있다

# #시트 탭 색깔 변경(하늘색)
# ws_copy.sheet_properties.tabColor='00FFFF'

# #다시 저장하여 확인
# wb.save(r"after.xlsx")