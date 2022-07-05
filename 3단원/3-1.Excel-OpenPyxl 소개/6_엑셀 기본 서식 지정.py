#엑셀 서식 지정하기-기본

#1.Font(글꼴)
# import openpyxl as op
# from openpyxl.styles.fonts import Font

# wb=op.Workbook()
# ws=wb.active

# #font test1: 직접 font 설정하기
# ws["A1"].value="Font test1"
# ws["A1"].font=Font(size=20,italic=True,bold=True)   #italic:글자 기울기

# #font test2: format을 정해놓고 font 설정하기
# ws["A2"].value="Font test2"
# font_format=Font(size=12,name="굴림",color="FF000000")
# ws["A2"].font=font_format

# #Workbook(엑셀) 저장 및 객체 담기
# wb.save("test_result.xlsx")
# wb.close

#2.Border,Side(테두리)
#액셀의 셀 테두리를 설정할 수 있는 모듈
#Border의 경우 "선택 Cell에 상하좌우 어떤 부분에 테두리를 설정할 것인가?" 에 대한 모듈
#Side는 "각 테두리에 어떤 테두리 형식을 적용할 것인가"에 관한 모듈

# import openpyxl as op
# from openpyxl.styles import Border, Side
# from openpyxl.styles.colors import Color

# wb=op.Workbook()
# ws=wb.active

# ws["C3"].value="1개 Cell"

# #border test1: 위에는 실선, 아래에는 이중선 적용 예시 코드
# ws["C3"].border=Border(top=Side(border_style="thin"))
# ws["C3"].border=Border(bottom=Side(border_style="double"))

# #테두리 색상 코드 설정 예시
# ws["C3"].border=Border(top=Side(border_style="thin",color="000000"))
# #color의 속성 값 '000000': openpyxl에서 제공하는 IndexedColours값

# #객체 저장 및 닫기
# wb.save("style_result.xlsx")
# wb.close

# #3.Alignment(정렬)
# import openpyxl as op
# from openpyxl.styles import Alignment

# wb=op.Workbook()
# ws=wb.active

# #'C2'와 'C4'에 Text 입력
# ws["C2"].value="Alignment test1"
# ws["C4"].value="Alignment test2"

# #셀 너비,높이 설정하기
# ws.row_dimensions[2].height=50  #2행의 높이 50으로
# ws.row_dimensions[4].height=50 #4행의 높이 50으로
# ws.column_dimensions['C'].width=50 #C열의 너비 50으로

# #Alignment test1
# ws["C2"].alignment=Alignment(horizontal='left',vertical='center')

# #Alignment test2
# format1=Alignment(horizontal='center',vertical='center')
# ws["C4"].alignment=format1

# wb.save("result.xlsx")
# wb.close

#4.PatternFill(채우기)

# import greenlet
# import openpyxl as op
# from openpyxl.styles import PatternFill  #셀 채우기를 설정하는 모듈

# wb=op.Workbook()
# ws=wb.active

# #PatternFill test1 : green
# ws["C3"].fill=PatternFill(fill_type='solid',fgColor="00FF00")

# #PatternFill test2 : Black
# ws["C5"].fill=PatternFill(fill_type='solid',fgColor="000000")

# wb.save("result.xlsx")
# wb.close

#5.Protection(셀 숨김,보호)

# import openpyxl as op
# from openpyxl.styles import Protection

# wb=op.Workbook()
# ws=wb.active

# ws["C3"].value="Protection test1 : locked"
# ws["C5"].value="Protection test2 : hidden"

# #Protection 속성 설정하기(숨김/잠금)
# ws["C3"].protection=Protection(locked=True,hidden=True)
# ws["C5"].protection=Protection(locked=False,hidden=False)

# #액셀 파일 저장 및 객체 닫기
# wb.save("result.xlsx")
# wb.close
