#VBA란?

#Visual Basic for Application의 약자로, Microsoft사에서 제공하는 Microsoft 응용 프로그램을 위한 프로그래밍 언어이다

#엑셀 VBA

#Excel에서 사용자가 원하는 기능을 프로그래밍 언어를 통해 개발할 수 있는 도구이다

#파이썬과 VBA의 연동

#파이썬의 xlwings 라이브러리를 사용하면 엑셀의 VBA와 연동이 가능해진다

#엑셀에서 파이썬 코드를 불러와서 사용할 수 있고, 반대로 파이썬에서 엑셀 VBA를 실행시키는 것도 가능하다

#엑셀 VBA 사용하기

#매크로 파일 생성하기

#엑셀에는 여러 확장자 파일이 있다
#VBA를 동작시키기 위해서는 .xlsm 확장자 파일로 생성한다

#VBA 개발 도구 실행해보기

#메뉴 선택 바 중 개발 도구 - Visual Basic을 선택 or 단축키(Alt+F11)을 누른다

#모듈 추가

#엑셀 VBA에는 모듈이라는 개념이 있다
#모듈은 VBA에서 프로젝트를 구성하는 기본 단위이다
#프로시저(특정 기능을 실행하기 위한 코드 집합)의 집합

#프로시저 1개를 정의하고 메시지박스를 띄우는 VBA코드를 작성해본다

#엑셀 VBA 환경에서 코드 작성했다고 가정
from ast import Sub


Sub PrintMsg()
  Text="VBA 테스트"
  MsgBox Text
End Sub