# 주소록 자동생성 프로그램으로 시작

#AutoAddress #VBA #Excel #Macro#주소록 #자동작성

범위를 선택하고 주소록 작성 버튼을 누르면, 
선택한 범위에 포함된 행의 주소 데이터가 엑셀의 다른 시트에 중복 없이 정렬되는 프로그램입니다

Sheet2에 있는 CommandButton Code에

-----------------------------------------
Private Sub CommandButton1_Click()

Call Comp

End Sub

------------------------------------------

를 입력해야 합니다


# 업데이트 계획

엑셀에서 쓰던 파일을 그대로 따론 프로토타입이라
Sheet와  Range가 특정되어 있습니다. 

현재버전에서
Sheet 2의 선택된 행의 A:J 열 데이터가
Sheet 4의 A:J 열에 누적됩니다.

향후, 주소록 버튼에 옵션설정 버튼을 추가하고, 옵션안에서

복사할 시트, 붙여질 시트, 복사할 범위( 시작 : 끝 )  으로 설정이 가능하도록 업데이트할 계획입니다.
최종적으로 릴리즈 전에 모듈로 설치할 수 있는 형태로 완성 계획