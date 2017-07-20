Attribute VB_Name = "주소록_생성"

Sub Comp()


Dim Name
Dim InputAddr As Range, compAddr As Range, FirstcompAddr As Range
Set DB_Range = Sheets(4).Range("A:A")


Dim keyCell As Range
Dim ro
Dim Rng As Range

Set Rng = Selection

'범위 내 계속 실행
i = 0
For Each ro In Rng.Rows


    '찾을 행
    Set keyCell = Cells(i + Rng.Row, 3)
    Name = keyCell.Value
    Set InputAddr = keyCell.Resize(, 8)

    If Name = "" Then End

    With DB_Range

        Set Finder = .Find(Name, Lookat:=xlWhole)
        Debug.Print (Name & "찾음")

    '이름이 있음
    If Not Finder Is Nothing Then

        '비교 위해서
        Set FirstcompAddr = Finder.Resize(, 8)
        Set compAddr = Finder.Resize(, 8)

        Debug.Print ("비교 대상 " & compAddr.Address)
        Debug.Print (Rng2List(compAddr)(0))

        Do


        '같은 거 나오면 이번행은 끝내기
            If Rng2List(compAddr)(0) = Rng2List(InputAddr)(0) Then
            GoTo EndLine

            Else  '주소 같은거 계속 찾음

                Set Finder = .FindNext(Finder)
                Set compAddr = Finder.Resize(, 8)

            End If


        '동명이인 같은 주소 못 찾음
        Loop While Not Finder Is Nothing And Rng2List(compAddr)(0) <> Rng2List(FirstcompAddr)(0)



        '공백 넣고
        Sheets(4).Range("K7:R7").Copy
        compAddr.Insert Shift:=xlDown


        '복사
        InputAddr.Copy
        Finder.Offset(-1, 0).PasteSpecial xlPasteValues
        Application.CutCopyMode = False

    Else ' 이름이 없는 경우


        '끝에 새로 복사
        InputAddr.Copy
        Sheets(4).Cells(Rows.Count, "A").End(3)(2).PasteSpecial xlPasteValues
        Application.CutCopyMode = False

    End If

    End With

EndLine:

i = i + 1
Next ro


End Sub
