Attribute VB_Name = "�ּҷ�_����"

Sub Comp()


Dim Name
Dim InputAddr As Range, compAddr As Range, FirstcompAddr As Range
Set DB_Range = Sheets(4).Range("A:A")


Dim keyCell As Range
Dim ro
Dim Rng As Range

Set Rng = Selection

'���� �� ��� ����
i = 0
For Each ro In Rng.Rows


    'ã�� ��
    Set keyCell = Cells(i + Rng.Row, 3)
    Name = keyCell.Value
    Set InputAddr = keyCell.Resize(, 8)

    If Name = "" Then End

    With DB_Range

        Set Finder = .Find(Name, Lookat:=xlWhole)
        Debug.Print (Name & "ã��")

    '�̸��� ����
    If Not Finder Is Nothing Then

        '�� ���ؼ�
        Set FirstcompAddr = Finder.Resize(, 8)
        Set compAddr = Finder.Resize(, 8)

        Debug.Print ("�� ��� " & compAddr.Address)
        Debug.Print (Rng2List(compAddr)(0))

        Do


        '���� �� ������ �̹����� ������
            If Rng2List(compAddr)(0) = Rng2List(InputAddr)(0) Then
            GoTo EndLine

            Else  '�ּ� ������ ��� ã��

                Set Finder = .FindNext(Finder)
                Set compAddr = Finder.Resize(, 8)

            End If


        '�������� ���� �ּ� �� ã��
        Loop While Not Finder Is Nothing And Rng2List(compAddr)(0) <> Rng2List(FirstcompAddr)(0)



        '���� �ְ�
        Sheets(4).Range("K7:R7").Copy
        compAddr.Insert Shift:=xlDown


        '����
        InputAddr.Copy
        Finder.Offset(-1, 0).PasteSpecial xlPasteValues
        Application.CutCopyMode = False

    Else ' �̸��� ���� ���


        '���� ���� ����
        InputAddr.Copy
        Sheets(4).Cells(Rows.Count, "A").End(3)(2).PasteSpecial xlPasteValues
        Application.CutCopyMode = False

    End If

    End With

EndLine:

i = i + 1
Next ro


End Sub
