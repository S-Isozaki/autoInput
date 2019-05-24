Sub foo()
    Dim buf As String, cnt As Long
    Dim numberedColumn As Long
    numberedColumn = 3
    Dim Path As String
    Path = "C:\Users\cmm_user\Desktop\“ü—ÍƒeƒXƒg\"
    buf = Dir(Path & "*chr.txt")
    cnt = 0
    Do While buf <> ""
        Open Path & buf For Input As #1
        Do
            Dim buf2 As String
            Line Input #1, buf2
            If buf2 = "END" Then Exit Do
            Dim stringArray() As String
            stringArray = Split(buf2, vbTab)
            Dim numberOfResultAsString As String
            numberOfResultAsString = Left(stringArray(2), 2)
            If IsNumeric(numberOfResultAsString) = False Then GoTo CONTINUE
            Dim Result As Double
            Result = Val(stringArray(5))
            Dim numberOfResult As Long
            numberOfResult = Val(numberOfResultAsString)
            Dim i As Long
            For i = 1 To 2
                Worksheets(i).Select
                Dim j As Long
                For j = 2 To 100
                    Dim s As String
                    s = Worksheets(i).Cells(j, numberedColumn).Value
                    If IsNumeric(s) = False Then GoTo CONTINUE2
                    Dim v As Long
                    v = Val(s)
                    If Worksheets(i).Cells(j, numberedColumn + cnt + 9).Borders(xlDiagonalUp).LineStyle = xlContinuous Then GoTo CONTINUE2
                    If numberOfResult = v Then
                        Worksheets(i).Cells(j, numberedColumn + cnt + 9).Value = WorksheetFunction.Round(Result, 4)
                    End If
CONTINUE2:
                Next j
            Next i
CONTINUE:
        Loop
        Close #1
        cnt = cnt + 1
        buf = Dir()
    Loop
End Sub
