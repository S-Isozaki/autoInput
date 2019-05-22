Sub foo()
    Dim buf As String, cnt As Long
    Dim numberedColumn As Long
    numberedColumn = 0
    buf = Dir("./*chr.txt")
    cnt = 0
    Do While buf <> ""
        cnt = cnt + 1
        Open buf for Input As #1
        Do
            Dim buf2 As String
            Line Input #1, buf2
            If buf2 = "END" Then Exit Do End If
            Dim stringArray() As String
            stringArray = Split(buf2, "\t")
            Dim numberOfResultAsString As String
            numberOfResultAsString = Left(stringArray(4), 2)
            If isNumeric(numberOfResult) = False Then GoTo CONTINUE End If
            Dim numberOfResult As Long
            numberOfResult = Val(numberOfResultAsString)
            Dim i As Long
            For i = 0 To 1
                Worksheets(i).select
                Dim j As Long
                For j = 0 To 50
                    Dim s As String
                    s = Cells(numberedColumn, j)
                    If isNumeric(s) = False Then GoTo CONTINUE2 End If
                    CONTINUE2:
            CONTINUE:
        Loop