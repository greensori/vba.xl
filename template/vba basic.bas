Sub finder()
    Dim i As Integer
    Dim temp As Integer
    Dim pt(3) As Range
    Dim eCol As Integer
    Dim totalstr As String
    Dim maxcount As Integer
    Dim count As Integer
    Dim stime As Single
    
    Dim row As Integer
    Dim tempdata As String
    
    stime = Timer
    
    count = 0
    i = 3
    eCol = eColproc()
    Sheets(1).Cells(2, 3).Value = eCol
    
    Do While Sheets(1).Cells(i, 2).Value <> ""
        Sheets(1).Cells(i, 1).Value = (i - 2)
        Set pt(0) = Range(Sheets(1).Cells(i, 2), Sheets(1).Cells(i, eCol))
        Set pt(1) = Range(Sheets(2).Cells(i, 4), Sheets(2).Cells(i, (eCol + 2)))
        pt(0).Copy pt(1)
        totalstr = sumstr(i, eCol)
        Sheets(2).Cells(i, 2).Value = sumstr(i, eCol)
        Sheets(2).Cells(i, 1).Value = (i - 2)
        i = i + 1
    Loop
    
    
    i = 3
    Do While Sheets(2).Cells(i, 2).Value <> ""
        tempdata = Sheets(2).Cells(i, 2).Value
        row = i + 1
        Do While Sheets(2).Cells(row, 2).Value <> ""
            If Sheets(2).Cells(row, 2).Value = tempdata Then
                Sheets(2).Cells(row, 2).EntireRow.Delete
                count = count + 1
                row = row - 1
            End If
            row = row + 1
        Loop
        i = i + 1
    Loop
    
    MsgBox ("총 " & count & " 개의 중복 데이터 제거에 " & (Timer - stime) & "초가 걸림")
    
End Sub


Function eColproc() As Integer
    eColproc = Sheets(1).Cells(3, Columns.count).End(xlToLeft).Column
End Function

Function sumstr(rowno As Integer, max As Integer) As String
    Dim icol As Integer
        
    icol = 2
    
    Do While icol <= max
        sumstr = sumstr & Sheets(1).Cells(rowno, icol).Value
        icol = icol + 1
    Loop
End Function

