Sub Find()
    Dim i(1) As Integer
    
    i(0) = 3
    i(1) = 2
    Do While Sheets(1).Cells(i(0), 3) <> ""
        tempdata = Sheets(1).Cells(i(0), 3).Value
        i(1) = 2
        '28 regnum 29 certnum
        Do While Sheets(1).Cells(i(1), 28) <> ""
            If Sheets(1).Cells(i(1), 28).Value = tempdata Then
                Sheets(1).Cells(i(0), 1) = Sheets(1).Cells(i(1), 27)
                Sheets(1).Cells(i(0), 2) = Sheets(1).Cells(i(1), 29)
            End If
            i(1) = i(1) + 1
        Loop
        i(0) = i(0) + 1
    Loop
End Sub
