Sub abc()
    Dim row As Integer
    Dim pt(1) As Range
    Dim lastrow As Integer
    
    
    lastrow = Sheets(1).Cells(Rows.Count, 3).End(xlUp).row

    row = 2
    Do While Sheets(1).Cells(row, 3) <> ""
        Sheets(1).Cells(row, 11).Value = Sheets(1).Cells(row, 3).Value & Sheets(1).Cells(row, 4).Value
        row = row + 1
    Loop
    
    Set pt(0) = Range(Sheets(1).Cells(2, 11), Sheets(1).Cells(lastrow, 11))
    row = 2
    Do While Sheets(1).Cells(row, 11) <> ""
        Sheets(1).Cells(row, 10).Value = Application.WorksheetFunction.CountIf(pt(0), Sheets(1).Cells(row, 11))
        row = row + 1
    Loop

    Set pt(1) = Range(Sheets(1).Cells(2, 11), Sheets(1).Cells(Rows.Count, 11))
    pt(1).ClearContents
End Sub
