Sub test_mean():
    Dim rng(2) As Range
    Dim Lnumber As Double
    
    Set rng(0) = Sheets(1).Range(Sheets(1).Cells(3, 1), Sheets(1).Cells(14648, 1))
    tempdata = Application.WorksheetFunction.Average(rng(0))
    Sheets(1).Cells(1, 3).Value = tempdata
    
    i = 3
    Do While Sheets(1).Cells(i, 1) <> ""
        tempdata1 = Sheets(1).Cells(i, 1).Value
        Sheets(1).Cells(i, 2).Value = Abs(tempdata - tempdata1)
        'tempdata4 = Sheets(1).Cells(i, 2).Value
        i = i + 1
    Loop
    
    Set rng(1) = Sheets(1).Range(Sheets(1).Cells(3, 2), Sheets(1).Cells(14648, 2))
    tempdata = Application.WorksheetFunction.Average(rng(1))
    Sheets(1).Cells(1, 4).Value = tempdata
    
End Sub

