
'myCountif(sheetnumber, startrownumber, endrownumber, targetcolnumber, outputoffset)
Function myCountif(ParamArray myParam() As Variant)
    Dim pts(1) As Range
    Dim pt As Range
    
    With Sheets(myParam(0))
        Set pts(0) = Range(Cells(myParam(1), myParam(3)), Cells(myParam(2), myParam(3)))
        For Each pt In pts(0)
            pt.Offset(,myParam(4)) = Application.WorksheetFunction.countif(pts(0), pt)
        Next
    End With
End Function


Function sigmoid(x As Double) As Double
    sigmoid = 1 / (1 + Exp(x))
End Function


Function ecol(Optional n = 1, Optional m = 1) As Double
    ecol = Sheets(n).Cells(m, Columns.count).End(xlToLeft).Column
End Function


Function erow(Optional n = 1, Optional m = 1) As Double
    erow = Sheets(n).Cells(Rows.count, m).End(xlUp).row
End Function
