Public count As Integer
Public sht As Worksheet

Sub name()
    Set sht = Sheets(1)
    count = 1
    
    Call countif(1, 2)
End Sub


Function countif(ParamArray col() As Variant)
    Dim obs As Integer
    Dim pts(1) As Range
    Dim pt As Range
    
    Dim param As Variant
    
    'setting start rows number
    obs = count
    
    'setting col values col(0) is target range and col(1) is result range
    With sht
        Set pts(0) = Range(Cells(obs, col(0)), Cells(erow(), col(0)))
        For Each pt In pts(0)
            Cells(obs, col(1)).Value = Application.WorksheetFunction.countif(pts(0), pt)
            obs = obs + 1
        Next
    End With
End Function
    

Function ecol(Optional n = 1, Optional m = 1) As Double
    ecol = Sheets(n).Cells(m, Columns.count).End(xlToLeft).Column
    
End Function

Function erow(Optional n = 1, Optional m = 1) As Double
    erow = Sheets(n).Cells(Rows.count, m).End(xlUp).row
End Function
