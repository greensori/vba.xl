Sub mainname()
    
    If Sheets(2).Cells(69, 5).Value > 1 Then
        Call abc
    ElseIf Sheets(2).Cells(71, ecol(2, 71)) > 1 Then
        Call efg
    End If
End Sub

Sub abc()
    Dim clt As New Collection
    Dim inputdata As Long
    
    Dim rwReceive As Integer
    
    rwReceive = getrw(2, 82, inputnumber(2, 69, 5))
    
End Sub

Sub efg()
    MsgBox "seconds calc"
End Sub


Function inputnumber(ParamArray par() As Variant) As Long
    inputnumber = Sheets(par(0)).Cells(par(1), par(2)).Value

End Function

Function getrate(ParamArray par() As Variant) As Integer


End Function

'getrw(sheetsno, startrwonuber, inputdata)
'now input row number is 82
Function getrw(ParamArray par() As Variant) As Integer
    Dim pts(1) As Range
    Dim pt As Range
    
    'for using return obs number
    Dim rwReturn As Integer
    
    rwReturn = (par(1) - 1)
    
    'par(0) is sheets number
    'par(1) is start obs number of range pts(0)
    
    'range must set lower than (maximum range - 1)
    Set pts(0) = Range(Sheets(par(0)).Cells(par(1), 1), Sheets(par(0)).Cells((par(1) + 15), 1))

    'par(2) is inputdata, this data must return obs number of pts(0) range
    For Each pt In pts(0)
        rwReturn = (rwReturn + 1)
        If par(2) > pt And par(2) < pt.Offset(1, 0) Then
            getrw = rwReturn
            MsgBox getrw
        End If
    Next

End Function


Function ecol(Optional n = 1, Optional m = 1) As Double

    ecol = Sheets(n).Cells(m, Columns.Count).End(xlToLeft).Column

End Function



Function erow(Optional n = 1, Optional m = 1) As Double

    erow = Sheets(n).Cells(Rows.Count, m).End(xlUp).Row

End Function


Sub myMerge()
    Dim pts(1) As Range
    Dim pt As Range
    
    
    Set pts(0) = Range(Sheets(2).Cells(53, 14), Sheets(2).Cells(63, 14))
    
    For Each pt In pts(0)
        Range(pt, pt.Offset(, 3)).Merge
    Next

End Sub
