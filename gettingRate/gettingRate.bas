Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    Dim pts(1) As Range
    
    Set pts(0) = Range(Sheets(1).Cells(1, 1), Sheets(1).Cells(3, 1))
    

    If Not Intersect(pts(0), Range(Target.Address)) Is Nothing Then
        Call myauto2
    End If

End Sub

Sub Auto_Open()
    Application.OnKey "~", "myauto"
End Sub


Sub mainnameSht1()
    
    Call abc
    
    If Sheets(1).Cells(4, 6).Value >= Sheets(1).Cells(43, 6).Value Then
        Sheets(1).Cells(16, 23).Value = Sheets(1).Cells(4, 6).Value
    ElseIf Sheets(1).Cells(4, 6).Value < Sheets(1).Cells(43, 6).Value Then
        Sheets(1).Cells(16, 23).Value = Sheets(1).Cells(4, 6).Value
    End If
    
    
End Sub

Sub mainnnameSht2()
    Call EFG
    
    With Sheets(2)
        If Cells(4, 7).Value >= Cells(42, 7).Value Then
            Cells(8, 24).Value = Cells(4, 7)
        ElseIf Cells(4, 7).Value < Cells(42, 7).Value Then
            Cells(8, 24).Value = Cells(42, 7)
        End If
    End With
    

End Sub



'rate per buildig price(sheets(1))
Sub abc()
    Dim clt As New Collection
    
    Dim rwReceive As Integer
    Dim clReceive As Integer
    
    Dim NumReceive As Double
    Dim resultY As Double
    
    NumReceive = inputnumber(1, 5, 3)
    
    'getrw(sheet, start obs, inputnumber
    rwReceive = getrw(1, 17, NumReceive)
    'getrate(clnumber, inputstr_this must 4 lengh string
    clReceive = getrate(4, inputStringData(1, 6, 3))
    

    resultY = finalRate(NumReceive, rwReceive, clReceive)
    'this represent multiple rate coef_
    Sheets(1).Cells(7, 3).Value = resultY & "%"
    MsgBox "my result:" & NumReceive * (resultY * 0.01)
    'below line represent result of total value
    'Sheets(1).Cells(16, 23).Value = NumReceive * (resultY * 0.01)
    'MsgBox resultY
    'complete getting rw, cl
    'Sheets(1).Cells(1, 23).Value = "t"
End Sub

'0 = X, 1 = x1, 2 = x2, 3 = y1, 4 = y2
'0 = X, 1 = rw, 2 = cl
Function finalRate(ParamArray par() As Variant) As Double
    Dim pt(2) As Range
    Dim temp(5) As Double
    
    Set pt(0) = Sheets(1).Cells(par(1), par(2))
    
    temp(0) = (pt(0) - pt(0).Offset(1, 0))
    
    Set pt(1) = Sheets(1).Cells(par(1), 2)
    
    temp(1) = par(0) - pt(1).Value
    
    temp(2) = (temp(0) * temp(1))
    temp(4) = pt(1).Offset(1, 0) - pt(1)
    
    finalRate = pt(0) - (temp(2) / temp(4))
    
    
    
End Function

Function inputStringData(ParamArray par() As Variant) As String
    inputStringData = Sheets(par(0)).Cells(par(1), par(2)).Value

End Function

Function inputnumber(ParamArray par() As Variant) As Double
    inputnumber = Sheets(par(0)).Cells(par(1), par(2)).Value

End Function

'it returns column values
'must input 2 values on this proc
'before using this proc then cutting strings with 2words
'getrate(cl value, totalstring value
'standard cl value is 4
Function getrate(ParamArray par() As Variant) As Integer
    Dim cl As Integer
    Dim divStr(2) As String
    
    cl = par(0)
    divStr(0) = Left(par(1), 2)
    divStr(1) = Right(par(1), 2)
    
    'par(1) goto proc clGetRate()
    'it returns final cl value
    'clGetRate(targetstring, present cl number)
    If divStr(0) = "3종" Then
        getrate = clGetRate(cl, divStr(1))
    ElseIf divStr(0) = "2종" Then
        cl = cl + 3
        getrate = clGetRate(cl, divStr(1))
        cl = cl + 6
    ElseIf divStr(0) = "1종" Then
        getrate = clGetRate(cl, divStr(1))
    End If

End Function

'clgetrate(string from xldata, temporaly cl number)
'this proc must return seperate column numbers
Function clGetRate(ParamArray par() As Variant) As Integer

    If par(1) = "상급" Then
        clGetRate = par(0)
    ElseIf par(1) = "중급" Then
        clGetRate = par(0) + 1
    ElseIf par(1) = "기본" Then
        clGetRate = par(0) + 2
    End If
End Function

'getrw(sheetsno, startrwonuber, inputdata, colnumber)
'now input row number is 17
Function getrw(ParamArray par() As Variant) As Integer
    Dim pts(1) As Range
    Dim pt As Range
    
    'for using return obs number
    Dim rwreturn As Integer
    
    rwreturn = (par(1) - 1)
    
    'par(0) is sheets number
    'par(1) is start obs number of range pts(0)
    
    'range must set lower than (maximum range - 1)
    Set pts(0) = Range(Sheets(par(0)).Cells(par(1), 2), Sheets(par(0)).Cells((par(1) + 15), 2))

    'par(2) is inputdata, this data must return obs number of pts(0) range
    For Each pt In pts(0)
        rwreturn = (rwreturn + 1)
        If par(2) > pt And par(2) < pt.Offset(1, 0) Then
            getrw = rwreturn
        End If
    Next

End Function


Function ecol(Optional n = 1, Optional m = 1) As Double

    ecol = Sheets(n).Cells(m, Columns.Count).End(xlToLeft).Column

End Function



Function erow(Optional n = 1, Optional m = 1) As Double

    erow = Sheets(n).Cells(Rows.Count, m).End(xlUp).Row

End Function


Sub EFG()
    Dim rwReceive As Integer
    Dim clReceive As Integer

    Dim NumReceive As Double
    
    Dim resultY As Double
    
    NumReceive = inputnumber(2, 4, 7)
        
    rwReceive = getrwEFG(2, 16, NumReceive)
    resultY = finalRateEFG(NumReceive, rwReceive, inputStringData(2, 4, 11))
    
    Debug.Print "resultY value:" & resultY
    'below line represent result of EFG proc
    'Sheets(2).Cells(8, 24).Value = (NumReceive * (resultY / 100))
    'Sheets(2).Cells(1, 24).Value = "t"
End Sub

Function getrwEFG(ParamArray par() As Variant) As Integer
    Dim pts(1) As Range
    Dim pt As Range
    
    'for using return obs number
    Dim rwreturn As Integer
    
    rwreturn = (par(1) - 1)
    
    'par(0) is sheets number
    'par(1) is start obs number of range pts(0)
    
    'range must set lower than (maximum range - 1)
    Set pts(0) = Range(Sheets(par(0)).Cells(par(1), 3), Sheets(par(0)).Cells((par(1) + 15), 3))

    'par(2) is inputdata, this data must return obs number of pts(0) range
    For Each pt In pts(0)
        rwreturn = (rwreturn + 1)
        If par(2) > pt And par(2) < pt.Offset(1, 0) Then
            getrwEFG = rwreturn
        End If
    Next

End Function

'0 = X, 1 = rw, 2 = string
Function finalRateEFG(ParamArray par() As Variant) As Double
    Dim temp(5) As Double
    Dim str As String
    
    Dim offsetno As Integer
    
    Dim pt As Range
    
    str = Mid(par(2), 2, 2)
    
    If str = "3종" Then
        offsetno = 1
    ElseIf str = "2종" Then
        offsetno = 4
    ElseIf str = "1종" Then
        offsetno = 7
    End If
    'Debug.Print "offsetno:" & offsetno

    Set pt = Sheets(2).Cells(par(1), 3)
    
    temp(0) = par(0) - pt
    temp(1) = ((pt.Offset(0, offsetno) - pt.Offset(1, offsetno)) / 100)
    temp(2) = pt.Offset(1, 0) - pt
    temp(3) = ((temp(0) * temp(1)) / temp(2))
    finalRateEFG = pt.Offset(0, offsetno) - temp(3)
    
    Debug.Print "standard:" & pt
    Debug.Print "offset result2:" & (pt.Offset(0, offsetno) - pt.Offset(1, offsetno))

End Function




Sub myMerge()
    Dim pts(1) As Range
    Dim pt As Range
    
    Dim cl As Integer
    
    cl = 7
    
    With Sheets(1)
        Set pts(0) = Range(Cells(46, cl), Cells(54, cl))
    End With
    
    For Each pt In pts(0)
        Range(pt, pt.Offset(, 1)).Merge
    Next

End Sub

Sub rwHeight()

    Dim pts(1) As Range
    Sheets(1).Cells(2, 26).Value = "3"
    
    Set pts(0) = Range(Sheets(1).Cells(59, 1), Sheets(1).Cells(59, 100))
    pts(0).RowHeight = 17
End Sub

Sub clwidhth()

End Sub

Sub restrictedWord()

    Dim clt As New Collection

    
End Sub

Sub dateFormatChanger()
    Dim temp As String
    
    temp = Sheets(2).Cells(18, 29).Value
    
    Sheets(2).Cells(18, 29).Value = Left(temp, 4) & "년 " & Mid(temp, 5, 2) & "월 " & Right(temp, 2) & "일"

End Sub

Sub returnActiveCell()

    Dim x, y As Integer
    
    x = ActiveCell.row
    y = ActiveCell.Column
    
    MsgBox "x_" & x & "y_" & y
    

End Sub
