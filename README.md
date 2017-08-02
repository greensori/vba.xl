# xl.vba


Public c As Integer

Sub total()
    Dim i As Integer
    Dim row(1 To 2) As Double
    Dim data(1 To 25) As String
    Dim temp_data As Long
    Dim Len_targetname As Integer
    Dim Len_totalname As Integer
    Dim count As Integer
    'row(1) is set data on sheet1, and row(2) set data on sheet2
    row(1) = c
    row(2) = c

    b = 1
    '#1 input registration number
    tempdata = Sheets(1).Cells(row(1), 1).Value
    Sheets(2).Cells(row(2), b).Value = tempdata
    b = b + 1
    '#2~3 input korean name
    For i = 5 To 6
        tempdata = Sheets(1).Cells(row(1), i).Value
        Sheets(2).Cells(row(2), b).Value = tempdata
        b = b + 1
    Next
    '#4 input eng name
    i = 7
    tempdata = Sheets(1).Cells(row(1), i).Value
    Len_totalname = Len(tempdata)
    Len_targetname = InStr(1, tempdata, " ", 0)
    Ltempdata = Left(tempdata, Len_targetname - 1)
    Mtempdata = Right(tempdata, Len_totalname - Len_targetname)
    Rtempdata = Mtempdata
    count = Len(Mtempdata)
    Len_targetname = InStr(1, Mtempdata, " ", 0)
    If Len_targetname = 0 Then
        Len_targetname = Len_targetname + 1
        Sheets(2).Cells(row(2), b).Interior.Color = vbRed
        Sheets(2).Cells(row(2), b + 1).Interior.Color = vbRed

    End If
    Mtempdata = Left(Mtempdata, Len_targetname - 1)
    Rtempdata = Right(Rtempdata, count - Len_targetname)
    Ltempdata = Application.WorksheetFunction.Proper(Ltempdata)
    Mtempdata = Application.WorksheetFunction.Proper(Mtempdata)
    Rtempdata = Application.WorksheetFunction.Proper(Rtempdata)
    Utempdata = UCase(Ltempdata)
    tempdata = (Ltempdata & " " & Mtempdata & " " & Rtempdata)
    Sheets(2).Cells(row(2), b).Value = tempdata
    b = b + 1
    tempdata = (Utempdata & " " & Mtempdata & " " & Rtempdata)
    Sheets(2).Cells(row(2), b).Value = tempdata
    b = b + 1
    'certification number
    i = 13
    tempdata = Sheets(1).Cells(row(1), i).Value
    Sheets(2).Cells(row(2), b).Value = tempdata
    b = b + 1
    'date of getting certification
    i = 14
    tempdata = Sheets(1).Cells(row(1), i).Value
    Ltempdata = Left(tempdata, 4)
    Mtempdata = Mid(tempdata, 5, 2)
    Rtempdata = Right(tempdata, 2)
    tempdata = Ltempdata & ". " & Mtempdata & ". " & Rtempdata
    Sheets(2).Cells(row(2), b).Value = tempdata
    b = b + 1
    'input birth date
    i = 8
    tempdata = Sheets(1).Cells(row(1), i).Value
    tempdata = Replace(tempdata, "-", ". ")
    Sheets(2).Cells(row(2), b).Value = tempdata
    b = b + 1
    'to modify picture file name, delete .jpg or .png
    tempdata = Sheets(1).Cells(row(1), 25).Value
    tempdata = Replace(tempdata, ".jpg", "")
    tempdata = Replace(tempdata, ".png", "")
    tempdata = Replace(tempdata, ".JPG", "")
    tempdata = Replace(tempdata, ".PNG", "")
    Sheets(2).Cells(row(2), b).Value = tempdata
    b = b + 1
    'date of regiration
    i = 15
    tempdata = Sheets(1).Cells(row(1), i).Value
    Sheets(2).Cells(row(2), b).Value = tempdata
    b = b + 1
    'expire data
    i = 16
    tempdata = Sheets(1).Cells(row(1), i).Value
    tempdata = Replace(tempdata, ".", "")
    Ltempdata = Left(tempdata, 4)
    Mtempdata = Mid(tempdata, 5, 2)
    Rtempdata = Right(tempdata, 2)
    tempdata = Ltempdata & "년 " & Mtempdata & "월 " & Rtempdata & "일까지"
    Sheets(2).Cells(row(2), b).Value = tempdata
    b = b + 1
    'input zip code, len(zipcode) = 6, zipcode(3) = zipcode(3)+'-'
    i = 9
    tempdata = Sheets(1).Cells(row(1), i).Value
    If Len(tempdata) = 6 Then
        Ltempdata = Left(tempdata, 3)
        Rtempdata = Right(tempdata, 3)
        tempdata = Ltempdata & "-" & Rtempdata
    End If
    Sheets(2).Cells(row(2), b).Value = tempdata
    b = b + 1
    'input address
    i = 10
    tempdata = Sheets(1).Cells(row(1), i).Value
    Sheets(2).Cells(row(2), b).Value = tempdata
    b = b + 1
End Sub



Sub Main()
    Dim i As Integer
    Dim lastrow As Long
    
    i = 3
    c = 3
    Do While Sheets(1).Cells(i, 1).Value <> ""
        i = i + 1
        Call total
        c = c + 1
    Loop
End Sub

Sub Saves()
    Dim strPath(2) As String
    Dim obs(1) As Double
    Dim countrow As Double
    
    obs(0) = Sheets(1).Cells(3, 1).Value
    countrow = Sheets(1).Cells(Rows.count, 1).End(xlUp).row
    obs(1) = Sheets(1).Cells(countrow, 1).Value
    strPath(0) = "C:\Users\건축사등록원6887\Desktop\공유\1. 건축사등록원\1__1 건축사등록_발급(등록증 등록카드)\1) 로우데이터 만들기\"
    ChDir strPath(0)
    ActiveWorkbook.SaveAs Filename:=strPath(0) & Date & "기준 " & " #" & obs(0) & " ~ " & obs(1) & ".xlsx"
    MsgBox "Path : " & strPath(0) & vbCrLf & "Sheets 1 ~ 4 Saves complete"

End Sub

Sub Clear()
    Dim a As Double
    Dim row_len(1) As Double
    Dim col_len(1) As Double
    row_len(0) = 3
    col_len(0) = 1
    Dim rng As Range
    Dim n As Integer
    
    n = ActiveCell.row
    Set rng = Range(Cells(3, 1), Cells(7, 4))
    
    Sheets(1).Range(Cells(row_len(0), col_len(0)), Cells(Rows.count, Columns.count)).ClearContents
End Sub




'##색상 바꾸기
'Sheets(1).Cells(12, 12).Interior.Color = vbYellow
