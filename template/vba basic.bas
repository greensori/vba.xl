# change cell color example
Sheets(1).Cells(12, 12).Interior.Color = vbYellow

# to avoid '1004'error when using vlookup in vba
On Error Resume Next
  pt(0) = Application.WorksheetFunction.VLookup(pt(1), pt(2), 3, 0)

#erase #N/A
Do While Sheets(1).Cells(rno(0), 1) <> ""
    If Sheets(1).Cells(rno(0), 1).Value = "#N/A" Then
        Sheets(1).Cells(rno(0), 1).Value = ""
    End If

#replace word
Sheets(1).Cells(row(1), 25) = Replace(Sheets(1).Cells(row(1), 25), ".jpg", "")

#drawing strikethrough
Range(Sheets(2).Cells(i, 2), Sheets(2).Cells(i, 5)).Font.Strikethrough = True

#clearcontents
Set pt = Range(Sheets(1).Cells(5, 5), Sheets(1).Cells(Rows.Count, Columns.Count))
pt.ClearContents

#getting last row value
Sheets(1).Cells(1, 3).Value = Sheets(1).Cells(Rows.Count, 1).End(xlUP).Row

# show or hidden rows and columns
Sheets(1).Rows("1:4").Hidden = True / False
Sheets(1).Columns("A:D").Hidden = True / False

#count blank
Application.WorksheetFunction.CountBlank(Range(Sheets(1).Cells(3, 2), Sheets(1).Cells(Sheets(1).Cells(Rows.Count, 2).End(xlUp).Row, 2)))

#used time
Dim stime As Single
stime = Timer
Sheets(1).Cells(1, 1).Value = Format(Timer - stime, "#0.00")

#copy data
Sub test()
    Dim pt(1) As Range
    Dim sheetno(1) As Integer
    Dim i(2) As Double
    Dim count As Integer
    
    count = 0
    sheetno(0) = 1
    sheetno(1) = 2
    i(0) = 1
    i(1) = 1000
    i(2) = 1000
    
    Do While count < 7
        Set pt(0) = Range(Sheets(sheetno(0)).Cells(i(0), 3), Sheets(sheetno(0)).Cells(i(1), 5))
        Set pt(1) = Range(Sheets(sheetno(1)).Cells(1, 3), Sheets(sheetno(1)).Cells(1000, 5))
        pt(0).Copy pt(1)
        i(0) = i(0) + i(2)
        i(1) = i(1) + i(2)
        count = count + 1
        sheetno(1) = sheetno(1) + 1
    Loop
End Sub

#preventdoubledata
    Do While Sheets(1).Cells(row, 2) <> ""
        Sheets(1).Cells(row, 14).Value = Application.WorksheetFunction.CountIf(pt(0), Sheets(1).Cells(row, 2))
        row = row + 1
    Loop


# delete specific data

    Dim pt As Range
    Dim row As Long
    Dim lstrow As Long
    row = 1
    
    lstrow = Sheets(1).Cells(Rows.Count, 2).End(xlUp).row
    Do While row <= lstrow
        If Sheets(1).Cells(row, 2).Value = "" Then
            Sheets(1).Cells(row, 2).EntireRow.Delete
        End If
        row = row + 1
    Loop


# changing date format

Dim Dvalue As String
Dvalue = Format(Date, "yyyy. mm. dd")
