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