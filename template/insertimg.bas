Sub insert_pic()
    Dim pic As Picture
    Dim shp As Shape
    Dim Path As String
    
    Dim i As Integer
    
    i = 0
    Path = "C:\cv2crop_" & i & ".png"
    Set pic = Sheets(1).Pictures.Insert(Path)
    pic.Name = "mypicture"
    
    Set shp = Sheets(1).Shapes("mypicture")
    With shp
        .Height = 100
        .Width = 75
        .LockAspectRatio = msoCTrue
        .Placement = 1
        .Top = 100
        .Left = 100
    End With
End Sub
