Attribute VB_Name = "modObjectPicture"
Sub ExportWorksheetPicturesTest(x)
    ExportWorksheetPictures Activesheet
End Sub

Sub ExportWorksheetPictures(sht As Worksheet)
    Dim tmpChart As Chart
    Dim shp As Shape, iShp As Long
    Dim imgPath As String

    Application.ScreenUpdating = False
        
    For Each shp In sht.Shapes
        iShp = iShp + 1
        imgPath = sht.Parent.path & "\pict " & sht.Name & " - " & iShp & " " & shp.Name & ".jpg"
        
        'create chart as a canvas for saving this picture
        Set tmpChart = Charts.Add
        tmpChart.Name = "TemporaryPictureChart"
        'move chart to the sheet where the picture is
        Set tmpChart = tmpChart.Location(Where:=xlLocationAsObject, Name:=sht.Name)

        'resize chart to picture size
        tmpChart.ChartArea.Width = shp.Width
        tmpChart.ChartArea.Height = shp.Height
        tmpChart.Parent.Border.LineStyle = 0 'remove shape container border

        'copy picture
        shp.Copy

        'paste picture into chart
        tmpChart.ChartArea.Select
        tmpChart.Paste

        'save chart as jpg
        tmpChart.Export Filename:=imgPath, FilterName:="jpg"
        Debug.Print imgPath
        
        'delete chart
        sht.cells(1, 1).Activate
        sht.ChartObjects(sht.ChartObjects.Count).Delete
    Next
    
    Application.ScreenUpdating = True
End Sub
