Attribute VB_Name = "modWbk2Dict"
Option Explicit

' Options
Const bExportAllValuesAsSingleRange = False

' Validation
Const xlValidateList = 3   ' Value must be present in a specified list.
Const xlValidateCustom = 7   ' Data is validated using an arbitrary formula.

' Special Cells
Const xlCellTypeConstants = 2 ' Cells containing constants (text, numbers, ...)
Const xlCellTypeBlanks = 4 ' Empty cells
Const xlCellTypeLastCell = 11 ' The last cell in the used range
Const xlCellTypeFormulas = -4123 ' Cells containing formulas (formulas and array formulas)
Const xlCellTypeComments = -4144 ' Cells containing notes
Const xlCellTypeAllFormatConditions = -4172 ' Cells of any format
Const xlCellTypeSameFormatConditions = -4173 ' Cells having the same format
Const xlCellTypeAllValidation = -4174 ' Cells having validation criteria
Const xlCellTypeSameValidation = -4175 ' Cells having the same validation criteria

Dim globalChartTypeMap As Dictionary

Sub export_workbook_to_file_test()
    Dim wbk As Workbook
    Set wbk = ActiveWorkbook
    
    Dim FilePath As String: FilePath = wbk.FullName & ".json"
    Call export_workbook_to_file(wbk, FilePath)
    Debug.Print FilePath
End Sub

Function export_workbook_to_file(wbk As Workbook, export_path As String)
    Dim d As Dictionary
    Set d = export_workbook(wbk)
    
    Dim json As String
    json = vbs_dict_to_json(d)
    
    'Debug.Print json
    Debug.Print "#json: " & Round(Len(json) / 1024) & "kB " & Round(Len(json) / FileLen(wbk.path & "\" & wbk.Name) * 100) & "%"
    
    Dim FileNumber: FileNumber = FreeFile
    Open export_path For Output Access Write As #FileNumber
    Print #FileNumber, json
    Close #FileNumber
End Function

Function export_workbook(wbk As Workbook)
    Dim nbExcelLinks As Long
    On Error Resume Next
    nbExcelLinks = -1
    nbExcelLinks = Len(wbk.LinkSources(xlExcelLinks))  ' xlExcelLinks, xlOLELinks
    On Error GoTo 0
    
    Dim d As Dictionary
    Set d = New Dictionary
    d.Add "LastSaved", Now()
    d.Add "Author", wbk.Author
    d.Add "Name", wbk.Name
    d.Add "Path", wbk.path
    d.Add "FileSize", FileLen(wbk.path & "\" & wbk.Name)
    d.Add "CodeName", wbk.CodeName
    d.Add "HasVBProject", wbk.HasVBProject
    ' d.Add "ReferenceStyle", IIf(Application.ReferenceStyle = xlA1, "A1", "R1C1") ' always A1
    d.Add "#ExcelLinks", nbExcelLinks
    d.Add "#Styles", wbk.styles.Count
    d.Add "#Charts", wbk.Charts.Count
    
    If nbExcelLinks <> 0 Then
        Set export_workbook = d
        Debug.Print "!! ExcelLinks !!"
        Exit Function
    End If

    Call d.Add("Names", export_names(wbk.names))
    
    'ReDim sht_array(1 To wbk.Sheets.Count) As Dictionary
    Dim sht_list: Set sht_list = CreateObject("System.Collections.ArrayList")
    
    Dim sht_pos As Long
    Dim sht As Worksheet
    For Each sht In wbk.Sheets
        sht_pos = sht_pos + 1
        'Set sht_array(sht_pos) = export_worksheet(sht)
        sht_list.Add export_worksheet(sht)
    Next sht
    'Call d.Add("Worksheets", sht_array)
    Call d.Add("Worksheets", sht_list)
    
    Set export_workbook = d
End Function


Function export_worksheet(sht As Worksheet)
    'sht.Select
    'sht.Activate
    sht.Unprotect
    
    Dim d As Dictionary
    Set d = New Dictionary
    
    d.Add "Name", sht.Name
    d.Add "CodeName", sht.CodeName
    d.Add "Visible", sht.Visible
    'd.Add "DisplayGridlines", ActiveWindow.DisplayGridlines
    d.Add "#Shapes", sht.Shapes.Count
    d.Add "#Charts", sht.ChartObjects.Count ' chart.SourceData
    
    Call d.Add("formulas", export_worksheet_formulas(sht))
    
    If bExportAllValuesAsSingleRange Then
        Call d.Add("values", export_worksheet_values(sht))
    End If
    
    Call d.Add("comments", export_worksheet_comments(sht))
    
    Call d.Add("formats", export_worksheet_format(sht))
    
    Call d.Add("charts", export_worksheet_charts(sht))
    
    Set export_worksheet = d
End Function


Function export_names(nms As names)
    If nms.Count = 0 Then Exit Function
    
    Dim names_dict As Dictionary
    Set names_dict = New Dictionary
    
    Dim nm As Name
    For Each nm In nms
        
        ' Naive implementation: stores Worksheet and Name together separated by "!"
        ' names_dict.Add nm.name, nm.RefersTo
        
        ' Is it a WorkSheet name ? If so, ask Excel to extract it
        Dim nm_sheet As String, nm_bare As String
        Dim nm_sheet_ends: nm_sheet_ends = InStrRev(nm.Name, "!", , vbBinaryCompare)
        If nm_sheet_ends > 0 Then
            ' does not work when RefersTo is not a Range address (e.g. a formula)
            'nm_sheet = Range(nm.name).Worksheet.name
            nm_sheet = Range(Left(nm.Name, nm_sheet_ends) & "A1").Worksheet.Name
            nm_bare = Mid(nm.Name, nm_sheet_ends + 1)
        Else
            nm_sheet = ""
            nm_bare = nm.Name
        End If
        
        If Not names_dict.Exists(nm_sheet) Then
            names_dict.Add nm_sheet, New Dictionary
        End If
        names_dict.Item(nm_sheet).Add nm_bare, nm.RefersTo ' nm.RefersTo or nm.RefersToR1C1
        
        ' Check whether we loose Named range comment
        Debug.Assert nm.Comment = ""
    Next
    
    Set export_names = names_dict
End Function

Function export_worksheet_range(sht As Worksheet) As Range
    Set export_worksheet_range = sht.Range(sht.Range("A1"), sht.cells.SpecialCells(xlCellTypeLastCell))
End Function

Function export_worksheet_values(sht As Worksheet)
    Dim rng As Range
    Set rng = export_worksheet_range(sht)
    
    Dim vals
    vals = rng.Value
    
    Dim json As String, json_row As String, i As Long, j As Long, n As Long, m As Long
    Dim empty_row As Boolean, last_i As Long, last_j As Long
    n = UBound(vals) - LBound(vals) + 1
    m = UBound(vals, 2) - LBound(vals, 2) + 1
    
    For i = LBound(vals) To UBound(vals)
        For j = UBound(vals, 2) To LBound(vals, 2) Step -1
            If Not IsEmpty(vals(i, j)) Then
                last_i = i
                last_j = IIf(j > last_j, j, last_j)
                Exit For
            End If
        Next j
    Next i
        
    If last_i = UBound(vals) And last_j = UBound(vals, 2) Then
        export_worksheet_values = rng.Value
    Else
        export_worksheet_values = sht.Range("A1").Resize(last_i, last_j)
    End If
End Function
    
    
Function export_worksheet_formulas(sht As Worksheet)
    
    Dim last_cell As Range, last_row As Long, last_col As Long
    Set last_cell = sht.cells.SpecialCells(xlCellTypeLastCell)
    last_row = last_cell.Row
    last_col = last_cell.Column
    
    Dim formulas As Dictionary: Set formulas = New Dictionary
    
    Dim i As Long, i2 As Long, j As Long, c As Range, c2 As Range
    Dim already_done_address As Dictionary
    Set already_done_address = New Dictionary
    
    ' Formulas
    For i = 1 To last_row
        For j = 1 To last_col
            
            Set c = sht.cells(i, j)
            
            If already_done_address.Exists(c.Address) Then
                GoTo next_cell
            End If
            
            If c.HasArray Then
                formulas(c.CurrentArray.Address) = "{" & c.FormulaArray & "}"
            ElseIf c.HasFormula Then
                For i2 = i + 1 To last_row
                    Set c2 = sht.cells(i2, j)
                    If c.FormulaR1C1 <> c2.FormulaR1C1 Then
                        Exit For
                    Else
                        already_done_address(c2.Address) = True
                    End If
                Next
                i2 = i2 - 1
                
                If i = i2 Then
                    
                    formulas(c.Address) = c.Formula
                Else
                    Set c2 = sht.cells(i2, j)
                    formulas(c.Address & ":" & c2.Address) = c.Formula
                End If
            ElseIf Not bExportAllValuesAsSingleRange Then
                If IsError(c) Then
                    ' do nothing (an error value)
                ElseIf Not IsEmpty(c.Value) Then
                    For i2 = i + 1 To last_row
                        Set c2 = sht.cells(i2, j)
                        If IsError(c2) Then
                            Exit For
                        ElseIf c.Value <> c2.Value Then
                            Exit For
                        Else
                            already_done_address(c2.Address) = True
                        End If
                    Next
                    i2 = i2 - 1
                    
                    If i = i2 Then
                        formulas(c.Address) = c.Value
                    Else
                        Set c2 = sht.cells(i2, j)
                        formulas(c.Address & ":" & c2.Address) = c.Value
                    End If
                End If
            End If
next_cell:
        Next
    Next
    
    Set export_worksheet_formulas = formulas
End Function

Sub calc_info_test()
    Dim rng As Range: Set rng = Selection
    
    Dim depth As Long, recursive As Boolean, has_deps As Boolean
    depth = calculation_graph_depth(rng)
    recursive = range_calculation_graph_recursive(rng)
    has_deps = range_has_dependence(rng)
    Debug.Print rng.Address & ", depth: " & depth & ", recursive: " & recursive & ", has_deps: " & has_deps
    
End Sub

Function calculation_graph_depth(rng As Range) As Long
    ' using rng.DirectPrecedents to calculate the calculation graph depth
    calculation_graph_depth = 0
    
    Dim prec_rng As Range, direct_prec_rng As Range, prev_nb_prec As Long
    Set prec_rng = rng
    Do
        Set direct_prec_rng = Nothing
        On Error Resume Next
        Set direct_prec_rng = prec_rng.DirectPrecedents
        On Error GoTo 0
        
        If Not direct_prec_rng Is Nothing Then
            prev_nb_prec = prec_rng.Count
            Set prec_rng = Union(prec_rng, direct_prec_rng)
            If prec_rng.Count > prev_nb_prec Then
                calculation_graph_depth = calculation_graph_depth + 1
            End If
        End If
    Loop While (Not direct_prec_rng Is Nothing) And (prec_rng.Count > prev_nb_prec)
End Function

Function range_calculation_graph_recursive(rng As Range) As Boolean
    ' using rng.Precedents to detect whether the range uses itelf as input
    Dim prec_rng As Range
    Set prec_rng = Nothing
    On Error Resume Next
    Set prec_rng = rng.Precedents
    On Error GoTo 0
    If Not prec_rng Is Nothing Then
        If Not Intersect(rng, prec_rng) Is Nothing Then
            range_calculation_graph_recursive = True  ' incorrect when no precedents
        End If
    End If
End Function
    
Function range_has_dependence(rng As Range) As Boolean
    Dim tmp_rng As Range
    Set tmp_rng = Nothing
    On Error Resume Next
    Set tmp_rng = rng.DirectDependents
    On Error GoTo 0
    If Not tmp_rng Is Nothing Then
        If Union(rng, tmp_rng).Count > rng.Count Then
            range_has_dependence = True
        End If
    End If
End Function


Function export_worksheet_comments(sht As Worksheet)
    Dim cell As Range, cells As Range, shp As Shape
    Set cells = sht.cells
    On Error GoTo no_comments
    Set cells = cells.SpecialCells(xlCellTypeComments) ' same technique could be use to go through sparse xlCellTypeConstants & xlCellTypeFormulas
    On Error GoTo 0
    
    Dim comments As Dictionary: Set comments = New Dictionary
    For Each cell In cells
        comments.Add cell.Address, cell.Comment.Text
    Next
no_comments:
End Function

Function export_worksheet_format(sht As Worksheet)
    Dim formats As Dictionary: Set formats = New Dictionary
    
    'Call export_worksheet_columns_format(formats, sht)
    Call export_worksheet_cells_format(formats, sht)
    
    Set export_worksheet_format = formats
End Function

Function export_worksheet_columns_format(d As Dictionary, sht As Worksheet)
    Dim rng As Range
    Set rng = export_worksheet_range(sht)
    
    ReDim column_width(1 To rng.Columns.Count)
    Dim col As Range, iCol As Long
    iCol = 0
    For Each col In rng.Columns
        iCol = iCol + 1
        column_width(iCol) = col.ColumnWidth
    Next
    d.Add "column_width", column_width
End Function
    
Function export_worksheet_cells_format(d As Dictionary, sht As Worksheet)
    Dim last_cell As Range, last_row As Long, last_col As Long
    Set last_cell = sht.cells.SpecialCells(xlCellTypeLastCell)
    On Error Resume Next
    ' Go one cell past the end(mainly for borders
    last_cell = last_cell.Offset(1, 1)
    On Error GoTo 0
    last_row = last_cell.Row
    last_col = last_cell.Column
    
    If last_row > 256 Then last_row = 256
    If last_col > 64 Then last_col = 64
    
    Dim formats As Dictionary
    Dim styles As Dictionary: Set styles = New Dictionary
    
    Dim i As Long, i2 As Long, j As Long, j2 As Long, c As Range, c2 As Range
    
    Dim already_done_address As Dictionary
    
    Dim fmtName As String, fmtNames As Variant, iFmtName As Long, fmt As String, fmt2 As String
    fmtNames = Array("color", "font", "alignment", "number_format", "borders")
    fmtNames = Array("cells", "number_format", "borders")
    
    For iFmtName = LBound(fmtNames) To UBound(fmtNames)
        fmtName = fmtNames(iFmtName)
        Set formats = New Dictionary
        Set already_done_address = New Dictionary
        
        For j = 1 To last_col
            For i = 1 To last_row
            
                If already_done_address.Exists(sht.cells(i, j).Address) Then
                    GoTo next_cell
                End If
                
                ' HorizontalAlignment
                ' Borders
                ' Number format
                Set c = sht.cells(i, j)
                fmt = export_worksheet_cell_format(fmtName, c, styles)
                
                If Len(fmt) > 0 Then
                    For i2 = i + 1 To last_row
                        Set c2 = sht.cells(i2, j)
                        fmt2 = export_worksheet_cell_format(fmtName, c2, styles)
                        If fmt <> fmt2 Then
                            Exit For
                        Else
                            already_done_address(c2.Address) = True
                        End If
                    Next i2
                    i2 = i2 - 1
                            
                    j2 = j
                    If i = i2 Then
                    For j2 = j + 1 To last_col
                        Set c2 = sht.cells(i, j2)
                        fmt2 = export_worksheet_cell_format(fmtName, c2, styles)
                        If fmt <> fmt2 Then
                            Exit For
                        Else
                            already_done_address(c2.Address) = True
                        End If
                    Next j2
                    j2 = j2 - 1
                    End If
                
                    If i = i2 And j = j2 Then
                        formats(c.Address) = fmt
                    Else
                        Set c2 = sht.cells(i2, j2)
                        formats(c.Address & ":" & c2.Address) = fmt
                    End If
                End If
                
next_cell:
            Next i
        Next j
        
        d.Add fmtName, formats
    Next iFmtName
    
    d.Add "styles", styles
End Function

Function export_worksheet_cell_format(fmtName As String, c As Range, styles As Dictionary)
    Dim tmp, fmt As String, styleName As String
    
    Dim bCellFmt As Boolean: bCellFmt = (fmtName = "cells")
    If bCellFmt Then
    
        If fmtName = "cells" Or fmtName = "color" Then
            ' Test color with vbBlack/vbWhite or ColorIndex with 0/1
            ' Special color indices:
            ' xlColorIndexAutomatic: -4105
            ' xlColorIndexNone: -4142
            If c.Font.ColorIndex <> xlColorIndexAutomatic And c.Font.ColorIndex <> xlColorIndexNone And c.Font.Color <> vbBlack Then
                tmp = color_to_rgb_string(c.Font.Color)
                styleName = "fc" & tmp
                styles(styleName) = "color: #" & tmp & ";"
                fmt = fmt & styleName & " "
            End If
            If c.Interior.ColorIndex <> xlColorIndexAutomatic And c.Interior.ColorIndex <> xlColorIndexNone And c.Interior.Color <> vbWhite Then
                tmp = color_to_rgb_string(c.Interior.Color)
                styleName = "bc" & tmp
                styles(styleName) = "background-color: #" & tmp & ";"
                fmt = fmt & styleName & " "
            End If
        End If
        
        If fmtName = "cells" Or fmtName = "font" Then
        
            If c.Font.FontStyle <> "Regular" Then
                fmt = fmt & LCase(c.Font.FontStyle) & " "
            End If
            If c.Font.Underline <> xlUnderlineStyleNone Then fmt = fmt & "underline "
        End If
    
        If fmtName = "cells" Or fmtName = "alignment" Then
            tmp = c.HorizontalAlignment
            If tmp = xlCenter Then fmt = fmt & "center "
            If tmp = xlDistributed Then fmt = fmt & "distributed "
            If tmp = xlJustify Then fmt = fmt & "justify "
            If tmp = xlLeft Then fmt = fmt & "left "
            If tmp = xlRight Then fmt = fmt & "right "
        End If
            
    ElseIf fmtName = "number_format" Then
        If c.NumberFormat <> "General" Then
             If c.NumberFormat = "d-mmm-yy" Then
                 fmt = "dd-mmm-yyyy"
             Else
                 fmt = c.NumberFormat
             End If
        End If
        
        On Error Resume Next
        tmp = -1
        tmp = c.Validation.Type
        On Error GoTo 0
        If tmp >= 0 Then ' InCellDropdown
            ' ' Enforce constraint ?
            fmt = fmt & " { xlValidateList: " & IIf(tmp = xlValidateList, 1, 0) & ", showError: " & IIf(c.Validation.ShowError, 1, 0) & ", list: '" & c.Validation.Formula1 & "' }"
        End If
        
        If c.FormatConditions.Count >= 0 Then ' InCellDropdown
            For Each tmp In c.FormatConditions
                fmt = fmt & " { FormatConditions: " & tmp.Type
                On Error Resume Next
                fmt = fmt & ", condition: '" & tmp.Formula1 & "'"
                On Error GoTo 0
                fmt = fmt & " }"
            Next
        End If
        
    ElseIf fmtName = "borders" Then
        ' xlBordersIndex constants: xlDiagonalDown, xlDiagonalUp, xlEdgeBottom, xlEdgeLeft, xlEdgeRight, or xlEdgeTop, xlInsideHorizontal, or xlInsideVertical.
        If c.Borders(xlEdgeTop).LineStyle <> xlLineStyleNone Then fmt = fmt & "border-top "
        'If c.Borders(xlEdgeBottom).LineStyle <> xlLineStyleNone Then fmt = fmt & "border-bottom "
        If c.Borders(xlEdgeLeft).LineStyle <> xlLineStyleNone Then fmt = fmt & "border-left "
        'If c.Borders(xlEdgeBottom).LineStyle <> xlLineStyleNone Then fmt = fmt & "border-right "
        
    Else
        Debug.Assert False
    End If
    
    export_worksheet_cell_format = Trim(fmt)
End Function

Function color_to_rgb_string(c)
    'c = RGB(255, 128, 0)
    
    ' This produces GreenRedBlue colors
    'Debug.Print "BGR   #" & Hex(c)
    
    ' So we must do it the hard way
    
    Dim r, g, b
    r = c And 255
    g = ((c - r) / 256) And 255
    b = ((c - r - 256 * g) / 65536) And 255

    r = Hex(r)
    g = Hex(g)
    b = Hex(b)
    
    'Debug.Print "Red   #" & r
    'Debug.Print "Green #" & g
    'Debug.Print "Blue  #" & b
    
    Dim txt As String
    txt = IIf(Len(r) = 1, "0", "") & r _
        & IIf(Len(g) = 1, "0", "") & g _
        & IIf(Len(b) = 1, "0", "") & b
    'Debug.Print "RGB   #" & txt
    
    color_to_rgb_string = txt
End Function

Function test_color_index()
    Dim x
    x = RGB(255, 128, 1)
    Debug.Print "Red-ish", color_to_rgb_string(x)
    
    Dim color_index As Long
    For color_index = 1 To 56
        x = ActiveWorkbook.Colors(color_index)
        Debug.Print "color " & color_index & " : " & color_to_rgb_string(x) & " (" & x; ")"
    Next color_index
    
End Function

Function export_worksheet_charts(sht As Worksheet)
    Dim chartList: Set chartList = CreateObject("System.Collections.ArrayList")
    Dim chartDict
    Dim chartObj
    Dim imgPath As String
    On Error Resume Next
    For Each chartObj In sht.ChartObjects
        Set chartDict = New Dictionary
        chartDict("type") = ChartTypeName(chartObj.Chart.ChartType) ' xlLine, xlXYScatter, ...
        chartDict("title") = chartObj.Chart.ChartTitle
        chartDict("formula1") = chartObj.Chart.SeriesCollection(1).Formula
        chartDict("range") = Range(chartObj.TopLeftCell, chartObj.BottomRightCell).Address
        chartDict("width") = chartObj.Width
        chartDict("height") = chartObj.Height
        chartList.Add chartDict
        
        imgPath = "C:\temp\wbk2dict " & chartObj.Name & ".jpg" ' jpg or gif
        chartObj.Chart.Export Filename:=imgPath, FilterName:="jpg"
        Debug.Print imgPath
    Next
    On Error GoTo 0
    Set export_worksheet_charts = chartList
End Function

Function ChartTypeName(ChartTypeEnum)
    If IsEmpty(globalChartTypeMap) Then
        Set globalChartTypeMap = New Dictionary
        
        globalChartTypeMap(xlArea) = "xlArea"  ' 1
        globalChartTypeMap(xlLine) = "xlLine"  ' 4
        globalChartTypeMap(xlLineMarkers) = "xlLineMarkers"  ' 65
        globalChartTypeMap(xlLineStacked) = "xlLineStacked"  ' 63
        globalChartTypeMap(xlPie) = "xlPie"  ' 5
        globalChartTypeMap(xlPieExploded) = "xlPieExploded"  ' 69
        globalChartTypeMap(xlSurface) = "xlSurface" ' 83
        globalChartTypeMap(xlXYScatter) = "xlXYScatter" ' -4169
        globalChartTypeMap(xlXYScatterLines) = "xlXYScatterLines" ' 74
        globalChartTypeMap(xlXYScatterSmooth) = "xlXYScatterSmooth" ' 72
    End If
    
    If globalChartTypeMap.Exists(ChartTypeEnum) Then
        ChartTypeName = globalChartTypeMap(ChartTypeEnum)
    Else
        ChartTypeName = ChartTypeEnum
    End If
End Function

Function export_worksheet_shapes(sht As Worksheet)
    'On Error Resume Next
    Dim shp As Shape, imgPath As String
    For Each shp In sht.Shapes
        imgPath = "C:\temp\wbk2dict_shape_" & shp.Name & ".GIF" ' jpg or gif
        ' shp.SaveAsPicture imgPath
        ' PictureFromObjectToFile shp, imgPath
    Next
    On Error GoTo 0
End Function
