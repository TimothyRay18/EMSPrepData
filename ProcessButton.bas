Attribute VB_Name = "ProcessButton"
Function getMaxRow(col As Integer) As Double
    getMaxRow = ActiveSheet.Cells(Rows.Count, col).End(xlUp).row
End Function

Function getMaxCol(row As Integer) As Double
    getMaxCol = ActiveSheet.Cells(row, Columns.Count).End(xlToLeft).Column
End Function

Function GetFilenameFromPath(ByVal strPath As String) As String
    If Right$(strPath, 1) <> "\" And Len(strPath) > 0 Then
        GetFilenameFromPath = GetFilenameFromPath(Left$(strPath, Len(strPath) - 1)) + Right$(strPath, 1)
    End If
End Function

Function findCellInColumn(row As Integer, str As String) As Double
    Dim i As Double
    i = 1
    Dim m As Double
    m = getMaxCol(row)
    While LCase(ActiveSheet.Cells(row, i).Value) <> LCase(str) And i <= m
        i = i + 1
    Wend
    findCellInColumn = i
End Function

Sub Process()
Attribute Process.VB_ProcData.VB_Invoke_Func = " \n14"
'
'
'

'
    OpenAndClean
    Pivot
    
End Sub
Sub Process2()
    AfterRearrangeDate
    BomProcess
End Sub

Sub Process3()
    SEMBSOH
    FLEXSOH
End Sub

Sub OpenAndClean()
    Dim so_file As String
    so_file = Range("B1").Value
    Workbooks.OpenText Filename:=so_file _
        , Origin:=xlWindows, StartRow:=1, DataType:=xlDelimited, TextQualifier _
        :=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:= _
        False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array _
        (1, 1), Array(2, 1), Array(3, 1), Array(4, 2), Array(5, 1), Array(6, 1), Array(7, 1), Array(8 _
        , 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), _
        Array(15, 1), Array(16, 1), Array(17, 1), Array(18, 1), Array(19, 1), Array(20, 1)), _
        TrailingMinusNumbers:=True
    Columns("E:E").Select
    Selection.Delete Shift:=xlToLeft
    Columns("G:G").Select
    Selection.Delete Shift:=xlToLeft
    Columns("H:H").Select
    Selection.Delete Shift:=xlToLeft
    Columns("I:I").Select
    Selection.Delete Shift:=xlToLeft
    Columns("K:K").Select
    Selection.Delete Shift:=xlToLeft
    Columns("L:L").Select
    Selection.Delete Shift:=xlToLeft
    Rows("10:10").Select
    Selection.Delete Shift:=xlUp
    Range("G9").Select
    Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveCell.FormulaR1C1 = "Week"
    
    Columns("F:F").Select
    Selection.TextToColumns Destination:=Range("F1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 4), TrailingMinusNumbers:=True
    Columns("F:F").EntireColumn.AutoFit
    Selection.NumberFormat = "[$-en-US]d-mmm-yy;@"
    Columns("E:E").Select
    Selection.TextToColumns Destination:=Range("E1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 4), TrailingMinusNumbers:=True
    Columns("E:E").EntireColumn.AutoFit
    Selection.NumberFormat = "[$-en-US]d-mmm-yy;@"
    Range("G10").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=YEAR(RC[-1])&""-""&WEEKNUM(RC[-1])"
    Range("G10").Select
    
    Dim max_row As String
    max_row = getMaxRow(3)
    Selection.AutoFill Destination:=Range("G10:G" + CStr(max_row))
End Sub

Sub Pivot()
    Dim ThisFileName As String
    Dim BaseFileName As String
    
    Dim FileNameArray() As String
    
    ThisFileName = ActiveWorkbook.Name
    FileNameArray = Split(ThisFileName, ".")
    BaseFileName = FileNameArray(0)
    
    Dim max_row As String
    max_row = getMaxRow(3)
    
    Dim source As String
    source = BaseFileName + "!R9C3:R" + CStr(max_row) + "C13"
    Sheets.Add.Name = "Sheet1"
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        source, Version:=6).CreatePivotTable TableDestination _
        :="Sheet1!R3C1", TableName:="PivotTable1", DefaultVersion:=6
    Sheets("Sheet1").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("PivotTable1")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable1").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Week")
        .Orientation = xlColumnField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Material")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("    Recpt/reqd"), "Sum of     Recpt/reqd", xlSum
        
    MsgBox "Please check week order and rearrange manually. When finished, please click the process 2 button"
End Sub
Sub AfterRearrangeDate()
Attribute AfterRearrangeDate.VB_ProcData.VB_Invoke_Func = " \n14"
'
'
'

'
    Workbooks(GetFilenameFromPath(Range("B1").Value)).Activate
    Dim so As String
    so = ActiveWorkbook.Name
    Dim max_col As Double
    max_col = getMaxCol(4)
    
    Dim weekNum As Double
    weekNum = CInt(Format(Date, "ww", 2))
    
    Dim weekNow As Double
    weekNow = findCellInColumn(4, CStr(Year(Now)) + "-" + CStr(weekNum - 1))
    
    Cells(4, max_col + 1).Select
    ActiveCell.FormulaR1C1 = "Back order"
    
    Cells(5, max_col + 1).Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-" + CStr(max_col - 1) + "]:RC[-" + CStr(max_col + 2 - weekNow) + "])"
    Range("CG5").Select
    
    max_row = getMaxRow(1)
    Selection.AutoFill Destination:=Range("CG5:CG" + CStr(max_row))
    
    max_row = getMaxRow(1)
    
    For i = 0 To 7 Step 1
        Cells(4, max_col + 2 + i).Value = "W" + CStr(weekNum + i - 1)
        Cells(5, max_col + 2 + i).FormulaR1C1 = "=SUM(RC[-" + CStr(max_col + 2 - weekNow) + "])"
        Cells(5, max_col + 2 + i).Select
        Selection.AutoFill Destination:=Range(Cells(5, max_col + 2 + i), Cells(max_row, max_col + 2 + i))
    Next
    
    ThisWorkbook.Activate
    Dim wb As Workbook
    Set wb = Workbooks.Open(Range("B7").Value)
    wb.Worksheets("MPP").Activate
    Columns("E:E").Select
    Selection.TextToColumns Destination:=Range("E1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 2), TrailingMinusNumbers:=True
    Dim mpp  As String
    mpp = ActiveWorkbook.Name
    Dim col_m As Double
    Dim col_qman As Double
    col_m = findCellInColumn(4, "Material")
'    col_qman = findCellInColumn(4, "Qman Family")
    
    i = 6
    Dim m As Double
    m = getMaxCol(4)
    While LCase(ActiveSheet.Cells(4, i).Value) <> LCase("Qman Family") And i <= m
        i = i + 1
    Wend
    col_qman = i
    
    Workbooks(so).Activate
    Cells(4, max_col + 10).Select
    ActiveCell.FormulaR1C1 = "Product Family"
    Cells(5, max_col + 10).Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-" + CStr(max_col + 9) + "],'[" + mpp + "]MPP'!C" + CStr(col_m) + ":C" + CStr(col_qman) + "," + CStr(col_qman - col_m + 1) + ",0)"
    Cells(5, max_col + 10).Select
    Selection.AutoFill Destination:=Range(Cells(5, max_col + 10), Cells(max_row - 1, max_col + 10))
    
'    Range("CG4:CP1180").Select
    Range(Cells(4, max_col + 1), Cells(getMaxRow(1) - 1, max_col + 10)).Select
'    max_row = getMaxRow(1)
    
    Sheets.Add.Name = "Sheet2"
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Sheet1!R4C" + CStr(max_col + 1) + ":R" + CStr(max_row - 1) + "C" + CStr(max_col + 10), Version:=7).CreatePivotTable TableDestination:= _
        "Sheet2!R3C1", TableName:="PivotTable3", DefaultVersion:=7
    Sheets("Sheet2").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("PivotTable3")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("PivotTable3").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable3").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("Product Family")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable3").AddDataField ActiveSheet.PivotTables( _
        "PivotTable3").PivotFields("Back order"), "Sum of Back order", xlSum
    
    For i = 0 To 7 Step 1
        ActiveSheet.PivotTables("PivotTable3").AddDataField ActiveSheet.PivotTables( _
            "PivotTable3").PivotFields("W" + CStr(weekNum - 1 + i)), "Sum of W" + CStr(weekNum - 1 + i), xlSum
    Next
End Sub
Sub BomProcess()
    ThisWorkbook.Activate
    Dim bom As String
    bom = GetFilenameFromPath(Range("B11").Value)
    Dim mpp As String
    mpp = GetFilenameFromPath(Range("B7").Value)
    Dim so As String
    so = GetFilenameFromPath(Range("B1").Value)
    Workbooks.OpenText Filename:= _
        Range("B11").Value, Origin:=xlWindows, _
        StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False _
        , Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 2), _
        Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), _
        Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15, 1), Array( _
        16, 1), Array(17, 1), Array(18, 1), Array(19, 1), Array(20, 1), Array(21, 1), Array(22, 1), _
        Array(23, 1), Array(24, 1), Array(25, 1), Array(26, 1), Array(27, 1), Array(28, 1), Array( _
        29, 1), Array(30, 1), Array(31, 1), Array(32, 1), Array(33, 1), Array(34, 1), Array(35, 1), _
        Array(36, 1), Array(37, 1), Array(38, 1), Array(39, 1), Array(40, 1), Array(41, 1), Array( _
        42, 1), Array(43, 1)), TrailingMinusNumbers:=True
    Rows("1:5").Select
    Selection.Delete Shift:=xlUp
    Rows("3:3").Select
    Selection.Delete Shift:=xlUp
    Rows("4:4").Select
    Selection.Delete Shift:=xlUp
    Columns("C:C").Select
    Selection.Delete Shift:=xlToLeft
    Columns("P:P").Select
    Selection.Delete Shift:=xlToLeft
    Columns("Q:R").Select
    Selection.Delete Shift:=xlToLeft
    Columns("AF:AF").Select
    Selection.Delete Shift:=xlToLeft
    Columns("AG:AG").Select
    Selection.Delete Shift:=xlToLeft
    
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "Product Family"
    Columns("C:D").Select
    Selection.EntireColumn.Hidden = True
    Columns("H:AC").Select
    Selection.EntireColumn.Hidden = True
    Columns("AI:AJ").Select
    Selection.Delete Shift:=xlToLeft
    
    Workbooks(mpp).Activate
    Dim max_col_mpp As Double
    max_col_mpp = getMaxCol(4)
    Dim m_col_mpp As Double
    m_col_mpp = findCellInColumn(4, "Material")
    
    Workbooks(bom).Activate
    Range("A4").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[1],'[" + mpp + "]MPP'!C" + CStr(m_col_mpp) + ":C" + CStr(max_col_mpp) + "," + CStr(max_col_mpp - m_col_mpp + 1) + ",0)"
    Range("A4").Select
    Selection.AutoFill Destination:=Range("A4:A" + CStr(getMaxRow(2)))
    
    Columns("AH:AH").Select
    Selection.Delete Shift:=xlToLeft
    
    Workbooks(so).Activate
    Sheets("Sheet1").Select
    Range(Cells(4, findCellInColumn(4, "Back order")), Cells(4, findCellInColumn(4, "Product Family") - 1)).Select
    Selection.Copy
    Windows(bom).Activate
    Cells(3, getMaxCol(3) + 1).Select
    ActiveSheet.Paste
    Range("AH5").Select
    Selection.AutoFilter
    
    Dim bo_col As Double
    bo_col = findCellInColumn(4, "Back Order")
    
    Cells(4, bo_col).Select
    
    Workbooks(so).Activate
    Dim max_col_so As Double
    max_col_so = getMaxCol(4)
    Workbooks(bom).Activate
    
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC2,'[" + so + "]Sheet1'!C1:C" + CStr(max_col_so - 1) + ",R1C,),)"
    Cells(1, bo_col).Select
    ActiveCell.FormulaR1C1 = "85"
    Cells(1, bo_col + 1).Select
    ActiveCell.FormulaR1C1 = "86"
    Range(Cells(1, bo_col), Cells(1, bo_col + 1)).Select
    Selection.AutoFill Destination:=Range(Cells(1, bo_col), Cells(1, bo_col + 8)), Type:=xlFillDefault
    Range("AI1:AQ1").Select
    Cells(4, bo_col).Select
    Selection.AutoFill Destination:=Range(Cells(4, bo_col), Cells(4, bo_col + 8)), Type:=xlFillDefault
    Range(Cells(4, bo_col), Cells(4, bo_col + 8)).Select
    Selection.AutoFill Destination:=Range("AI4:AQ" + CStr(getMaxRow(1)))
    ActiveWorkbook.BreakLink Name:= _
        mpp, _
        Type:=xlExcelLinks
    ActiveWorkbook.BreakLink Name:= _
        so, Type:=xlExcelLinks
        
    Sheets.Add.Name = "SEMB SOH"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "Site"
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "Part Number"
    
    Dim weekNum As Double
    weekNum = CInt(Format(Date, "ww", 2))
    
    Range("C3").Select
    ActiveCell.FormulaR1C1 = "SOH W" + CStr(weekNum - 1)
    
    ThisWorkbook.Activate
    Set wb = Workbooks.Open(Range("B19").Value)
    Dim oh As String
    oh = ActiveWorkbook.Name
    Dim oh_gt_col As Double
    oh_gt_col = findCellInColumn(4, "Grand Total")
    
    ThisWorkbook.Activate
    
    Set wb = Workbooks.Open(Range("B15").Value)
    Dim priority As String
    priority = ActiveWorkbook.Name
    
    Dim Sh As Worksheet
    Dim sh_name
    For Each Sh In Worksheets
        If UCase(Left(Sh.Name, 2)) = "WK" Then sh_name = Sh.Name
    Next Sh
    Worksheets(sh_name).Activate
    
    Dim max_col_prior As Double
    max_col_prior = getMaxCol(2)
    
    Cells(2, max_col_prior).Value = "SO W" + CStr(weekNum - 1)
    Cells(2, max_col_prior + 1).Value = "X"
    Cells(2, max_col_prior + 2).Value = "X"
    
    Workbooks(priority).Activate
    Cells(3, max_col_prior + 1).Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC2,'[" + oh + "]PivotTable'!C1:C" + CStr(oh_gt_col) + "," + CStr(oh_gt_col) + ",0),0)"
    Cells(3, max_col_prior + 1).Select
    
    Dim max_row_prior As Double
    max_row_prior = getMaxRow(1)
    
    Selection.AutoFill Destination:=Range(Cells(3, max_col_prior + 1), Cells(max_row_prior, max_col_prior + 1))
    
    Cells(3, max_col_prior + 2).Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC3,'[" + oh + "]PivotTable'!C1:C" + CStr(oh_gt_col) + "," + CStr(oh_gt_col) + ",0),0)"
    Cells(3, max_col_prior + 2).Select
    Selection.AutoFill Destination:=Range(Cells(3, max_col_prior + 2), Cells(max_row_prior, max_col_prior + 2))
    
    Cells(3, max_col_prior).Select
    ActiveCell.Formula2R1C1 = "=RC[1]+RC[2]"
    Cells(3, max_col_prior).Select
    Selection.AutoFill Destination:=Range(Cells(3, max_col_prior), Cells(max_row_prior, max_col_prior))
    
    Range("B3:B" + CStr(getMaxRow(1))).Select
    Selection.Copy
    Workbooks(bom).Activate
    Range("B4").Select
    ActiveSheet.Paste
    Workbooks(priority).Activate
    Cells(3, max_col_prior + 1).Select
    Range(Cells(3, max_col_prior + 1), Cells(getMaxRow(1), max_col_prior + 1)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Workbooks(bom).Activate
    Range("C4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Workbooks(priority).Activate
    MsgBox "Please untick EOL, blank, Urgent, and Check sama downstream in Follow Up Material Column. Then click process 3"
End Sub

Sub SEMBSOH()
    ThisWorkbook.Activate
    Dim bom As String
    bom = GetFilenameFromPath(Range("B11").Value)
    Dim mpp As String
    mpp = GetFilenameFromPath(Range("B7").Value)
    Dim so As String
    so = GetFilenameFromPath(Range("B1").Value)
    Dim priority As String
    priority = GetFilenameFromPath(Range("B15").Value)
    
    Workbooks(priority).Activate
    Range("C3:C" + CStr(getMaxRow(2))).Select
    Selection.Copy
    Workbooks(bom).Activate
    Worksheets("SEMB SOH").Activate
    Range("B" + CStr(getMaxRow(2) + 1)).Select
    ActiveSheet.Paste
    
    Workbooks(priority).Activate
    Range(Cells(3, getMaxCol(2)), Cells(getMaxRow(getMaxCol(2)), getMaxCol(2))).Select
    Selection.Copy
    Workbooks(bom).Activate
    Worksheets("SEMB SOH").Activate
    Range("C" + CStr(getMaxRow(3) + 1)).Select
    ActiveSheet.Paste
    
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "SEMB"
    Range("A4").Select
    Selection.AutoFill Destination:=Range("A4:A" + CStr(getMaxRow(2)))
    
    
End Sub
Sub FLEXSOH()
    ThisWorkbook.Activate
    Dim bom As String
    bom = GetFilenameFromPath(Range("B11").Value)
    Dim mpp As String
    mpp = GetFilenameFromPath(Range("B7").Value)
    Dim so As String
    so = GetFilenameFromPath(Range("B1").Value)
    Dim priority As String
    priority = GetFilenameFromPath(Range("B15").Value)
    Dim req As String
    req = GetFilenameFromPath(Range("B23").Value)
    
    Workbooks(bom).Activate
    
    Sheets.Add.Name = "FLEX SOH + BCD"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "Site"
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "Part Number"
    
    Dim weekNum As Double
    weekNum = CInt(Format(Date, "ww", 2))
    
    For i = 0 To 7 Step 1
        Cells(3, i + 3).Value = "W" + CStr(weekNum + i)
    Next
    
    Range("A3").Select
    Selection.AutoFilter
    Range("C4").Select
    ActiveWindow.FreezePanes = True
    
    ThisWorkbook.Activate
    Dim wb As Workbook
    Set wb = Workbooks.Open(Range("B23").Value)
    ActiveSheet.Range("$A$7:$AA$" + CStr(getMaxRow(1))).AutoFilter Field:=7, Criteria1:= _
        "SE Manufacturing Batam-PEL"
    Range("A8").Select
    Range(Cells(8, 1), Cells(getMaxRow(1), 1)).Select
    Selection.Copy
    Workbooks(bom).Activate
    Range("B4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Workbooks(req).Activate
    Cells(8, getMaxCol(7) - 8).Select
    Range(Cells(8, getMaxCol(7) - 8), Cells(getMaxRow(1), getMaxCol(7) - 1)).Select
    Selection.Copy
    Workbooks(bom).Activate
    Range("C4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "Flex Batam"
    Range("A4").Select
    Selection.AutoFill Destination:=Range("A4:A" + CStr(getMaxRow(2)))
    
    Sheets.Add.Name = "SOH + BCD Combine"
    Range("A3").Select
    Sheets("FLEX SOH + BCD").Select
    Range("A3:J3").Select
    Selection.Copy
    Sheets("SOH + BCD Combine").Select
    ActiveSheet.Paste
    Range("B3").Select
    Selection.AutoFilter
    
    Sheets("SEMB SOH").Select
    Range("A4:C" + CStr(getMaxRow(1))).Select
    Selection.Copy
    Sheets("SOH + BCD Combine").Select
    Range("A4").Select
    ActiveSheet.Paste
    
    Sheets("FLEX SOH + BCD").Select
    Range("A4:J" + CStr(getMaxRow(1))).Select
    Selection.Copy
    Sheets("SOH + BCD Combine").Select
    Range("A" + CStr(getMaxRow(1) + 1)).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    
End Sub
Sub LastPivot()
    Dim max_row_com As Double
    max_row_com = getMaxRow(1)
'    ActiveSheet.Paste
    Dim source As String
    source = "SOH + BCD Combine!R3C1:R" + CStr(max_row_com) + "C10"

    Range("A3:J" + CStr(getMaxRow(1))).Select
    Sheets.Add.Name = "SOH + BCD Pivot"
    
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("PivotTable4")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("PivotTable4").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable4").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTable4").PivotFields("Part Number")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable4").AddDataField ActiveSheet.PivotTables( _
        "PivotTable4").PivotFields("W34"), "Sum of W34", xlSum
    ActiveSheet.PivotTables("PivotTable4").AddDataField ActiveSheet.PivotTables( _
        "PivotTable4").PivotFields("W35"), "Sum of W35", xlSum
    ActiveSheet.PivotTables("PivotTable4").AddDataField ActiveSheet.PivotTables( _
        "PivotTable4").PivotFields("W36"), "Sum of W36", xlSum
    ActiveSheet.PivotTables("PivotTable4").AddDataField ActiveSheet.PivotTables( _
        "PivotTable4").PivotFields("W37"), "Sum of W37", xlSum
    ActiveSheet.PivotTables("PivotTable4").AddDataField ActiveSheet.PivotTables( _
        "PivotTable4").PivotFields("W38"), "Sum of W38", xlSum
    ActiveSheet.PivotTables("PivotTable4").AddDataField ActiveSheet.PivotTables( _
        "PivotTable4").PivotFields("W39"), "Sum of W39", xlSum
    ActiveSheet.PivotTables("PivotTable4").AddDataField ActiveSheet.PivotTables( _
        "PivotTable4").PivotFields("W40"), "Sum of W40", xlSum
    ActiveSheet.PivotTables("PivotTable4").AddDataField ActiveSheet.PivotTables( _
        "PivotTable4").PivotFields("W41"), "Sum of W41", xlSum
    Sheets("Sheet38").Select
    Sheets("Sheet38").Move After:=Sheets(5)
    Sheets("Sheet38").Select
    Sheets("Sheet38").Name = "SOH + BCD Pivot"
    Range("J13").Select
End Sub
