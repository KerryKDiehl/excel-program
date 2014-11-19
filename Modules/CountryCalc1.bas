Attribute VB_Name = "CountryCalc1"

Sub CountryCalc()


'Filter data - remove input data by filtering out anything that doesn't begin with SKU, filter out Shoe Care, Gardening, Calderea, and Total Product supply/export/dom subsid
    SetFilter
    

'Separate SKU number from SKU description
    Sheets("FilteredDataDump").Select
    Columns("B:B").Select
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("R1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="_", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
    Range("T2").Select
    x = Range(Selection, Selection.End(xlDown)).Cells.Count + 1
    Range("U2").Select
    Selection.Copy
    Range("U2:U" & x).Select
    Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("V2").Select
    Selection.Copy
    Range("V2:V" & x).Select
    Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("W2").Select
    Selection.Copy
    Range("W2:W" & x).Select
    Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
'Adjust PivotTable data range
    Sheets("FilteredDataDump").Select
    Range("A1").Select
    Selection.CurrentRegion.Select
    PivotData = "FilteredDataDump!R1C1:R" & Selection.Rows.Count & "C" & Selection.Columns.Count
    Sheets("PivotTable").Select
    Range("C1").Select
    ActiveSheet.PivotTables("PivotTable3").ChangePivotCache ActiveWorkbook. _
        PivotCaches.Create(SourceType:=xlDatabase, SourceData:=PivotData, _
        Version:=xlPivotTableVersion14)
    ActiveSheet.PivotTables("PivotTable3").PivotCache.Refresh
    
'Copy PivotTable data to PreTail tab
    Columns("A:F").Select
    Selection.Copy
    Sheets("PreTail").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
   
    
'Sort GS smallest to largest and filter SKUs with <1,000 of GS
    Columns("A:F").Select
    ActiveWorkbook.Worksheets("PreTail").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("PreTail").Sort.SortFields.Add Key:=Range( _
        "E2:E15073"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("PreTail").Sort
        .SetRange Range("A1:F15073")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range("A2").Select
    x = Range(Selection, Selection.End(xlDown)).Cells.Count + 1
    Range("G2").Select
    Selection.Copy
    Range("G2:G" & x).Select
    Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveSheet.Range("$A$1:$G$796").AutoFilter Field:=5, Criteria1:=">1000", _
        Operator:=xlAnd
    Range("A:G").Select
        
'Copy data to Calculation tab
    Selection.Copy
    Sheets("Calculation").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("A1").Select
    x = Range(Selection, Selection.End(xlDown)).Cells.Count
    If x > 3 Then
        Range("G3:J3").Select
        Selection.Copy
        Range("G3:G" & x).Select
        Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Sheets("Sheet2").Select
    Else
    End If
End Sub

