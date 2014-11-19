Attribute VB_Name = "DataCleanse"
Sub SingleCountryDataCleanse()
Attribute SingleCountryDataCleanse.VB_ProcData.VB_Invoke_Func = " \n14"
'
' SingleCountryDataCleanse Macro
'

'
'Clear prior data from spreadsheet
    Sheets("BPCPull").Select
    Range("P19").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Sheets("PreTail").Select
    Cells.Select
    Selection.AutoFilter
    Range("G3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Columns("A:F").Select
    Selection.ClearContents
    Sheets("FilteredDataDump").Select
    Columns("A:T").Select
    Selection.ClearContents
    Range("U3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("V3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("W3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Sheets("DataDump").Select
    Cells.Select
    Selection.ClearContents
    Selection.ClearContents
    Sheets("Calculation").Select
    Columns("A:G").Select
    Selection.ClearContents
    Range("H4:K4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Sheets("CountryResults").Select
    Range(Range("A4"), Range("A4").SpecialCells(xlLastCell)).Select
    Selection.Delete
    Columns("C:C").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete
    Range("A3:C4").Value = "Blank text on purpose - Do not delete!"
    Range("A3:C4").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    
'Copy BPC data to DataDump tab
    Sheets("BPCPull").Select
    Range("G18").Select
    x = Range(Selection, Selection.End(xlDown)).Cells.Count + 17
    Range("P18:W18").Select
    Selection.Copy
    Range("P18:P" & x).Select
    Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Range("G17").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("DataDump").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
'Copy titles to DataDump tab
    Rows("1:1").Select
    Selection.ClearContents
    Application.Goto Reference:="DataDumpTitles"
    Selection.Copy
    Sheets("DataDump").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    CountryCalc
    
End Sub

Sub MultiCountryDataCleanse()
'Clear prior data from spreadsheet
    
    Sheets("PreTail").Select
    Cells.Select
    Selection.AutoFilter
    Range("G3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Columns("A:F").Select
    Selection.ClearContents
    Sheets("FilteredDataDump").Select
    Range("U3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("V3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("W3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Columns("A:T").Select
    Selection.ClearContents
    Sheets("Calculation").Select
    Columns("A:G").Select
    Selection.ClearContents
    Range("H4:K4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    
    CountryCalc


End Sub



