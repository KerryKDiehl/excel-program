Attribute VB_Name = "CountrySummary"

Sub CountrySummary1()
Attribute CountrySummary1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' CountrySummary Macro
'

'
'    Clear any prior data
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
    
'    Copy data from BPCPull
    Sheets("BPCPull").Select
    Range("P19").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Sheets("DataDump").Select
    Cells.Select
    Selection.ClearContents
    Selection.ClearContents
    
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
        
'    Count countries in region
    Sheets("Inputs").Select
    Range("Q1").Select
    x = Range(Selection, Selection.End(xlDown)).Cells.Count - 1
    For i = 1 To x
       
'    Calculate each country/region and print on results tab
    Sheets("Inputs").Range("L1").Value = i
    MultiCountryDataCleanse
    CopyResults
    Next i
    
' Add formatting
    Columns("C:Z").ColumnWidth = 25
    Rows(1).RowHeight = 105
    Rows(2).RowHeight = 40
    Rows(3).RowHeight = 30
    Columns("C:Z").HorizontalAlignment = xlCenter
    

    
    
End Sub

