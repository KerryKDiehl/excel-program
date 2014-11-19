Attribute VB_Name = "ResultsSummary"
Sub CopyResults()
'   Copy info to results tab
    Sheets("Calculation").Select
    Range("N1:O6").Select
    Selection.Copy
    Sheets("CountryResults").Select
    Range("A3").Select
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
' Copy Active SKU formula to results tab
    Sheets("Calculation").Select
    Range("O6").Select
    Selection.Copy
    Sheets("CountryResults").Select
    Range("A2").Select
    Selection.End(xlDown).Select
    Selection.Offset(0, 1).Select
    Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
     
'    Print country/cluster SKUs on results tab
    Sheets("Calculation").Select
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("CountryResults").Select
    Range("A3").Select
    Selection.End(xlToRight).Select
    Selection.Offset(0, 1).Select
    Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Inputs").Select
    Range("L3").Select
    Selection.Copy
    Sheets("CountryResults").Select
    Range("A3").Select
    Selection.End(xlToRight).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
'Print cluster GS, DP, and Category on results tab
    Sheets("Calculation").Select
    Range("E1:G1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("CountryResults").Select
    Range("A3").Select
    Selection.End(xlToRight).Select
    Selection.Offset(0, 1).Select
    Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Inputs").Select
    Range("L4:N4").Select
    Selection.Copy
    Sheets("CountryResults").Select
    Range("A3").Select
    Selection.End(xlToRight).Select
    Selection.Offset(0, -2).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

End Sub
