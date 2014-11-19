Attribute VB_Name = "CountryFilter"
Sub CountryFilter1()
'
' CountryFilter Macro
'

'
    Sheets("Inputs").Select
    Selection.AutoFilter
    Columns("P:R").Select
    Selection.ClearContents
    Range("K18:M59").Select
    Selection.AutoFilter
    ActiveSheet.Range("$K$18:$M$59").AutoFilter Field:=1, Criteria1:=Range("B7")
    Range("K18").Select
    Selection.CurrentRegion.Select
    Selection.Copy
    ActiveWindow.SmallScroll Down:=-9
    Range("P1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Sheet2").Select
End Sub




