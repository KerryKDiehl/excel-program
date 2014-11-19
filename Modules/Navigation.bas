Attribute VB_Name = "Navigation"

Sub Main()
Attribute Main.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Main Macro

    Sheets("Sheet2").Select
    Range("A15").Select
End Sub
Sub DisplayResults()
Attribute DisplayResults.VB_ProcData.VB_Invoke_Func = " \n14"
'
' DisplayResults Macro

    Sheets("CountryResults").Select
    Range("A1").Select
End Sub
