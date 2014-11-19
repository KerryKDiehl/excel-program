Attribute VB_Name = "Filters"
Sub SetFilter()

'Filter data - remove input data by filtering out anything that doesn't begin with SKU, filter out Shoe Care, Gardening, Calderea, and Total Product supply/export/dom subsid
    Sheets("DataDump").Select
    Columns("A:Q").Select
    Application.CutCopyMode = False
    Selection.AutoFilter
'Filter data so only SAP SKUs show up (i.e., remove P_ and anything with more than 2 _"'
    ActiveSheet.Range("$A$1:$Q$346686").AutoFilter Field:=2, Criteria1:="=SKU*", Criteria2:="<>*_*_*_*", Operator:=xlAnd
    
'Filter country specific information'
    If Sheets("Inputs").Range("M1").Value <> 0 Then
        Sheets("DataDump").Select
    ActiveSheet.Range("$A$1:$Q$347432").AutoFilter Field:=1, Criteria1:=Range("CountryCode")
    End If
    
'Filter Business lines based on choice made by user'
    If Sheets("Inputs").Range("V1").Value = 1 Then
        ActiveSheet.Range("$A$1:$Q$346686").AutoFilter Field:=11, Criteria1:="<>GBL_221800", Operator:=xlAnd
        ActiveSheet.Range("$A$1:$Q$346686").AutoFilter Field:=16, Criteria1:="<>Shoe Care", Operator:=xlAnd, Criteria2:="<>Gardening", Operator:=xlAnd
        ActiveSheet.Range("$A$1:$Q$346686").AutoFilter Field:=17, Criteria1:="<>Caldrea Business", Operator:=xlAnd, Criteria2:="<>TOTAL PRODUCT SUPPLY/EXPORT/DOM SUBSID.", Operator:=xlAnd
        CopyFilteredData
    ElseIf Sheets("Inputs").Range("V1").Value = 2 Then
        CopyAcquitionData
            
    ElseIf Sheets("Inputs").Range("V1").Value = 3 Then
        CopyFilteredData
    End If
    

    
End Sub

Sub CopyFilteredData()
'Copy filtered data
    Sheets("DataDump").Select
    Columns("A:Q").Select
    Selection.Copy
    Sheets("FilteredDataDump").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

        

End Sub

Sub CopyAcqFilteredData()
'Copy filtered data
    Sheets("DataDump").Select
    Range("A1").Select
    Range(ActiveCell.Offset(1), ActiveCell.End(xlDown)).SpecialCells(12)(1).Select
    Range(Selection, Selection.Offset(0, 16)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("FilteredDataDump").Select
    If Sheets("FilteredDataDump").Range("A2").Value = Empty Then
        Sheets("FilteredDataDump").Range("A2").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    Else
        Range("A1").Select
        x = Range(Selection, Selection.End(xlDown)).Cells.Count + 1
        Range("A" & x).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
    
    Sheets("DataDump").Select
    ActiveSheet.Range("$A$1:$Q$346686").AutoFilter Field:=11
    ActiveSheet.Range("$A$1:$Q$346686").AutoFilter Field:=16
End Sub


Sub CopyAcquitionData()
' Copy column headings to FilteredDataDump tab'
    Sheets("DataDump").Select
    Range("A1:Q1").Select
    Application.CutCopyMode = False
    Selection.Cut
    Sheets("FilteredDataDump").Select
    Range("A1").Select
    ActiveSheet.Paste
'Copy Acquisition data to FilteredDataDump tab'
    Sheets("DataDump").Select
    Range("$K$1:$K$346686").Select
        Set cell1 = Selection.Find(What:="GBL_221800")
    If cell1 Is Nothing Then
    Else
         ActiveSheet.Range("$A$1:$Q$346686").AutoFilter Field:=11, Criteria1:="=GBL_221800", Operator:=xlAnd
         CopyAcqFilteredData
    End If
    
    Range("$P$1:$P$346686").Select
        Set cell2 = Selection.Find(What:="Shoe Care")
    If cell2 Is Nothing Then
    Else
         ActiveSheet.Range("$A$1:$Q$346686").AutoFilter Field:=16, Criteria1:="=Shoe Care", Operator:=xlAnd
         CopyAcqFilteredData
    End If
    
    Range("$P$1:$P$346686").Select
        Set cell3 = Selection.Find(What:="Gardening")
    If cell3 Is Nothing Then
    Else
         ActiveSheet.Range("$A$1:$Q$346686").AutoFilter Field:=16, Criteria1:="=Gardening", Operator:=xlAnd
         CopyAcqFilteredData
    End If
    
    Range("$P$1:$P$346686").Select
        Set cell4 = Selection.Find(What:="Caldrea Business")
    If cell4 Is Nothing Then
    Else
         ActiveSheet.Range("$A$1:$Q$346686").AutoFilter Field:=16, Criteria1:="=Caldrea Business", Operator:=xlAnd
         CopyAcqFilteredData
    End If
     
    If cell1 Is Nothing And cell2 Is Nothing And cell3 Is Nothing And cell4 Is Nothing Then
    MsgBox "This entity does not have any Acquisition SKUs"
    End If
End Sub
