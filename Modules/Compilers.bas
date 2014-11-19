Attribute VB_Name = "Compilers"
Sub CountrySKUCalc()
SingleCountryDataCleanse
CopyResults
        
' Add formatting
    Columns("C:Z").ColumnWidth = 25
    Rows(1).RowHeight = 105
    Rows(2).RowHeight = 40
    Rows(3).RowHeight = 30
    Columns("C:Z").HorizontalAlignment = xlCenter

End Sub


