'Most Users will not need this
'ME2L_Clean is a very specific macro dedicated to a task I had to perform at work, may not be needed by everybody
Sub ME2L_Clean()
    'Sort by F column descending
    Call ApplySort(Range("F1"), xlDescending)

    ActiveSheet.Range("$A:$Z").AutoFilter Field:=6, Criteria1:="<>"
    ActiveSheet.Range("2:" & ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row).Delete Shift:=xlUp
    ActiveSheet.ShowAllData
    Range("I:I,M:M,R:U,Y:Z").ClearContents
    
    'Sort by A column ascending
    Call ApplySort(Range("A:A"), xlAscending)
    
    'Sort by O column ascending
    Call ApplySort(Range("O:O"), xlAscending)

    'Selects the entire Range for copy / paste to ME2L Sheet
    Range("A2:Z2").Select
    Range(Selection, Selection.End(xlDown)).Select
End Sub

'Helper sub to avoid repeated sort code
Sub ApplySort(sortKey As Range, sortOrder As XlSortOrder)
    ActiveSheet.AutoFilter.Sort.SortFields.Clear
    ActiveSheet.AutoFilter.Sort.SortFields.Add Key:=sortKey, SortOn:=xlSortOnValues, Order:=sortOrder, DataOption:=xlSortNormal
    
    With ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub