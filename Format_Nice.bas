Attribute VB_Name = "Format_Nice"
Sub Format_Better()
Attribute Format_Better.VB_Description = "Formats my Spredsheet"
Attribute Format_Better.VB_ProcData.VB_Invoke_Func = "M\n14"
'
' Format_Better Macro
' Formats my Spredsheet
'
' Keyboard Shortcut: Ctrl+Shift+M
'

'Clears the Borders
    Cells.Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
 'Dissables Word Wrap
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
'Clears Color
    Cells.Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
'Sets top Row to Grey and Bold color
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
'Sets Top row to Wrap Text
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
'Sets the Filter
    Selection.AutoFilter
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    Selection.Font.Bold = True
     
'Freezes the Top Row
    Range("A1").Select
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    
'Auto Fits the Cells
    Cells.Select
    Cells.EntireColumn.AutoFit
    
End Sub

Sub Col_Numbers()
Attribute Col_Numbers.VB_ProcData.VB_Invoke_Func = "N\n14"
'
' Col_Numbers Macro
' Sets the Column to numbers
'
' Keyboard Shortcut: Ctrl+Shift+N
'
        
        Selection.TextToColumns
    
End Sub

Sub Col_DATE()
Attribute Col_DATE.VB_ProcData.VB_Invoke_Func = "D\n14"
'
' Col_DATE Macro
' Sets the Col to Numbers
'
' Keyboard Shortcut: Ctrl+Shift+D
'
    Selection.NumberFormat = "mm/dd/yyyy"
    
End Sub

Sub Text_Col_Bar()
Attribute Text_Col_Bar.VB_ProcData.VB_Invoke_Func = "T\n14"
' Part of Ctrl + Shift + T
' Text_Col_Bar Macro
'
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
End Sub


Sub ME2L_Clean()
Attribute ME2L_Clean.VB_ProcData.VB_Invoke_Func = "R\n14"
'
' ME2L_Clean Macro
' Cleanup the ME2l
' Working
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.AutoFilter.Sort.SortFields.Add Key:=Range _
        ("F1"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveSheet.Range("$A:$Z").AutoFilter Field:=6, Criteria1:="<>"
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    ActiveSheet.ShowAllData
    Range("I:I,M:M,R:U,Y:Z").Select
    Selection.ClearContents
    ActiveSheet.AutoFilter.Sort.SortFields.Clear
        ActiveSheet.AutoFilter.Sort.SortFields.Add Key:=Range _
        ("A:A"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveSheet.AutoFilter.Sort.SortFields.Clear
    ActiveSheet.AutoFilter.Sort.SortFields.Add Key:=Range _
        ("O:O"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
'Selects the entire Range for copy / paste to ME2l Sheet
    Range("A2:Z2").Select
    Range(Selection, Selection.End(xlDown)).Select
    
End Sub


Sub Text000()
Attribute Text000.VB_ProcData.VB_Invoke_Func = "L\n14"
'
' Text000 Macro Sets the column to TEXT = 000 format

'Declare a space for the cell we will reutrn too
    Dim CellLocation As String
    
'Selects a column and inserts a column to the left
    Selection.EntireColumn.Offset(0, 0).Select
    Selection.Insert Shift:=xlToRight
    ActiveCell.Offset(1, 0).Select
'Sets the Formula to TEXT(CELL,"000")
    ActiveCell.FormulaR1C1 = "=IF(RC[1]<999,TEXT(RC[1],""000""),RC[1])"
    
    
'Saving the Active Cell Location to return later a few more times
    CellLocation = ActiveCell.Address
    
    
'Recall the Cell Location Saved
    Range(CellLocation).Select
'Copy the Cell
    Selection.Copy
'Moves active sell one Right and moves to bottom of selected Row
    ActiveCell.Offset(0, 1).Select
    Selection.End(xlDown).Select
'Moves active sell one Left and pastes the formula the whole range
    ActiveCell.Offset(0, -1).Select
    ActiveSheet.Paste
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste

'Copy the full values over the OG Data
'Recall the Cell Location Saved
    Range(CellLocation).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
'Move the Cell one Right
    ActiveCell.Offset(0, 1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

'Clear the Formula COlumn
    Range(CellLocation).Select
    Selection.EntireColumn.Offset(0, 0).Select
    Selection.Delete Shift:=xlToLeft
    
End Sub

