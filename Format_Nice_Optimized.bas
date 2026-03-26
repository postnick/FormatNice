'This is something I use for work but would be useful to many people who use Excel a lot!
Sub Format_Better()
    ' Formats my Spreadsheet
    ' Keyboard Shortcut Suggestion: Ctrl+Shift+M

    'Clears the Borders
        With Cells.Borders
            .Item(xlDiagonalDown).LineStyle = xlNone
            .Item(xlDiagonalUp).LineStyle = xlNone
            .Item(xlEdgeLeft).LineStyle = xlNone
            .Item(xlEdgeTop).LineStyle = xlNone
            .Item(xlEdgeBottom).LineStyle = xlNone
            .Item(xlEdgeRight).LineStyle = xlNone
            .Item(xlInsideVertical).LineStyle = xlNone
            .Item(xlInsideHorizontal).LineStyle = xlNone
        End With

    'Disables Word Wrap
        With Cells
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

    'Clears Color from every cell
        With Cells.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With

    'Sets top Row to Grey and Bold color
        Dim topRow As Range
        Set topRow = Range("A1").Resize(1, Range("A1").End(xlToRight).Column)

    'Sets Top row to Wrap Text
        With topRow
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
        topRow.AutoFilter
        With topRow.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.249977111117893
            .PatternTintAndShade = 0
        End With
        topRow.Font.Bold = True

    'Freezes the Top Row
        With ActiveWindow
            .SplitColumn = 0
            .SplitRow = 1
        End With
        ActiveWindow.FreezePanes = True

    'Auto Fits the Cells
        Cells.EntireColumn.AutoFit
End Sub

' Sets the Column to numbers
Sub Col_Numbers()
    ' Keyboard Shortcut: Ctrl+Shift+N
    Selection.TextToColumns
End Sub

' Sets the Col to Numbers
Sub Col_DATE()
    ' Keyboard Shortcut: Ctrl+Shift+D
    Selection.NumberFormat = "mm/dd/yyyy"
End Sub

'Used for PIPE Delimited Data
Sub Text_Col_Bar()
    ' Part of Ctrl + Shift + T
    Columns("A:A").TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
End Sub

' Text000 Macro Sets the column to TEXT = 000 format
Sub Text000()
    'Declare variables
        Dim colNum As Long
        Dim formulaCell As Range
        Dim lastRow As Long
        Dim formulaRange As Range
        Dim dataRange As Range
    
    'Get the column number before inserting
        colNum = Selection.Column
    
    'Insert a column to the left
        Selection.EntireColumn.Insert Shift:=xlToRight
    
    'Sets the Formula to TEXT(CELL,"000") starting in row 2
        Set formulaCell = Cells(2, colNum)
        formulaCell.FormulaR1C1 = "=IF(RC[1]<999,TEXT(RC[1],""000""),RC[1])"
    
    'Find the last row with data in the column to the right
        lastRow = Cells(Rows.Count, colNum + 1).End(xlUp).Row
    
    'Copy the formula down to all rows
        Set formulaRange = Range(formulaCell, Cells(lastRow, colNum))
        formulaCell.Copy
        formulaRange.PasteSpecial Paste:=xlPasteFormulas
    
    'Copy the formula values over the original data
        formulaRange.Copy
        Range(Cells(2, colNum + 1), Cells(lastRow, colNum + 1)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    'Clear the Formula Column
        Columns(colNum).Delete Shift:=xlToLeft
        
        Application.CutCopyMode = False
End Sub

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

Sub CenterAcross()
    ' Center Highlighted Text across rather than Merge and Center
    ' Keyboard Shortcut: Ctrl+Shift+J
    With Selection
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub

Sub Better_Merge_Center()
    With Selection
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection.Font
        .FontStyle = "Bold"
        .ThemeFont = xlThemeFontMinor
    End With
End Sub
