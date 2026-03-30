'This is something I use for work but would be useful to many people who use Excel a lot!
Sub Format_Better()
    ' Formats my Spreadsheet
    ' Keyboard Shortcut Suggestion: Ctrl+Shift+M

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

    'Disables Word Wrap
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

    'Clears Color from every cell
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


' Text000 Macro Sets the column to TEXT = 000 format
Sub Text000()
    'Declare a space for the cell we will return to
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
    'Clear the Formula Column
        Range(CellLocation).Select
        Selection.EntireColumn.Offset(0, 0).Select
        Selection.Delete Shift:=xlToLeft
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
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
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