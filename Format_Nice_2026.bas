' Nick Post had these Ideas at one point and made them manualy before AI
' But March 2026 I asked copilot how to help me make them faster and more effecient.
' The reasults are here and they're better

Option Explicit

'========================
' Helper: fast/safe run wrapper
'========================
Private Sub RunWithExcelOptimizations(ByVal doWork As Boolean)
    Static prevCalc As XlCalculation
    Static prevScreen As Boolean
    Static prevEvents As Boolean
    Static prevAlerts As Boolean

    If doWork Then
        prevCalc = Application.Calculation
        prevScreen = Application.ScreenUpdating
        prevEvents = Application.EnableEvents
        prevAlerts = Application.DisplayAlerts

        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.DisplayAlerts = False
        Application.Calculation = xlCalculationManual
    Else
        Application.ScreenUpdating = prevScreen
        Application.EnableEvents = prevEvents
        Application.DisplayAlerts = prevAlerts
        Application.Calculation = prevCalc
    End If
End Sub


'========================
' Formats Spreadsheet
' Keyboard Shortcut Suggestion: Ctrl+Shift+M
'========================
Public Sub Format_Better()

    Dim ws As Worksheet
    Dim topRow As Range
    Dim lastCol As Long

    On Error GoTo CleanFail
    RunWithExcelOptimizations True

    Set ws = ActiveSheet
    
    'Clear borders fast
    ws.Cells.Borders.LineStyle = xlNone

    'Normalize alignment / wrap on all cells (lightweight)
    With ws.Cells
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = False
        .MergeCells = False
    End With

    'Clear fill color
    ws.Cells.Interior.Pattern = xlNone

    'Find last used column in Row 1
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If Application.WorksheetFunction.CountA(ws.Rows(1)) = 0 Then
        'Row 1 empty — nothing to format/filter/freeze meaningfully
        GoTo CleanExit
    End If

    Set topRow = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))

    'Top row formatting
    With topRow
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .Font.Bold = True

        'Apply fill (your gray theme style)
        With .Interior
            .Pattern = xlSolid
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.249977111117893
        End With

        'Filter (toggle off then on to avoid "already has AutoFilter" weirdness)
        If ws.AutoFilterMode Then ws.AutoFilterMode = False
        .AutoFilter
    End With

    'Freeze the Top Row (ALWAYS Row 1)
    With ActiveWindow
        .FreezePanes = False
        .ScrollRow = 1
        .ScrollColumn = 1
    End With

    ws.Activate
    ws.Range("A2").Select   ' <-- Key: select A2 to freeze Row 1
    ActiveWindow.FreezePanes = True

    'AutoFit
    ws.Cells.EntireColumn.AutoFit

CleanExit:
    RunWithExcelOptimizations False
    Exit Sub

CleanFail:
    'Always restore settings
    RunWithExcelOptimizations False
    Err.Raise Err.Number, "Format_Better", Err.Description
End Sub


'========================
' Text000 Macro
' Converts selected column values into 3-digit TEXT (000) when numeric < 1000.
' Leaves values >= 1000 unchanged.
' Works even when original values are text with leading zeros.
' Select One cell not entire column.
'========================
Public Sub Text000_Plant_Format()
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

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    ' Restore Excel state even if error
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Err.Raise Err.Number, "Text000", Err.Description
End Sub

'========================
' Sets selected column(s) to NUMBERS
' Keyboard Shortcut: Ctrl+Shift+N
' Select one or more ENTIRE columns
'========================
Public Sub Col_Numbers()

    Dim ws As Worksheet
    Dim col As Range

    On Error GoTo CleanFail
    RunWithExcelOptimizations True

    Set ws = ActiveSheet
    If TypeName(Selection) <> "Range" Then GoTo CleanExit

    'Loop each column in selection
    For Each col In Selection.Columns
        col.TextToColumns _
            Destination:=col.Cells(1, 1), _
            DataType:=xlDelimited, _
            TextQualifier:=xlTextQualifierDoubleQuote, _
            ConsecutiveDelimiter:=False, _
            Tab:=False, _
            Semicolon:=False, _
            Comma:=False, _
            Space:=False, _
            Other:=False, _
            FieldInfo:=Array(Array(1, xlGeneralFormat))
    Next col

CleanExit:
    RunWithExcelOptimizations False
    Exit Sub

CleanFail:
    RunWithExcelOptimizations False
    Err.Raise Err.Number, "Col_Numbers", Err.Description
End Sub


'========================
' Sets selected range to date format
' Keyboard Shortcut: Ctrl+Shift+D
'========================
Public Sub Col_DATE()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Selection.NumberFormat = "mm/dd/yyyy"
End Sub


'========================
' Used for PIPE Delimited Data (Column A)
'========================
Public Sub PIPE_Delimited_Data()
    Dim ws As Worksheet
    On Error GoTo CleanFail
    RunWithExcelOptimizations True

    Set ws = ActiveSheet

    ws.Columns("A:A").TextToColumns _
        Destination:=ws.Range("A1"), _
        DataType:=xlDelimited, _
        TextQualifier:=xlTextQualifierDoubleQuote, _
        ConsecutiveDelimiter:=False, _
        Tab:=False, _
        Semicolon:=False, _
        Comma:=False, _
        Space:=False, _
        Other:=True, _
        OtherChar:="|", _
        FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True

CleanExit:
    RunWithExcelOptimizations False
    Exit Sub

CleanFail:
    RunWithExcelOptimizations False
    Err.Raise Err.Number, "Text_Col_Bar", Err.Description
End Sub

'========================
' Clears literal "NULL" text from entire sheet
' Replaces with blank
'========================
Public Sub Clear_NULLs()

    Dim ws As Worksheet

    On Error GoTo CleanFail
    RunWithExcelOptimizations True

    Set ws = ActiveSheet

    ws.Cells.Replace _
        What:="NULL", _
        Replacement:="", _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        MatchCase:=False

CleanExit:
    RunWithExcelOptimizations False
    Exit Sub

CleanFail:
    RunWithExcelOptimizations False
    Err.Raise Err.Number, "Clear_NULLs", Err.Description
End Sub


'========================
' Better "Merge Center" (Center Across Selection)
'========================
Public Sub Better_Merge_Center()
    If TypeName(Selection) <> "Range" Then Exit Sub

    With Selection
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlBottom
        .WrapText = False
        .MergeCells = False
        .Font.Bold = True
    End With
End Sub
