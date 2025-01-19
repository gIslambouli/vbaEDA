Attribute VB_Name = "Module11"
    
Sub Macro3()

Dim originalSheet As Worksheet
Dim dataSheet As Worksheet

Set originalSheet = ActiveSheet
Dim originalSheetName As String

originalSheetName = originalSheet.Name

'
' Get Dimensions of Data
'

Dim cols As Long
Dim rows As Long

Selection.CurrentRegion.Select
cols = Selection.Columns.Count
rows = Selection.rows.Count - 1

Set DataRange = Range("A2").Resize(rows, cols)

'
' Create Correlation Matrix
'


Selection.CurrentRegion.Select
Application.Run "ATPVBAEN.XLAM!Mcorrel", DataRange, _
"Data Report", "C", False

Set dataSheet = ActiveSheet

Selection.Offset(1, 1).Select
Selection.Copy
Selection.Cells(Selection.rows.Count + 1, Selection.Columns.Count + 1).Select
Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
Selection.Copy
[B2].Select
Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        True, Transpose:=False

Selection.Cells(Selection.rows.Count + 1, Selection.Columns.Count + 1).Select
Selection.CurrentRegion.Select
Selection.Clear
[A1].Select
Selection.CurrentRegion.Select

'
' Label Rows and Columns by Original Column Names
'

originalSheet.Activate
Range(Cells(1, 1), Cells(1, cols)).Select
Selection.Copy
dataSheet.Activate
[B1].Select
ActiveSheet.Paste
Selection.Font.Bold = True
Selection.Copy
ActiveSheet.Cells(2, 1).PasteSpecial Paste:=xlPasteAll, _
Operation:=xlNone, SkipBlanks:=False, Transpose:=True


'
' Make correlation matrix title and add borders
'

ActiveSheet.rows(1).Select
Selection.Insert
Range("A1").Select
ActiveCell.FormulaR1C1 = "Correlation Matrix"
Range("A1").Select
Selection.Font.Bold = True
Range(Cells(1, 1), Cells(1, cols + 1)).Select
Selection.HorizontalAlignment = xlCenterAcrossSelection

Range(Cells(1, 1), Cells(cols + 2, cols + 1)).Select
Selection.Borders.LineStyle = xlContinuous

'
' Color the correlation matrix
'

Selection.FormatConditions.AddColorScale ColorScaleType:=3
Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
Selection.FormatConditions(1).ColorScaleCriteria(1).Type = _
    xlConditionValueLowestValue
With Selection.FormatConditions(1).ColorScaleCriteria(1).FormatColor
    .Color = 13011546
    .TintAndShade = 0
End With
Selection.FormatConditions(1).ColorScaleCriteria(2).Type = _
    xlConditionValuePercentile
Selection.FormatConditions(1).ColorScaleCriteria(2).Value = 50
With Selection.FormatConditions(1).ColorScaleCriteria(2).FormatColor
    .Color = 16776444
    .TintAndShade = 0
End With
Selection.FormatConditions(1).ColorScaleCriteria(3).Type = _
    xlConditionValueHighestValue
With Selection.FormatConditions(1).ColorScaleCriteria(3).FormatColor
    .Color = 7039480
    .TintAndShade = 0
End With

'
' Make Some Space
'

ActiveSheet.Columns(1).Select
Selection.Insert
ActiveSheet.rows(1).Select
Selection.Insert
Selection.Insert

'
' Make the title
'

[A1].Select
Selection.Value = "Data Report For " & originalSheetName
With Selection.Font
    .Bold = True
    .Size = 24
End With

Range(Cells(1, 1), Cells(1, cols + 3)).Select
Selection.HorizontalAlignment = xlCenterAcrossSelection



'
' Set up single variable data table labels
'

Cells(cols + 7, 2).Value = "Missing"
Cells(cols + 7, 3).Value = "Min"
Cells(cols + 7, 4).Value = "Max"
Cells(cols + 7, 5).Value = "Mean"
Cells(cols + 7, 6).Value = "Median"
Cells(cols + 7, 7).Value = "Std Dev"
Range(Cells(cols + 7, 2), Cells(cols + 7, 7)).Font.Bold = True

Range(Cells(5, 2), Cells(5 + cols - 1, 2)).Select
Selection.Copy
Cells(cols + 8, 1).Select
Selection.PasteSpecial xlPasteValues
Selection.Font.Bold = True


'
' Fill in Single Variable data
'

Dim curRow As Long
curRow = cols + 8
For I = 1 To cols
    Cells(curRow, 2).Value = Application.WorksheetFunction.CountBlank(DataRange.Columns(I))
    Cells(curRow, 3).Value = Application.WorksheetFunction.Min(DataRange.Columns(I))
    Cells(curRow, 4).Value = Application.WorksheetFunction.Max(DataRange.Columns(I))
    Cells(curRow, 5).Value = Application.WorksheetFunction.Average(DataRange.Columns(I))
    Cells(curRow, 6).Value = Application.WorksheetFunction.Median(DataRange.Columns(I))
    Cells(curRow, 7).Value = Application.WorksheetFunction.StDev(DataRange.Columns(I))
    curRow = curRow + 1
    Next I

'
' Format single variable table data
'

Cells(cols + 6, 1).Select
Selection.Value = "Single Variable Data"
Selection.Font.Bold = True
Range(Cells(cols + 6, 1), Cells(cols + 6, 7)).Select
Selection.HorizontalAlignment = xlCenterAcrossSelection
Selection.CurrentRegion.Select
Selection.Borders.LineStyle = xlContinuous
Range(Cells(cols + 6, 1), Cells(2 * cols + 7, 1)).Select
Selection.Insert

'
' Highlight columns with nonzero amounts of missing data
'

Range(Cells(cols + 8, 3), Cells(2 * cols + 7, 3)).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
'
' Color the rest of the cells a gradient
'

For I = 1 To cols - 1

    Range(Cells(cols + 8, I + 3), Cells(2 * cols + 7, I + 3)).Select
    Selection.FormatConditions.AddColorScale ColorScaleType:=3
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).ColorScaleCriteria(1).Type = _
        xlConditionValueLowestValue
    With Selection.FormatConditions(1).ColorScaleCriteria(1).FormatColor
        .Color = 13011546
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).ColorScaleCriteria(2).Type = _
        xlConditionValuePercentile
    Selection.FormatConditions(1).ColorScaleCriteria(2).Value = 50
    With Selection.FormatConditions(1).ColorScaleCriteria(2).FormatColor
        .Color = 16776444
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).ColorScaleCriteria(3).Type = _
        xlConditionValueHighestValue
    With Selection.FormatConditions(1).ColorScaleCriteria(3).FormatColor
        .Color = 7039480
        .TintAndShade = 0
    End With
        
Next I
    


End Sub

