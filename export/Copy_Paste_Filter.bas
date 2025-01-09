Attribute VB_Name = "Copy_Paste_Filter"
Sub FilteredPasteValuesInPlace()
'   15/08/2012
'   Ron Campbell


'   When using a filtered excel worksheet if you have formulas and wish to do
'   copy - paste values you must un-filter the worksheet first or the data will not go where you intend it to
'   This macro will go cell by cell through a selected range and replace the formula with the cell value
'   There is no outward sign of operation of this macro.

' Copy and paste values within a filtered range
' within the range go to each cell and paste its value
' Started with
'        cell.Copy
'        cell.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=True, Transpose:=False
'   Each step updated screen
'   found cell.formula = cell.value
If ActiveWorkbook Is Nothing Then
    Exit Sub
End If

Dim fixDates As Boolean
Dim SetText As Boolean
Dim response As Integer

On Error GoTo errorHandler
'
'
runOption = GetDataType
If runOption = 0 Then
    Exit Sub
Else
    Select Case runOption
    Case -1: Exit Sub
    Case 1: 'date
            SetText = False
            fixDates = True
    Case 2: 'numeric
            SetText = False
            fixDates = False
    Case 3: 'text
            SetText = True
            fixDates = False
    Case Else: Exit Sub
    End Select
End If

SuspendAutoCalc

Dim currCell As Range
    If Selection.Count > 1 Then
        Selection.SpecialCells(xlCellTypeVisible).Select
    End If
    For Each currCell In Selection.Cells
' VBA inherently treads all dates as mm/dd/yyyy ignoring local settings.
' Added this section when I found that dates stored as text got clobbered
        
        If IsDate(currCell.value) And fixDates Then
            currCell.NumberFormat = "@"
            currCell.Formula = Format(currCell.value, "mm/dd/yy")
            currCell.NumberFormat = "dd/mm/yyyy"
        ElseIf SetText Then
            currCell.NumberFormat = "@"
        Else: currCell.NumberFormat = "General"
        End If
            currCell.Formula = Trim(currCell.value)
    Next
    'restore the user's calculation setting
    ResumeAutoCalc
    Exit Sub
ErrorExit:
    'restore the user's calculation setting
    ResumeAutoCalc
    
    'the ErrorHandler code should only be executed if there is an error
    Exit Sub
errorHandler:
    Debug.Print Err.Number & vbLf & Err.Description
    Resume ErrorExit
        
End Sub

Sub TrimInPlace()
'   2/05/2013
'   Ron Campbell

'Highlight a section of cells
'step through cell by cell trimming the contents
'included in removal of HTML non-breaking space ascii 160 as this was showing up in some supplier data
'included removal of double spaces
'
If ActiveWorkbook Is Nothing Then
    Exit Sub
End If

On Error GoTo errorHandler

SuspendAutoCalc
Dim currCell As Range
Dim NACount As Long
NACount = 0

    If Selection.Count > 1 Then
        Selection.SpecialCells(xlCellTypeVisible).Select
    End If
    For Each currCell In Selection.Cells
        If Not Application.IsNA(currCell.value) Then
            currCell.Formula = Application.WorksheetFunction.Trim(Replace(Replace(Replace(Replace(Replace(Replace(Trim(currCell.value), Chr(9), " "), Chr(160), ""), Chr(10), ""), Chr(13), ""), ChrW(8195), " "), "  ", " "))
            'currCell.value = Trim(Replace(Replace(Replace(Replace(Trim(currCell.value), Chr(160), ""), Chr(10), ""), Chr(13), ""), "  ", " "))
        Else
            NACount = NACount + 1
        End If
    Next
    'restore the user's calculation setting
    ResumeAutoCalc
    If NACount > 0 Then
        MsgBox "Your data contains " & NACount & " cells with #N/A"
    End If
    Exit Sub
    
ErrorExit:
    
    'restore the user's calculation setting
    ResumeAutoCalc
    'the ErrorHandler code should only be executed if there is an error
    Exit Sub
errorHandler:
        Debug.Print Err.Number & vbLf & Err.Description
        Resume ErrorExit
        

End Sub


Sub FilteredPasteValues()
'   New version can copy multiple columns at a time
'   24/09/2014 - Ron Campbell for HPV
'   select only visible cells
'   Start at nominated destination starting cell then incrementally move down through visible rows populating with the source
'   Can copy from non-filtered to filtered or from filtered to filtered range

If ActiveWorkbook Is Nothing Then
    Exit Sub
End If

On Error GoTo ErrorExit

    Dim c As Range
    Dim RowCount As Long
    Dim ColCount As Integer
    Dim y As Long
    Dim xOffset As Integer
    Dim yOffset As Long
    Dim xStart As Integer
    Dim applyTextFormat As Integer
    Dim prevCol As Long
    datasrc = Selection.Address
    If Range(datasrc).Count > 1 Then
        Range(datasrc).SpecialCells(xlCellTypeVisible).Select
    Else
        Range(datasrc).Select
    End If
    RowCount = Selection.SpecialCells(xlVisible).Rows.Count
    ColCount = Selection.SpecialCells(xlVisible).Columns.Count
    
    dataDest = Application.InputBox("Select range to copy selected cells to", Type:=8).Address
    applyTextFormat = MsgBox("Apply Text formatting to destination cells?", vbYesNo, "Text format")
    xStart = Selection.Column
    yStart = Selection.Row

    y = 0
    
SuspendAutoCalc
    For Each c In Selection.Cells
            If c.Column = xStart And c.Row <> yStart Then
                y = y + 1
            End If
            If Range(dataDest).Offset(y, 0).RowHeight = 0 Then
                Do
                    y = y + 1
                Loop While Range(dataDest).Offset(y, 0).RowHeight = 0
            End If
                Application.StatusBar = "Processing " & String((19 * (y / RowCount)) Mod 20, ".")
                With Range(dataDest).Offset(y, c.Column - xStart)
                    If applyTextFormat = vbYes Then
                        .NumberFormat = "@"
                    End If
                    .value = Trim(c.value)
                    '.Interior.ColorIndex = 6
                End With
    Next c
    'leave selected cell as start of destination for copy-paste
    Range(dataDest).Select
 'restore the user's calculation setting
    ResumeAutoCalc
    Exit Sub
ErrorExit:
    
   'restore the user's calculation setting
    ResumeAutoCalc
    'the ErrorHandler code should only be executed if there is an error
    Exit Sub
errorHandler:
        Debug.Print Err.Number & vbLf & Err.Description
        Resume ErrorExit
        

End Sub


Sub FilterSelectValue()
FilterSelectedValue False
End Sub
Sub FilterHideValue()
FilterSelectedValue True
End Sub

Sub FilterSelectedValue(hideVal As Boolean)

'Select some cells in a column
'Filter that column by all vaules in selected range - visible cells only
'idea for future
'reverse this function, filter a list by all except selected
'create a second array - populate with each unique value in column
'remove values that occur in selection array
'filter failed with some numeric formats
'cell. options of text, value and value2 have differing behaviours
'cell.text returns the value with formatting as shown
'cell.value returns the cells contents without the format except for date and currency
'cell.value2 returns the cells contents without the format

On Error GoTo hideSelectedErr
If ActiveWorkbook Is Nothing Then
    Exit Sub
End If

Dim arrListSelected() As String ' array of selected cell values
Dim lngCnt As Long
Dim filterRng As Range
Dim arrElement As String
Dim lastRow As Long


lngCnt = 0


'load values into an array
lngCnt = 0
If Selection.Count > 1 Then
    Selection.SpecialCells(xlCellTypeVisible).Select
End If

ReDim arrListSelected(Selection.Count)
        For Each cell In Selection
            arrListSelected(lngCnt) = cell.Value2
            lngCnt = lngCnt + 1
        Next

filterColumn = Selection.Column
lastRow = Cells.SpecialCells(xlCellTypeLastCell).Row

With ActiveWorkbook.ActiveSheet
    Set filterRng = Range(Cells(2, filterColumn), Cells(lastRow, filterColumn)).SpecialCells(xlCellTypeVisible)
End With

'load all values present into a second array

lngCnt = 0
ReDim arrListColumn(lastRow)
SuspendAutoCalc

    For Each cell In filterRng
        arrElement = cell.Value2
        lngCnt = lngCnt + 1
            
        If elementInArray(arrListSelected, arrElement) Then
            cell.EntireRow.Hidden = hideVal
        Else
            cell.EntireRow.Hidden = Not hideVal
        End If
    Next
ResumeAutoCalc
hideSelectedErr:

End Sub




Public Sub FilterColumns()
'Public Sub columnFilter(filterRow As Integer)

If ActiveWorkbook Is Nothing Then
    Exit Sub
End If

Dim col As Integer
Dim startCol As Integer
Dim hideCol As Boolean
Dim filterValue As String
'Dim filterRow As Integer

hideCol = False
filterRow = ActiveCell.Row
startCol = ActiveCell.Column
unhideAllColumns ActiveSheet
Application.ScreenUpdating = False

filterValue = Application.InputBox("What value would you like to filter this row for?", "Filter", ActiveCell.value, , , , , 2)
If filterValue <> "False" Then
    For col = startCol To Cells.SpecialCells(xlCellTypeLastCell).Column
        'If (Cells(filterRow, col).Value <> "") Then
            If Cells(filterRow, col).value = filterValue Then
                hideCol = False
            Else
                hideCol = True
            End If
        'End If
        Cells(filterRow, col).EntireColumn.Hidden = hideCol
    Next col
Else
    unhideAllColumns ActiveSheet
End If

Application.ScreenUpdating = True
End Sub

Sub unhideAllColumns(ws As Worksheet)

With ws.Cells
    .EntireColumn.Hidden = False
    '.EntireRow.Hidden = False
End With
End Sub

Sub unhideAllRows(Optional ws As Worksheet)

With ws.Cells
    '.EntireColumn.Hidden = False
    .EntireRow.Hidden = False
End With
End Sub

Sub PivotClearFilters()
'prevents unused items in non-OLAP PivotTables
'pivot table tutorial by contextures.com
'also refreshes all pivot tables in a workbook in one go

On Error GoTo errPivotClearFilter
Dim pt As PivotTable
Dim ws As Worksheet
Dim pc As PivotCache

'change the settings
For Each ws In ActiveWorkbook.Worksheets
  For Each pt In ws.PivotTables
    pt.PivotCache.MissingItemsLimit = xlMissingItemsNone
    'pt.RepeatAllLabels xlDoNotRepeatLabels
    pt.RepeatAllLabels xlRepeatLabels
    pt.DisplayContextTooltips = False
    pt.RowAxisLayout xlTabularRow
    pt.ColumnGrand = False
    pt.RowGrand = False
    pt.HasAutoFormat = False
    For Each Field In pt.PivotFields
       If Field.Name <> "Values" Then
        If Not Field.IsCalculated Then
             Field.Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        End If
       End If
    Next Field
  Next pt
Next ws
PivotRefresh

normalExit:
Exit Sub

errPivotClearFilter:
Resume Next

End Sub

Public Sub PivotRefresh()
'refresh all the pivot tables within a workbook
Dim pt As PivotTable
Dim pc As PivotCache

For Each pc In ActiveWorkbook.PivotCaches
  On Error Resume Next
  pc.Refresh
Next pc

End Sub


Function copyCells(sourceWS As Worksheet, destWS As Worksheet, sourceColumn As Integer, sourceRow As Long, destColumn As Integer, destRow As Long) As Boolean
'copies the value of a specified cell on one worksheet to another worksheet
'use from within a loop to copy a range or cells between worksheets
'14/11/14

'Application.ScreenUpdating = False
If ActiveWorkbook Is Nothing Then
    Exit Function
End If

    destWS.Cells(destRow, destColumn).NumberFormat = "@"
    destWS.Cells(destRow, destColumn).value = sourceWS.Cells(sourceRow, sourceColumn).value
    

'Application.ScreenUpdating = True

End Function

Function unHideAll(ws As Worksheet)
    'unhide all rows and columns in provided worksheet
If ActiveWorkbook Is Nothing Then
    Exit Function
End If

    With ws.Cells
        .EntireColumn.Hidden = False
        .EntireRow.Hidden = False
    End With
    
End Function

Sub unhideAllDataAllSheets(Optional wbk As Workbook)
    
    'unhide all rows and columns in current worksheet

On Error Resume Next

If wbk Is Nothing Then
    Set wbk = ActiveWorkbook
End If

If ActiveWorkbook Is Nothing Then
    Exit Sub
End If
    
    For Each ws In wbk.Worksheets
        With ws
            If .FilterMode Then
               .ShowAllData
            End If
            '.Activate
            For Each t In .ListObjects
                If t.ShowAutoFilter = True Then
                    t.AutoFilter.ShowAllData
                End If
            Next t
            
            .Cells.EntireColumn.Hidden = False
            .Cells.EntireRow.Hidden = False
        End With
    Next ws
    


End Sub
Sub unHideAllCurrent()
    'unhide all rows and columns in current worksheet
If ActiveWorkbook Is Nothing Then
    Exit Sub
End If
        With ActiveSheet
            If .FilterMode Then
               .ShowAllData
            End If
            Cells.EntireColumn.Hidden = False
            Cells.EntireRow.Hidden = False
        End With
    
End Sub

Sub unHideAllSheets(Optional wbk As Workbook)

On Error Resume Next
If wbk Is Nothing Then
    Set wbk = ActiveWorkbook
End If

If ActiveWorkbook Is Nothing Then
    Exit Sub
End If

On Error GoTo errorHandler
    
    For Each ws In wbk.Worksheets
        ws.Visible = xlSheetVisible
    Next

errorHandler:

End Sub

