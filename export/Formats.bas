Attribute VB_Name = "Formats"
Sub FormatAsABN()
SuspendAutoCalc
On Error GoTo AormatAsABN_Err
    For Each cell In Selection.Cells
        cell.value = Format(cell.value, "00 000 000 000")
    Next
AormatAsABN_Err:
ResumeAutoCalc
End Sub

Sub rounding()

If ActiveWorkbook Is Nothing Then
    Exit Sub
End If

On Error GoTo errorHandler
'

SuspendAutoCalc

Dim currCell As Range
    If Selection.Count > 1 Then
        Selection.SpecialCells(xlCellTypeVisible).Select
    End If
    For Each currCell In Selection.Cells
        With currCell
            .Formula = "=round(" & Right(.Formula, Len(.Formula) - 1) & ",2)"
        End With
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


Sub LeadingZeroAdd()
If ActiveWorkbook Is Nothing Then
    Exit Sub
End If

Dim response As Integer

On Error GoTo errorHandler
'
askAgain:
response = InputBox("How many zeros would you like to have?", "Leading zero", 1)

If response = 0 Then
    Exit Sub
End If

If Not IsNumeric(response) Then
    MsgBox "You must enter a number"
    GoTo askAgain
End If

SuspendAutoCalc

Dim currCell As Range
    If Selection.Count > 1 Then
        Selection.SpecialCells(xlCellTypeVisible).Select
    End If
    For Each currCell In Selection.Cells
            currCell.NumberFormat = "@"
            If IsNumeric(currCell.value) Then
                currCell.Formula = String(response, "0") & 0 + currCell.value
            Else
                currCell.Formula = String(response, "0") & currCell.value
            End If
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


Sub LeadingZeroPad()
If ActiveWorkbook Is Nothing Then
    Exit Sub
End If

Dim response As Integer

On Error GoTo errorHandler
'
askAgain:
response = InputBox("How many characters would you like to have including zero padding?", "Leading zero", 11)

If response = 0 Then
    Exit Sub
End If

If Not IsNumeric(response) Then
    MsgBox "You must enter a number"
    GoTo askAgain
End If

SuspendAutoCalc

Dim currCell As Range
    If Selection.Count > 1 Then
        Selection.SpecialCells(xlCellTypeVisible).Select
    End If
    For Each currCell In Selection.Cells
            currCell.NumberFormat = "@"
            If IsNumeric(currCell.value) Then
                currCell.Formula = Right(String(response, "0") & 0 + currCell.value, response)
            Else
                currCell.Formula = Right(String(response, "0") & currCell.value, response)
            End If
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

Private Sub userFormTest()
    frmDataType.Show
    
End Sub


Sub UnWrapText()
    Selection.WrapText = False
End Sub
Sub setTextCaseUpper()
    setTextCase 1
End Sub

Sub setTextCaseLower()
    setTextCase 2
End Sub
Sub setTextCaseTitle()
    setTextCase 3
End Sub

Private Sub setTextCase(caseToSet As Integer)
If ActiveWorkbook Is Nothing Then
    Exit Sub
End If

On Error GoTo errorHandler
'

SuspendAutoCalc

Dim currCell As Range
    If Selection.Count > 1 Then
        Selection.SpecialCells(xlCellTypeVisible).Select
    End If
    For Each currCell In Selection.Cells
        With currCell
            If caseToSet = 1 Then
                .value = UCase(.value)
            ElseIf caseToSet = 2 Then
                .value = LCase(.value)
            ElseIf caseToSet = 3 Then
                .value = StrConv(.value, vbProperCase)
            End If
        End With
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


Sub SupplierNameSimplify()
'   28/07/2015
'   Ron Campbell

'Highlight a section of cells
'step through cell by cell
'remove complications such as PTY/LTD etc
'use exclude words instead, can have additional words removed
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
        If Not Application.IsNA(currCell) Then
            currCell.Formula = Trim(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(UCase(currCell.value), "PTY", ""), "LTD", ""), "LIMITED", ""), "P/L", ""), "(VIC)", ""), "(AUST)", ""), "(AUSTRALIA)", ""), "THE TRUSTEE FOR ", ""))
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

Sub fillDown()
'fill the gaps in a column with the value from the cell above when a cell is blank

SuspendAutoCalc
    For Each cell In Selection.Cells
        If Len(cell.value) = 0 Then
            cell.value = cell.Offset(-1, 0).value
        End If
    Next
ResumeAutoCalc
End Sub


Sub copyFromAbove()
'Copy row from line above current selection
copyRow -1
End Sub

Sub copyFromBelow()
'Copy row from line below current selection
copyRow 1
End Sub

Sub copyRow(offsetRows As Integer)
'copy the row above or below by given number of rows. copy from current column to right extent of data
With ActiveCell
    Range(.Offset(offsetRows, 0), .Offset(offsetRows, (.SpecialCells(xlLastCell).Column) - ActiveCell.Column)).Copy
    ActiveSheet.Paste
    Application.CutCopyMode = False
End With

End Sub
Sub BorderThickInnerThin()
'Set up basic table format with heavy outer border and light cell border

   Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

End Sub

Sub clearDataValidation()
' remove data validation from all cells on all worksheets in the current workbook
On Error Resume Next
For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
    Cells.SpecialCells(xlCellTypeAllValidation).Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator _
        :=xlBetween
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
Next
End Sub

Sub clearConditionalFormatting()
On Error GoTo nextWS
For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
    Cells.FormatConditions.Delete
nextWS:
Next

End Sub


