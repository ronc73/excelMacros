Attribute VB_Name = "Macros_Functions"
'Global dataSrc
'Global dataDest

Dim intDataType  As Integer


Sub mergeSame()

If ActiveWorkbook Is Nothing Then
    Exit Sub
End If

Dim c As Range
Dim i As Integer
Application.DisplayAlerts = False
SuspendAutoCalc
    i = 0
    For Each c In Selection.Cells
        Do While c.value = c.Offset(i + 1, 0).value And c.value <> ""
            i = i + 1
        Loop
        If i > 1 Then
            Range(c, c.Offset(i, 0)).Merge
        End If
        i = 0
    Next
    
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Font.Bold = True
    End With

Application.DisplayAlerts = True

        

End Sub


Sub StyleKill()
' ripped from the net
' Removes text styles that are not built in styles

If ActiveWorkbook Is Nothing Then
    Exit Sub
End If

     Dim n As Long
     Dim d   As Object
     Dim styT As Style
     Dim intRet As Integer
     Dim StartTime As Variant
     Dim Endtime As Variant
     Dim StyleCount As Long
     Dim StyleRemoval As Long
    
StartTime = Now
    StyleRemoval = 0
    Set d = CreateObject("scripting.dictionary")
        d.CompareMode = 1
    StyleCount = 0

    With ActiveWorkbook
'        n = .Styles.Count
        StyleCount = .Styles.Count
'        'get all the non-built styles
'        For i = 1 To n
'            If Not .Styles(i).BuiltIn Then
'                d.Item(.Styles(i).NameLocal) = False
'                StyleCount = StyleCount + 1
'                Application.StatusBar = "Counting Styles: " & StyleCount
'            End If
'        Next
        
     ' MsgBox "Counting Styles: " & StyleCount
     On Error GoTo errorHandler
     SuspendAutoCalc

     For Each styT In ActiveWorkbook.Styles
         If Not styT.BuiltIn Then
             If styT.Name <> "1" Then styT.Delete
             StyleRemoval = StyleRemoval + 1
             'Application.StatusBar = "Removing Style: " & StyleRemoval & " of " & StyleCount
         End If
     Next styT
        Endtime = Now
           MsgBox "This file initially contained: " & StyleCount & " Styles." & vbCr & _
         "The Macro removed: " & StyleRemoval & vbCr & _
         "The macro took: " & Format(Endtime - StartTime, "hh:nn:ss") & vbCr & _
         "Start: " & Format(StartTime, "dd/mm/yy hh:nn:ss") & "  End: " & Format(Endtime, "dd/mm/yy hh:nn:ss"), vbOKOnly, "Score Card"

    End With
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
        If Err.Number = 1004 Then
            Resume Next
        End If
        Resume ErrorExit
        

 End Sub

Sub RemoveUnusedStyles()
    
    '// Author  : Admin @ ExcelFox.com
    '// Purpose : Delete all unused styles from a workbook.
'alternate way of removing unused styles
'search through the file listing the styles in use
'list styles in file
'cycle through available styles comparing to in-use list
'remove listed styles not in use

'Additional functionality required 18/06/13
'If any worksheet is protected the styles cannot be deleted
'Cycle through all worksheets
'Any protected sheets record and unprotect
'at end of macro re-protect these sheets
If ActiveWorkbook Is Nothing Then
    Exit Sub
End If

    Dim i   As Long
    Dim c   As Long
    Dim n   As Long
    Dim r   As Long
    Dim d   As Object
    Dim s   As Style
    Dim a
    Dim StyleCount As Long
    Dim StyleRemoval As Long
    Dim StartTime As Variant
    Dim Endtime As Variant
    Dim availableStyle(64000) As String
    Dim availableStylecount As Long
    Dim usedStyle(64000) As String
    Dim usedStyleCount As Long
    Dim foundStyle As Boolean
    'Dim ws As Worksheet
    Dim worksheetList As Object
    Dim pass As String
    Dim ColCount As Long
    Dim RowCount As Long
    Dim WorksheetCount As Integer
    StartTime = Now()
            
    Set d = CreateObject("scripting.dictionary")
        d.CompareMode = 1
    
    Set worksheetList = CreateObject("scripting.dictionary")
    counter = 0
    
    '--------------------------
    'Count up the available styles and advise the user of the result and ask if they wish to continue
        
    availableStylecount = 0

    With ActiveWorkbook
        StyleCount = .Styles.Count
        'get all the non-built styles
        'add names of custom styles to array availableStyle
        For i = 1 To StyleCount
            If Not .Styles(i).BuiltIn Then
                Application.StatusBar = "Processing" & String(i Mod 10, ".")
                d.Item(.Styles(i).NameLocal) = False
                availableStylecount = availableStylecount + 1
                Application.StatusBar = "Counting Styles: " & availableStylecount & " of " & StyleCount
            End If
        Next
                
    If MsgBox("There are " & availableStylecount & _
             " styles available in this workbook. " & _
             "Do you wish to continue to remove the unused styles?", vbYesNo, "Continue") = vbNo Then
        GoTo endNow
    End If
        
        
    'turn off protection
'    For Each ws In ActiveWorkbook.Worksheets
'        If ws.ProtectContents = True Then
'            worksheetList.Add ws.CodeName, True
'            ws.Unprotect
'        End If
'    Next
    pass = unprotectWorksheet(worksheetList)
    If pass = "-1" Then ' password required but not known. exit
        End
    End If
    
    
    SuspendAutoCalc

        
        n = 0
        usedStyleCount = 0
        For i = 1 To .Worksheets.Count
            WorksheetCount = .Worksheets.Count
            With .Worksheets(i).UsedRange
                Worksheets(i).Activate
                RowCount = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row
                ColCount = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Column
                For c = 1 To ColCount
                    Application.StatusBar = "Processing" & String(c Mod 10, ".")
                    For r = 1 To RowCount
                        Set s = .Cells(r, c).Style
                        If Not s.BuiltIn Then
                            'check if this style has been added to the array of usedStyleCount
                            'if not then add it
                            foundStyle = False
                            n = 0
                            While n <= usedStyleCount And foundStyle = False
                                If usedStyle(n) = s.Name Then
                                    foundStyle = True
                                End If
                                n = n + 1
                            Wend ' checking style of current cell with list of used styles
                            If Not foundStyle Then
                                usedStyle(usedStyleCount) = s.Name
                                d.Item(ActiveWorkbook.Styles(s.Name).NameLocal) = True
                                usedStyleCount = usedStyleCount + 1
                                Application.StatusBar = "Counting Used Styles: " & usedStyleCount
                            End If ' foundstyle
                        End If 'not built in style
                        Application.StatusBar = "Counting Used Styles: Sheet: " & i & "/" & WorksheetCount & _
                                    ". Column: " & c & "/" & ColCount & _
                                    ". Row: " & r & "/" & RowCount & _
                                    ". Style: " & s & "/" & usedStyleCount
                    Next ' row
                    Application.StatusBar = "Counting Used Styles: Sheet: " & i & "/" & WorksheetCount & _
                                    ". Column: " & c & "/" & ColCount & _
                                    ". Row: " & r & "/" & RowCount & _
                                    ". Style: " & s & "/" & usedStyleCount
                Next ' column
                Application.StatusBar = "Counting Used Styles: Sheet: " & i & "/" & WorksheetCount & _
                                    ". Column: " & c & "/" & ColCount & _
                                    ". Row: " & r & "/" & RowCount & _
                                    ". Style: " & s & "/" & usedStyleCount
            End With ' used range
            Application.StatusBar = "Counting Used Styles: Sheet: " & i & "/" & WorksheetCount & _
                                    ". Column: " & c & "/" & ColCount & _
                                    ". Row: " & r & "/" & RowCount & _
                                    ". Style: " & s & "/" & usedStyleCount
        Next 'Worksheet
        
        StyleRemoval = 0
        'Cycle through list of styles, delete unused styles
        a = Array(d.keys, d.items)
        For i = LBound(a) To UBound(a(0))
            'delete unused styles
            If Not CBool(a(1)(i)) Then
                If Not .Styles(a(0)(i)).BuiltIn Then
                    .Styles(a(0)(i)).Locked = False
                    .Styles(a(0)(i)).Delete
                    StyleRemoval = StyleRemoval + 1
                     Application.StatusBar = "Removing Style: " & StyleRemoval & " of " & StyleCount
                End If
            End If
        Next


   End With
   
   

   
      'turn on protection
    For Each ws In ActiveWorkbook.Worksheets
        If worksheetList.Item(ws.CodeName) Then
            ws.Protect pass
        End If
    Next
    
    
   Endtime = Now()
   
   
   MsgBox "This file initially contained: " & availableStylecount & " Styles." & vbCr & _
         "The number of styles in use: " & usedStyleCount & vbCr & _
         "The Macro removed: " & StyleRemoval & vbCr & _
         "The macro took: " & Format(Endtime - StartTime, "hh:nn:ss") & vbCr & _
         "Start: " & Format(StartTime, "dd/mm/yy hh:nn:ss") & vbCr & _
         "End: " & Format(Endtime, "dd/mm/yy hh:nn:ss"), vbOKOnly, "Score Card"


endNow:
    ResumeAutoCalc
    Application.StatusBar = ""
    Exit Sub

ErrorExit:
    'restore the user's calculation setting
    ResumeAutoCalc
    Application.StatusBar = ""
    'the ErrorHandler code should only be executed if there is an error
    Exit Sub

errorHandler:
    Debug.Print Err.Number & vbLf & Err.Description
    Resume ErrorExit
        
End Sub

Sub CurrencyFormat()
'   7/05/2013
'   Ron Campbell

'Highlight a section of cells
'setting to currency format with negatives as red


'
If ActiveWorkbook Is Nothing Then
    Exit Sub
End If

SuspendAutoCalc
    'Selection.Style = "Currency"
    Selection.NumberFormat = "$#,##0.00;[Red]-$#,##0.00"
    
    ResumeAutoCalc
    Exit Sub
    
ErrorExit:
    On Error Resume Next
    'restore the user's calculation setting
    ResumeAutoCalc
    'the ErrorHandler code should only be executed if there is an error
    Exit Sub
errorHandler:
        Debug.Print Err.Number & vbLf & Err.Description
        Resume ErrorExit
        
End Sub

Sub HPVDateFormat()
'   12/08/2014
'   Ron Campbell

'Highlight a section of cells
'Set date format to yyy-mm-dd


If ActiveWorkbook Is Nothing Then
    Exit Sub
End If

'
    SuspendAutoCalc
    'Selection.Style = "Currency"
    Selection.NumberFormat = "yyyy-mm-dd hh:mm:ss"
    
    ResumeAutoCalc
    Exit Sub
    
ErrorExit:
    On Error Resume Next
    'restore the user's calculation setting
    ResumeAutoCalc
    'the ErrorHandler code should only be executed if there is an error
    Exit Sub
errorHandler:
        Debug.Print Err.Number & vbLf & Err.Description
        Resume ErrorExit
        
End Sub



Sub InteriorColourIndex()
    
If ActiveWorkbook Is Nothing Then
    Exit Sub
End If

Dim c As Range
    
    For Each c In Selection.Cells
        'Range(dataDest).Offset(y, 0).Value = c.Value
        c.Interior.ColorIndex = c.value
    Next c
End Sub

Function SetBGColour(colourSet As Integer)
If ActiveWorkbook Is Nothing Then
    Exit Function
End If

Dim c As Range

    For Each c In Selection.Cells
        'Only set the colour is the cell is visible ie height >0
        If c.Height > 0 Then
            c.Interior.ColorIndex = colourSet
        End If
    Next c
End Function

Sub SetBGColourYellow()
           SetBGColour (27)
End Sub

Sub SetBGColourGreen()
           SetBGColour (50)
End Sub

Sub SetBGColourRed()
           SetBGColour (30)
End Sub

Sub ClearBGColour()
If ActiveWorkbook Is Nothing Then
    Exit Sub
End If

Dim c As Range

'Selection.Cells.Interior.Pattern = xlNone

For Each c In Selection.Cells
        'Range(dataDest).Offset(y, 0).Value = c.Value
        If c.Height > 0 Then
            c.Interior.Pattern = xlNone
        End If
    Next c
End Sub


Sub TextColourIndex()
If ActiveWorkbook Is Nothing Then
    Exit Sub
End If

Dim c As Range
    
    For Each c In Selection.Cells
        'Range(dataDest).Offset(y, 0).Value = c.Value
        c.Font.ColorIndex = c.value
    Next c
End Sub


Sub TextInterionColourIndex()
If ActiveWorkbook Is Nothing Then
    Exit Sub
End If
Dim c As Range
    
    For Each c In Selection.Cells
        'Range(dataDest).Offset(y, 0).Value = c.Value
        'c.Font.ColorIndex = c.value
        c.Font.ColorIndex = (c.value + 100 * Rnd) Mod 56
        c.Interior.ColorIndex = (c.value + 100 * Rnd) Mod 56
    Next c
End Sub

Sub CloseThisAddIn()
   ' ThisWorkbook.IsAddin = False
   MsgBox "Closing Macro file"
    ThisWorkbook.Close False
End Sub


Sub UnpivotElectricityData()
If ActiveWorkbook Is Nothing Then
    Exit Sub
End If

Dim DataRows As Long
Dim DataCols As Long
Dim destRow As Long '
Dim destCol As Long
Dim sourceCol As Integer    'will  be 3 by default after adding a column the data headers start at column 3
Dim sourceRow As Integer    'will be 1 by default as the source data begins in the first row
Dim dateStampRange As Range
Dim dataRange As Range

If MsgBox("You are about to re-organise your data for import into the electricity consumption database. Continue? ", vbInformation + vbYesNo, "Continue? ") = vbYes Then
    DataRows = Application.ActiveSheet.UsedRange.Rows.Count - 1
    DataCols = Application.ActiveSheet.UsedRange.Columns.Count - 1
    
    'insert blank column at A
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    
    'populate A column
    'repeat first column header onces for DataRows from 2 to DataRows + startRow
    'Set testRange = .Range(.Cells(5, 5), .Cells(10, 10))
    
    destCol = 1
    sourceRow = 1
    destRow = 2
    
    'dateStampRange = Range(Cells(2, 2), Cells(2 + DataRows - 1, 2))
    For i = 3 To DataCols + 2
        sourceCol = i
        'dataRange = Range(Cells(2, sourceCol), Cells(DataRows + 1, sourceCol))
        Range(Cells(destRow, destCol), Cells(DataRows + destRow - 1, destCol)).value = Cells(sourceRow, sourceCol).value
        If i > 3 Then
            Range(Cells(destRow, destCol + 1), Cells(DataRows + destRow - 1, destCol + 1)).value = Range(Cells(2, 2), Cells(2 + DataRows - 1, 2)).value ' dateStampRange
            Range(Cells(destRow, destCol + 2), Cells(DataRows + destRow - 1, destCol + 2)).value = Range(Cells(2, sourceCol), Cells(DataRows + 1, sourceCol)).value ' dataRange
        End If
        destRow = destRow + DataRows
    Next i
    
    Range("A1").value = "MeterNo"
    Range("B1").value = "DateTimeofUse"
    Range("C1").value = "Consumption"
    
    'remove excess data
    Range(Range("D1"), Cells.SpecialCells(xlLastCell)).Delete
    'save file
    SaveWorkbookAsNewFile Format(Date, "yyyy-mm-dd") & "-ElectricityImport"
    'open access
    OpenDatabase "H:\Shared\Operational\Tender Program\Electricity 2013\Usage Data\Electricity.accdb", "frmMenu"
Else
    MsgBox "OK No harm done"
End If

End Sub

Sub SaveWorkbookAsNewFile(NewFileName As String)
If ActiveWorkbook Is Nothing Then
    Exit Sub
End If

Dim ActSheet As Worksheet
    Dim ActBook As Workbook
    Dim currentFile As String
    Dim NewFileType As String
    Dim newFile As String

Application.ScreenUpdating = False    ' Prevents screen refreshing.

    currentFile = ThisWorkbook.FullName
 
    NewFileType = "Excel Workbook (*.xlsx), *.xlsx," & _
               "All files (*.*), *.*"
 
    newFile = Application.GetSaveAsFilename( _
        InitialFileName:=NewFileName, _
        fileFilter:=NewFileType)
 
    If newFile <> "" And newFile <> "False" Then
        ActiveWorkbook.SaveAs fileName:=newFile, _
            FileFormat:=xlOpenXMLWorkbook, _
            Password:="", _
            WriteResPassword:="", _
            ReadOnlyRecommended:=False, _
            CreateBackup:=False
 
'        Set ActBook = ActiveWorkbook
'        Workbooks.Open CurrentFile
'        ActBook.Close
        ActiveWorkbook.Close
    End If
 
    Application.ScreenUpdating = True
End Sub

Sub OpenDatabase(strdb As String, strOpenForm As String)
If ActiveWorkbook Is Nothing Then
    Exit Sub
End If

'strdb = "c:\demo.mdb"
Set appAccess = CreateObject("Access.Application")
appAccess.Visible = True
appAccess.OpenCurrentDatabase strdb
appAccess.docmd.OpenForm strOpenForm
Set appAccess = Nothing

End Sub



Sub turnOffRemovePersonalInformation()

If ActiveWorkbook Is Nothing Then
    Exit Sub
End If

If Not ActiveWorkbook Is Nothing Then
    ActiveWorkbook.RemovePersonalInformation = False
    MsgBox "Remove Personal Information Setting has been turned off"
End If
End Sub

Sub AdditionalComments()
'Adds a comment to all cells in a selected range
'if the cell already contains something the comment is appended after a comma

If ActiveWorkbook Is Nothing Then
    Exit Sub
End If

Dim AddedComment As String

AddedComment = InputBox("Please type your comment", "AdditionalComment")
If AddedComment = "" Then   'AddedComment = vbCancel Or
    Exit Sub
End If
SuspendAutoCalc
'load values into an array
For Each cell In Selection.SpecialCells(xlCellTypeVisible)
    If Len(cell.value) > 0 Then
        cell.value = cell.value & ", " & AddedComment
    Else
        cell.value = AddedComment
    End If
Next
ResumeAutoCalc

End Sub


Sub UnpivotGeneric()

If ActiveWorkbook Is Nothing Then
    Exit Sub
End If

Dim DataRows As Long
Dim DataCols As Long
Dim destRow As Long '
Dim destCol As Long
Dim sourceCol As Integer    'will  be 3 by default after adding a column the data headers start at column 3
Dim sourceRow As Integer    'will be 1 by default as the source data begins in the first row
Dim dateStampRange As Range
Dim dataRange As Range
'Dim DataStartCol As Integer
Dim HeaderRowStart As Integer
Dim HeaderRowEnd As Integer
Dim DataColStart As Integer
Dim DataColEnd As Integer
Dim i As Integer    'default loop counter

If MsgBox("You are about to re-organise your data. Continue? ", vbInformation + vbYesNo, "Continue? ") = vbYes Then
    DataRows = Application.ActiveSheet.UsedRange.Rows.Count - 1
    DataCols = Application.ActiveSheet.UsedRange.Columns.Count - 1
    
'RowCount = Selection.SpecialCells(xlVisible).Rows.Count
'ColCount = Selection.SpecialCells(xlVisible).Columns.Count
'dataDest = Application.InputBox("Select range to copy selected cells to", Type:=8).Address

    Do
        tmp = Application.InputBox("Enter the row number of the first row of the header", "Header Row Start", Type:=8).Address
        HeaderRowStart = Range(tmp).Row
    Loop Until IsNumeric(HeaderRowStart)
    
    
    Do
        tmp = Application.InputBox("Enter the columns number of the first column of the Data", "Data Column Start", Type:=8).Address
        DataColStart = Range(tmp).Column
    Loop Until IsNumeric(DataColStart)
'    Do
'        tmp = Application.InputBox("Enter the row number of the Last column of the Data", "Data Column End", Type:=8).Address
'        DataColEnd = Range(tmp).Column
'    Loop Until IsNumeric(DataColEnd)
    DataColEnd = Cells.SpecialCells(xlLastCell).Column + 1
    
    'insert blank column at A for each row of header
    Columns("A:A").Select
'    For i = 1 To HeaderRowEnd - HeaderRowStart + 1
        Selection.Insert Shift:=xlToRight
'    Next i
    
    'populate A column
    'repeat first column header onces for DataRows from 2 to DataRows + startRow
    'Set testRange = .Range(.Cells(5, 5), .Cells(10, 10))
    
    destCol = 1
    sourceRow = HeaderRowStart
    destRow = HeaderRowStart + 1
    'DataStartCol = HeaderRowEnd - HeaderRowStart + 1 + 2
    
    
    'dateStampRange = Range(Cells(2, 2), Cells(2 + DataRows - 1, 2))
    For i = DataColStart To DataColEnd
        sourceCol = i
        Range(Cells(destRow, destCol), Cells(DataRows + destRow - 1, destCol)).value = Cells(sourceRow, sourceCol).value
        If i > DataColStart Then
            Range(Cells(destRow, destCol + DataColStart - 1), Cells(DataRows + destRow - 1, destCol + DataColStart - 1)).value _
                = Range(Cells(2, 2), Cells(2 + DataRows - 1, DataColStart)).value  ' dateStampRange
            Range(Cells(destRow, destCol + DataColStart), Cells(DataRows + destRow - 1, destCol + DataColStart)).value _
                = Range(Cells(2, sourceCol), Cells(DataRows + 1, sourceCol)).value ' dataRange
        End If
        destRow = destRow + DataRows
    Next i
    
'    Range("A1").Value = "Date"
'    Range("B1", "E1").Value = Range("B2", "E2").Value
'    Range("F1").Value = "Hours"
    
    'remove excess data
'    Range(Cells(1, DataColStart + 1), Cells.SpecialCells(xlLastCell)).Delete
    'save file
    'SaveWorkbookAsNewFile Format(Date, "yyyy-mm-dd") & "-ElectricityImport"
    'open access
    'OpenDatabase "H:\Shared\Operational\Tender Program\Electricity 2013\Usage Data\Electricity.accdb", "frmMenu"
Else
    MsgBox "OK No harm done"
End If

End Sub

Sub FilteredHarshTrim()
Attribute FilteredHarshTrim.VB_Description = "The result will be the contents of the reference cell with all characters stripped out except 0-9, A-Z and a-z."
Attribute FilteredHarshTrim.VB_ProcData.VB_Invoke_Func = " \n20"
'   14/08/2014
'   Ron Campbell

Application.MacroOptions _
        Macro:="FilteredHarshTrim", _
        Description:="The result will be the contents of the reference cell with all characters stripped out except 0-9, A-Z and a-z.", _
        Category:="Custom Worksheet Functions", _
        ArgumentDescriptions:=Array( _
            "inputStr is the first argument.  The string to be operated on, usually a cell reference", _
            "OtherChr is an optional argument.  Allows you to specify additional characters to keep")

If ActiveWorkbook Is Nothing Then
    Exit Sub
End If

Dim response As Integer
Dim CharToKeep As String
On Error GoTo errorHandler
'
CharToKeep = InputBox("List any characters to keep beyond letters and numbers ie. ./x etc ?", "Char to Keep", "/")
'NumericFilter(strCode As String, Optional OtherChr As String)
If CharToKeep = "" Then
    If MsgBox("Continue and keep just letters and numbers?", vbYesNo) = vbNo Then
        Exit Sub
    End If
End If

SuspendAutoCalc

Dim currCell As Range
    If Selection.Count > 1 Then
        Selection.SpecialCells(xlCellTypeVisible).Select
    End If
    For Each currCell In Selection.Cells
' VBA inherently treads all dates as mm/dd/yyyy ignoring local settings.
' Added this section when I found that dates stored as text got clobbered
            currCell.value = HarshTrim(currCell.value, CharToKeep)
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

Sub FilteredNumericExtractor()
'   14/08/2014
'   Ron Campbell

If ActiveWorkbook Is Nothing Then
    Exit Sub
End If

Dim response As Integer
Dim CharToKeep As String
On Error GoTo errorHandler
'
CharToKeep = InputBox("List any characters to keep beyond numbers ie. ./x etc ?", "Char to Keep", "/")
'NumericFilter(strCode As String, Optional OtherChr As String)
If CharToKeep = "" Then
    If MsgBox("Continue and keep just numbers?", vbYesNo) = vbNo Then
        Exit Sub
    End If
End If

SuspendAutoCalc

Dim currCell As Range
    If Selection.Count > 1 Then
        Selection.SpecialCells(xlCellTypeVisible).Select
    End If
    For Each currCell In Selection.Cells
' VBA inherently treats all dates as mm/dd/yyyy ignoring local settings.
' Added this section when I found that dates stored as text got clobbered
            currCell.value = NumericFilter(currCell.value, CharToKeep)
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

Sub FilteredTextExtractor()
'   19/04/2016
'   Ron Campbell

If ActiveWorkbook Is Nothing Then
    Exit Sub
End If

Dim response As Integer
Dim CharToKeep As String
On Error GoTo errorHandler
'
CharToKeep = InputBox("List any characters to keep beyond numbers ie. ./x etc ?", "Char to Keep", "/")
'NumericFilter(strCode As String, Optional OtherChr As String)
If CharToKeep = "" Then
    If MsgBox("Continue and keep just numbers?", vbYesNo) = vbNo Then
        Exit Sub
    End If
End If

SuspendAutoCalc

Dim currCell As Range
    If Selection.Count > 1 Then
        Selection.SpecialCells(xlCellTypeVisible).Select
    End If
    For Each currCell In Selection.Cells
' VBA inherently treads all dates as mm/dd/yyyy ignoring local settings.
' Added this section when I found that dates stored as text got clobbered
            currCell.value = TextFilter(currCell.value, CharToKeep)
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

Public Sub ColumnEnd()
Attribute ColumnEnd.VB_ProcData.VB_Invoke_Func = "C\n14"
' Attach macro to shift-ctrl-E
'highlight to the last cell in the current Column
'use Row number from specialcells(xlLastCell)

If ActiveWorkbook Is Nothing Then
    Exit Sub
End If

Dim SelectionColStart As Long
Dim SelectionRowStart As Long
Dim SelectionColCount As Long

If Selection.Count > 1 Then
    SelectionColStart = Selection.Columns(1).Column
    SelectionRowStart = Selection.Rows(1).Row
    SelectionColCount = Selection.Columns.Count - 1
Else
    SelectionColStart = ActiveCell.Column
    SelectionRowStart = ActiveCell.Row
    SelectionColCount = 0
End If

Range(Cells(SelectionRowStart, SelectionColStart), Cells(Cells.SpecialCells(xlLastCell).Row, SelectionColStart + SelectionColCount)).Select
End Sub

Public Sub RowEnd()
Attribute RowEnd.VB_ProcData.VB_Invoke_Func = "m\n14"
' Attach macro to ctrl-m
'highlight to the last cell in the current row
'use column number from specialcells(xlLastCell)

'Range(ActiveCell, Cells(ActiveCell.Row, Cells.SpecialCells(xlLastCell).Column)).Select
If ActiveWorkbook Is Nothing Then
    Exit Sub
End If

Dim SelectionColStart As Long
Dim SelectionRowStart As Long
Dim SelectionRowCount As Long

If Selection.Count > 1 Then
    SelectionColStart = Selection.Columns(1).Column
    SelectionRowStart = Selection.Rows(1).Row
    SelectionRowCount = Selection.Rows.Count - 1
Else
    SelectionColStart = ActiveCell.Column
    SelectionRowStart = ActiveCell.Row
    SelectionRowCount = 0
End If

Range(Cells(SelectionRowStart, SelectionColStart), Cells(SelectionRowStart + SelectionRowCount, Cells.SpecialCells(xlLastCell).Column)).Select

End Sub


Sub sdProcessReportUpdate()
' perform data transform for Greg J monthly reporting
' 13/12/14 Ron Campbell
'Update status summary line (column B)
'Update reporting stats numbers, (three columns "Report Periods outstanding", "Periods awaiting reports" and "Periods awaiting data cleaning) )
'   columns c, d and e into priority list sheet columns f, g and h)
'Insert any additional rows with period, qtr, date range, received and cleaned numbers
'   Columns f to j into c to g
'Source sheet is named "Latest Data" sheet8 use sheet8 reference
'Destination sheet named "Priority Listing" ID sheet3

If ActiveWorkbook Is Nothing Then
    Exit Sub
End If

Dim sourceDataSheet As Worksheet
Dim destDataSheet As Worksheet
Dim sourceCell As Range
Dim destCell As Range
Dim destRow As Long
Dim lastSourceContractRow As Integer
Dim sourceContracts() As String
Dim contractCount As Integer

On Error GoTo sdProcessReportUpdateErr
contractCount = 0

Set sourceDataSheet = GetWsFromCodeName(ActiveWorkbook, "Sheet8")
Set destDataSheet = GetWsFromCodeName(ActiveWorkbook, "Sheet3")

unHideAll destDataSheet
destDataSheet.AutoFilter.ShowAllData
SuspendAutoCalc
Application.StatusBar = "Starting ....."

With sourceDataSheet
    For Each sourceCell In Range(.Cells(2, 1), .Cells(.Cells.SpecialCells(xlLastCell).Row, 1))
        If sourceCell.Column = 1 And sourceCell.value <> "" Then
            ReDim Preserve sourceContracts(contractCount)  ' preserve keeps the content of the array when redimensioning it
            sourceContracts(contractCount) = sourceCell.value
            contractCount = contractCount + 1
           lastSourceContractRow = sourceCell.Row
            With destDataSheet.Cells
                Set c = .Find(Left(sourceCell.value, 12))
                    If Not c Is Nothing Then
                        destRow = c.Row
                        Application.StatusBar = "Row " & c.Row
                    Else 'contract name not found
                        'copy the contract to the bottom of the sheet
                        destRow = destDataSheet.Cells.SpecialCells(xlLastCell).Row
                    End If
                    destCol = 1 ' copy the contract name and status summary text
                        For i = 0 To 1
                            copyCells sourceDataSheet, destDataSheet, sourceCell.Column + i, sourceCell.Row, destCol + i, destRow
                        Next i
                    destCol = 6 ' copy the contract statistics
                        For i = 0 To 2
                            copyCells sourceDataSheet, destDataSheet, sourceCell.Column + 2 + i, sourceCell.Row, destCol + i, destRow
                        Next i
            End With
        End If
    Next
    
    For i = 0 To UBound(sourceContracts)
    Application.StatusBar = "Updating period data for " & sourceContracts(i)
        If i < UBound(sourceContracts) Then
            Set c = .Cells.Find(sourceContracts(i + 1))
            detailSetEnd = c.Row - 1
        Else
            Set c = .Cells.SpecialCells(xlCellTypeLastCell)
            detailSetEnd = c.Row
        End If
        Set c = .Cells.Find(sourceContracts(i))
        detailSetStart = c.Row
        
        If Not c Is Nothing Then
            'destRow is still looking at the contract header row
            'check if next destRow has a contract name, if not ctrl-down will find the next contract header row
            'count rows between source contract lines
            'count rows between dest contract lines
            With destDataSheet
            'find contract in dest sheet
            'add any rows needed in the destination sheet to ensure there are as many as in the source sheet
                Set c = .Cells.Find(sourceContracts(i))
                If Not c Is Nothing Then
                    destRow = c.Row
                End If
                For r = 1 To (detailSetEnd - detailSetStart)
                    If .Cells(destRow + r, 1).value <> "" Then
                        ' insert a new row
                        .Cells(destRow + r, 1).EntireRow.Insert
                    End If
                    'copy columns f to j into columns c to g (5 columns)
                sourceRow = detailSetStart
                sourceCol = 6
                destCol = 3
                    For col = 0 To 4
                        copyCells sourceDataSheet, destDataSheet, sourceCol + col, sourceRow + r, destCol + col, destRow + r
                    Next col
                Next r
            End With
        End If
    Next i
End With

With destDataSheet
    .Range(.Cells(1, 3), .Cells(1, 5)).EntireColumn.Hidden = True
    .Outline.ShowLevels rowlevels:=1
    .Range(.Cells(1, 1), .Cells(.Cells.SpecialCells(xlLastCell).Row, 8)).SpecialCells(xlCellTypeVisible).Copy Worksheets("StatusByContractData").Range("a1")
End With


sdProcessReportUpdateErr:
 'Application.ScreenUpdating = True
ResumeAutoCalc
Application.StatusBar = ""
End Sub




Sub testRowCount()
 MsgBox ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Column
 MsgBox ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row
 
End Sub


Sub getSheetList_PlainText()
    getSheetList False
End Sub

Sub getSheetList_Links()
    getSheetList True
End Sub

Sub getSheetList(hyperLink As Boolean)
'list the sheets within the current workbook. used for making an index page
Dim wSheet As Worksheet
Dim wBook As Workbook
Dim wCount As Integer

Set wBook = ActiveWorkbook
wCount = 0
For Each cell In Selection.Cells
    For Each wSheet In wBook.Worksheets
        With cell.Offset(wCount, 0)
            .value = wSheet.Name
            .ClearFormats
            .ClearHyperlinks
            If hyperLink Then
                ActiveSheet.Hyperlinks.Add Anchor:=cell.Offset(wCount, 0), Address:="", SubAddress:= _
                "'" & wSheet.Name & "'!A1", TextToDisplay:=wSheet.Name '& ", Color: " & wSheet.Tab.Color & ", ColorIndex: " & wSheet.Tab.ColorIndex
                .Interior.ColorIndex = wSheet.Tab.ColorIndex
                If wSheet.Tab.Color Then
                    If Right(wSheet.Tab.ColorIndex, 1) = 9 Or Right(wSheet.Tab.ColorIndex, 1) = 3 Then
                        .Font.Color = vbWhite
                    Else
                        .Font.Color = vbBlack
                    End If
                End If
            End If
            wCount = wCount + 1
        End With
    Next
Exit Sub
Next
End Sub

Sub SheetSelectA1()
'list the sheets within the current workbook. used for making an index page
Dim wSheet As Worksheet
Dim wBook As Workbook
Dim wCount As Integer
Dim hideSheet As Boolean
hideSheet = False

Set wBook = ActiveWorkbook
'SuspendAutoCalc
    For Each wSheet In wBook.Worksheets
        wSheet.Activate
        Range("A1").Select
        If wSheet.Name = "Hidden Graphs -->" Then
            hideSheet = True
        End If
        If hideSheet Then
            wSheet.Visible = xlSheetHidden
        End If
    Next
    
    ActiveWorkbook.Sheets(1).Activate
'ResumeAutoCalc
End Sub





Public Sub MoveSlicer(ws As String, pivot As String, slicer As String)
    Dim wsPT As Worksheet
    Dim pt As PivotTable
    Dim sh As Shape
    Dim rngSh As Range
    Dim lColPT As Long
    Dim lCol As Long
    Dim lPad As Long
    Dim tRow As Long
    
'    Set wsPT = Worksheets("Pivot")
'    Set pt = wsPT.PivotTables("ExcSpend")
'    Set sh = wsPT.Shapes("AccName")
    
    Set wsPT = Worksheets(ws)
    Set pt = wsPT.PivotTables(pivot)
    Set sh = wsPT.Shapes(slicer)
    
    lPad = 10
    lColPT = pt.TableRange2.Columns.Count
    lCol = pt.TableRange2.Columns(lColPT).Column
    tRow = ActiveWindow.ScrollRow ' added line for picking out the top row on the screen

    Set rngSh = wsPT.Cells(tRow, lCol + 1) ' included the top row on the screen to the range setting
        sh.Left = rngSh.Left + lPad
        sh.Top = rngSh.Top ' adjusted the top of the slicer accordingly.
    
    'ActiveWindow.ScrollRow * Selection.Cells.RowHeight
    'cells(Activewindow.Scrollrow,activewindow.ScrollColumn).address

'include this code in the worksheet code window where you have a pivot table
'Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'Dim pivotTableName As String
'Dim slicerName As String
'
'pivotTableName = ""
'slicerName = "" ' with multiple slicers, group the slicers and use the group name
'    Application.Run "HPV_Macros.xlam!MoveSlicer", ActiveSheet.Name, pivotTableName, slicerName
'End Sub
End Sub


Public Function countUnique(rng As Range, CaseSensitive As Integer) As Long
Dim uniqueList() As String
Dim listCount As Long
Dim strSearch As String

On Error GoTo countUnique_err

listCount = 0
ReDim uniqueList(listCount)

For Each c In rng
    If CaseSensitive = vbNo Then
        strSearch = UCase(c.value)
    Else
        strSearch = c.value
    End If
    If Not (elementInArray(uniqueList, strSearch)) Then
        ReDim Preserve uniqueList(listCount)
        uniqueList(listCount) = strSearch
        listCount = listCount + 1
    End If
Next
'countUnique = listCount
countUnique = listCount
countUnique_err:
End Function

Public Sub countUniqueMC()
Dim listCount As Long
Dim rng As Range
Dim CaseSensitive As Integer

On Error GoTo countUniqueMC_err

 CaseSensitive = MsgBox("Would you like your count to be case sensitive?", vbYesNoCancel, "Case Sensitive")
 If CaseSensitive <> vbCancel Then
    Set rng = Selection.SpecialCells(xlCellTypeVisible).Cells
    
    listCount = countUnique(rng, CaseSensitive)
    'countUnique = listCount
    MsgBox "There are " & listCount & " unique values in the selected cells"
 End If
countUniqueMC_err:
End Sub


Public Sub copySum()
Dim rng As Range
Dim total As Single
'Dim dataDest As Variant
On Error GoTo copySumErr
Set rng = Selection.SpecialCells(xlCellTypeVisible).Cells
total = 0

For Each c In rng
    If IsNumeric(c.value) Then
        total = total + c.value
    End If
Next
dataDest = Application.InputBox("Select range to copy selected cells to", Type:=8).Address(external:=True)
If Not (IsEmpty(dataDest)) Then
    Range(dataDest).value = total
End If

copySumErr:

End Sub

Public Sub openMacroFile()
MsgBox "Macro File is now open"

End Sub

'********************************************************************************************
'********************************************************************************************
'********************************************************************************************
'Functions
'********************************************************************************************
'********************************************************************************************
'********************************************************************************************



Function unprotectWorksheet(worksheetList As Object)
If ActiveWorkbook Is Nothing Then
    Exit Function
End If

    Dim ws As Worksheet
   ' Dim worksheetList As Object
    Dim pass As String
    
On Error GoTo errorHandler
'    Set d = CreateObject("scripting.dictionary")
'        d.comparemode = 1
    
'    Set worksheetList = CreateObject("scripting.dictionary")
    counter = 0
    'turn off protection
    
    pass = ""
    
    For Each ws In ActiveWorkbook.Worksheets
        If ws.ProtectContents = True Then
            worksheetList.Add ws.CodeName, True
TryAgain:
            ws.Unprotect (pass)
        End If
    Next
    
endNow:
    unprotectWorksheet = pass
    Exit Function
    
ErrorExit:
    ''On Error Resume Next
    unprotectWorksheet = pass
    Exit Function
errorHandler:
        Debug.Print Err.Number & vbLf & Err.Description
        If Err.Number = 1004 Then
        pass = InputBox("Please enter the protection Password for work sheet " & ws.Name & ". Cancel if you don't know the password")
            If pass <> "" Then
                Resume TryAgain
            Else
                MsgBox "Unable to continue as there are password protected worksheets. Please contact the workbook owner for the password"
                pass = "-1" ' return pass= -1 as flag to advise calling sub that unprotecting has failed
            End If
        End If
    Resume ErrorExit
End Function



Function GetWsFromCodeName(wb As Workbook, CodeName As String) As Excel.Worksheet
' from http://yoursumbuddy.com/using-worksheet-codenames-in-other-workbooks/
' Allows code to address worksheets by codename rather than sheet name or index
' Benefit, sheet code names do not change when sheets are added/removed/rearranged
' downside is sheet codenames cannot be directly referenced within another workbook only from code directly in the same workbook as the code is running
'
If ActiveWorkbook Is Nothing Then
    Exit Function
End If

Dim ws As Excel.Worksheet

On Error GoTo GetWsFromCodeNameErr
For Each ws In wb.Worksheets
    If ws.CodeName = CodeName Then
        Set GetWsFromCodeName = ws
        Exit For
    End If
Next ws

Exit Function
GetWsFromCodeNameErr:
MsgBox "Worksheet : " & CodeName & " not found within workbook : " & wb

End Function

Sub WorksheetPasswordBreaker()
    'Breaks worksheet password protection.
If ActiveWorkbook Is Nothing Then
    Exit Sub
End If

For Each ws In ActiveWorkbook.Worksheets
If ws.ProtectContents = True Then
    Dim i As Integer, j As Integer, k As Integer
        Dim l As Integer, m As Integer, n As Integer
        Dim i1 As Integer, i2 As Integer, i3 As Integer
        Dim i4 As Integer, i5 As Integer, i6 As Integer
        On Error Resume Next
        For i = 65 To 66:
            For j = 65 To 66:
                For k = 65 To 66
                    For l = 65 To 66:
                        For m = 65 To 66:
                            For i1 = 65 To 66
                                For i2 = 65 To 66:
                                    For i3 = 65 To 66:
                                        For i4 = 65 To 66
                                            For i5 = 65 To 66:
                                                For i6 = 65 To 66:
                                                    For n = 32 To 126
                                                        ws.Unprotect Chr(i) & Chr(j) & Chr(k) & _
                                                            Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _
                                                            Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
                                                        If ws.ProtectContents = False Then
                                                            MsgBox "One usable password is " & Chr(i) & Chr(j) & _
                                                                Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & _
                                                                Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
                                                            Exit Sub
                                                        End If
                                                    Next: 'n
                                                Next: '6
                                            Next: 'i5
                                        Next: 'i4
                                    Next: 'i3
                                Next: 'i2
                            Next: 'i1
                        Next: 'm
                    Next: 'l
                Next: 'k
            Next: 'j
        Next: 'i
    End If 'ws.ProtectContents = true
Next 'WS
End Sub


Sub switchFileMode()
    ThisWorkbook.IsAddin = Not ThisWorkbook.IsAddin
    
End Sub


Sub SaveHPV_Macros()
Dim newFile As String
'edit 13/07/2018 4:15
'added _tmp to file name since I switched to using add in from z: drive rather than H: drive
If MsgBox("Have you closed any worksheets with custom worksheet functions?", vbYesNo) = vbYes Then
' newFile = "Z:\HPV_Macros"
' newFile = "C:\Users\campbellr\macro\HPV_Macros_"
 newFile = "C:\Users\campbellr\OneDrive - HealthShare Victoria\Macros\HPV_Macros_"
 Application.DisplayAlerts = False
    ThisWorkbook.IsAddin = True
    ThisWorkbook.SaveAs newFile, xlOpenXMLAddIn
'    ThisWorkbook.Save

    ThisWorkbook.IsAddin = False
    ThisWorkbook.SaveAs newFile, xlOpenXMLWorkbookMacroEnabled
    
'save backup
 newFile = "C:\Users\campbellr\OneDrive - HealthShare Victoria\Macros\HPV_Macros_backup"
 Application.DisplayAlerts = False
    ThisWorkbook.IsAddin = True
    ThisWorkbook.SaveAs newFile, xlOpenXMLAddIn
'    ThisWorkbook.Save


'    updateHPVMacro
 
  newFile = "C:\Users\campbellr\OneDrive - HealthShare Victoria\Macros\HPV_Macros_tmp"
    ThisWorkbook.IsAddin = False
    ThisWorkbook.SaveAs newFile, xlOpenXMLWorkbookMacroEnabled
 Application.DisplayAlerts = True
    
 Call updateHPVMacro
 ThisWorkbook.Close False
End If
    
End Sub


Sub updateHPVMacro()
 '   Shell ("C:\Users\campbellr\OneDrive - HealthShare Victoria\Macros\HPV_Macros_Update.bat")
 'Set Wshshell = CreateObject("WScript.Shell")
 '   strCommand = "Powershell -ExecutionPolicy Bypass -File C:\Users\campbellr\OneDrive - HealthShare Victoria\Macros\HPV_Macro_Update.ps1"
 '   Wshshell.Run strCommand, 2, True '2 =hidden popup, true = wait for script before continuation
MsgBox "Powershell script is currently being blocked. copy the macro file and ribbon file to the H: drive"
    
End Sub


Function setDataType(Optional retVal As Integer)
    intDataType = retVal
    
End Function

Function GetDataType() As Integer
'to get a value from a form must first show the form
'when a button on the form is clicked the form calls setDataType with a number for the button pushed
'setdataType then gives a value to the global variable intDataType
'the function GetDataType then returns the global variable

    frmDataType.Show
    GetDataType = intDataType
End Function

Sub testGetDataType()
    MsgBox GetDataType
End Sub

Public Function elementInArray(arr() As String, info As String) As Boolean
'return true if the array arr contains an element equal to info
Dim i As Long
Dim found As Boolean

found = False
For i = 0 To UBound(arr)
    If arr(i) = info Then
        found = True ' value is already in the array
        Exit For
    End If
Next i
elementInArray = found
End Function

Public Sub turnOnalerts()
    Application.DisplayAlerts = True
End Sub

Function sheetExists(sheetToFind As String) As Boolean
    sheetExists = False
    For Each sheet In Worksheets
        If sheetToFind = sheet.Name Then
            sheetExists = True
            Exit Function
        End If
    Next sheet
End Function

Sub readHyperLink()
For Each lnk In Sheets("Notes").Hyperlinks
'display link address
MsgBox lnk.SubAddress
Next
End Sub


Sub fixMacroRef()
''\\FS-HD-01\Common\Shared\HPV templates\Operational Templates\HPV_Macros.xlam'!
        Cells.Replace What:= _
        "'\\FS-HD-01\Common\Shared\HPV templates\Operational Templates\HPV_Macros.xlam'!" _
        , Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:= _
        False, SearchFormat:=False, ReplaceFormat:=False
        
        'https://hpvau-my.sharepoint.com/personal/r_campbell_healthsharevic_org_au/Documents/HPV_Macros.xlsm'!
        Cells.Replace What:= _
        "'https://hpvau-my.sharepoint.com/personal/r_campbell_healthsharevic_org_au/Documents/HPV_Macros.xlsm'!" _
        , Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:= _
        False, SearchFormat:=False, ReplaceFormat:=False
        ''Z:\HPV_Macros.xlsm'!
        Cells.Replace What:= _
        "'Z:\HPV_Macros.xlsm'!" _
        , Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:= _
        False, SearchFormat:=False, ReplaceFormat:=False
        
        'C:\Users\campbellr\OneDrive - Health Purchasing Victoria\HPV_Macros.xlam'!
        Cells.Replace What:= _
        "'C:\Users\campbellr\OneDrive - Health Purchasing Victoria\HPV_Macros.xlam'!" _
        , Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:= _
        False, SearchFormat:=False, ReplaceFormat:=False
        

End Sub


Sub ResetRightClickMenus()
    Dim myCommandBar As CommandBar
    For Each myCommandBar In Application.CommandBars
        If myCommandBar.Type = msoBarTypePopup Then
            myCommandBar.Reset
            myCommandBar.Enabled = True
        End If
    Next myCommandBar
    Application.DisplayStatusBar = True
End Sub


Sub SplitFile(Optional dataListSheetName As String = "", _
                Optional filterColumn As Integer = 0, _
                Optional filterRow As Integer = 0, _
                Optional filterListSheetName As String = "", _
                Optional outputFileName As String = "", _
                Optional outputFileType As String = "", _
                Optional outputFileSuffix As String = "", _
                Optional outputFilePath As String = "", _
                Optional columnsToExclude As Integer = 0 _
                )

'Dim filterListSheetName As String
'Dim dataListSheetName As String
Dim filterList As Worksheet
Dim dataList As Worksheet
'Dim outputFileName As String
'Dim outputFilePath As String
'Dim outputFileSuffix As String
'Dim outputFileType As String
Dim currentFile As Workbook
Dim newFile As Workbook
Dim newSheet As Worksheet
'Dim filterColumn As Integer ' column in data that contains the selection criteria, ie. supplier name
'Dim columnsToExclude As Integer
Dim fileCount As Integer
Dim strDefaultPath As String




'default output path for delta re-runs
strDefaultPath = "H:\Shared\Operational\DataSystems\SCIT\CommonCatalogue\PROD\01-Processed\ReRun\DeltaFiles"

On Error GoTo SplitFile_Err

If dataListSheetName = "" Then ' if not passed
    dataListSheetName = SelectWorkSheetName("Please Select the Sheet with the Data to Split", "Data List")
End If

If dataListSheetName = "" Then
    Exit Sub
End If

If filterColumn = 0 Then
    filterColumn = SelectAddress("Select the Column to be filtered", "Filter Column", "Column") ' column in data that contains the selection criteria, ie. supplier name
End If
If filterColumn < 1 Then
    Exit Sub
End If

If filterRow = 0 Then
    filterRow = SelectAddress("Select the first row to be filtered", "Filter Row", "Row") ' column in data that contains the selection criteria, ie. supplier name
End If
If filterRow < 1 Then
    Exit Sub
End If

If filterListSheetName = "" Then
    filterListSheetName = SelectWorkSheetName("Please Select the Sheet with the list to filter by, Will use first column", "Filter List")
End If
If filterListSheetName = "" Then
    Exit Sub
End If

If columnsToExclude = 0 Then
    columnsToExclude = InputBox("Select how many columns to exclude from the right hand end of the data", "Excluded Column Count", 0)
End If

If outputFileType = "" Then
    outputFileType = UCase(InputBox("Default output File Suffix, CSV or XLSX or TXT or Worksheet Tabs", "Sufix", "CSV|XLSX|TXT|TAB"))
End If
If outputFileType = "" Or outputFileType = "CSV|XLSX|TXT|TAB" Then
    Exit Sub
End If

If outputFileType <> "TAB" Then
    If outputFileName = "" Then
        outputFileName = InputBox("Default output File prefix", "Prefix", "EGO_")
    End If
    If outputFileName = "" Then
        Exit Sub
    End If
    
    If outputFileSuffix = "" Then
        outputFileSuffix = InputBox("Default output File Suffix, before file extension", "Sufix", Format(Now(), "YYMMDD") & "_HPV.CC")
    End If
    If outputFileSuffix = "" Then
        Exit Sub
    End If
    
    If outputFilePath = "" Then
        outputFilePath = InputBox("Output file Location", "Location", strDefaultPath)
    End If
    If outputFilePath = "" Then
        Exit Sub
    End If
    
    If Right(outputFilePath, 1) <> "\" Then
        outputFilePath = outputFilePath & "\"
    End If
End If

Set currentFile = ThisWorkbook


        Set filterList = ActiveWorkbook.Sheets(filterListSheetName)
        Set dataList = ActiveWorkbook.Sheets(dataListSheetName)
        
'For Each ws In ActiveWorkbook.Sheets
'    If ws.Name = filterListSheetName Then  'sheet with list of suppliers
'        Set filterList = ws
'    End If
'    If ws.Name = dataListSheetName Then
'       Set dataList = ws
'    End If
'Next ws

If filterList Is Nothing Then
    MsgBox "Supplier list not found, looking for sheet name: " & filterListSheetName
    Exit Sub
End If
 
If dataList Is Nothing Then
    MsgBox "Data Sheet not found, looking for sheet name: " & dataListSheetName
    Exit Sub
End If
fileCount = 0
    For Each supplierNameCell In filterList.Range(filterList.Cells(1, 1), filterList.Cells(filterList.Cells.SpecialCells(xlLastCell).Row, 1))
        If Len(supplierNameCell.value) > 0 Then
            supplierName = supplierNameCell.value
     
                'apply filter to data sheet
                'copy contents to new file with supplier name in file name
                dataList.AutoFilterMode = False
    '            dataList.ListObjects("Table13").Range.AutoFilter Field:=1, Criteria1:=supplierName
                dataList.Range(dataList.Cells(filterRow, 1), dataList.Cells.SpecialCells(xlLastCell)).AutoFilter Field:=filterColumn, Criteria1:=supplierName
                'select all
                dataList.Range("A1", dataList.Cells(dataList.Cells.SpecialCells(xlLastCell).Row, dataList.Cells.SpecialCells(xlLastCell).Column - columnsToExclude)).Copy
                
            If outputFileType = "CSV" Or outputFileType = "XLSX" Then
               
                    'create new file
                    Set newFile = Workbooks.Add
                    With newFile
                        Set newSheet = .Sheets(1)
               
                    newSheet.Range("A1").PasteSpecial xlPasteColumnWidths, xlPasteSpecialOperationNone, False, False
                    newSheet.Range("A1").PasteSpecial xlPasteValues, xlPasteSpecialOperationNone, False, False
                    newSheet.Range("A1").PasteSpecial xlPasteFormats, xlPasteSpecialOperationNone, False, False
                    newSheet.Range("A1", newSheet.Cells.SpecialCells(xlLastCell)).EntireColumn.AutoFit
                    
                    Application.DisplayAlerts = False
                    
                    If outputFileType = "CSV" Then
                        .SaveAs fileName:=outputFilePath & outputFileName & Replace(Replace(supplierName, "/", ""), "|", " - ") & outputFileSuffix & ".csv", FileFormat:=xlCSV, ConflictResolution:=xlLocalSessionChanges, Local:=True
                    Else
                       .SaveAs fileName:=outputFilePath & outputFileName & Replace(Replace(supplierName, "/", ""), "|", " - ") & outputFileSuffix & ".xlsx", ConflictResolution:=xlLocalSessionChanges
                    End If
                    
                    Application.DisplayAlerts = True
                    newFile.Close False
                    
                End With
            End If 'not TXT
            
                If outputFileType = "TXT" Then
                    Application.Wait (Now + TimeValue("0:00:01"))
                    Call copyToNotepad(outputFilePath & outputFileName & Replace(Replace(supplierName, "/", ""), "|", " - ") & outputFileSuffix & ".txt")
                End If
                
                If outputFileType = "TAB" Then
                    'Dim newSheet As Worksheet
                    Set newSheet = ActiveWorkbook.Worksheets.Add
                    With newSheet
                        .Name = Left(supplierName, 10)
                        .Range("A1").PasteSpecial xlPasteColumnWidths, xlPasteSpecialOperationNone, False, False
                        .Range("A1").PasteSpecial xlPasteValues, xlPasteSpecialOperationNone, False, False
                        .Range("A1").PasteSpecial xlPasteFormats, xlPasteSpecialOperationNone, False, False
                        .Range("A1", .Cells.SpecialCells(xlLastCell)).EntireColumn.AutoFit
                    End With
                    
                End If
                        
            fileCount = fileCount + 1
       End If
    Next supplierNameCell
dataList.ShowAllData
MsgBox "File Split complete, " & fileCount & " files created"
Exit Sub

SplitFile_Err:
MsgBox "Error occured" & vbCr & Err.Number & vbCr & Err.Description & vbCr & outputFilePath & outputFileName & Replace(Replace(supplierName, "/", ""), "|", " - ") & ".xlsx", vbCritical + vbOKOnly
End Sub

Function SelectWorkSheetName(InputDescription As String, InputTitle As String) As String
Dim ws As Worksheet

Set ws = Application.InputBox(InputDescription, InputTitle, Type:=8).Worksheet
SelectWorkSheetName = ws.Name
End Function

Function SelectAddress(InputDescription As String, InputTitle As String, portion As String) As Long
Dim tempAddress As Range

SelectAddress = 0
Set tmpAddress = Application.InputBox(InputDescription, InputTitle, Type:=8)
If portion = "Column" Then
    SelectAddress = tmpAddress.Column
End If
If portion = "Row" Then
    SelectAddress = tmpAddress.Row
End If
End Function

Sub CSV_SetPipe()
If MsgBox("Switch CSV output separator to |", vbYesNo, "Set CSV |") = vbYes Then
    Shell ("Z:\CSVDelimeterPipe.bat")
    MsgBox "Restart Excel for the change to take effect"
End If

End Sub

Sub CSV_SetComma()

If MsgBox("Switch CSV output separator to ,", vbYesNo, "Set CSV ,") = vbYes Then
    Shell ("Z:\CSVDelimeterComma.bat")
    MsgBox "Restart Excel for the change to take effect"
End If
End Sub




Sub freezeFilter()
 Range("a1", Cells.SpecialCells(xlCellTypeLastCell)).Select
    Selection.AutoFilter
'    Selection.EntireColumn.AutoFit
    Range("a1").Select
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
End Sub


Function IsWorkBookOpen(Name As String) As Boolean
    Dim xWb As Workbook
    On Error Resume Next
    Set xWb = Application.Workbooks.Item(Name)
    IsWorkBookOpen = (Not xWb Is Nothing)
End Function
 
Function openFile(Path As String, fileName As String)
    Dim xRet As Boolean
    xRet = IsWorkBookOpen(fileName)
    If xRet Then
        
    Else
        Application.Workbooks.Open (Path & fileName)
    End If
End Function

Sub copyToNotepad(fileName As String) ', textToSend As String)

With Application
'Selection.Copy
Shell "Notepad.exe ", vbNormalFocus
Application.Wait (Now + TimeValue("0:00:01"))
SendKeys " ", True
Application.Wait (Now + TimeValue("0:00:01"))
SendKeys "^v", True
Application.Wait (Now + TimeValue("0:00:01"))
Application.SendKeys "^s", True
Application.Wait (Now + TimeValue("0:00:01"))
Application.SendKeys fileName, True
Application.SendKeys "{ENTER}", True
Application.Wait (Now + TimeValue("0:00:01"))
'VBA.AppActivate .Caption
'.CutCopyMode = False
Application.Wait (Now + TimeValue("0:00:01"))
Application.SendKeys "%{F4}", True
Application.Wait (Now + TimeValue("0:00:01"))
SendKeys " ", True
End With

End Sub

Public Sub send_email(emailTo As String, _
                        emailCC As String, _
                        emailFrom As String, _
                        emailSubject As String, _
                        emailMsg As String, _
                        emailToName As String, _
                        Optional emailAtt As String _
                        )
'send e-mail without the need for outlook
'need to turn on reference: microsoft CDO for windows 2000 library
   Dim NewMail As CDO.message
   Dim mailConfig As CDO.Configuration
   Dim fields As Variant
   Dim msConfigURL As String
   On Error GoTo Err:

   'early binding
   Set NewMail = New CDO.message
   Set mailConfig = New CDO.Configuration

   'load all default configurations
   mailConfig.Load -1

   Set fields = mailConfig.fields
    
   'Set All Email Properties
   With NewMail
        .From = emailFrom
        .To = emailTo
        .CC = emailCC
        .BCC = emailFrom
        .Subject = emailSubject
        .TextBody = emailMsg
        .HTMLBody = emailMsg
        If emailAtt <> "" Then
            .Addattachment outFilePath & outFileName
        End If
   End With

   msConfigURL = "http://schemas.microsoft.com/cdo/configuration"

   With fields
       .Item(msConfigURL & "/smtpusessl") = True 'Enable SSL Authentication
       .Item(msConfigURL & "/smtpauthenticate") = 1 'SMTP authentication Enabled
       .Item(msConfigURL & "/smtpserver") = "healthsharevic-org-au.mail.protection.outlook.com" 'Set the SMTP server details
       .Item(msConfigURL & "/smtpserverport") = 25 'Set the SMTP port Details
       .Item(msConfigURL & "/sendusing") = 2 'Send using default setting
       .Item(msConfigURL & "/sendusername") = emailFrom 'Your gmail address
       '.Item(msConfigURL & "/sendpassword") = "password" 'Your password or App Password
       .Update 'Update the configuration fields
   End With
   NewMail.Configuration = mailConfig
   NewMail.Send
   
   'MsgBox "Your email has been sent", vbInformation

Exit_Err:
   'Release object memory
   Set NewMail = Nothing
   Set mailConfig = Nothing
   Exit Sub

Err:
   Select Case Err.Number
   Case -2147220973 'Could be because of Internet Connection
       MsgBox "Check your internet connection." & vbNewLine & Err.Number & ": " & Err.Description
   Case -2147220975 'Incorrect credentials User ID or password
       MsgBox "Check your login credentials and try again." & vbNewLine & Err.Number & ": " & Err.Description
   Case Else 'Report other errors
       MsgBox "Error encountered while sending email." & vbNewLine & Err.Number & ": " & Err.Description
   End Select

   Resume Exit_Err

End Sub


Sub DBRefresh(Optional quiet As Boolean)
SuspendAutoCalc
Dim wb As Workbook
Dim bgr As Boolean
Set wb = ActiveWorkbook
On Error Resume Next
For Each conn In wb.Connections
    bgr = conn.OLEDBConnection.BackgroundQuery
    conn.OLEDBConnection.BackgroundQuery = False
    conn.OLEDBConnection.Refresh
    conn.OLEDBConnection.BackgroundQuery = bgr
Next conn
ResumeAutoCalc
'refresh pivot tables
Application.Calculate
Call PivotRefresh
If Not quiet Then
    MsgBox "Database Extract Refresh Complete"
End If
End Sub
