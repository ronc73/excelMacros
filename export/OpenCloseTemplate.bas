Attribute VB_Name = "OpenCloseTemplate"

Sub prepareScreen(startStop As Boolean, Optional switch As Integer)
Dim onoff As Boolean
Dim showLevel As Integer
Dim lastRow As Long
Dim formulaRemovalRowFrom As Integer
Dim formulaRemovalRowTo As Integer
Dim formulaColumnFrom As Integer
Dim formulaColumnTo As Integer

If startStop Then
    onoff = True
    showLevel = 2
Else
    onoff = False
    showLevel = 1
End If
'set sheet protection
'reduce to outline level 1
'replace values in some columns
Application.Worksheets("Input Template").Activate
unlockWorksheet ("hpv")
Application.Worksheets("Input Template").Outline.ShowLevels rowlevels:=1, columnlevels:=showLevel
Application.Worksheets("Input Template").Activate
ActiveWindow.DisplayOutline = onoff
ActiveWindow.DisplayHeadings = onoff
Application.DisplayFormulaBar = onoff
Application.EnableEvents = False
If startStop = False Then
'Set formulas to values for A9:F<End of data>, T9:Y<end of data>
'AM9:AU<end of data>
    lastRow = ActiveSheet.Cells.SpecialCells(xlLastCell).Row
    firstRow = 226
    With ActiveSheet
        removeFormula firstRow, lastRow, .Range("A1").Column, .Range("F1").Column 'A:F
        removeFormula firstRow, lastRow, .Range("T1").Column, .Range("Y1").Column  'T:Y
        removeFormula firstRow, lastRow, .Range("AK1").Column, .Range("AY1").Column  'AM:AU
        setPercent firstRow, lastRow, .Range("An1").Column, .Range("An1").Column  'AN:AN
        setPercent firstRow, lastRow, .Range("Ap1").Column, .Range("Ap1").Column  'AN:Ap
        setPercent firstRow, lastRow, .Range("AS1").Column, .Range("AS1").Column  'As:As
        unlockCells firstRow, lastRow, .Range("c1").Column, .Range("f1").Column  'c:f
    End With
    
    If IsNull(switch) Or switch = 0 And sheetExists("TemplateData") Then
        Application.DisplayAlerts = False
        Application.Worksheets("TemplateData").Delete
        Application.DisplayAlerts = True
    End If
    ActiveSheet.Range("C226").Select
End If
If Not startStop Then
    lockWorksheet ("hpv")
End If


Application.Worksheets("Instructions").Activate
ActiveSheet.Range("A1").Select
Application.EnableEvents = True
End Sub


Sub SCTemplateStart()
prepareScreen (True)
End Sub

Sub SCTemplateFinshed()
prepareScreen (False)
End Sub

Sub removeFormula(rowFrom, RowTo, columnFrom, columnTo)
Range(Cells(rowFrom, columnFrom), Cells(RowTo, columnTo)).Copy
Range(Cells(rowFrom, columnFrom), Cells(RowTo, columnTo)).PasteSpecial xlPasteValues
End Sub

'
'Sub removeFormula(rowFrom, RowTo, columnFrom, columnTo)
'too slow compared to built in copy paste values
'For Each cell In Range(Cells(rowFrom, columnFrom), Cells(RowTo, columnTo))
'    cell.Formula = cell.Value
'
'Next
'End Sub


Sub setPercent(rowFrom, RowTo, columnFrom, columnTo)
Range(Cells(rowFrom, columnFrom), Cells(RowTo, columnTo)).Style = "Percent"
End Sub

Sub unlockCells(rowFrom, RowTo, columnFrom, columnTo)
Range(Cells(rowFrom, columnFrom), Cells(RowTo, columnTo)).Locked = False
End Sub

Public Sub lockWorksheet(unlockPasswd As String)
'
' lockWorksheet Macro
'
   ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFiltering:=True, Password:=unlockPasswd
   
End Sub

Public Sub unlockWorksheet(unlockPasswd As String)
'If ActiveSheet.Protect Then
    ActiveSheet.Unprotect Password:=unlockPasswd
'End If
End Sub

Sub fixPercentages()
Dim showLevel As Integer
Dim lastRow As Long

prepareScreen True
Application.Worksheets("Input Template").Activate
    lastRow = ActiveSheet.Cells.SpecialCells(xlLastCell).Row
    firstRow = 226
    With ActiveSheet
        removeFormula firstRow, lastRow, .Range("A1").Column, .Range("F1").Column 'A:F
        removeFormula firstRow, lastRow, .Range("T1").Column, .Range("Y1").Column  'T:Y
        removeFormula firstRow, lastRow, .Range("AL1").Column, .Range("Aw1").Column  'AM:AU
        setPercent firstRow, lastRow, .Range("An1").Column, .Range("An1").Column  'AN:AN
        setPercent firstRow, lastRow, .Range("Ap1").Column, .Range("Ap1").Column  'AN:Ap
        setPercent firstRow, lastRow, .Range("AS1").Column, .Range("AS1").Column  'As:As
        unlockCells firstRow, lastRow, .Range("c1").Column, .Range("f1").Column  'c:f
    End With

prepareScreen False, 1
End Sub


Sub CollateCategorisationSheets()
'Purpose: Copy the contents of all of the categorisation template sheets from each file in a folder into a new file
'By Ron Campbell
'Last Updated: 6/10/2016
'
'Process: Requests a folder from the user. (Source Folder)
'       Requests an excel file from the user (Destination File)
'       systematically goes through each sheet in each excel file in the source folder copying contents to destination file
'
'
'
'
' Path - modify as needed but keep trailing backslash


  Dim sPath As String
  Dim sFile As String
  Dim outFile As String
  Dim wbkSource As Workbook
  Dim wSource As Worksheet
  Dim wTarget As Worksheet
  Dim lMaxSourceRow As Long
  Dim lMaxTargetRow As Long
  Dim sourceStartLine As Integer
  Dim sourceLastRow As Long
  
  sourceStartLine = 225
  ' advise user what they are about to do. Hitting No button will stop the macro running

On Error GoTo errorHandler
  If MsgBox("You are about to import all of the excel files in a folder into one excel file." & vbCr & _
            "First select a folder with the data to be copied then the file to import the data into." & vbCr & _
            "Create a new excel file with right click - new - excel file" & vbCr & _
            "Do you wish to continue?", vbYesNo) = vbNo Then
    Exit Sub
  End If
  
  'outFile = "c:\workingData\testout.xlsx"
  'set default starting path
  'sPath = GetFolder("H:\Shared\Operational\Tender Program\Reporting - Current Contracts\")
  sPath = GetFolder("\\FS-HD-01\Common\Shared\Operational\Process Review\34 P049 HS Spend Analysis\05 Delivery\Supplier Categorisation\AllVic")
  '"Z:\SalesData-dirty\2013-14"
  'sPath = GetFolder("Z:\SalesData-dirty\2013-14")
    If Len(sPath) < 5 Then
        Exit Sub
    End If
  outFile = GetOutputFile(sPath)
    If Len(outFile) < 5 Then
        Exit Sub
    End If
  
  'Confirm what the macro is about to do and folders and file selected. again user can stop by pressing No
  If MsgBox("You have Selected to import files from:" & vbCr & sPath & vbCr & _
            "and you have selected " & outFile & " as your destination file. is this correct? Are you ready to continue?", vbYesNo) = vbNo Then
        Exit Sub
  End If
    
  'set source data set to all files in provided folder
  Set wTarget = Workbooks.Open(outFile).Sheets(1)
  sFile = Dir(sPath & "*.xls*")
  SuspendAutoCalc
    progCount = 0
'process all files in the folder
  Do While Not sFile = ""
    Set wbkSource = Workbooks.Open(fileName:=sPath & sFile, UpdateLinks:=0, AddToMRU:=False)
'               updatelinks Value Meaning
'       0       External references (links) will not be updated when the workbook is opened.
'       3       External references (links) will be updated when the workbook is opened.
    
    SheetCount = 1
    
'process all worksheets in each file
    For Each wSource In wbkSource.Worksheets
        If wSource.Name = "Input Template" Then 'copy the data otherwise ignore. exit for loop afer copying Input Template to dest file
            wSource.Unprotect ("hpv")
              Application.StatusBar = "Processing" & String(progCount Mod 10, ".")
              'turn off the filter
              wSource.Activate
              wSource.AutoFilterMode = False
              
              'find the last row in the target file with data.
              lMaxTargetRow = wTarget.Cells.SpecialCells(xlLastCell).Row
              sourceLastRow = wSource.Cells.SpecialCells(xlLastCell).Row
              'Insert File and sheet name before data
              wTarget.Cells(lMaxTargetRow + 1, 1).value = "File: " & wbkSource.Name & " - Sheet: " & wSource.Name
              'copy the source sheet - all used cells into target sheet starting 1 line below last entry
              'This line does copy and paste
              'wSource.UsedRange.Copy wTarget.Cells(lMaxTargetRow + 2, 1)
              
              'This block does copy paste values leaving behind formlas
              wSource.Range(Cells(sourceStartLine, 1), Cells(sourceLastRow, 38)).Copy
              wTarget.Cells(lMaxTargetRow + 2, 1).PasteSpecial xlPasteValues
    '          Selection.PasteSpecial Paste:=xlPasteValues, _
    '            Operation:=xlNone, _
    '            SkipBlanks:=False, _
    '            Transpose:=False
              Application.CutCopyMode = False
              'end of copy paste values
              
              
            s = s & sFile & " - Sheet #" & SheetCount & vbCr & vbLf
            SheetCount = SheetCount + 1
            progCount = progCount + 1
            Exit For
        End If
    Next
    wbkSource.Close SaveChanges:=False
    sFile = Dir
  Loop
  wTarget.Activate
  'find the last row in the target file with data.
  lMaxTargetRow = wTarget.Cells.SpecialCells(xlLastCell).Row
  wTarget.Cells(lMaxTargetRow + 1, 1).value = s
  MsgBox "Process Complete" & vbCrLf & s
  Application.StatusBar = ""
  setupSD wTarget
  'restore the user's calculation setting
    ResumeAutoCalc
  Exit Sub
  
ErrorExit:
    
    'restore the user's calculation setting
    ResumeAutoCalc
    MsgBox "Macro Failed: Last file processed: " + vbCr + s
    'the ErrorHandler code should only be executed if there is an error
    Exit Sub
errorHandler:
        Debug.Print Err.Number & vbLf & Err.Description
        MsgBox "An error occured during the copy process. " & vbCr & _
        "The last file open was : " & wbkSource.Name & " - Sheet: " & wSource.Name & Err.Number & vbLf & Err.Description
        
        Resume ErrorExit
        


End Sub


