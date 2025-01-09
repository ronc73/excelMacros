Attribute VB_Name = "TRW"
'Attribute VB_Name = "TRW"
Sub ConcatenateTRW()
'Purpose: Copy the contents of all xls and xlsx files in a folder into one xls or xlsx file.
'By Ron Campbell
'Last Updated: 27/07/2021
'
'Process: Requests a folder from the user. (Source Folder)
'       Requests an excel file from the user (Destination File)
'       systematically goes through each sheet in each excel file in the source folder copying contents to destination file
'       For TRWs with Clinical and commercial sheet
'       Copy all content of commercial sheets into one commercial sheet
'       copy all Clinical X sheets to Clinical X

  Dim sPath         As String
  Dim sFile         As String
  Dim outFile       As String
  Dim wbkSource     As Workbook
  Dim wbkDest       As Workbook
  Dim wSource       As Worksheet
  Dim wTarget       As Worksheet
  Dim lMaxSourceRow As Long
  Dim lMaxSourceCol As Long
  Dim lMaxTargetRow As Long
  Dim shDestName    As String
  Dim firstRun      As Boolean
  Dim headerRows    As Integer
  Dim headerRowsAdditional    As Integer
  
  
  ' advise user what they are about to do. Hitting No button will stop the macro running
firstRun = True
headerRows = 13
headerRowsAdditional = 1


On Error GoTo errorHandler
  If MsgBox("You are about to import all of the excel files in a folder into one excel file." & vbCr & _
            "First select the file to import the data into then a folder with the data to be copied ." & vbCr & _
            "Create a new excel file with right click - new - excel file" & vbCr & _
            "Do you wish to continue?", vbYesNo) = vbNo Then
    Exit Sub
  End If
  
  sPath = "c:\temp\"
  outFile = GetOutputFile(sPath)
    If Len(outFile) < 5 Then
        Exit Sub
    End If
    
  sPath = GetFolder(sPath)
  If Len(sPath) < 5 Then
        Exit Sub
  End If
  
 
  'Confirm what the macro is about to do and folders and file selected. again user can stop by pressing No
  If MsgBox("You have Selected to import files from:" & vbCr & sPath & vbCr & _
            "and you have selected " & outFile & " as your destination file. is this correct? Are you ready to continue?", vbYesNo) = vbNo Then
        Exit Sub
  End If
    
    sFile = Dir(sPath & "*.xls*")
    Set wbkDest = Workbooks.Open(outFile)
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
    Call unHideAllSheets(wbkSource)
    Call unhideAllDataAllSheets(wbkSource)
    For Each wSource In wbkSource.Worksheets
          Application.StatusBar = "Processing" & String(progCount Mod 10, ".")
          ActiveSheet.AutoFilterMode = False
          
        If wSource.Name Like "Category*(Commercial)" Then
            shDestName = "Commercial"
        ElseIf wSource.Name Like "Category*(Clinical)" Then
            shDestName = wSource.Name
        Else
            GoTo skipSheet
        End If
        Set wTarget = setDestSheet(wbkDest, shDestName)


          'find the last row in the target file with data.
          lMaxTargetRow = wTarget.Cells.SpecialCells(xlLastCell).Row
          lMaxSourceRow = wSource.UsedRange.Rows.Count
          lMaxSourceCol = wSource.UsedRange.Columns.Count
          'Insert File and sheet name before data
        If lMaxTargetRow = 1 Then
            lMaxTargetRow = 0
        End If
        wTarget.Cells(lMaxTargetRow + 1, 1).value = "File: " & wbkSource.Name & " - Sheet: " & wSource.Name
            
            If lMaxTargetRow = 0 Then
                'copy the source sheet header otherwise skip the header and copy the data as values
                wSource.Range(wSource.Cells(1, 1), wSource.Cells(headerRows, lMaxSourceCol)).Copy (wTarget.Cells(lMaxTargetRow + 2, 1))
                lMaxTargetRow = lMaxTargetRow + 2 + headerRows
            Else
                lMaxTargetRow = lMaxTargetRow + 2
            End If
        'set destination cells to text format for paste values to not clobber number values
        wTarget.Range(wTarget.Cells(lMaxTargetRow + 1, 1), wTarget.Cells(lMaxSourceRow, lMaxSourceCol)).NumberFormat = "@"
          
          wSource.Range(wSource.Cells(headerRows + 1 + headerRowsAdditional, 1), wSource.Cells(lMaxSourceRow, lMaxSourceCol)).Copy
          wTarget.Cells(lMaxTargetRow, 1).PasteSpecial xlPasteValues
          Application.CutCopyMode = False
          'end of copy paste values
          
          
        s = s & sFile & " - Sheet #" & SheetCount & vbCr & vbLf
        SheetCount = SheetCount + 1
        progCount = progCount + 1
skipSheet:

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
  
    For Each wSource In wbkDest.Sheets
        setupSD wSource
    Next
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


Function setDestSheet(wbkDest As Workbook, newShName As String) As Worksheet
Dim wTarget As Worksheet

On Error Resume Next
    Set wTarget = wbkDest.Sheets(newShName)
On Error GoTo Err_setDestSheet
    If wTarget Is Nothing Then
        Set wTarget = wbkDest.Sheets.Add(After:=wbkDest.Sheets(wbkDest.Sheets.Count))
            wTarget.Name = newShName
    End If

Err_setDestSheet:
    
Exit_setDestSheet:
Set setDestSheet = wTarget
End Function


Sub TRW_Cleanup()

Dim headerRow As Integer

headerRow = 6
For Each ws In ActiveWorkbook.Sheets
'after c1 cell has been reduced to basic filename
'replace formulas with values for columns a and b
'filter to remove data where column c is blank, contains file:* or End or |||
ws.Activate
ws.AutoFilterMode = False
With ws.Range(ws.Cells(1, 1), ws.Cells(ws.Cells.SpecialCells(xlLastCell).Row, 2))
    .Copy
    .PasteSpecial xlPasteValues
End With

With ws.Range(ws.Cells(headerRow, 1), ws.Cells.SpecialCells(xlLastCell))
    .AutoFilter Field:=3, Criteria1:=Array("0", "", "|||", "END"), Operator:=xlFilterValues
    ws.Range(ws.Cells(headerRow + 1, 1), ws.Cells.SpecialCells(xlLastCell)).EntireRow.Delete
    .AutoFilter Field:=3, Field:=3, Criteria1:="File: *"
    ws.Range(ws.Cells(headerRow + 1, 1), ws.Cells.SpecialCells(xlLastCell)).EntireRow.Delete
End With
'remove lines below the header row
ws.AutoFilterMode = False
Next

End Sub


Sub Combine_All_Worksheets_to_one()
'Purpose: Copy the contents of all xls and xlsx files in a folder into one xls or xlsx file.
'By Ron Campbell
'Last Updated: 27/07/2021
'
'Process:   systematically goes through each sheet in the current excel file copying contents to new sheet

  Dim wbkSource     As Workbook
  Dim wbkDest       As Workbook
  Dim wSource       As Worksheet
  Dim wTarget       As Worksheet
  Dim lMaxSourceRow As Long
  Dim lMaxSourceCol As Long
  Dim lMaxTargetRow As Long
  Dim firstRun      As Boolean
  Dim headerRows    As Integer
  Dim shDestName    As String
  
  
  ' advise user what they are about to do. Hitting No button will stop the macro running
firstRun = True
headerRows = 2

On Error GoTo errorHandler
  
  SuspendAutoCalc
    progCount = 0
'process all files in the folder
 
    Set wbkSource = Application.ActiveWorkbook
'               updatelinks Value Meaning
'       0       External references (links) will not be updated when the workbook is opened.
'       3       External references (links) will be updated when the workbook is opened.
    SheetCount = 1
    shDestName = "Consolidated"
'process all worksheets in each file
    For Each wSource In wbkSource.Worksheets
          Application.StatusBar = "Processing" & String(progCount Mod 10, ".")
          ActiveSheet.AutoFilterMode = False
        
        Set wTarget = setDestSheet(wbkSource, shDestName)


          'find the last row in the target file with data.
          lMaxTargetRow = wTarget.Cells.SpecialCells(xlLastCell).Row
          lMaxSourceRow = wSource.UsedRange.Rows.Count
          lMaxSourceCol = wSource.UsedRange.Columns.Count
          'Insert File and sheet name before data
        If lMaxTargetRow = 1 Then
            lMaxTargetRow = 0
        End If
        wTarget.Cells(lMaxTargetRow + 1, 1).value = "File: " & wbkSource.Name & " - Sheet: " & wSource.Name
            
            If lMaxTargetRow = 0 Then
                'copy the source sheet header otherwise skip the header and copy the data as values
                wSource.Range(wSource.Cells(1, 1), wSource.Cells(headerRows, lMaxSourceCol)).Copy (wTarget.Cells(lMaxTargetRow + 2, 1))
                lMaxTargetRow = lMaxTargetRow + 2 + headerRows
            Else
                lMaxTargetRow = lMaxTargetRow + 2
            End If
        'set destination cells to text format for paste values to not clobber number values
        wTarget.Range(wTarget.Cells(lMaxTargetRow + 1, 1), wTarget.Cells(lMaxSourceRow, lMaxSourceCol)).NumberFormat = "@"
          
          wSource.Range(wSource.Cells(headerRows + 1, 1), wSource.Cells(lMaxSourceRow, lMaxSourceCol)).Copy
          wTarget.Cells(lMaxTargetRow, 1).PasteSpecial xlPasteValues
          Application.CutCopyMode = False
          'end of copy paste values
          
          
        s = s & sFile & " - Sheet #" & SheetCount & vbCr & vbLf
        SheetCount = SheetCount + 1
        progCount = progCount + 1
skipSheet:

    Next
 
  wTarget.Activate
  'find the last row in the target file with data.
  lMaxTargetRow = wTarget.Cells.SpecialCells(xlLastCell).Row
  wTarget.Cells(lMaxTargetRow + 1, 1).value = s
  MsgBox "Process Complete" & vbCrLf & s
  Application.StatusBar = ""
  
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


'Attribute VB_Name = "TRW"
Sub ConcatenateTRW_02()
'Purpose: Copy the contents of all xls and xlsx files in a folder into one xls or xlsx file.
'By Ron Campbell
'Last Updated: 27/07/2021
'
'Process: Requests a folder from the user. (Source Folder)
'       Requests an excel file from the user (Destination File)
'       systematically goes through each sheet in each excel file in the source folder copying contents to destination file
'       For TRWs with Clinical and commercial sheet
'       Copy all content of commercial sheets into one commercial sheet
'       copy all Clinical X sheets to Clinical X

  Dim sPath         As String
  Dim sFile         As String
  Dim outFile       As String
  Dim wbkSource     As Workbook
  Dim wbkDest       As Workbook
  Dim wSource       As Worksheet
  Dim wTarget       As Worksheet
  Dim lMaxSourceRow As Long
  Dim lMaxSourceCol As Long
  Dim lMaxTargetRow As Long
  Dim shDestName    As String
  Dim firstRun      As Boolean
  Dim headerRows    As Integer
  Dim headerRowsAdditional    As Integer
  
  
  ' advise user what they are about to do. Hitting No button will stop the macro running
firstRun = True
headerRows = 13
headerRowsAdditional = 1


On Error GoTo errorHandler
  If MsgBox("You are about to import all of the excel files in a folder into one excel file." & vbCr & _
            "First select the file to import the data into then a folder with the data to be copied ." & vbCr & _
            "Create a new excel file with right click - new - excel file" & vbCr & _
            "Do you wish to continue?", vbYesNo) = vbNo Then
    Exit Sub
  End If
  
  sPath = "c:\temp\"
  outFile = GetOutputFile(sPath)
    If Len(outFile) < 5 Then
        Exit Sub
    End If
    
  sPath = GetFolder(sPath)
  If Len(sPath) < 5 Then
        Exit Sub
  End If
  
 
  'Confirm what the macro is about to do and folders and file selected. again user can stop by pressing No
  If MsgBox("You have Selected to import files from:" & vbCr & sPath & vbCr & _
            "and you have selected " & outFile & " as your destination file. is this correct? Are you ready to continue?", vbYesNo) = vbNo Then
        Exit Sub
  End If
    
    sFile = Dir(sPath & "*.xls*")
    Set wbkDest = Workbooks.Open(outFile)
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
    Call unHideAllSheets(wbkSource)
    Call unhideAllDataAllSheets(wbkSource)
    For Each wSource In wbkSource.Worksheets
          Application.StatusBar = "Processing" & String(progCount Mod 10, ".")
          ActiveSheet.AutoFilterMode = False
          
        If wSource.Name Like "Category*(Commercial)" Then
            shDestName = "Commercial"
        ElseIf wSource.Name Like "Category*(Clinical)" Then
            shDestName = wSource.Name
        Else
            GoTo skipSheet
        End If
        Set wTarget = setDestSheet(wbkDest, shDestName)


          'find the last row in the target file with data.
          lMaxTargetRow = wTarget.Cells.SpecialCells(xlLastCell).Row
          lMaxSourceRow = wSource.UsedRange.Rows.Count
          lMaxSourceCol = wSource.UsedRange.Columns.Count
          'Insert File and sheet name before data
        If lMaxTargetRow = 1 Then
            lMaxTargetRow = 0
        End If
 '       wTarget.Cells(lMaxTargetRow + 1, 1).value = "File: " & wbkSource.Name & " - Sheet: " & wSource.Name
            
            If lMaxTargetRow = 0 Then
                'copy the source sheet header otherwise skip the header and copy the data as values
                wSource.Range(wSource.Cells(1, 1), wSource.Cells(headerRows, lMaxSourceCol)).Copy (wTarget.Cells(lMaxTargetRow + 2, 1))
                lMaxTargetRow = lMaxTargetRow + 2 + headerRows
            Else
                lMaxTargetRow = lMaxTargetRow + 2
            End If
        'set destination cells to text format for paste values to not clobber number values
        wTarget.Range(wTarget.Cells(lMaxTargetRow + 1, 1), wTarget.Cells(lMaxSourceRow, lMaxSourceCol)).NumberFormat = "@"
          
          wSource.Range(wSource.Cells(headerRows + 1 + headerRowsAdditional, 1), wSource.Cells(lMaxSourceRow, lMaxSourceCol)).Copy
          wTarget.Cells(lMaxTargetRow, 1).PasteSpecial xlPasteAll 'xlPasteValues
          Application.CutCopyMode = False
          'end of copy paste values
          
          
        s = s & sFile & " - Sheet #" & SheetCount & vbCr & vbLf
        SheetCount = SheetCount + 1
        progCount = progCount + 1
skipSheet:

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
  
'    For Each wSource In wbkDest.Sheets
'        setupSD wSource
'    Next
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

