Attribute VB_Name = "Concat_Files"
Sub ConcatenateSalesData()
'Purpose: Copy the contents of all xls and xlsx files in a folder into one xls or xlsx file.
'By Ron Campbell
'Last Updated: 11/07/2012
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
  
  ' advise user what they are about to do. Hitting No button will stop the macro running

On Error GoTo errorHandler
  If MsgBox("You are about to import all of the excel files in a folder into one excel file." & vbCr & _
            "First select the file to import the data into then a folder with the data to be copied ." & vbCr & _
            "Create a new excel file with right click - new - excel file" & vbCr & _
            "Do you wish to continue?", vbYesNo) = vbNo Then
    Exit Sub
  End If
  
  'outFile = "c:\workingData\testout.xlsx"
  'set default starting path
  sPath = "H:\Shared\Operational\HSSDA (P049)\P049 HSSDA\05 Delivery\HPV Sales Data"
  outFile = GetOutputFile(sPath)
    If Len(outFile) < 5 Then
        Exit Sub
    End If
    
  sPath = GetFolder(outFile)
  '"Z:\SalesData-dirty\2013-14"
  'sPath = GetFolder("Z:\SalesData-dirty\2013-14")
  
  
  If Len(sPath) < 5 Then
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
          Application.StatusBar = "Processing" & String(progCount Mod 10, ".")
          'turn off the filter
          ActiveSheet.AutoFilterMode = False
          
          'find the last row in the target file with data.
          lMaxTargetRow = wTarget.Cells.SpecialCells(xlLastCell).Row
          
          'Insert File and sheet name before data
          wTarget.Cells(lMaxTargetRow + 1, 1).value = "File: " & wbkSource.Name & " - Sheet: " & wSource.Name
          'copy the source sheet - all used cells into target sheet starting 1 line below last entry
          'This line does copy and paste
          'wSource.UsedRange.Copy wTarget.Cells(lMaxTargetRow + 2, 1)
          
          'This block does copy paste values leaving behind formlas
          wSource.UsedRange.Copy
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



Function GetFolder(strPath As String) As String
'
'code taken from an Excel forum and changed slightly
'Returns the full path for a folder from a dialog box
'
Dim fldr As FileDialog
Dim sItem As String
Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
With fldr
    .title = "Select a Folder"
    .AllowMultiSelect = False
    .InitialFileName = strPath
    If .Show <> -1 Then GoTo NextCode
    sItem = .SelectedItems(1)
End With
NextCode:
GetFolder = sItem & "\"
Set fldr = Nothing


        

End Function

Function GetOutputFile(strPath As String) As String
'
'code taken from an Excel forum and changed slightly
'Returns a file name from a dialog box

Dim outFile As FileDialog
Dim sItem As String
Set outFile = Application.FileDialog(msoFileDialogFilePicker)
With outFile
    .title = "Select an output File"
    .AllowMultiSelect = False
    .InitialFileName = strPath
    If .Show <> -1 Then GoTo NextCode
    sItem = .SelectedItems(1)
End With
NextCode:
GetOutputFile = sItem
Set fldr = Nothing


End Function

Sub setupSD_manual()
setupSD
End Sub
Sub setupSD(Optional shName As Worksheet)
'
' When making a new data set Add two columns and put in a line count in the first and the file name in the second
'
If ActiveWorkbook Is Nothing Then
    Exit Sub
End If

If shName Is Nothing Then
   Set shName = ActiveSheet
End If


Dim endRow As Long
Dim endCol As Long

    shName.Activate
    Columns("A:B").Select
    Selection.Insert Shift:=xlToRight
    Range("A3").value = "sort"
    Range("B3").value = "file"
    Range("B4").value = Range("$C$1")
    
    Range("A4").value = "1"
    Range("A5").Formula = "=A4+1"
    Range("B5").Formula = "=IF(LEFT(C5,LEN($c$1))=$c$1,C5,B4)"
    
    endRow = ActiveCell.SpecialCells(xlLastCell).Row
    endCol = ActiveCell.SpecialCells(xlLastCell).Column
    
    Range(Cells(3, 1), Cells(3, endCol)).WrapText = True
    Range(Cells(3, 1), Cells(3, endCol)).Style = "Good"
    
    Range("A5:B5").Select
    Range("B5").Activate
    Selection.Copy
    
    Range(Cells(endRow, 1), Cells(endRow, 2)).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("C2").Activate
    
End Sub

