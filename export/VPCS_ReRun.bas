Attribute VB_Name = "VPCS_ReRun"

Sub FileArrange_Item()
'step 1
'copy each section of the output file into a separate sheet
'step 2
'insert columns where necessary to fit final template
'Item output:
'Section 1
'MANUFACTURER_NAME  | MANUFACTURER_PART_NUM | MANUFACTURER_GLN | GTIN_BASE_UNIT |OHIS_HPV_CAT_DESCRIPTION | OHIS_HPV_SUBCAT_DESCRIPTION | OHIS_UNSPSC | OHIS_UNSPSC_DESCRIPTION | STATUS | ERROR_CODE

'Section 2
'MANUFACTURER_NAME | MANUFACTURER_PART_NUM | MANUFACTURER_GLN | GTIN_BASE_UNIT | OHIS_HPV_SUBCAT_DESCRIPTION | OHIS_UNSPSC_DESCRIPTION|BASE_GTIN|ITEM_NUMBER| STATUS| ERROR_CODE
'
'Stats Block:
'FILE_NAME | STATUS | ERROR_CODE | COUNT
'
'Lines to remove
'
'=========================  VALIDATE ITEMS  =========================

'--  VALIDATE_ITEMS  --
'========================================
'Record Count:

Dim section01Marker As String
Dim section02Marker As String
Dim section02Markera As String
Dim statsBlockMarker As String
Dim linesToRemove() As String
Dim copyDest As Worksheet
Dim copySource As Worksheet
Dim ws01 As Worksheet
Dim ws02 As Worksheet
Dim wsItemOutput As Worksheet
Dim wsStats As Worksheet
Dim itemHeader As String

itemHeader = "File|MANUFACTURER_NAME|MANUFACTURER_PART_NUM|MANUFACTURER_GLN|GTIN_BASE_UNIT|OHIS_HPV_CAT_DESCRIPTION|OHIS_HPV_SUBCAT_DESCRIPTION|OHIS_UNSPSC|OHIS_UNSPSC_DESCRIPTION|ITEM_NUMBER|STATUS|ERROR_CODE"

Set copySource = ActiveSheet

Set ws01 = ActiveWorkbook.Sheets.Add(After:=Worksheets(Worksheets.Count))
    ws01.Name = "Sect 01"
Set ws02 = ActiveWorkbook.Sheets.Add(After:=Worksheets(Worksheets.Count))
    ws02.Name = "Sect 02"
Set wsItemOutput = ActiveWorkbook.Sheets.Add(After:=Worksheets(Worksheets.Count))
    wsItemOutput.Name = "Item Output"
Set wsStats = ActiveWorkbook.Sheets.Add(After:=Worksheets(Worksheets.Count))
    wsStats.Name = "Stats"

    
linesToRemove() = Split("=========================  VALIDATE ITEMS  =========================," & _
                    "--  VALIDATE_ITEMS  --," & _
                    "========================================,", ",")
section01Marker = "MANUFACTURER_NAME  | MANUFACTURER_PART_NUM | MANUFACTURER_GLN | GTIN_BASE_UNIT |OHIS_HPV_CAT_DESCRIPTION | OHIS_HPV_SUBCAT_DESCRIPTION | OHIS_UNSPSC | OHIS_UNSPSC_DESCRIPTION | STATUS | ERROR_CODE"
section02Marker = "MANUFACTURER_NAME | MANUFACTURER_PART_NUM | MANUFACTURER_GLN | GTIN_BASE_UNIT | OHIS_HPV_SUBCAT_DESCRIPTION | OHIS_UNSPSC_DESCRIPTION|BASE_GTIN|ITEM_NUMBER| STATUS| ERROR_CODE | ERROR_DESC"
section02Markera = "MANUFACTURER_NAME | MANUFACTURER_PART_NUM | MANUFACTURER_GLN | GTIN_BASE_UNIT | OHIS_HPV_SUBCAT_DESCRIPTION | OHIS_UNSPSC_DESCRIPTION|BASE_GTIN|ITEM_NUMBER| STATUS| ERROR_CODE"
statsBlockMarker = "Record Count:"

For Each c In copySource.UsedRange.Cells
    If Trim(c.value) = section01Marker Then
        'copy the following lines to sect 1 sheet
        Set copyDest = ws01
    End If
    If Trim(c.value) = section02Marker Or Trim(c.value) = section02Markera Then
        'copy the following lines to sect 1 sheet
        Set copyDest = ws02
    End If
    If Trim(c.value) = statsBlockMarker Then
        Set copyDest = wsStats
    End If
    If Not elementInArray(linesToRemove(), c.value) Then
        If copyDest Is Nothing Then
        Else
            With copyDest
             If .Cells.SpecialCells(xlCellTypeLastCell).value = "" Then
                .Cells(.Cells.SpecialCells(xlCellTypeLastCell).Row, 1).value = c.value
             Else
                .Cells(.Cells.SpecialCells(xlCellTypeLastCell).Row + 1, 1).value = c.value
             End If
            End With
        End If
    End If
Next c

    Call FileArrange_Item_Split
    Call FileArrange_Header(wsItemOutput, itemHeader)
    Call FileArrange_copySheets(wsItemOutput)
    Call FileArrange_fileName(wsItemOutput, wsStats)
End Sub


Sub FileArrange_GTIN()
'step 1
'copy each section of the output file into a separate sheet
'step 2
'insert columns where necessary to fit final template
'Item output:
'Section 1
'GTIN OF BASE UNIT | GTIN |OHIS UOM |OHIS NUMBER OF EACH |STATUS| ERROR CODE
'
'Section 2
'GTIN_BASE_UNIT | GTIN | MANUFACTURER_NAME | MANUFACTURE_PART_NUM | STATUS | ERROR_CODE
'
'Stats Block:
'FILE_NAME | STATUS | ERROR_CODE | COUNT
'
'Lines to remove
'
'========================================
'Record Count:

Dim section01Marker As String
Dim section02Marker As String
Dim statsBlockMarker As String
Dim linesToRemove() As String
Dim copyDest As Worksheet
Dim copySource As Worksheet
Dim ws01 As Worksheet
Dim ws02 As Worksheet
Dim wsGTINOutput As Worksheet
Dim wsStats As Worksheet
Dim GTINHeader As String

GTINHeader = "File|GTIN OF BASE UNIT|GTIN|OHIS UOM|OHIS NUMBER OF EACH|MANUFACTURER_NAME|MANUFACTURE_PART_NUM|STATUS|ERROR CODE"

Set copySource = ActiveSheet

Set ws01 = ActiveWorkbook.Sheets.Add(After:=Worksheets(Worksheets.Count))
    ws01.Name = "Sect 01"
Set ws02 = ActiveWorkbook.Sheets.Add(After:=Worksheets(Worksheets.Count))
    ws02.Name = "Sect 02"
Set wsGTINOutput = ActiveWorkbook.Sheets.Add(After:=Worksheets(Worksheets.Count))
    wsGTINOutput.Name = "GTIN Output"
Set wsStats = ActiveWorkbook.Sheets.Add(After:=Worksheets(Worksheets.Count))
    wsStats.Name = "Stats"

    
linesToRemove() = Split("========================================,", ",")
section01Marker = "GTIN OF BASE UNIT | GTIN |OHIS UOM |OHIS NUMBER OF EACH |STATUS| ERROR CODE"
section02Marker = "GTIN_BASE_UNIT | GTIN | MANUFACTURER_NAME | MANUFACTURE_PART_NUM | STATUS | ERROR_CODE"
statsBlockMarker = "Record Count:"

For Each c In copySource.UsedRange.Cells
    If Trim(c.value) = section01Marker Then
        'copy the following lines to sect 1 sheet
        Set copyDest = ws01
    End If
    If Trim(c.value) = section02Marker Then
        'copy the following lines to sect 1 sheet
        Set copyDest = ws02
    End If
    If Trim(c.value) = statsBlockMarker Then
        Set copyDest = wsStats
    End If
    If Not elementInArray(linesToRemove(), c.value) Then
        If copyDest Is Nothing Then
        Else
            With copyDest
             If .Cells.SpecialCells(xlCellTypeLastCell).value = "" Then
                .Cells(.Cells.SpecialCells(xlCellTypeLastCell).Row, 1).value = c.value
             Else
                .Cells(.Cells.SpecialCells(xlCellTypeLastCell).Row + 1, 1).value = c.value
             End If
            End With
        End If
    End If
Next c

    Call FileArrange_GTIN_Split
    Call FileArrange_Header(wsGTINOutput, GTINHeader)
    Call FileArrange_copySheets(wsGTINOutput)
    Call FileArrange_fileName(wsGTINOutput, wsStats)
End Sub


Sub FileArrange_Price()
'step 1
'copy each section of the output file into a separate sheet
'step 2
'insert columns where necessary to fit final template
'Price output:
'Section 1
'
'
'Section 2
'
'
'Stats Block:
'FILE_NAME | STATUS | ERROR_CODE | COUNT
'
'Lines to remove
'========================================
'Report Check ABN
'End of Report Check ABN
'Report Missing GLN
'End of Missing GLN
'=======================================================================================

Dim section01Marker As String
Dim section02Marker As String
Dim statsBlockMarker As String
Dim linesToRemove() As String
Dim copyDest As Worksheet
Dim copySource As Worksheet
Dim ws01 As Worksheet
Dim ws02 As Worksheet
Dim wsPriceOutput As Worksheet
Dim wsStats As Worksheet
Dim PriceHeader As String

PriceHeader = "file|BASE GTIN|OHIS GTIN|SUPPLIER ABN|SUPPLIER|MANUFACTURER NAME|MANUFACTURER PART NUM|MISSING GLN|OHIS UOM|OHIS NUMBER OF EACH|OHIS SUPPLIER ABN|HPV_PART_NUM|STATUS|ERROR CODE"


Set copySource = ActiveSheet

Set ws01 = ActiveWorkbook.Sheets.Add(After:=Worksheets(Worksheets.Count))
    ws01.Name = "Sect 01"
Set ws02 = ActiveWorkbook.Sheets.Add(After:=Worksheets(Worksheets.Count))
    ws02.Name = "Sect 02"
Set ws03 = ActiveWorkbook.Sheets.Add(After:=Worksheets(Worksheets.Count))
    ws03.Name = "Sect 03"
Set ws04 = ActiveWorkbook.Sheets.Add(After:=Worksheets(Worksheets.Count))
    ws04.Name = "Sect 04"
Set wsPriceOutput = ActiveWorkbook.Sheets.Add(After:=Worksheets(Worksheets.Count))
    wsPriceOutput.Name = "Price Output"
Set wsStats = ActiveWorkbook.Sheets.Add(After:=Worksheets(Worksheets.Count))
    wsStats.Name = "Stats"

    
linesToRemove() = Split("========================================," & _
                    "Report Check ABN," & _
                    "End of Report Check ABN," & _
                    "Report Missing GLN," & _
                    "End of Missing GLN," & _
                    "=======================================================================================,", ",")
section01Marker = "BASE GTIN | SUPPLIER ABN | SUPPLIER  | MANUFACTURER NAME | MANUFACTURER PART NUM | STATUS| ERROR CODE"
section02Marker = "BASE GTIN | SUPPLIER ABN | SUPPLIER  | MANUFACTURER NAME | MANUFACTURER PART NUM | MISSING GLN | STATUS| ERROR CODE"
section03Marker = "OHIS GTIN |OHIS UOM |OHIS NUMBER OF EACH | OHIS SUPPLIER ABN |STATUS| ERROR CODE"
section04Marker = "GTIN_BASE_UNIT | GTIN | OHIS_SUPPLIER_ABN | OHIS_SUPPLIER | OHIS_HPV_MANUFACTURER_NAME | OHIS_HPV_MANUFACTURER_PART_NUM | HPV_PART_NUM | STATUS | ERROR_CODE"
statsBlockMarker = "Record Count:"

For Each c In copySource.UsedRange.Cells
    If Trim(c.value) = section01Marker Then
        'copy the following lines to sect 1 sheet
        Set copyDest = ws01
    End If
    If Trim(c.value) = section02Marker Then
        'copy the following lines to sect 1 sheet
        Set copyDest = ws02
    End If
    If Trim(c.value) = section03Marker Then
        'copy the following lines to sect 1 sheet
        Set copyDest = ws03
    End If
    If Trim(c.value) = section04Marker Then
        'copy the following lines to sect 1 sheet
        Set copyDest = ws04
    End If
    If Trim(c.value) = statsBlockMarker Then
        Set copyDest = wsStats
    End If
    If Not elementInArray(linesToRemove(), c.value) Then
        If copyDest Is Nothing Then
        Else
            With copyDest
             If .Cells.SpecialCells(xlCellTypeLastCell).value = "" Then
                .Cells(.Cells.SpecialCells(xlCellTypeLastCell).Row, 1).value = c.value
             Else
                .Cells(.Cells.SpecialCells(xlCellTypeLastCell).Row + 1, 1).value = c.value
             End If
            End With
        End If
    End If
Next c

    Call FileArrange_Price_Split
    Call FileArrange_Header(wsPriceOutput, PriceHeader)
    Call FileArrange_copySheets(wsPriceOutput)
    Call FileArrange_fileName(wsPriceOutput, wsStats)
End Sub



Sub FileArrange_Item_Split()
'step 2
'split columns, insert columns and add header row
Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets
    With ws
    If Left(.Name, 4) = "Sect" Or .Name = "Stats" Then
        ws.Activate
        Range(Cells(1, 1), ws.Cells(ws.Cells.SpecialCells(xlCellTypeLastCell).Row, 1)).Select
        Call TextToCols(2, "|")
        
        If ws.Name = "Sect 01" Then
            ws.Columns("A:A").Select
            Selection.Insert Shift:=xlToRight
            ws.Columns("J:J").Select
            Selection.Insert Shift:=xlToRight
        End If
        If ws.Name = "Sect 02" Then
            ws.Columns("A:A").Select
            Selection.Insert Shift:=xlToRight
            ws.Columns("F:F").Select
            Selection.Insert Shift:=xlToRight
            ws.Columns("H:H").Cut (ws.Columns("I:I"))
        End If

    End If
    End With
Next ws

End Sub
Sub FileArrange_GTIN_Split()
'step 2
'split columns, insert columns and add header row
Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets
    With ws
    If Left(.Name, 4) = "Sect" Or .Name = "Stats" Then
        ws.Activate
        Range(Cells(1, 1), ws.Cells(ws.Cells.SpecialCells(xlCellTypeLastCell).Row, 1)).Select
        Call TextToCols(2, "|")
        
        If ws.Name = "Sect 01" Then
            ws.Columns("A:A").Select
            Selection.Insert Shift:=xlToRight
            ws.Columns("F:G").Select
            Selection.Insert Shift:=xlToRight
        End If
        If ws.Name = "Sect 02" Then
            ws.Columns("A:A").Select
            Selection.Insert Shift:=xlToRight
            ws.Columns("D:E").Select
            Selection.Insert Shift:=xlToRight
        End If

    End If
    End With
Next ws

End Sub
Sub FileArrange_Price_Split()
'step 2
'split columns, insert columns and add header row
Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets
    With ws
    If Left(.Name, 4) = "Sect" Or .Name = "Stats" Then
        ws.Activate
        Range(Cells(1, 1), ws.Cells(ws.Cells.SpecialCells(xlCellTypeLastCell).Row, 1)).Select
        Call TextToCols(2, "|")
        
        If ws.Name = "Sect 01" Then
            ws.Columns("A:A").Select
            Selection.Insert Shift:=xlToRight
            ws.Columns("C:C").Select
            Selection.Insert Shift:=xlToRight
            ws.Columns("B:B").Copy (ws.Columns("C:C"))
            ws.Columns("H:L").Select
            Selection.Insert Shift:=xlToRight
        End If
        If ws.Name = "Sect 02" Then
            ws.Columns("A:A").Select
            Selection.Insert Shift:=xlToRight
            ws.Columns("C:C").Select
            Selection.Insert Shift:=xlToRight
            ws.Columns("B:B").Copy (ws.Columns("C:C"))
            ws.Columns("I:L").Select
            Selection.Insert Shift:=xlToRight
        End If
        If ws.Name = "Sect 03" Then
            ws.Columns("A:A").Select
            Selection.Insert Shift:=xlToRight
            ws.Columns("C:C").Select
            Selection.Insert Shift:=xlToRight
            ws.Columns("B:B").Copy (ws.Columns("C:C"))
            ws.Columns("D:H").Select
            Selection.Insert Shift:=xlToRight
            ws.Columns("L:L").Select
            Selection.Insert Shift:=xlToRight
        End If
        If ws.Name = "Sect 04" Then
            ws.Columns("A:A").Select
            Selection.Insert Shift:=xlToRight
            ws.Columns("H:K").Select
            Selection.Insert Shift:=xlToRight
        End If

    End If
    End With
Next ws

End Sub


Sub FileArrange_Header(ws As Worksheet, header As String)
ws.Range("a1").value = header
ws.Range("a1").TextToColumns Range("a1"), xlDelimited, xlTextQualifierNone, False, False, False, False, False, True, "|"
ws.Range("a1").TextToColumns Range("a1"), xlDelimited, xlTextQualifierNone, False, False, False, False, False, False

        
End Sub


Sub FileArrange_copySheets(wsItemOutput As Worksheet)

For Each ws In ActiveWorkbook.Worksheets
    With ws
        If Left(.Name, 4) = "Sect" Then
            .Activate
            If .Cells.SpecialCells(xlCellTypeLastCell).Row > 1 Then
                If wsItemOutput.Cells(1, 1) = "" Then
                    .Range(Cells(2, 1), .Cells.SpecialCells(xlCellTypeLastCell)).Copy (wsItemOutput.Cells(wsItemOutput.Cells.SpecialCells(xlCellTypeLastCell).Row, 1))
                Else
                    .Range(Cells(2, 1), .Cells.SpecialCells(xlCellTypeLastCell)).Copy (wsItemOutput.Cells(wsItemOutput.Cells.SpecialCells(xlCellTypeLastCell).Row + 1, 1))
                End If
                wsItemOutput.Select
            End If
        End If
    End With
Next ws
End Sub

Sub FileArrange_fileName(wsItemOutput As Worksheet, wsStats As Worksheet)
    wsItemOutput.Activate
    With wsItemOutput
        If .Cells.SpecialCells(xlCellTypeLastCell).Row > 1 Then
            .Range(.Cells(2, 1), Cells(.Cells.SpecialCells(xlCellTypeLastCell).Row, 1)).value = wsStats.Cells(3, 1).value
        End If
    End With
End Sub

Sub FileArrange_FinalFormatting()
'hide top row
'hide vpcs match column
'in success tables hide error columns
Dim wbkConsolidated As Workbook
Call unhideAllDataAllSheets

Set wbkConsolidated = ActiveWorkbook
''open the vpcs catalogue if it's not already open
Dim xRet As Boolean
    xRet = IsWorkBookOpen("VPCS_Catalogue.xlsx")
    If xRet = False Then
        Application.Workbooks.Open ("H:\Shared\Operational\DataSystems\SCIT\CommonCatalogue\PROD\VPCS_Catalogue.xlsx")
    End If

    wbkConsolidated.Activate

For Each ws In wbkConsolidated.Worksheets
    ws.Activate
        If InStr(1, ws.Name, "Output Errors") > 0 Then
            Call HPVContractDataLookupFormulas(ws.Name)
            Application.CalculateFull
            Call replaceFormulasWithValues(ws.Name)
        End If
Next ws

    Call moveSuccessful

For Each ws In wbkConsolidated.Worksheets
    ws.Activate
    If ws.Name = "Item Output Errors" Then
        ws.Columns("N:N").EntireColumn.Hidden = True
           lastRow = ws.Cells(ws.Cells.SpecialCells(xlCellTypeLastCell).Row, 1).End(xlUp).Row
        For Each c In Range(ws.Cells(3, 11), ws.Cells(lastRow, 11)).Cells
            If c.value = "" Then
                c.value = "Error"
            End If
        Next c
    End If
    If ws.Name = "GTIN Output Errors" Then
        ws.Columns("J:J").EntireColumn.Hidden = True
    End If
    If ws.Name = "Price Output Errors" Then
        ws.Columns("O:O").EntireColumn.Hidden = True
    End If
    If ws.Name = "Item Output Success" Then
        ws.Columns("L:N").EntireColumn.Hidden = True
    End If
    If ws.Name = "GTIN Output Success" Then
        ws.Columns("I:J").EntireColumn.Hidden = True
    End If
    If ws.Name = "Price Output Success" Then
        ws.Columns("N:O").EntireColumn.Hidden = True
    End If
    If ws.Name <> "Summary" Then
        ws.Rows("1:1").EntireRow.Hidden = True
    End If
    
    ws.Range("A2").Select
Next ws
    
    wbkConsolidated.Sheets("Summary").Activate
    Range("A1").Select
    Call PivotRefresh
End Sub


Sub replaceFormulasWithValues(wsToClear As String)
    For Each ws In ActiveWorkbook.Sheets
        If ws.Name = wsToClear Then
            With ws
                .Activate
                For Each t In ActiveSheet.ListObjects
                    'If t.ShowAutoFilter = True Then
                    If t.ShowAutoFilter Then
                         t.ShowAutoFilter = False
                         t.ShowAutoFilter = True
                    End If
                    t.DataBodyRange.Copy
                    t.DataBodyRange.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                    'End If
                    Exit For
                    Exit For
               Next t
            End With
        End If
    Next ws
End Sub

Sub consolidateOutput(wsType As String, wbkSrc As Workbook, wbkDest As Workbook)
Dim wsSrc, wsDest As Worksheet
Dim destRow As Long

For Each wsSrc In wbkSrc.Worksheets
    If InStr(1, wsSrc.Name, wsType) > 0 Then  'copy content to consolidated file
        'row source data from a2 to last cell
        wsSrc.Range(wsSrc.Cells(2, 1), wsSrc.Cells.SpecialCells(xlCellTypeLastCell)).Copy
        For Each wsDest In wbkDest.Worksheets
            If InStr(1, wsDest.Name, wsType) > 0 And InStr(1, wsDest.Name, "Error") > 0 Then
                wsDest.Activate
                destRow = wsDest.Cells.SpecialCells(xlCellTypeLastCell).Row
                If destRow >= 4 Then
                    destRow = destRow + 1
                End If
                wsDest.Cells(destRow, 1).PasteSpecial xlPasteAll
            End If
        Next wsDest
    End If
Next wsSrc

End Sub

Sub ProcessFMISOutput()
'
'Open output files in a folder
'run the file arrangement applicable to the type, item, gtin or price
'Ron Campbell 16/07/2019
'
'
'

  Dim sPath As String
  Dim sFile As String
  Dim oPath As String
  Dim dpath As String
  Dim wbkDestName As String
  
  Dim wbkSource As Workbook
  Dim wSource As Worksheet
  
  Dim wbkDest As Workbook
  Dim wsDest As Worksheet
  
  
  ' advise user what they are about to do. Hitting No button will stop the macro running

On Error GoTo errorHandler
'  If MsgBox("You are about to import all of the excel files in a folder into one excel file." & vbCr & _
'            "First select the file to import the data into then a folder with the data to be copied ." & vbCr & _
'            "Create a new excel file with right click - new - excel file" & vbCr & _
'            "Do you wish to continue?", vbYesNo) = vbNo Then
'    Exit Sub
'  End If
  
  sPath = "H:\Shared\Operational\DataSystems\SCIT\CommonCatalogue\PROD\02-Output\testin\"
  oPath = "H:\Shared\Operational\DataSystems\SCIT\CommonCatalogue\PROD\02-Output\testout\"
  dpath = "H:\Shared\Operational\DataSystems\SCIT\CommonCatalogue\PROD\02-Output\"
  
    If MsgBox("Select consolidated output file", vbOKCancel) = vbOK Then
        wbkDestName = GetOutputFile(dpath)
        If Len(wbkDestName) < Len(dpath) + 3 Then
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    
    Set wbkDest = Workbooks.Open(fileName:=wbkDestName, UpdateLinks:=0, AddToMRU:=False)
    
  'sPath = GetFolder(sPath)
    If Len(sPath) < 5 Then
        Exit Sub
    End If
    
 ' oPath = GetFolder(oPath)
    If Len(oPath) < 5 Then
        Exit Sub
    End If
   
  
  
  'Confirm what the macro is about to do and folders and file selected. again user can stop by pressing No
  If MsgBox("You have Selected to import files from:" & vbCr & sPath & vbCr & _
            "and you have selected " & outFile & " as your destination file. is this correct? Are you ready to continue?", vbYesNo) = vbNo Then
        Exit Sub
  End If
    
  'set source data set to all files in provided folder
  sFile = Dir(sPath & "*CC*.txt")
'  SuspendAutoCalc
    progCount = 0
'process all files in the folder
  Do While Not sFile = ""
    Set wbkSource = Workbooks.Open(fileName:=sPath & sFile, UpdateLinks:=0, AddToMRU:=False)
    
    'check for number of lines in input file, process if > x. headers only
   ' MsgBox wbkSource.Sheets(1).Cells.SpecialCells(xlLastCell).Row
          'Identify file type
            'run appropriate file arrange
            If InStr(1, sFile, "CCITEM") > 0 Then
                If wbkSource.Sheets(1).Cells.SpecialCells(xlLastCell).Row > 8 Then
                    Call FileArrange_Item
                    'copy to collection file
                    Call consolidateOutput("Item", wbkSource, wbkDest)
                End If
            End If
            If InStr(1, sFile, "CCGTIN") > 0 Then
                If wbkSource.Sheets(1).Cells.SpecialCells(xlLastCell).Row > 6 Then
                    Call FileArrange_GTIN
                    'copy to collection file
                    Call consolidateOutput("GTIN ", wbkSource, wbkDest)
                End If
            End If
            If InStr(1, sFile, "CCPRICE") > 0 Then
                If wbkSource.Sheets(1).Cells.SpecialCells(xlLastCell).Row > 16 Then
                    Call FileArrange_Price
                    'copy to collection file
                    Call consolidateOutput("Price", wbkSource, wbkDest)
                End If
            End If
    
    wbkSource.SaveAs oPath & Replace(sFile, ".txt", ".xlsx"), xlWorkbookDefault
    wbkSource.Close False
    sFile = Dir
  Loop
  
  'restore the user's calculation setting
    ResumeAutoCalc
    MsgBox "Process Complete"
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

Sub LoadProgressUpdate()
'open the file "H:\Shared\Operational\DataSystems\SCIT\CommonCatalogue\PROD\01-Processed\processed.txt"
'add any new lines into the file "H:\Shared\Operational\DataSystems\SCIT\CommonCatalogue\PROD\LoadProgress.xlsx"
'
'process:
'open the file LoadProgress.xlsx
'open the file processed.txt
'split text to columns
'look for EGO_ filenames in Load Progress.xlsx
'add any new file entries from processed.txt into loadProgress

Dim wsFound  As Boolean: wsFound = False
Dim ws, sheetDest As Worksheet
Dim wkbSource, wkbDest As Workbook
Dim longDestRow As Long
Dim dpath As String
Dim dFile As String

'SuspendAutoCalc
'get today's files from FTP server -- have set daily task to do this
'    strCommand = "cmd /c H:\Shared\Operational\DataSystems\SCIT\CommonCatalogue\PROD\FTP_Get_PROD_Today.bat"
'    Set Wshshell = CreateObject("WScript.Shell")
'    Wshshell.Run strCommand, 2, True '2 =hidden popup, true = wait for script before continuation

'run the rename files script to rename .CCxxx to .CCxxx.txt
'    strCommand = "Powershell -File H:\Shared\Operational\DataSystems\SCIT\CommonCatalogue\PROD\01-Processed\Rename.ps1"
    Set Wshshell = CreateObject("WScript.Shell")
'    Wshshell.Run strCommand, 2, True '2 =hidden popup, true = wait for script before continuation
     
'run powershell script to create processed.txt for import
    strCommand = "Powershell -ExecutionPolicy Bypass -File H:\Shared\Operational\DataSystems\SCIT\CommonCatalogue\PROD\01-Processed\Line_DateAllFiles.ps1"
    Wshshell.Run strCommand, 2, True '2 =hidden popup, true = wait for script before continuation

  dpath = "H:\Shared\Operational\DataSystems\SCIT\CommonCatalogue\PROD\"
  dFile = "LoadProgress.xlsx"
  Call openFile(dpath, dFile)
  Set wbkDest = Workbooks.Item(dFile)
    
  For Each ws In wbkDest.Worksheets
    If ws.Name = "Input" Then
      Set sheetDest = ws
      wsFound = True
    End If
  Next ws
'  sFile = "H:\Shared\Operational\DataSystems\SCIT\CommonCatalogue\PROD\01-Processed\processed.txt"
'  'read last update datetime of sFile
'
'  Set wbkSource = Workbooks.Open(fileName:=sFile, UpdateLinks:=0, AddToMRU:=False)
'
'  'select all
'  'call text to columns with general field types
'  wbkSource.Activate
'  Set wsSource = ActiveSheet
'  With wsSource
'    .Range("a1", Cells.SpecialCells(xlCellTypeLastCell)).Select
'  Call TextToCols(1, "|")
'    .Cells.Replace What:=".txt", Replacement:="", LookAt:=xlPart, SearchOrder _
'        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'  End With
'
  'replace file dates with ones from FTP list
  'open FullList.txt
  'add lookup formula into processed.txt
 Dim tz As Integer 'time zone. doing a hack. if current dat between first sunday of april and first sunday on october then tz=+10 otherwise +11
 Dim DSTStart As Date
 Dim DSTEnd As Date
 
 lFile = "H:\Shared\Operational\DataSystems\SCIT\CommonCatalogue\PROD\01-Processed\fullList.txt"

'tz calc
DSTStart = DateSerial(Year(Date), 4, 1 + 8 - Weekday(DateSerial(Year(Date), 4, 1), vbSunday))
DSTEnd = DateSerial(Year(Date), 10, 1 + 8 - Weekday(DateSerial(Year(Date), 10, 1), vbSunday))

If Date >= DSTStart And Date <= DSTEnd Then
    tz = 10
Else
    tz = 11
End If

  Set wbkList = Workbooks.Open(fileName:=lFile, UpdateLinks:=0, AddToMRU:=False)

  wbkList.Activate
  Set wsList = ActiveSheet
  With wsList
    .Range("a1", Cells.SpecialCells(xlCellTypeLastCell)).Select
  Selection.TextToColumns Destination:=Range("A1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 9), Array(42, 1), Array(55, 1)), TrailingMinusNumbers _
        :=True
    .Range("c1").Formula = "=a1 + time(" & tz & ",0,0)"
    .Range("c1").Copy
    ActiveSheet.Range("c2", Cells(Cells.SpecialCells(xlCellTypeLastCell).Row, 3)).Select
    ActiveSheet.Paste
    
  End With
  
  sFile = "H:\Shared\Operational\DataSystems\SCIT\CommonCatalogue\PROD\01-Processed\processed.txt"
  'read last update datetime of sFile
  
  Set wbkSource = Workbooks.Open(fileName:=sFile, UpdateLinks:=0, AddToMRU:=False)
  
  wbkSource.Activate
  Set wsSource = ActiveSheet
  With wsSource
    .Range("a1", Cells.SpecialCells(xlCellTypeLastCell)).Select
    Call TextToCols(1, "|")
    .Cells.Replace What:=".txt", Replacement:="", LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    .Range("c2").Formula = "=text(XLOOKUP(A2,FullList.txt!$B:$B,fullList.txt!$C:$C,""""),""dd-mmm-yyyy hh:mm"")"
    .Range("c2").Copy
    ActiveSheet.Range("c2", Cells(Cells.SpecialCells(xlCellTypeLastCell).Row, 3)).Select
    ActiveSheet.Paste
    'ActiveSheet.Range("c2", Cells(Cells.SpecialCells(xlCellTypeLastCell).Row, 3)).NumberFormat = "dd/mm/yyyy hh:mm"
End With

  
  For Each srcC In wsSource.Range("a2", wsSource.Cells(Cells.SpecialCells(xlCellTypeLastCell).Row, 1)).Cells
    'if c.value in loadprogress list then move next else add row
    'do search in dest sheet?
    valuefound = False
    For Each destC In sheetDest.Range(sheetDest.Cells(1, 1), sheetDest.Cells(sheetDest.Cells.SpecialCells(xlCellTypeLastCell).Row, 1)).Cells
        If destC.value = srcC.value Then
            valuefound = True
        End If
    Next destC
        If valuefound = False Then 'add the new row
            wsSource.Range(srcC, srcC.Offset(0, 2)).Copy
            longDestRow = sheetDest.Cells.SpecialCells(xlCellTypeLastCell).Row + 1
            sheetDest.Range(sheetDest.Cells(longDestRow, 1), sheetDest.Cells(longDestRow, 3)).PasteSpecial xlPasteValues
        End If
  Next srcC
  
  wbkSource.Close False
  wbkList.Close False
  wbkDest.Activate
    Sheets("RunTimingExport").Select
    Range("VPCSLoadReportTiming_1[[#Headers],[REQUESTED_START_DATE]]").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
  Sheets("Input").Select
  Call copyLatestDuration
  ResumeAutoCalc
'  Call writeOutputFile("C:\Users\campbellr\Scheduled\output" & Format(Now(), "ddmmyyhhmmss") & ".txt", "LoadProgressUpdate Complete")
  Application.ActiveWorkbook.Close True
  'MsgBox "Load Progress has been updated"
  

End Sub
Sub copyLatestDuration()
'where the processing column is blank (G), check the apportioned duration column (L) for a value.
'if there is a value copy the value otherwise do nothing\
Dim tbl As ListObject
Dim rng As Range
    Set tbl = ActiveSheet.ListObjects("tblInput")
    Set rng = tbl.ListColumns("Processing Time").DataBodyRange

For Each c In rng.Cells
    If c.value = "" Then
        If Cells(c.Row, 12).value <> "" Then
            c.value = Cells(c.Row, 12).value
        End If
    End If
Next c

End Sub

Sub writeOutputFile(outFile As String, message As String)

'outFile = "C:\Users\campbellr\Scheduled\output" & Format(Now(), "ddmmyyhhmmss") & ".txt"
Set newFile = Workbooks.Add
With newFile
    Set newSheet = .Sheets(1)
            
    newSheet.Range("A1").value = message
    newSheet.Range("a2").value = Date + Time
    newFile.SaveAs outFile, xlCSV, , , , , , xlLocalSessionChanges
    .Close
End With


End Sub

Sub HPVContractDataLookupFormulas(strWS As String)

Dim pastVPCS As Boolean
pastVPCS = False

For Each ws In ActiveWorkbook.Worksheets
    If ws.Name = strWS Then
        Exit For
    End If
Next ws
If ws Is Nothing Then
    Exit Sub
End If

ws.Activate
With ws
    For Each c In .Range("a2", .Cells(2, .Cells.SpecialCells(xlCellTypeLastCell).Column)).Cells
        If c.value = "VPCS Catalogue Match" Or pastVPCS Then
           c.Offset(1, 0).Formula = c.Offset(-1, 0).value
            pastVPCS = True
        End If
    Next c
End With
End Sub


Sub moveSuccessful() 'ws As Worksheet)
'for each of the error worksheets copy the successful rows to the success sheet and delete the rows from the error sheet
 'filter error sheet for "Error" in the "Status" column
 Dim strDest As String
 Dim wsDest As Worksheet
 Dim noRows As Boolean
 
 noRows = False
 
 On Error GoTo moveSuccessful_err
 For Each ws In ActiveWorkbook.Worksheets
    If InStr(1, ws.Name, "Error") > 0 Then
        strDest = Replace(ws.Name, "Errors", "Success")
            For Each wsd In ActiveWorkbook.Worksheets
                If wsd.Name = strDest Then
                    Set wsDest = wsd
                    Exit For
                End If
            Next wsd
    ws.Activate
            With ws
               For Each c In .Range("A2", .Cells(2, .Cells.SpecialCells(xlCellTypeLastCell).Column)).Cells
                   If c.value = "STATUS" Then 'filter the table by this column
                       For Each t In .ListObjects
                        If t.ShowAutoFilter Then
                             t.ShowAutoFilter = False
                        End If
                           t.Range.AutoFilter Field:=c.Column, Criteria1:="Success"
                            If t.DataBodyRange.SpecialCells(xlCellTypeVisible).Rows.Count > 1 Then 'range of table =1 if no data. row count for databodyrange gives an error if no data
                            'If .Range(Cells(1, c.Column), Cells(.Cells.SpecialCells(xlLastCell).Row, c.Column)).SpecialCells(xlCellTypeVisible).Rows.Count > 1 Then 'range of table =1 if no data. row count for databodyrange gives an error if no data
                                If noRows = False Then
                                    t.DataBodyRange.SpecialCells(xlCellTypeVisible).Copy ' selects all available visible data to copy
                                    destRow = wsDest.Cells.SpecialCells(xlCellTypeLastCell).Row
                                     If destRow >= 4 Then     ' if there is already data in the dest table go to the next row after last row. empty table the last row is the first empty row
                                         destRow = destRow + 1
                                     End If
                                    wsDest.Cells(destRow, 1).PasteSpecial xlPasteValues
                                    Application.DisplayAlerts = False
                                    t.DataBodyRange.Delete
                                    Application.DisplayAlerts = True
                                End If
                                noRows = False
                            End If
                            t.Range.AutoFilter
                            t.ShowAutoFilter = True
                           Exit For
                       Next t
                   'copy rows to success sheet
                   'remove rows from the table
                   End If
               Next c
            End With
    End If
Next ws
'MsgBox "Move successful Operation complete"
Exit Sub

moveSuccessful_err:
'MsgBox Err.Description
If Err.Description = "No cells were found." Then
    noRows = True
    Resume Next
Else
    MsgBox Err.Description
End If

End Sub
'
'
'    For Each ws In ActiveWorkbook.Sheets
'        If ws.Name = wsToClear Then
'            With ws
'                .Activate
'                For Each t In ActiveSheet.ListObjects
'                    If t.ShowAutoFilter = True Then
'                        t.AutoFilter.Range.Select
'                        Selection.Copy
'                        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'                    End If
'               Next t
'            End With
'        End If
'    Next ws

Sub reRunPrep_01()
'text to columns for each of the three sheets ITEM, GTIN, Price
'copy formula in last column through whole column
'each export sheet, copy the formulas the length of the related sheet, exportItem, ExportGTIN, ExportPrice
'do export of data
Dim wbk As Workbook
Dim RelSheet As Worksheet
Dim RelRowCount  As Integer
Dim fileTypes(2) As String

fileTypes() = Array("Item", "GTIN", "Price")
Set wbk = ActiveWorkbook
For Each sh In wbk.Sheets
    With sh
        If .Name = "Item" Or .Name = "GTIN" Or .Name = "Price" Then
            .Activate
            .Range(.Cells(1, 1), .Cells(.Cells.SpecialCells(xlLastCell).Row, 1)).Select
            Application.DisplayAlerts = False
            Call TextToCols(2, "|")
            Application.DisplayAlerts = True
            .Range(.Cells(1, .Cells.SpecialCells(xlCellTypeLastCell).Column), .Cells(1, .Cells.SpecialCells(xlCellTypeLastCell).Column)).Copy
            .Range(.Cells(2, .Cells.SpecialCells(xlCellTypeLastCell).Column), .Cells(.Cells.SpecialCells(xlCellTypeLastCell).Row, .Cells.SpecialCells(xlCellTypeLastCell).Column)).Select
            .Paste
        End If
        If .Name Like "*Export" Then
            .Activate
            Set RelSheet = ActiveWorkbook.Sheets(Replace(.Name, "Export", ""))
            RelRowCount = RelSheet.Cells.SpecialCells(xlCellTypeLastCell).Row
            .Range(.Cells(1, 1), .Cells(1, 3)).Copy
            .Range(.Cells(2, 1), .Cells(RelRowCount, 3)).Select
            .Paste
        End If

    End With
Next


End Sub


Sub reRunPrep()
'text to columns for each of the three sheets ITEM, GTIN, Price
'copy formula in last column through whole column
'each export sheet, copy the formulas the length of the related sheet, exportItem, ExportGTIN, ExportPrice
'do export of data
Dim wbk As Workbook
Dim ws As Worksheet
Dim expSheet As Worksheet
Dim RelRowCount  As Integer
Dim fileTypes(2) As String
Dim outFolder As String
Dim saveFolder As String
Dim saveFile As String

fileTypes(0) = "Item"
fileTypes(1) = "GTIN"
fileTypes(2) = "Price"
outFolder = "H:\Shared\Operational\DataSystems\SCIT\CommonCatalogue\PROD\01-Processed\ReRun\DeltaFiles"
saveFolder = "H:\Shared\Operational\DataSystems\SCIT\CommonCatalogue\PROD\01-Processed\ReRun\"

Set wbk = ActiveWorkbook
Set ws = wbk.Sheets("Date Selector")

'saveFile = "VPCS_Load_ReRun_2022-03-12_2022-03-18.xlsx"
saveFile = "VPCS_Load_ReRun_" & Format(ws.Range("d2").value, "YYYY-MM-DD") & "_" & Format(ws.Range("d3").value, "YYYY-MM-DD") & ".xlsx"
wbk.SaveAs saveFolder & saveFile, xlWorkbookDefault
'wbk.RefreshAll

If MsgBox("Have you completed the Data Refresh?", vbYesNo) <> vbYes Then
    Exit Sub
End If

For Each ftype In fileTypes
    Set sh = wbk.Sheets(ftype)
    With sh
        .Activate
'        .Range(.Cells(1, 1), .Cells(.Cells.SpecialCells(xlLastCell).Row, 1)).Select
'        Application.DisplayAlerts = False
'        Call TextToCols(2, "|")
        Application.DisplayAlerts = True
        .Range(.Cells(1, .Cells.SpecialCells(xlCellTypeLastCell).Column), .Cells(1, .Cells.SpecialCells(xlCellTypeLastCell).Column)).Copy
        .Range(.Cells(2, .Cells.SpecialCells(xlCellTypeLastCell).Column), .Cells(.Cells.SpecialCells(xlCellTypeLastCell).Row, .Cells.SpecialCells(xlCellTypeLastCell).Column)).Select
        .Paste

        Set expSheet = ActiveWorkbook.Sheets(.Name & "Export")
            RelRowCount = .Cells.SpecialCells(xlCellTypeLastCell).Row
        With expSheet
            .Activate
            .Range(.Cells(1, 1), .Cells(1, 3)).Copy
            .Range(.Cells(2, 1), .Cells(RelRowCount, 3)).Select
            .Paste
        End With 'relsheet

    End With 'sh
Next ftype

'call the exports
      
For Each ftype In fileTypes
    Set sh = wbk.Sheets(ftype & "Export")
    With sh
        .Activate
        Call SplitFile(sh.Name, 2, 1, ftype & "Filter", "EGO_", "TXT", Format(Now(), "YYMMDD") & "_HPV.CC" & UCase(ftype), _
            outFolder, 2)
    End With
Next ftype

wbk.Save
'run powershell script "H:\Shared\Operational\DataSystems\SCIT\CommonCatalogue\PROD\01-Processed\ReRun\DeltaFiles\RENAME_-TXT.ps1"
'    strCommand = "Powershell -File H:\Shared\Operational\DataSystems\SCIT\CommonCatalogue\PROD\01-Processed\Rename.ps1"
    Set Wshshell = CreateObject("WScript.Shell")
'    Wshshell.Run strCommand, 2, True '2 =hidden popup, true = wait for script before continuation
     
'run powershell script to create processed.txt for import
    strCommand = "Powershell -ExecutionPolicy Bypass -File H:\Shared\Operational\DataSystems\SCIT\CommonCatalogue\PROD\01-Processed\ReRun\DeltaFiles\RENAME_-TXT.ps1"
    Wshshell.Run strCommand, 2, True '2 =hidden popup, true = wait for script before continuation

'run powershell script upload delta files
    strCommand = "Powershell -ExecutionPolicy Bypass -File H:\Shared\Operational\DataSystems\SCIT\CommonCatalogue\PROD\01-Processed\ReRun\FMIS_FTP_Upload.ps1"
    Wshshell.Run strCommand, 2, True '2 =hidden popup, true = wait for script before continuation
    
'move uploaded files to done folder
    strCommand = "Powershell -ExecutionPolicy Bypass -File H:\Shared\Operational\DataSystems\SCIT\CommonCatalogue\PROD\01-Processed\ReRun\moveDone.ps1"
    Wshshell.Run strCommand, 2, True '2 =hidden popup, true = wait for script before continuation
    MsgBox "Process Complete"

End Sub


Sub reRunPrep_02()
'used array formulas to replace copy&paste across sheets
'used formulas in tables to replace need for copying formulas through the sheet
'do export of data
Dim wbk As Workbook
Dim ws As Worksheet
Dim expSheet As Worksheet
Dim RelRowCount  As Integer
Dim fileTypes(2) As String
Dim outFolder As String
Dim saveFolder As String
Dim saveFile As String

fileTypes(0) = "Item"
fileTypes(1) = "GTIN"
fileTypes(2) = "Price"
outFolder = "H:\Shared\Operational\DataSystems\SCIT\CommonCatalogue\PROD\01-Processed\ReRun\DeltaFiles"
saveFolder = "H:\Shared\Operational\DataSystems\SCIT\CommonCatalogue\PROD\01-Processed\ReRun\"

Set wbk = ActiveWorkbook
Set ws = wbk.Sheets("Date Selector")

'saveFile = "VPCS_Load_ReRun_2022-03-12_2022-03-18.xlsx"
saveFile = "VPCS_Load_ReRun_" & Format(ws.Range("d2").value, "YYYY-MM-DD") & "_" & Format(ws.Range("d3").value, "YYYY-MM-DD") & ".xlsx"
wbk.SaveAs saveFolder & saveFile, xlWorkbookDefault
'wbk.RefreshAll

If MsgBox("Have you completed the Data Refresh?", vbYesNo) <> vbYes Then
    Exit Sub
End If


'call the exports
      
For Each ftype In fileTypes
    Set sh = wbk.Sheets(ftype & "Export")
    With sh
        .Activate
        Call SplitFile(sh.Name, 2, 1, ftype & "Filter", "EGO_", "TXT", Format(Now(), "YYMMDD") & "_HPV.CC" & UCase(ftype), _
            outFolder, 2)
    End With
Next ftype

wbk.Save
'run powershell script "H:\Shared\Operational\DataSystems\SCIT\CommonCatalogue\PROD\01-Processed\ReRun\DeltaFiles\RENAME_-TXT.ps1"
'    strCommand = "Powershell -File H:\Shared\Operational\DataSystems\SCIT\CommonCatalogue\PROD\01-Processed\Rename.ps1"
    Set Wshshell = CreateObject("WScript.Shell")
'    Wshshell.Run strCommand, 2, True '2 =hidden popup, true = wait for script before continuation
     
'run powershell script to create processed.txt for import
    strCommand = "Powershell -ExecutionPolicy Bypass -File H:\Shared\Operational\DataSystems\SCIT\CommonCatalogue\PROD\01-Processed\ReRun\DeltaFiles\RENAME_-TXT.ps1"
    Wshshell.Run strCommand, 2, True '2 =hidden popup, true = wait for script before continuation

'run powershell script upload delta files
    strCommand = "Powershell -ExecutionPolicy Bypass -File H:\Shared\Operational\DataSystems\SCIT\CommonCatalogue\PROD\01-Processed\ReRun\FMIS_FTP_Upload.ps1"
    Wshshell.Run strCommand, 2, True '2 =hidden popup, true = wait for script before continuation
    
    MsgBox "Process Complete"

End Sub


