Attribute VB_Name = "Testing"

Function SelectWorkSheetNameTest() As String
Dim ws As Worksheet

Set ws = Application.InputBox("Please Select the Sheet with the list to filter by", "Filter List", supplierListSheetName, Type:=8).Worksheet
getSheetNameFN = ws.Name
End Function

Sub testGetSheetName()
    MsgBox SelectWorkSheetNameTest
End Sub



Function SelectAddressTest(InputDescription As String, InputTitle As String, portion As String) As Integer
Dim tempAddress As Range

SelectAddressTest = 0
Set tmpAddress = Application.InputBox(InputDescription, InputTitle, Type:=8)
If portion = "Column" Then
    SelectAddressTest = tmpAddress.Column
End If
If portion = "Row" Then
    SelectAddressTest = tmpAddress.Row
End If
End Function

Sub testSelectAddress()
filterColumn = SelectAddressTest("Select the Column to be filtered", "Filter Column", "Column")

MsgBox filterColumn

End Sub

Sub test_clearTableFilter()
For Each tbl In ActiveSheet.ListObjects
 'tbl.Range.FilterMode
 MsgBox tbl.ShowAutoFilter
  tbl.ShowAutoFilter = False
 MsgBox tbl.ShowAutoFilter
  tbl.ShowAutoFilter = True
   MsgBox tbl.ShowAutoFilter
Next tbl
End Sub


Sub ConvertRangeToColumn()
'Updateby20131126
Dim Range1 As Range, Range2 As Range, rng As Range
Dim rowIndex As Integer
xTitleId = "Multiple Rows to ONE column"
Set Range1 = Application.Selection
Set Range1 = Application.InputBox("Source Ranges:", xTitleId, Range1.Address, Type:=8)
Set Range2 = Application.InputBox("Convert to (single cell):", xTitleId, Type:=8)
rowIndex = 0
Application.ScreenUpdating = False
For Each rng In Range1.Rows
    Rng.Copy
    Range2.Offset(rowIndex, 0).PasteSpecial Paste:=xlPasteAll, Transpose:=True
    rowIndex = rowIndex + rng.Columns.Count
Next
Application.CutCopyMode = False
Application.ScreenUpdating = True
End Sub

Sub RunAndGetCmd()

    strCommand = "Powershell -ExecutionPolicy Bypass -File H:\Shared\Operational\DataSystems\SCIT\CommonCatalogue\PROD\01-Processed\Line_DateAllFiles.ps1"

    Set Wshshell = CreateObject("WScript.Shell")
    Set WshShellExec = Wshshell.exec(strCommand)
    strOutput = WshShellExec.StdOut.ReadAll
    MsgBox strOutput

End Sub

Sub testCopyFormula()
Dim pastVPCS As Boolean
pastVPCS = False

For Each c In Range("a2", Cells(2, Cells.SpecialCells(xlCellTypeLastCell).Column)).Cells
    If c.value = "VPCS Catalogue Match" Or pastVPCS Then
        c.Offset(1, 0).Formula = c.Offset(-1, 0).value
        pastVPCS = True
    End If
Next c

End Sub







'Option Explicit
'Sub notepad()
'Application.ScreenUpdating = False
'Dim FF As Integer
'Dim plik As String
'Dim tekst As String
'Dim kom As Range
'Dim intResult As Variant
'plik = ThisWorkbook.path & "\webpage1.txt"
'If FileOrDirExists(plik) Then
'Kill plik
'MsgBox "This File is Exists"
''Application.SendKeys "%{F4}", True 'close
'Else
'FF = FreeFile
'Open plik For Output As #FF
'For Each kom In Sheets(1).Range("A1:B16")
'tekst = kom.Text
'Print #FF, tekst
'Next
'Close #FF
'intResult = Shell("Notepad.exe " & plik, vbNormalFocus)
'Application.Wait
'Application.SendKeys "%{F4}", True 'close
''or CloseAPP_B "Notepad.exe"
'End If
'Application.ScreenUpdating = True
'End Sub
'Function FileOrDirExists(PathName As String) As Boolean
'
'Dim iTemp As Integer
'
'On Error Resume Next
'iTemp = GetAttr(PathName)
'
'Select Case Err.Number
'Case Is = 0
'FileOrDirExists = True
'Case Else
'FileOrDirExists = False
'End Select
'
'On Error GoTo 0
'End Function
''**************************************
'Sub KillTest()
'MsgBox IIf(CloseAPP("notepad.exe", _
'True, False), _
'"Killed", "Failed")
'End Sub
''**************************************
'Sub KillTest_B()
'CloseAPP_B "notepad.exe"
'End Sub
''Close Application
''CloseApp KillAll=False -Only first occurrence
'' KillAll=True -All occurrences
'' NeedYesNo=True -Prompt to kill
'' NeedYesNo=False -Silent kill
'Private Function CloseAPP _
'( _
'AppNameOfExp _
'As String, _
'Optional _
'KillAll _
'As Boolean = False, _
'Optional _
'NeedYesNo _
'As Boolean = True _
') _
'As Boolean
'Dim oProcList As Object
'Dim oWMI As Object
'Dim oProc As Object
'CloseAPP = False
'' step 1: create WMI object instance:
'Set oWMI = GetObject("winmgmts:")
'If IsNull(oWMI) = False Then
'' step 2: create object collection of Win32 processes:
'Set oProcList = oWMI.InstancesOf("win32_process")
'' step 3: iterate through the enumerated collection:
'For Each oProc In oProcList
''MsgBox oProc.Name
'' option to close a process:
'If UCase(oProc.Name) = UCase(AppNameOfExp) Then
'If NeedYesNo Then
'If MsgBox("Kill " & _
'oProc.Name & vbNewLine & _
'"Are you sure?", _
'vbYesNo + vbCritical) _
'= vbYes Then
'oProc.Terminate (0)
''no test to see if this is really true
'CloseAPP = True
'End If 'MsgBox("Kill "
'Else 'NeedYesNo
'oProc.Terminate (0)
''no test to see if this is really true
'CloseAPP = True
'End If 'NeedYesNo
''continue search for more???
'If Not KillAll And CloseAPP Then
'Exit For 'oProc In oProcList
'End If 'Not KillAll And CloseAPP
'End If 'IsNull(oWMI) = False
'Next 'oProc In oProcList
'Else 'IsNull(oWMI) = False
''report error
'End If 'IsNull(oWMI) = False
'' step 4: close log file; clear out the objects:
'Set oProcList = Nothing
'Set oWMI = Nothing
'End Function
''**************************************
''No frills killer
'Private Function CloseAPP_B(AppNameOfExp As String)
'Dim oProcList As Object
'Dim oWMI As Object
'Dim oProc As Object
'' step 1: create WMI object instance:
'Set oWMI = GetObject("winmgmts:")
'If IsNull(oWMI) = False Then
'' step 2: create object collection of Win32 processes:
'Set oProcList = oWMI.InstancesOf("win32_process")
'' step 3: iterate through the enumerated collection:
'For Each oProc In oProcList
'' option to close a process:
'If UCase(oProc.Name) = UCase(AppNameOfExp) Then
'oProc.Terminate (0)
'End If 'IsNull(oWMI) = False
'Next
'End If
'End Function
'

Sub fillArray()
'test of passing an array from one sub to another to see if modifying the array in the receiving sub would alter the original or a copy
'result is the receiving sub modifies the array contents. result array does not need to be passed back explicitly
Dim localArr() As String

For i = 0 To 10
    ReDim Preserve localArr(i)
    localArr(i) = i
Next i
showArray localArr, "Initial Array Content"
passArray localArr, 2
showArray localArr, "Array Content passed and modified"
End Sub

Sub showArray(localArr() As String, message As String)

MsgBox message & " : " & Join(localArr, ", ")
    
End Sub

Sub passArray(localArr() As String, intNum As Integer)

For i = 0 To UBound(localArr)
    localArr(i) = localArr(i) * intNum
Next i
End Sub

Sub setVar()
Dim localvar As String

localvar = "a string for now"
showVar localvar, "Original"
passVar localvar, "Additional Text"
showVar localvar, "After PassVar"

End Sub

Sub showVar(var As String, message As String)
MsgBox message & ": " & var
End Sub

Sub passVar(var As String, extra As String)

    var = var & " " & extra
    showVar var, "In passVar"
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

Sub moveDone()
 Set Wshshell = CreateObject("WScript.Shell")
 strCommand = "Powershell -ExecutionPolicy Bypass -File H:\Shared\Operational\DataSystems\SCIT\CommonCatalogue\PROD\01-Processed\ReRun\moveDone.ps1"
    Wshshell.Run strCommand, 2, True '2 =hidden popup, true = wait for script before continuation
End Sub


Sub getColLetter()
Dim subCatCol As Variant
Dim subCatColLetter As String

    subCatCol = Application.InputBox("Select the column containing the sub category name", Type:=8).Address
    subCatColLetter = subCatCol
    MsgBox TextFilter(subCatColLetter)
End Sub
