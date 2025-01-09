Attribute VB_Name = "Descriptions_WT"
' Define the SheetHeaderInfo structure
Type SheetHeaderInfo
    ws As Worksheet
    headerRow As Long
    colUniqueID As Long
    colKey As Long
    colProductDescription As Long
    colBrandName As Long
    colItemNumber As Long
    colManufacturerName As Long
    colManufacturerPartNumber As Long
    colProductNewSubcategory As Long
End Type



Function RemoveMultipleLineBreaks(text As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.Pattern = "(\r?\n){2,}"
    text = regex.Replace(text, vbCrLf)
    ' Trim any leading and trailing line breaks
    Do While Left(text, 1) = vbCrLf
        text = Mid(text, 2)
    Loop
    Do While Right(text, 1) = vbCrLf
        text = Left(text, Len(text) - 1)
    Loop
    RemoveMultipleLineBreaks = text
End Function

Function CleanUpText(text As String) As String
    ' Remove unwanted characters
    text = Replace(text, "_x000D_", "")
    text = Replace(text, vbCr, "")
    text = Replace(text, vbLf, " ")
    text = Trim(text)
    CleanUpText = text
End Function




Function GetSheetAndHeaderInfo_Subcategory() As SheetHeaderInfo
    Dim info As SheetHeaderInfo
    Dim sheetName1 As String
    Dim wb As Workbook
    Dim colLetter As String
    Dim headerRowInput As String
    Dim cell As Range
    Dim found As Range
    Dim firstAddress As String
    Dim response As VbMsgBoxResult
    Dim sheetList As String
    Dim sheet As Worksheet
    
    ' Set the active workbook
    Set wb = ActiveWorkbook
    
    ' Default sheet name
    sheetName1 = "Commercial Data (All)"
    
    ' Check if the default sheet exists
    On Error Resume Next
    Set info.ws = wb.Sheets(sheetName1)
    On Error GoTo 0
    
    ' If the default sheet does not exist, prompt the user to select a sheet
    If info.ws Is Nothing Then
        ' Create a list of all worksheet names in the active workbook
        sheetList = ""
        For Each sheet In wb.Sheets
            sheetList = sheetList & sheet.Name & vbCrLf
        Next sheet
        
        ' Prompt the user to select a sheet from the list
        sheetName1 = InputBox("Commercial Data (All) is not found. Please type in the name of the sheet which represents Commercial Data (All) sheet from the following list:" & vbCrLf & sheetList, "Select Sheet")
        
        ' Check if the user selected a sheet
        If sheetName1 = "" Then
            MsgBox "No sheet selected. Exiting.", vbCritical
            Exit Function
        End If
        
        ' Set worksheet with error handling
        On Error Resume Next
        Set info.ws = wb.Sheets(sheetName1)
        On Error GoTo 0
        
        ' Check if the sheet exists
        If info.ws Is Nothing Then
            MsgBox "Sheet '" & sheetName1 & "' not found. Please check the sheet name and try again.", vbCritical
            Exit Function
        End If
    End If
    
    ' Clear any filters in the sheet
    If info.ws.AutoFilterMode Then info.ws.AutoFilterMode = False
    
    ' Unhide all columns in the selected sheet
    info.ws.Cells.EntireColumn.Hidden = False
    
    ' Find the header row with "Product New Subcategory Number & Description" or "Product New Sub Category Number & Description" in the selected sheet
    On Error Resume Next
    For i = 1 To info.ws.Rows.Count
        Set found = info.ws.Rows(i).Find(Trim("Product New Subcategory Number & Description"), LookIn:=xlValues, LookAt:=xlPart)
        If found Is Nothing Then
            Set found = info.ws.Rows(i).Find(Trim("Product New Sub Category Number & Description"), LookIn:=xlValues, LookAt:=xlPart)
        End If
        If Not found Is Nothing Then
            info.colProductNewSubcategory = found.Column
            info.headerRow = i
            Exit For
        End If
    Next i
    On Error GoTo 0
    
    ' If not found, prompt the user for the column letter and header row
    If info.colProductNewSubcategory = 0 Then
        colLetter = InputBox("Column 'Product New Subcategory Number & Description' or 'Product New Sub Category Number & Description' not found. Please provide the column letter (e.g., AB, CK):", "Column Letter")
        If colLetter = "" Then
            MsgBox "No column letter provided. Exiting.", vbCritical
            Exit Function
        End If
        headerRowInput = InputBox("Please provide the row number of the header row:", "Header Row Number")
        If headerRowInput = "" Or Not IsNumeric(headerRowInput) Then
            MsgBox "Invalid row number provided. Exiting.", vbCritical
            Exit Function
        End If
        info.headerRow = CLng(headerRowInput)
        info.colProductNewSubcategory = Range(colLetter & info.headerRow).Column
    End If
    
    ' Return the result
    GetSheetAndHeaderInfo_Subcategory = info
End Function


Sub CheckSubcategoryMapping()
    Dim info As SheetHeaderInfo
    Dim wsIndex As Worksheet
    Dim rngSubcategory As Range
    Dim rngProductSubcategory As Range
    Dim cell As Range
    Dim found As Range
    Dim wb As Workbook
    Dim headerRow As Long
    Dim colSubcategory As Long
    Dim colProductSubcategory As Long
    Dim comments() As String
    Dim commentCount As Long
    Dim commentRows(1 To 50) As Long
    Dim i As Long, j As Long
    
    ' Disable screen updating and calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    ' Set the active workbook
    Set wb = ActiveWorkbook
    
    ' Get the sheet and header info for the main sheet
    info = GetSheetAndHeaderInfo_Subcategory()
    If info.ws Is Nothing Then Exit Sub
    
    ' Set range for "Product New Subcategory Number & Description" or "Product New Sub Category Number & Description" column in the main sheet
    Set rngProductSubcategory = info.ws.Range(info.ws.Cells(info.headerRow + 1, info.colProductNewSubcategory), info.ws.Cells(info.ws.Cells(info.ws.Rows.Count, info.colProductNewSubcategory).End(xlUp).Row, info.colProductNewSubcategory))
    
    ' Clean up the text in the range
    For Each cell In rngProductSubcategory
        If Not IsEmpty(cell.value) Then
            cell.value = CleanUpText(cell.value)
        End If
    Next cell
    
    ' Clear specific comments in column A
    For Each cell In rngProductSubcategory
        If InStr(cell.Offset(0, -info.colProductNewSubcategory + 1).value, "Subcategory not found") > 0 Then
            cell.Offset(0, -info.colProductNewSubcategory + 1).value = Replace(cell.Offset(0, -info.colProductNewSubcategory + 1).value, "Subcategory not found", "")
        End If
    Next cell
    
    ' Remove empty spaces within comments in column A using Substitute
    For Each cell In info.ws.Range(info.ws.Cells(info.headerRow + 1, 1), info.ws.Cells(info.ws.Cells(info.ws.Rows.Count, 1).End(xlUp).Row, 1))
        If Len(cell.value) = 0 Then
            cell.value = Application.WorksheetFunction.Substitute(cell.value, Chr(10), "")
        Else
            cell.value = Application.WorksheetFunction.Substitute(cell.value, Chr(10) & Chr(10), Chr(10))
        End If
    Next cell
    
    ' Find the "Cat_Subcat_Major_Minor" sheet
    On Error Resume Next
    Set wsIndex = wb.Sheets("Cat_Subcat_Major_Minor")
    On Error GoTo 0
    
    ' Check if the sheet exists
    If wsIndex Is Nothing Then
        MsgBox "Sheet 'Cat_Subcat_Major_Minor' not found. Please check the sheet name and try again.", vbCritical
        Exit Sub
    End If
    
    ' Clear any filters in the sheet
    If wsIndex.AutoFilterMode Then wsIndex.AutoFilterMode = False
    
    ' Unhide all columns in the selected sheet
    wsIndex.Cells.EntireColumn.Hidden = False
    
    ' Find the header row with "Sub cat # and Sub cat description" or "Subcat Number & Subcat Description" in the "Cat_Subcat_Major_Minor" sheet
    On Error Resume Next
    For i = 1 To wsIndex.Rows.Count
        colSubcategory = wsIndex.Rows(i).Find(Trim("Sub cat # and Sub cat description"), LookIn:=xlValues, LookAt:=xlWhole).Column
        If colSubcategory = 0 Then
            colSubcategory = wsIndex.Rows(i).Find(Trim("Subcat Number & Subcat Description"), LookIn:=xlValues, LookAt:=xlWhole).Column
        End If
        If colSubcategory > 0 Then
            headerRow = i
            Exit For
        End If
    Next i
    On Error GoTo 0
    
    ' If neither column name is found, use column E
    If colSubcategory = 0 Then
        colSubcategory = 5 ' Column E
        headerRow = 1 ' Assuming the header row is the first row
    End If
    
    ' Set range for "Sub cat # and Sub cat description" or "Subcat Number & Subcat Description" column in the "Cat_Subcat_Major_Minor" sheet
    Set rngSubcategory = wsIndex.Range(wsIndex.Cells(headerRow + 1, colSubcategory), wsIndex.Cells(wsIndex.Cells(wsIndex.Rows.Count, colSubcategory).End(xlUp).Row, colSubcategory))
    
    ' Initialize comments array
    ReDim comments(1 To 50)
    commentCount = 0
    
    ' Check if the values in "Product New Subcategory Number & Description" or "Product New Sub Category Number & Description" are in "Sub cat # and Sub cat description" or "Subcat Number & Subcat Description"
    For Each cell In rngProductSubcategory
        Set found = rngSubcategory.Find(Trim(cell.value), LookIn:=xlValues, LookAt:=xlWhole)
        If found Is Nothing Then
            commentCount = commentCount + 1
            comments(commentCount) = "Subcategory not found"
            cell.Interior.Color = RGB(255, 255, 0) ' Highlight in yellow if not found
            commentRows(commentCount) = cell.Row
            
            ' Populate comments in column A every 50 comments
            If commentCount = 50 Then
                For j = 1 To 50
                    If info.ws.Cells(commentRows(j), 1).value = "" Then
                        info.ws.Cells(commentRows(j), 1).value = comments(j)
                    Else
                        info.ws.Cells(commentRows(j), 1).value = info.ws.Cells(commentRows(j), 1).value & vbCrLf & comments(j)
                    End If
                Next j
                commentCount = 0
                ReDim comments(1 To 50)
            End If
        Else
            cell.Interior.ColorIndex = xlNone ' Remove any existing highlight if found
        End If
    Next cell
    
    ' Populate remaining comments in column A
    If commentCount > 0 Then
        For j = 1 To commentCount
            If info.ws.Cells(commentRows(j), 1).value = "" Then
                info.ws.Cells(commentRows(j), 1).value = comments(j)
            Else
                info.ws.Cells(commentRows(j), 1).value = info.ws.Cells(commentRows(j), 1).value & vbCrLf & comments(j)
            End If
        Next j
    End If
    
    ' Remove empty spaces within comments in column A again after processing using Substitute
    For Each cell In info.ws.Range(info.ws.Cells(info.headerRow + 1, 1), info.ws.Cells(info.ws.Cells(info.ws.Rows.Count, 1).End(xlUp).Row, 1))
        If Len(cell.value) = 0 Then
            cell.value = Application.WorksheetFunction.Substitute(cell.value, Chr(10), "")
        Else
            cell.value = RemoveMultipleLineBreaks(cell.value)
        End If
    Next cell
    
    ' Add filters to the header row if not already applied
    If Not info.ws.AutoFilterMode Then
        info.ws.Rows(info.headerRow).AutoFilter
    End If

    ' Freeze the first row
    info.ws.Activate
    ActiveWindow.FreezePanes = False
    info.ws.Rows("2:2").Select
    ActiveWindow.FreezePanes = True
    
    
    ' Display a message indicating that the macro has finished running
    MsgBox "The macro has finished running.", vbInformation, "Macro Finished"
    
    
    
    
Cleanup:
    ' Enable screen updating and calculation
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    
    ' Release object references
    Set rngSubcategory = Nothing
    Set rngProductSubcategory = Nothing
    Set found = Nothing
    Set wsIndex = Nothing
    Set wb = Nothing
End Sub


Function GetSheetAndHeaderInfo_UniqueID() As SheetHeaderInfo
    Dim info As SheetHeaderInfo
    Dim sheetName1 As String
    Dim wb As Workbook
    Dim colLetter As String
    Dim headerRowInput As String
    Dim cell As Range
    Dim found As Range
    Dim firstAddress As String
    Dim sheetList As String
    Dim sheet As Worksheet
    
    ' Set the active workbook
    Set wb = ActiveWorkbook
    
    ' Default sheet name
    sheetName1 = "Commercial Data (All)"
    
    ' Check if the default sheet exists
    On Error Resume Next
    Set info.ws = wb.Sheets(sheetName1)
    On Error GoTo 0
    
    ' If the default sheet does not exist, prompt the user to select a sheet
    If info.ws Is Nothing Then
        ' Create a list of all worksheet names in the active workbook
        sheetList = ""
        For Each sheet In wb.Sheets
            sheetList = sheetList & sheet.Name & vbCrLf
        Next sheet
        
        ' Prompt the user to select a sheet from the list
        sheetName1 = InputBox("Commercial Data (All) is not found. Please type in the name of the sheet which represents Commercial Data (All) sheet from the following list:" & vbCrLf & sheetList, "Select Sheet")
        
        ' Check if the user selected a sheet
        If sheetName1 = "" Then
            MsgBox "No sheet selected. Exiting.", vbCritical
            Exit Function
        End If
        
        ' Set worksheet with error handling
        On Error Resume Next
        Set info.ws = wb.Sheets(sheetName1)
        On Error GoTo 0
        
        ' Check if the sheet exists
        If info.ws Is Nothing Then
            MsgBox "Sheet '" & sheetName1 & "' not found. Please check the sheet name and try again.", vbCritical
            Exit Function
        End If
    End If
    
    ' Clear any filters in the sheet
    If info.ws.AutoFilterMode Then info.ws.AutoFilterMode = False
    
    ' Unhide all columns in the selected sheet
    info.ws.Cells.EntireColumn.Hidden = False
    
    ' Find the header row with "Unique Identification" in the selected sheet
    On Error Resume Next
    For i = 1 To info.ws.Rows.Count
        Set found = info.ws.Rows(i).Find("Unique Identification", LookIn:=xlValues, LookAt:=xlPart)
        If Not found Is Nothing Then
            info.colUniqueID = found.Column
            info.headerRow = i
            Exit For
        End If
    Next i
    On Error GoTo 0
    
    ' If not found, prompt the user for the column letter and header row
    If info.colUniqueID = 0 Then
        colLetter = InputBox("Column 'Unique Identification' not found. Please provide the column letter (e.g., AB, CK):", "Column Letter")
        If colLetter = "" Then
            MsgBox "No column letter provided. Exiting.", vbCritical
            Exit Function
        End If
        headerRowInput = InputBox("Please provide the row number of the header row:", "Header Row Number")
        If headerRowInput = "" Or Not IsNumeric(headerRowInput) Then
            MsgBox "Invalid row number provided. Exiting.", vbCritical
            Exit Function
        End If
        info.headerRow = CLng(headerRowInput)
        info.colUniqueID = Range(colLetter & info.headerRow).Column
    End If
    
    ' Return the result
    GetSheetAndHeaderInfo_UniqueID = info
End Function


Function GetSheetAndHeaderInfo_Key() As SheetHeaderInfo
    Dim info As SheetHeaderInfo
    Dim sheetName1 As String
    Dim wb As Workbook
    Dim colLetter As String
    Dim headerRowInput As String
    Dim cell As Range
    Dim found As Range
    Dim firstAddress As String
    Dim sheetList As String
    Dim sheet As Worksheet
    
    ' Set the active workbook
    Set wb = ActiveWorkbook
    
    ' Default sheet name
    sheetName1 = "Commercial Data (All)"
    
    ' Check if the default sheet exists
    On Error Resume Next
    Set info.ws = wb.Sheets(sheetName1)
    On Error GoTo 0
    
    ' If the default sheet does not exist, prompt the user to select a sheet
    If info.ws Is Nothing Then
        ' Create a list of all worksheet names in the active workbook
        sheetList = ""
        For Each sheet In wb.Sheets
            sheetList = sheetList & sheet.Name & vbCrLf
        Next sheet
        
        ' Prompt the user to select a sheet from the list
        sheetName1 = InputBox("Commercial Data (All) is not found. Please type in the name of the sheet which represents Commercial Data (All) sheet from the following list:" & vbCrLf & sheetList, "Select Sheet")
        
        ' Check if the user selected a sheet
        If sheetName1 = "" Then
            MsgBox "No sheet selected. Exiting.", vbCritical
            Exit Function
        End If
        
        ' Set worksheet with error handling
        On Error Resume Next
        Set info.ws = wb.Sheets(sheetName1)
        On Error GoTo 0
        
        ' Check if the sheet exists
        If info.ws Is Nothing Then
            MsgBox "Sheet '" & sheetName1 & "' not found. Please check the sheet name and try again.", vbCritical
            Exit Function
        End If
    End If
    
    ' Clear any filters in the sheet
    If info.ws.AutoFilterMode Then info.ws.AutoFilterMode = False
    
    ' Unhide all columns in the selected sheet
    info.ws.Cells.EntireColumn.Hidden = False
    
    ' Find the header row with "SCIT KEY", "Master Key", "MDT ID", or "MDT KEY" in the selected sheet
    On Error Resume Next
    For i = 1 To info.ws.Rows.Count
        Set found = info.ws.Rows(i).Find("SCIT KEY", LookIn:=xlValues, LookAt:=xlPart)
        If Not found Is Nothing Then
            info.colKey = found.Column
            info.headerRow = i
            Exit For
        End If
        Set found = info.ws.Rows(i).Find("Master Key", LookIn:=xlValues, LookAt:=xlPart)
        If Not found Is Nothing Then
            info.colKey = found.Column
            info.headerRow = i
            Exit For
        End If
        Set found = info.ws.Rows(i).Find("MDT ID", LookIn:=xlValues, LookAt:=xlPart)
        If Not found Is Nothing Then
            info.colKey = found.Column
            info.headerRow = i
            Exit For
        End If
        Set found = info.ws.Rows(i).Find("MDT KEY", LookIn:=xlValues, LookAt:=xlPart)
        If Not found Is Nothing Then
            info.colKey = found.Column
            info.headerRow = i
            Exit For
        End If
    Next i
    On Error GoTo 0
    
    ' If not found, prompt the user for the column letter and header row
    If info.colKey = 0 Then
        colLetter = InputBox("Column 'SCIT KEY', 'Master Key', 'MDT ID', or 'MDT KEY' not found. Please provide the column letter (e.g., AB, CK):", "Column Letter")
        If colLetter = "" Then
            MsgBox "No column letter provided. Exiting.", vbCritical
            Exit Function
        End If
        headerRowInput = InputBox("Please provide the row number of the header row:", "Header Row Number")
        If headerRowInput = "" Or Not IsNumeric(headerRowInput) Then
            MsgBox "Invalid row number provided. Exiting.", vbCritical
            Exit Function
        End If
        info.headerRow = CLng(headerRowInput)
        info.colKey = Range(colLetter & info.headerRow).Column
    End If
    
    ' Return the result
    GetSheetAndHeaderInfo_Key = info
End Function



Function GetSheetAndHeaderInfo_ProductInfo() As SheetHeaderInfo
    Dim info As SheetHeaderInfo
    Dim sheetName1 As String
    Dim wb As Workbook
    Dim colLetter As String
    Dim headerRowInput As String
    Dim cell As Range
    Dim found As Range
    Dim sheetList As String
    Dim sheet As Worksheet
    
    ' Set the active workbook
    Set wb = ActiveWorkbook
    
    ' Default sheet name
    sheetName1 = "Commercial Data (All)"
    
    ' Check if the default sheet exists
    On Error Resume Next
    Set info.ws = wb.Sheets(sheetName1)
    On Error GoTo 0
    
    ' If the default sheet does not exist, prompt the user to select a sheet
    If info.ws Is Nothing Then
        ' Create a list of all worksheet names in the active workbook
        sheetList = ""
        For Each sheet In wb.Sheets
            sheetList = sheetList & sheet.Name & vbCrLf
        Next sheet
        
        ' Prompt the user to select a sheet from the list
        sheetName1 = InputBox("Commercial Data (All) is not found. Please type in the name of the sheet which represents Commercial Data (All) sheet from the following list:" & vbCrLf & sheetList, "Select Sheet")
        
        ' Check if the user selected a sheet
        If sheetName1 = "" Then
            MsgBox "No sheet selected. Exiting.", vbCritical
            Exit Function
        End If
        
        ' Set worksheet with error handling
        On Error Resume Next
        Set info.ws = wb.Sheets(sheetName1)
        On Error GoTo 0
        
        ' Check if the sheet exists
        If info.ws Is Nothing Then
            MsgBox "Sheet '" & sheetName1 & "' not found. Please check the sheet name and try again.", vbCritical
            Exit Function
        End If
    End If
    
    ' Clear any filters in the sheet
    If info.ws.AutoFilterMode Then info.ws.AutoFilterMode = False
    
    ' Unhide all columns in the selected sheet
    info.ws.Cells.EntireColumn.Hidden = False
    
    ' Find the header row with "Full Product Description" in the selected sheet
    On Error Resume Next
    For i = 1 To info.ws.Rows.Count
        Set found = info.ws.Rows(i).Find("Full Product Description", LookIn:=xlValues, LookAt:=xlPart)
        If Not found Is Nothing Then
            info.colProductDescription = found.Column
            info.headerRow = i
            Exit For
        End If
    Next i
    On Error GoTo 0
    
    ' If not found, prompt the user for the column letter and header row
    If info.colProductDescription = 0 Then
        colLetter = InputBox("Column 'Full Product Description' not found. Please provide the column letter (e.g., AB, CK):", "Column Letter")
        If colLetter = "" Then
            MsgBox "No column letter provided. Exiting.", vbCritical
            Exit Function
        End If
        headerRowInput = InputBox("Please provide the row number of the header row:", "Header Row Number")
        If headerRowInput = "" Or Not IsNumeric(headerRowInput) Then
            MsgBox "Invalid row number provided. Exiting.", vbCritical
            Exit Function
        End If
        info.headerRow = CLng(headerRowInput)
        info.colProductDescription = Range(colLetter & info.headerRow).Column
    End If
    
    ' Find the header row with "Brand Name" in the selected sheet
    On Error Resume Next
    For i = 1 To info.ws.Rows.Count
        Set found = info.ws.Rows(i).Find("Brand Name", LookIn:=xlValues, LookAt:=xlPart)
        If Not found Is Nothing Then
            info.colBrandName = found.Column
            info.headerRow = i
            Exit For
        End If
    Next i
    On Error GoTo 0
    
    ' If not found, prompt the user for the column letter and header row
    If info.colBrandName = 0 Then
        colLetter = InputBox("Column 'Brand Name' not found. Please provide the column letter (e.g., AB, CK):", "Column Letter")
        If colLetter = "" Then
            MsgBox "No column letter provided. Exiting.", vbCritical
            Exit Function
        End If
        headerRowInput = InputBox("Please provide the row number of the header row:", "Header Row Number")
        If headerRowInput = "" Or Not IsNumeric(headerRowInput) Then
            MsgBox "Invalid row number provided. Exiting.", vbCritical
            Exit Function
        End If
        info.headerRow = CLng(headerRowInput)
        info.colBrandName = Range(colLetter & info.headerRow).Column
    End If
    
    ' Find the header row with "Item Number printed" in the selected sheet
    On Error Resume Next
    For i = 1 To info.ws.Rows.Count
        Set found = info.ws.Rows(i).Find("Item Number printed", LookIn:=xlValues, LookAt:=xlPart)
        If Not found Is Nothing Then
            info.colItemNumber = found.Column
            info.headerRow = i
            Exit For
        End If
    Next i
    On Error GoTo 0
    
    ' If not found, prompt the user for the column letter and header row
    If info.colItemNumber = 0 Then
        colLetter = InputBox("Column 'Item Number printed' not found. Please provide the column letter (e.g., AB, CK):", "Column Letter")
        If colLetter = "" Then
            MsgBox "No column letter provided. Exiting.", vbCritical
            Exit Function
        End If
        headerRowInput = InputBox("Please provide the row number of the header row:", "Header Row Number")
        If headerRowInput = "" Or Not IsNumeric(headerRowInput) Then
            MsgBox "Invalid row number provided. Exiting.", vbCritical
            Exit Function
        End If
        info.headerRow = CLng(headerRowInput)
        info.colItemNumber = Range(colLetter & info.headerRow).Column
    End If
    
    ' Find the header row with "Manufacturer Name" in the selected sheet (exact match)
    On Error Resume Next
    For i = 1 To info.ws.Rows.Count
        Set found = info.ws.Rows(i).Find("Manufacturer Name", LookIn:=xlValues, LookAt:=xlWhole)
        If Not found Is Nothing Then
            info.colManufacturerName = found.Column
            info.headerRow = i
            Exit For
        End If
    Next i
    On Error GoTo 0
    
    ' If not found, prompt the user for the column letter and header row
    If info.colManufacturerName = 0 Then
        colLetter = InputBox("Column 'Manufacturer Name' not found. Please provide the column letter (e.g., AB, CK):", "Column Letter")
        If colLetter = "" Then
            MsgBox "No column letter provided. Exiting.", vbCritical
            Exit Function

        End If
        headerRowInput = InputBox("Please provide the row number of the header row:", "Header Row Number")
        If headerRowInput = "" Or Not IsNumeric(headerRowInput) Then
            MsgBox "Invalid row number provided. Exiting.", vbCritical
            Exit Function
        End If
        info.headerRow = CLng(headerRowInput)
        info.colManufacturerName = Range(colLetter & info.headerRow).Column
    End If
    
    ' Find the header row with "Manufacturer Part Number" in the selected sheet (exact match)
    On Error Resume Next
    For i = 1 To info.ws.Rows.Count
        Set found = info.ws.Rows(i).Find("Manufacturer Part Number", LookIn:=xlValues, LookAt:=xlWhole)
        If Not found Is Nothing Then
            info.colManufacturerPartNumber = found.Column
            info.headerRow = i
            Exit For
        End If
    Next i
    On Error GoTo 0
    
    ' If not found, prompt the user for the column letter and header row
    If info.colManufacturerPartNumber = 0 Then
        colLetter = InputBox("Column 'Manufacturer Part Number' not found. Please provide the column letter (e.g., AB, CK):", "Column Letter")
        If colLetter = "" Then
            MsgBox "No column letter provided. Exiting.", vbCritical
            Exit Function
        End If
        headerRowInput = InputBox("Please provide the row number of the header row:", "Header Row Number")
        If headerRowInput = "" Or Not IsNumeric(headerRowInput) Then
            MsgBox "Invalid row number provided. Exiting.", vbCritical
            Exit Function
        End If
        info.headerRow = CLng(headerRowInput)
        info.colManufacturerPartNumber = Range(colLetter & info.headerRow).Column
    End If
    
    ' Return the result
    GetSheetAndHeaderInfo_ProductInfo = info
End Function

Sub CompareUniqueIdentification()
    Dim info As SheetHeaderInfo
    Dim ws As Worksheet
    Dim rng1 As Range, rng As Range
    Dim cell As Range
    Dim found As Range
    Dim isFound As Boolean
    Dim comments() As String
    Dim i As Long, j As Long
    Dim col As Long
    Dim wb As Workbook
    Dim headerRow As Long
    Dim cellValue As String, foundValue As String
    Dim partialMatch As Boolean
    Dim commentCount As Long
    Dim commentRows(1 To 50) As Long
    
    ' Disable screen updating and calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    ' Get the sheet and header info
    info = GetSheetAndHeaderInfo_UniqueID()
    If info.ws Is Nothing Then GoTo Cleanup
    
    ' Set the active workbook
    Set wb = ActiveWorkbook
    
    ' Set range for "Unique Identification" column in Sheet1
    Set rng1 = info.ws.Range(info.ws.Cells(info.headerRow + 1, info.colUniqueID), info.ws.Cells(info.ws.Cells(info.ws.Rows.Count, info.colUniqueID).End(xlUp).Row, info.colUniqueID))
    
    ' Clean up the text in the range
    For Each cell In rng1
        If Not IsEmpty(cell.value) Then
            cell.value = CleanUpText(cell.value)
        End If
    Next cell
    
    ' Unhighlight the column
    rng1.Interior.ColorIndex = xlNone
    
    ' Clear specific comments in column A
    For Each cell In rng1
        If InStr(cell.Offset(0, -info.colUniqueID + 1).value, "Item not found in clinical sheets") > 0 Then
            cell.Offset(0, -info.colUniqueID + 1).value = Replace(cell.Offset(0, -info.colUniqueID + 1).value, "Item not found in clinical sheets", "")
        End If
        If InStr(cell.Offset(0, -info.colUniqueID + 1).value, "Item found but only item no. matched") > 0 Then
            cell.Offset(0, -info.colUniqueID + 1).value = Replace(cell.Offset(0, -info.colUniqueID + 1).value, "Item found but only item no. matched", "")
        End If
    Next cell
    
    ' Remove empty spaces within comments in column A using Substitute
    For Each cell In info.ws.Range(info.ws.Cells(info.headerRow + 1, 1), info.ws.Cells(info.ws.Cells(info.ws.Rows.Count, 1).End(xlUp).Row, 1))
        If Len(cell.value) = 0 Then
            cell.value = Application.WorksheetFunction.Substitute(cell.value, Chr(10), "")
        Else
            cell.value = Application.WorksheetFunction.Substitute(cell.value, Chr(10) & Chr(10), Chr(10))
        End If
    Next cell
    
    ' Initialize comments array
    ReDim comments(1 To 50)
    commentCount = 0
    
    ' Compare items in Sheet1 with items in sheets starting with "Clinical Data (Category" or containing "(Clinical)"
    For Each cell In rng1
        isFound = False
        partialMatch = False
        cellValue = cell.value ' Use the entire value for the first comparison
        For Each ws In wb.Sheets
            If ws.Name Like "Clinical Data (Category*" Or ws.Name Like "*(Clinical)*" Then
                ' Clear any filters in the sheet
                If ws.AutoFilterMode Then ws.AutoFilterMode = False
                
                ' Unhide all columns in the current sheet
                ws.Cells.EntireColumn.Hidden = False
                
                ' Find the header row with "Unique Identification" in the current sheet
                On Error Resume Next
                For i = 1 To ws.Rows.Count
                    col = ws.Rows(i).Find("Unique Identification", LookIn:=xlValues, LookAt:=xlWhole).Column
                    If col > 0 Then
                        headerRow = i
                        Exit For
                    End If
                Next i
                On Error GoTo 0
                
                If col = 0 Then
                    MsgBox "Column 'Unique Identification' not found in " & ws.Name & ". Please check the column header and try again.", vbCritical
                    GoTo Cleanup
                End If
                
                ' Set range for "Unique Identification" column in the current sheet
                Set rng = ws.Range(ws.Cells(headerRow + 1, col), ws.Cells(ws.Cells(ws.Rows.Count, col).End(xlUp).Row, col))
                
                ' Check if the item is found in the current sheet
                Set found = rng.Find(cellValue, LookIn:=xlValues, LookAt:=xlWhole)
                If Not found Is Nothing Then
                    isFound = True
                    Exit For
                End If
            End If
            If isFound Then Exit For
        Next ws
        
        ' If not found, compare the value before the first pipe delimiter
        If Not isFound And InStr(cell.value, "\n") > 0 Then
            cellValue = Trim(Split(cell.value, "\n")(0)) ' Get the value before the first pipe delimiter and trim spaces
            For Each ws In wb.Sheets
                If ws.Name Like "Clinical Data (Category*" Or ws.Name Like "*(Clinical)*" Then
                    ' Clear any filters in the sheet
                    If ws.AutoFilterMode Then ws.AutoFilterMode = False
                    
                    ' Unhide all columns in the current sheet
                    ws.Cells.EntireColumn.Hidden = False
                    
                    ' Find the header row with "Unique Identification" in the current sheet
                    On Error Resume Next
                    For i = 1 To ws.Rows.Count
                        col = ws.Rows(i).Find("Unique Identification", LookIn:=xlValues, LookAt:=xlWhole).Column
                        If col > 0 Then
                            headerRow = i
                            Exit For
                        End If
                    Next i
                    On Error GoTo 0
                    
                    If col = 0 Then
                        MsgBox "Column 'Unique Identification' not found in " & ws.Name & ". Please check the column header and try again.", vbCritical
                        GoTo Cleanup
                    End If
                    
                    ' Set range for "Unique Identification" column in the current sheet
                    Set rng = ws.Range(ws.Cells(headerRow + 1, col), ws.Cells(ws.Cells(ws.Rows.Count, col).End(xlUp).Row, col))
                    
                    ' Check if the part number (value before the first pipe delimiter) is found
                    For Each found In rng
                        If InStr(found.value, "\n") > 0 Then
                            foundValue = Trim(Split(found.value, "\n")(0)) ' Trim spaces from the found value
                        Else
                            foundValue = Trim(found.value) ' Trim spaces from the found value
                        End If
                        If cellValue = foundValue Then
                            isFound = True
                            partialMatch = True
                            Exit For
                        End If
                    Next found
                End If
                If isFound Then Exit For
            Next ws
        End If
        
        ' Add comment based on the match result
        commentCount = commentCount + 1
        If Not isFound Then
            comments(commentCount) = "Item not found in clinical sheets"
            cell.Interior.Color = RGB(255, 210, 191) ' Highlight in RGB 255, 210, 191 if not found
        ElseIf partialMatch Then
            comments(commentCount) = "Item found but only item no. matched"
            cell.Interior.Color = RGB(158, 52, 235) ' Highlight in RGB 158, 52, 235 if only item no. matched
        End If
        commentRows(commentCount) = cell.Row
        
        ' Populate comments in column A every 50 comments
        If commentCount = 50 Then
            For j = 1 To 50
                If info.ws.Cells(commentRows(j), 1).value = "" Then
                    info.ws.Cells(commentRows(j), 1).value = comments(j)
                Else
                    info.ws.Cells(commentRows(j), 1).value = info.ws.Cells(commentRows(j), 1).value & vbCrLf & comments(j)
                End If
            Next j
            commentCount = 0
            ReDim comments(1 To 50)
        End If
    Next cell
    
    ' Populate remaining comments in column A
    If commentCount > 0 Then
        For j = 1 To commentCount
            If info.ws.Cells(commentRows(j), 1).value = "" Then
                info.ws.Cells(commentRows(j), 1).value = comments(j)
            Else
                info.ws.Cells(commentRows(j), 1).value = info.ws.Cells(commentRows(j), 1).value & vbCrLf & comments(j)
            End If
        Next j
    End If

    ' Remove empty spaces within comments in column A again after processing using Substitute
    For Each cell In info.ws.Range(info.ws.Cells(info.headerRow + 1, 1), info.ws.Cells(info.ws.Cells(info.ws.Rows.Count, 1).End(xlUp).Row, 1))
        If Len(cell.value) = 0 Then
            cell.value = Application.WorksheetFunction.Substitute(cell.value, Chr(10), "")
        Else
            cell.value = RemoveMultipleLineBreaks(cell.value)
        End If
    Next cell
    
    ' Add filters to the header row if not already applied
    If Not info.ws.AutoFilterMode Then
        info.ws.Rows(info.headerRow).AutoFilter
    End If

    ' Freeze the first row
    info.ws.Activate
    ActiveWindow.FreezePanes = False
    info.ws.Rows("2:2").Select
    ActiveWindow.FreezePanes = True
    
    
    
    ' Display a message indicating that the macro has finished running
    MsgBox "The macro has finished running.", vbInformation, "Macro Finished"
    
        
    
Cleanup:
    ' Enable screen updating and calculation
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    
    ' Release object references
    Set rng1 = Nothing
    Set rng = Nothing
    Set found = Nothing
    Set ws = Nothing
    Set wb = Nothing
End Sub
Sub CompareKeys()
    Dim info As SheetHeaderInfo
    Dim ws As Worksheet
    Dim rng1 As Range, rng As Range
    Dim cell As Range
    Dim found As Range
    Dim isFound As Boolean
    Dim comments() As String
    Dim i As Long, j As Long
    Dim col As Long
    Dim wb As Workbook
    Dim headerRow As Long
    Dim cellValue As String
    Dim commentCount As Long
    Dim commentRows(1 To 50) As Long
    
    ' Disable screen updating and calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    ' Get the sheet and header info
    info = GetSheetAndHeaderInfo_Key()
    If info.ws Is Nothing Then GoTo Cleanup
    
    ' Set the active workbook
    Set wb = ActiveWorkbook
    
    ' Set range for "SCIT KEY", "Master Key", "MDT ID", or "MDT KEY" column in Sheet1
    Set rng1 = info.ws.Range(info.ws.Cells(info.headerRow + 1, info.colKey), info.ws.Cells(info.ws.Cells(info.ws.Rows.Count, info.colKey).End(xlUp).Row, info.colKey))
    
    ' Clean up the text in the range
    For Each cell In rng1
        If Not IsEmpty(cell.value) Then
            cell.value = CleanUpText(cell.value)
        End If
    Next cell
    
    ' Unhighlight the column
    rng1.Interior.ColorIndex = xlNone
    
    ' Clear specific comments in column A
    For Each cell In rng1
        If InStr(cell.Offset(0, -info.colKey + 1).value, "Key not found in clinical sheets") > 0 Then
            cell.Offset(0, -info.colKey + 1).value = Replace(cell.Offset(0, -info.colKey + 1).value, "Key not found in clinical sheets", "")
        End If
    Next cell
    
    ' Remove empty spaces within comments in column A using Substitute
    For Each cell In info.ws.Range(info.ws.Cells(info.headerRow + 1, 1), info.ws.Cells(info.ws.Cells(info.ws.Rows.Count, 1).End(xlUp).Row, 1))
        If Len(cell.value) = 0 Then
            cell.value = Application.WorksheetFunction.Substitute(cell.value, Chr(10), "")
        Else
            cell.value = Application.WorksheetFunction.Substitute(cell.value, Chr(10) & Chr(10), Chr(10))
        End If
    Next cell
    
    ' Initialize comments array
    ReDim comments(1 To 50)
    commentCount = 0
    
    ' Compare items in Sheet1 with items in sheets starting with "Clinical Data (Category" or containing "(Clinical)"
    For Each cell In rng1
        isFound = False
        cellValue = cell.value ' Use the entire value for the comparison
        For Each ws In wb.Sheets
            If ws.Name Like "Clinical Data (Category*" Or ws.Name Like "*(Clinical)*" Then
                ' Clear any filters in the sheet
                If ws.AutoFilterMode Then ws.AutoFilterMode = False
                
                ' Unhide all columns in the current sheet
                ws.Cells.EntireColumn.Hidden = False
                
                ' Find the header row with "SCIT KEY", "Master Key", "MDT ID", or "MDT KEY" in the current sheet
                On Error Resume Next
                For i = 1 To ws.Rows.Count
                    col = ws.Rows(i).Find("SCIT KEY", LookIn:=xlValues, LookAt:=xlPart).Column
                    If col = 0 Then
                        col = ws.Rows(i).Find("Master Key", LookIn:=xlValues, LookAt:=xlPart).Column
                    End If
                    If col = 0 Then
                        col = ws.Rows(i).Find("MDT ID", LookIn:=xlValues, LookAt:=xlPart).Column
                    End If
                    If col = 0 Then
                        col = ws.Rows(i).Find("MDT KEY", LookIn:=xlValues, LookAt:=xlPart).Column
                    End If
                    If col > 0 Then
                        headerRow = i
                        Exit For
                    End If
                Next i
                On Error GoTo 0
                
                If col = 0 Then
                    MsgBox "Column 'SCIT KEY', 'Master Key', 'MDT ID', or 'MDT KEY' not found in " & ws.Name & ". Please check the column header and try again.", vbCritical
                    GoTo Cleanup
                End If
                
                ' Set range for "SCIT KEY", "Master Key", "MDT ID", or "MDT KEY" column in the current sheet
                Set rng = ws.Range(ws.Cells(headerRow + 1, col), ws.Cells(ws.Cells(ws.Rows.Count, col).End(xlUp).Row, col))
                
                ' Check if the item is found in the current sheet
                Set found = rng.Find(cellValue, LookIn:=xlValues, LookAt:=xlWhole)
                If Not found Is Nothing Then
                    isFound = True
                    Exit For
                End If
            End If
            If isFound Then Exit For
        Next ws
        
        ' Add comment based on the match result
        commentCount = commentCount + 1
        If Not isFound Then
            comments(commentCount) = "Key not found in clinical sheets"
            cell.Interior.Color = RGB(255, 255, 0) ' Highlight in yellow if not found
        End If
        commentRows(commentCount) = cell.Row
        
        ' Populate comments in column A every 50 comments
        If commentCount = 50 Then
            For j = 1 To 50
                If info.ws.Cells(commentRows(j), 1).value = "" Then
                    info.ws.Cells(commentRows(j), 1).value = comments(j)
                Else
                    info.ws.Cells(commentRows(j), 1).value = info.ws.Cells(commentRows(j), 1).value & vbCrLf & comments(j)
                End If
            Next j
            commentCount = 0
            ReDim comments(1 To 50)
        End If
    Next cell
    
    ' Populate remaining comments in column A
    If commentCount > 0 Then
        For j = 1 To commentCount
            If info.ws.Cells(commentRows(j), 1).value = "" Then
                info.ws.Cells(commentRows(j), 1).value = comments(j)
            Else
                info.ws.Cells(commentRows(j), 1).value = info.ws.Cells(commentRows(j), 1).value & vbCrLf & comments(j)
            End If
        Next j
    End If
    
    ' Remove empty spaces within comments in column A again after processing using Substitute
    For Each cell In info.ws.Range(info.ws.Cells(info.headerRow + 1, 1), info.ws.Cells(info.ws.Cells(info.ws.Rows.Count, 1).End(xlUp).Row, 1))
        If Len(cell.value) = 0 Then
            cell.value = Application.WorksheetFunction.Substitute(cell.value, Chr(10), "")
        Else
            cell.value = RemoveMultipleLineBreaks(cell.value)
        End If
    Next cell
    
    ' Add filters to the header row if not already applied
    If Not info.ws.AutoFilterMode Then
        info.ws.Rows(info.headerRow).AutoFilter
    End If

    ' Freeze the first row
    info.ws.Activate
    ActiveWindow.FreezePanes = False
    info.ws.Rows("2:2").Select
    ActiveWindow.FreezePanes = True
    
    ' Display a message indicating that the macro has finished running
    MsgBox "The macro has finished running.", vbInformation, "Macro Finished"
      
    
Cleanup:
    ' Enable screen updating and calculation
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    
    ' Release object references
    Set rng1 = Nothing
    Set rng = Nothing
    Set found = Nothing
    Set ws = Nothing
    Set wb = Nothing
End Sub


Sub Check_If_Key_Product_Infor_Blank()
    Dim info As SheetHeaderInfo
    Dim ws As Worksheet
    Dim rngProductDescription As Range
    Dim rngBrandName As Range
    Dim rngItemNumber As Range
    Dim rngManufacturerName As Range
    Dim rngManufacturerPartNumber As Range
    Dim cell As Range
    Dim comments() As String
    Dim commentCount As Long
    Dim commentRows(1 To 50) As Long
    Dim i As Long, j As Long
    
    ' Disable screen updating and calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    ' Get the sheet and header info
    info = GetSheetAndHeaderInfo_ProductInfo()
    If info.ws Is Nothing Then GoTo Cleanup
    
    ' Set the worksheet
    Set ws = info.ws
    
    ' Set range for "Full Product Description" column in Sheet1
    Set rngProductDescription = ws.Range(ws.Cells(info.headerRow + 1, info.colProductDescription), ws.Cells(ws.Cells(ws.Rows.Count, info.colProductDescription).End(xlUp).Row, info.colProductDescription))
    
    ' Set range for "Brand Name" column in Sheet1
    Set rngBrandName = ws.Range(ws.Cells(info.headerRow + 1, info.colBrandName), ws.Cells(ws.Cells(ws.Rows.Count, info.colBrandName).End(xlUp).Row, info.colBrandName))
    
    ' Set range for "Item Number printed on the packaging" column in Sheet1
    Set rngItemNumber = ws.Range(ws.Cells(info.headerRow + 1, info.colItemNumber), ws.Cells(ws.Cells(ws.Rows.Count, info.colItemNumber).End(xlUp).Row, info.colItemNumber))
    
    ' Set range for "Manufacturer Name" column in Sheet1
    Set rngManufacturerName = ws.Range(ws.Cells(info.headerRow + 1, info.colManufacturerName), ws.Cells(ws.Cells(ws.Rows.Count, info.colManufacturerName).End(xlUp).Row, info.colManufacturerName))
    
    ' Set range for "Manufacturer Part Number" column in Sheet1
    Set rngManufacturerPartNumber = ws.Range(ws.Cells(info.headerRow + 1, info.colManufacturerPartNumber), ws.Cells(ws.Cells(ws.Rows.Count, info.colManufacturerPartNumber).End(xlUp).Row, info.colManufacturerPartNumber))
    
    ' Exclude the last row if it contains "END" or "End"
    If UCase(rngProductDescription.Cells(rngProductDescription.Rows.Count).value) = "END" Then
        Set rngProductDescription = rngProductDescription.Resize(rngProductDescription.Rows.Count - 1)
    End If
    If UCase(rngBrandName.Cells(rngBrandName.Rows.Count).value) = "END" Then
        Set rngBrandName = rngBrandName.Resize(rngBrandName.Rows.Count - 1)
    End If
    If UCase(rngItemNumber.Cells(rngItemNumber.Rows.Count).value) = "END" Then
        Set rngItemNumber = rngItemNumber.Resize(rngItemNumber.Rows.Count - 1)
    End If
    If UCase(rngManufacturerName.Cells(rngManufacturerName.Rows.Count).value) = "END" Then
        Set rngManufacturerName = rngManufacturerName.Resize(rngManufacturerName.Rows.Count - 1)
    End If
    If UCase(rngManufacturerPartNumber.Cells(rngManufacturerPartNumber.Rows.Count).value) = "END" Then
        Set rngManufacturerPartNumber = rngManufacturerPartNumber.Resize(rngManufacturerPartNumber.Rows.Count - 1)
    End If
    
    ' Clean up the text in the ranges
    For Each cell In rngProductDescription
        If Not IsEmpty(cell.value) Then
            cell.value = CleanUpText(cell.value)
        End If
    Next cell
    For Each cell In rngBrandName
        If Not IsEmpty(cell.value) Then
            cell.value = CleanUpText(cell.value)
        End If
    Next cell
    For Each cell In rngItemNumber
        If Not IsEmpty(cell.value) Then
            cell.value = CleanUpText(cell.value)
        End If
    Next cell
    For Each cell In rngManufacturerName
        If Not IsEmpty(cell.value) Then
            cell.value = CleanUpText(cell.value)
        End If
    Next cell
    For Each cell In rngManufacturerPartNumber
        If Not IsEmpty(cell.value) Then
            cell.value = CleanUpText(cell.value)
        End If
    Next cell
    
    ' Unhighlight the columns
    rngProductDescription.Interior.ColorIndex = xlNone
    rngBrandName.Interior.ColorIndex = xlNone
    rngItemNumber.Interior.ColorIndex = xlNone
    rngManufacturerName.Interior.ColorIndex = xlNone
    rngManufacturerPartNumber.Interior.ColorIndex = xlNone
    
    ' Clear specific comments in column A
    For Each cell In rngProductDescription
        If InStr(cell.Offset(0, -info.colProductDescription + 1).value, "No Description") > 0 Then
            cell.Offset(0, -info.colProductDescription + 1).value = Replace(cell.Offset(0, -info.colProductDescription + 1).value, "No Description", "")
        End If
    Next cell
    For Each cell In rngBrandName
        If InStr(cell.Offset(0, -info.colBrandName + 1).value, "No Brand") > 0 Then
            cell.Offset(0, -info.colBrandName + 1).value = Replace(cell.Offset(0, -info.colBrandName + 1).value, "No Brand", "")
        End If
    Next cell
    For Each cell In rngItemNumber
        If InStr(cell.Offset(0, -info.colItemNumber + 1).value, "No Item Number") > 0 Then
            cell.Offset(0, -info.colItemNumber + 1).value = Replace(cell.Offset(0, -info.colItemNumber + 1).value, "No Item Number", "")
        End If
    Next cell
    For Each cell In rngManufacturerName
        If InStr(cell.Offset(0, -info.colManufacturerName + 1).value, "No MFG") > 0 Then
            cell.Offset(0, -info.colManufacturerName + 1).value = Replace(cell.Offset(0, -info.colManufacturerName + 1).value, "No MFG", "")
        End If
    Next cell
    For Each cell In rngManufacturerPartNumber
        If InStr(cell.Offset(0, -info.colManufacturerPartNumber + 1).value, "No MPN") > 0 Then
            cell.Offset(0, -info.colManufacturerPartNumber + 1).value = Replace(cell.Offset(0, -info.colManufacturerPartNumber + 1).value, "No MPN", "")
        End If
    Next cell
    
    ' Remove empty spaces within comments in column A using Substitute
    For Each cell In ws.Range(ws.Cells(info.headerRow + 1, 1), ws.Cells(ws.Cells(ws.Rows.Count, 1).End(xlUp).Row, 1))
        If Len(cell.value) = 0 Then
            cell.value = Application.WorksheetFunction.Substitute(cell.value, Chr(10), "")
        Else
            cell.value = RemoveMultipleLineBreaks(cell.value)
        End If
    Next cell
    
    ' Initialize comments array
    ReDim comments(1 To 50)
    commentCount = 0
    
    ' Check each cell in the "Full Product Description" column
    For Each cell In rngProductDescription
        If Len(cell.value) < 5 Then
            commentCount = commentCount + 1
            comments(commentCount) = "No Description"
            cell.Interior.Color = RGB(255, 255, 0) ' Highlight in yellow if missing or fewer than 5 characters
            commentRows(commentCount) = cell.Row
            
            ' Populate comments in column A every 50 comments
            If commentCount = 50 Then
                For j = 1 To 50
                    If ws.Cells(commentRows(j), 1).value = "" Then
                        ws.Cells(commentRows(j), 1).value = comments(j)
                    Else
                        ws.Cells(commentRows(j), 1).value = ws.Cells(commentRows(j), 1).value & vbCrLf & comments(j)
                    End If
                Next j
                commentCount = 0
                ReDim comments(1 To 50)
            End If
        End If
    Next cell
    
    ' Check each cell in the "Brand Name" column
    For Each cell In rngBrandName
        If Len(cell.value) = 0 Then
            commentCount = commentCount + 1
            comments(commentCount) = "No Brand"
            cell.Interior.Color = RGB(255, 255, 0) ' Highlight in yellow if blank
            commentRows(commentCount) = cell.Row
            
            ' Populate comments in column A every 50 comments
            If commentCount = 50 Then
                For j = 1 To 50
                    If ws.Cells(commentRows(j), 1).value = "" Then
                        ws.Cells(commentRows(j), 1).value = comments(j)
                    Else
                        ws.Cells(commentRows(j), 1).value = ws.Cells(commentRows(j), 1).value & vbCrLf & comments(j)
                    End If
                Next j
                commentCount = 0
                ReDim comments(1 To 50)
            End If
        End If
    Next cell
    
    ' Check each cell in the "Item Number printed on the packaging" column
    For Each cell In rngItemNumber
        If Len(cell.value) = 0 Then
            commentCount = commentCount + 1
            comments(commentCount) = "No Item Number"
            cell.Interior.Color = RGB(255, 255, 0) ' Highlight in yellow if blank
            commentRows(commentCount) = cell.Row
            
            ' Populate comments in column A every 50 comments
            If commentCount = 50 Then
                For j = 1 To 50
                    If ws.Cells(commentRows(j), 1).value = "" Then
                        ws.Cells(commentRows(j), 1).value = comments(j)
                    Else
                        ws.Cells(commentRows(j), 1).value = ws.Cells(commentRows(j), 1).value & vbCrLf & comments(j)
                    End If
                Next j
                commentCount = 0
                ReDim comments(1 To 50)
            End If
        End If
    Next cell
    
    ' Check each cell in the "Manufacturer Name" column
    For Each cell In rngManufacturerName
        If Len(cell.value) = 0 Then
            commentCount = commentCount + 1
            comments(commentCount) = "No MFG"
            cell.Interior.Color = RGB(255, 255, 0) ' Highlight in yellow if blank
            commentRows(commentCount) = cell.Row
            
            ' Populate comments in column A every 50 comments
            If commentCount = 50 Then
                For j = 1 To 50
                    If ws.Cells(commentRows(j), 1).value = "" Then
                        ws.Cells(commentRows(j), 1).value = comments(j)
                    Else
                        ws.Cells(commentRows(j), 1).value = ws.Cells(commentRows(j), 1).value & vbCrLf & comments(j)
                    End If
                Next j
                commentCount = 0
                ReDim comments(1 To 50)
            End If
        End If
    Next cell
    
    ' Check each cell in the "Manufacturer Part Number" column
    For Each cell In rngManufacturerPartNumber
        If Len(cell.value) = 0 Then
            commentCount = commentCount + 1
            comments(commentCount) = "No MPN"
            cell.Interior.Color = RGB(255, 255, 0) ' Highlight in yellow if blank
            commentRows(commentCount) = cell.Row
            
            ' Populate comments in column A every 50 comments
            If commentCount = 50 Then
                For j = 1 To 50
                    If ws.Cells(commentRows(j), 1).value = "" Then
                        ws.Cells(commentRows(j), 1).value = comments(j)
                    Else
                        ws.Cells(commentRows(j), 1).value = ws.Cells(commentRows(j), 1).value & vbCrLf & comments(j)
                    End If
                Next j
                commentCount = 0
                ReDim comments(1 To 50)
            End If
        End If
    Next cell
    
    ' Populate remaining comments in column A
    If commentCount > 0 Then
        For j = 1 To commentCount
            If ws.Cells(commentRows(j), 1).value = "" Then
                ws.Cells(commentRows(j), 1).value = comments(j)
            Else
                ws.Cells(commentRows(j), 1).value = ws.Cells(commentRows(j), 1).value & vbCrLf & comments(j)
            End If
        Next j
    End If
    
    ' Remove empty spaces within comments in column A again after processing using Substitute
    For Each cell In ws.Range(ws.Cells(info.headerRow + 1, 1), ws.Cells(ws.Cells(ws.Rows.Count, 1).End(xlUp).Row, 1))
        If Len(cell.value) = 0 Then
            cell.value = Application.WorksheetFunction.Substitute(cell.value, Chr(10), "")
        Else
            cell.value = RemoveMultipleLineBreaks(cell.value)
        End If
    Next cell
    
    ' Add filters to the header row if not already applied
    If Not info.ws.AutoFilterMode Then
        info.ws.Rows(info.headerRow).AutoFilter
    End If

    ' Freeze the first row
    info.ws.Activate
    ActiveWindow.FreezePanes = False
    info.ws.Rows("2:2").Select
    ActiveWindow.FreezePanes = True
    
    
    ' Display a message indicating that the macro has finished running
    MsgBox "The macro has finished running.", vbInformation, "Macro Finished"
    
    
Cleanup:
    ' Enable screen updating and calculation
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    
    ' Release object references
    Set rngProductDescription = Nothing
    Set rngBrandName = Nothing
    Set rngItemNumber = Nothing
    Set rngManufacturerName = Nothing
    Set rngManufacturerPartNumber = Nothing
    Set cell = Nothing
    Set ws = Nothing
End Sub


Function GetSheetAndHeaderInfo_Check_If_Item_on_Multiple_Subs() As SheetHeaderInfo
    Dim info As SheetHeaderInfo
    Dim sheetName1 As String
    Dim wb As Workbook
    Dim colLetter As String
    Dim headerRowInput As String
    Dim cell As Range
    Dim found As Range
    Dim sheetList As String
    Dim sheet As Worksheet
    
    ' Set the active workbook
    Set wb = ActiveWorkbook
    
    ' Default sheet name
    sheetName1 = "Commercial Data (All)"
    
    ' Check if the default sheet exists
    On Error Resume Next
    Set info.ws = wb.Sheets(sheetName1)
    On Error GoTo 0
    
    ' If the default sheet does not exist, prompt the user to select a sheet
    If info.ws Is Nothing Then
        ' Create a list of all worksheet names in the active workbook
        sheetList = ""
        For Each sheet In wb.Sheets
            sheetList = sheetList & sheet.Name & vbCrLf
        Next sheet
        
        ' Prompt the user to select a sheet from the list
        sheetName1 = InputBox("Commercial Data (All) is not found. Please type in the name of the sheet which represents Commercial Data (All) sheet from the following list:" & vbCrLf & sheetList, "Select Sheet")
        
        ' Check if the user selected a sheet
        If sheetName1 = "" Then
            MsgBox "No sheet selected. Exiting.", vbCritical
            Exit Function
        End If
        
        ' Set worksheet with error handling
        On Error Resume Next
        Set info.ws = wb.Sheets(sheetName1)
        On Error GoTo 0
        
        ' Check if the sheet exists
        If info.ws Is Nothing Then
            MsgBox "Sheet '" & sheetName1 & "' not found. Please check the sheet name and try again.", vbCritical
            Exit Function
        End If
    End If
    
    ' Clear any filters in the sheet
    If info.ws.AutoFilterMode Then info.ws.AutoFilterMode = False
    
    ' Unhide all columns in the selected sheet
    info.ws.Cells.EntireColumn.Hidden = False
    
    ' Find the header row with "Manufacturer Part Number" in the selected sheet (exact match)
    On Error Resume Next
    For i = 1 To info.ws.Rows.Count
        Set found = info.ws.Rows(i).Find(Trim("Manufacturer Part Number"), LookIn:=xlValues, LookAt:=xlWhole)
        If Not found Is Nothing Then
            info.colManufacturerPartNumber = found.Column
            info.headerRow = i
            Exit For
        End If
    Next i
    On Error GoTo 0
    
    ' If not found, prompt the user for the column letter and header row
    If info.colManufacturerPartNumber = 0 Then
        colLetter = InputBox("Column 'Manufacturer Part Number' not found. Please provide the column letter (e.g., AB, CK):", "Column Letter")
        If colLetter = "" Then
            MsgBox "No column letter provided. Exiting.", vbCritical
            Exit Function
        End If
        headerRowInput = InputBox("Please provide the row number of the header row:", "Header Row Number")
        If headerRowInput = "" Or Not IsNumeric(headerRowInput) Then
            MsgBox "Invalid row number provided. Exiting.", vbCritical
            Exit Function
        End If
        info.headerRow = CLng(headerRowInput)
        info.colManufacturerPartNumber = Range(colLetter & info.headerRow).Column
    End If
    
    ' Find the header row with "Product New Subcategory Number & Description" or "Product New Sub Category Number & Description" in the selected sheet
    On Error Resume Next
    For i = 1 To info.ws.Rows.Count
        Set found = info.ws.Rows(i).Find(Trim("Product New Subcategory Number & Description"), LookIn:=xlValues, LookAt:=xlPart)
        If found Is Nothing Then
            Set found = info.ws.Rows(i).Find(Trim("Product New Sub Category Number & Description"), LookIn:=xlValues, LookAt:=xlPart)
        End If
        If Not found Is Nothing Then
            info.colProductNewSubcategory = found.Column
            Exit For
        End If
    Next i
    On Error GoTo 0
    
    ' If not found, prompt the user for the column letter and header row
    If info.colProductNewSubcategory = 0 Then
        colLetter = InputBox("Column 'Product New Subcategory Number & Description' or 'Product New Sub Category Number & Description' not found. Please provide the column letter (e.g., AB, CK):", "Column Letter")
        If colLetter = "" Then
            MsgBox "No column letter provided. Exiting.", vbCritical
            Exit Function
        End If
        headerRowInput = InputBox("Please provide the row number of the header row:", "Header Row Number")
        If headerRowInput = "" Or Not IsNumeric(headerRowInput) Then
            MsgBox "Invalid row number provided. Exiting.", vbCritical
            Exit Function
        End If
        info.headerRow = CLng(headerRowInput)
        info.colProductNewSubcategory = Range(colLetter & info.headerRow).Column
    End If
    
    ' Return the result
    GetSheetAndHeaderInfo_Check_If_Item_on_Multiple_Subs = info
End Function



Sub Check_If_Item_on_Multiple_Subs()
    Dim info As SheetHeaderInfo
    Dim ws As Worksheet
    Dim rngPartNumber As Range
    Dim cell As Range
    Dim dict As Object
    Dim duplicates As Collection
    Dim subcategory As String
    Dim comments() As String
    Dim commentCount As Long
    Dim commentRows(1 To 50) As Long
    Dim i As Long, j As Long, k As Long
    
    ' Disable screen updating and calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    ' Get the sheet and header info
    info = GetSheetAndHeaderInfo_Check_If_Item_on_Multiple_Subs()
    If info.ws Is Nothing Then GoTo Cleanup
    
    ' Set the worksheet
    Set ws = info.ws
    
    ' Set range for "Manufacturer Part Number" column
    Set rngPartNumber = ws.Range(ws.Cells(info.headerRow + 1, info.colManufacturerPartNumber), ws.Cells(ws.Cells(ws.Rows.Count, info.colManufacturerPartNumber).End(xlUp).Row, info.colManufacturerPartNumber))
    
    ' Clean up the text in the range
    For Each cell In rngPartNumber
        If Not IsEmpty(cell.value) Then
            cell.value = CleanUpText(cell.value)
        End If
    Next cell
    
    ' Clear specific comments and unhighlight cells in column A
    For Each cell In rngPartNumber
        If InStr(cell.Offset(0, -info.colManufacturerPartNumber + 1).value, "Item may be tendered on multiple subcats") > 0 And Len(Trim(cell.value)) = 0 Then
            cell.Offset(0, -info.colManufacturerPartNumber + 1).value = Replace(cell.Offset(0, -info.colManufacturerPartNumber + 1).value, "Item may be tendered on multiple subcats", "")
            cell.Interior.ColorIndex = xlNone
        End If
    Next cell
    
    ' Remove empty spaces within comments in column A using Substitute
    For Each cell In ws.Range(ws.Cells(info.headerRow + 1, 1), ws.Cells(ws.Cells(ws.Rows.Count, 1).End(xlUp).Row, 1))
        If Len(cell.value) = 0 Then
            cell.value = Application.WorksheetFunction.Substitute(cell.value, Chr(10), "")
        Else
            cell.value = RemoveMultipleLineBreaks(cell.value)
        End If
    Next cell
    
    ' Initialize dictionary to track duplicates
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Loop through each cell in the "Manufacturer Part Number" column
    For Each cell In rngPartNumber
        If Len(Trim(cell.value)) > 0 Then ' Skip blank cells
            If Not dict.exists(cell.value) Then
                Set dict(cell.value) = New Collection
            End If
            dict(cell.value).Add cell.Row
        End If
    Next cell
    
    ' Initialize comments array
    ReDim comments(1 To 50)
    commentCount = 0
    
    ' Check for duplicates and compare subcategory values
    For Each Key In dict.keys
        If dict(Key).Count > 1 Then
            Set duplicates = dict(Key)
            subcategory = ws.Cells(duplicates(1), info.colProductNewSubcategory).value
            For i = 2 To duplicates.Count
                If ws.Cells(duplicates(i), info.colProductNewSubcategory).value <> subcategory Then
                    ' Highlight duplicates and add comment
                    For j = 1 To duplicates.Count
                        ws.Cells(duplicates(j), info.colManufacturerPartNumber).Interior.Color = RGB(173, 216, 230) ' Light blue
                        commentCount = commentCount + 1
                        comments(commentCount) = "Item may be tendered on multiple subcats"
                        commentRows(commentCount) = duplicates(j)
                        
                        ' Populate comments in column A every 50 comments
                        If commentCount = 50 Then
                            For k = 1 To 50
                                If InStr(ws.Cells(commentRows(k), 1).value, comments(k)) = 0 Then
                                    If ws.Cells(commentRows(k), 1).value = "" Then
                                        ws.Cells(commentRows(k), 1).value = comments(k)
                                    Else
                                        ws.Cells(commentRows(k), 1).value = ws.Cells(commentRows(k), 1).value & vbCrLf & comments(k)
                                    End If
                                End If
                            Next k
                            commentCount = 0
                            ReDim comments(1 To 50)
                        End If
                    Next j
                    Exit For
                End If
            Next i
        End If
    Next Key
    
    ' Populate remaining comments in column A
    If commentCount > 0 Then
        For j = 1 To commentCount
            If InStr(ws.Cells(commentRows(j), 1).value, comments(j)) = 0 Then
                If ws.Cells(commentRows(j), 1).value = "" Then
                    ws.Cells(commentRows(j), 1).value = comments(j)
                Else
                    ws.Cells(commentRows(j), 1).value = ws.Cells(commentRows(j), 1).value & vbCrLf & comments(j)
                End If
            End If
        Next j
    End If
    
    ' Additional loop to ensure comments and highlights are removed for blank cells
    For Each cell In rngPartNumber
        If InStr(cell.Offset(0, -info.colManufacturerPartNumber + 1).value, "Item may be tendered on multiple subcats") > 0 Then
            If Len(Trim(cell.value)) = 0 Then
                cell.Offset(0, -info.colManufacturerPartNumber + 1).value = Replace(cell.Offset(0, -info.colManufacturerPartNumber + 1).value, "Item may be tendered on multiple subcats", "")
                cell.Interior.ColorIndex = xlNone
            End If
        End If
    Next cell
    
    ' Remove empty spaces within comments in column A again after processing using Substitute
    For Each cell In ws.Range(ws.Cells(info.headerRow + 1, 1), ws.Cells(ws.Cells(ws.Rows.Count, 1).End(xlUp).Row, 1))
        If Len(cell.value) = 0 Then
            cell.value = Application.WorksheetFunction.Substitute(cell.value, Chr(10), "")
        Else
            cell.value = RemoveMultipleLineBreaks(cell.value)
        End If
    Next cell
    
    ' Add filters to the header row if not already applied
    If Not info.ws.AutoFilterMode Then
        info.ws.Rows(info.headerRow).AutoFilter
    End If

    ' Freeze the first row
    info.ws.Activate
    ActiveWindow.FreezePanes = False
    info.ws.Rows("2:2").Select
    ActiveWindow.FreezePanes = True
    
    ' Display a message indicating that the macro has finished running
    MsgBox "The macro has finished running.", vbInformation, "Macro Finished"
    
Cleanup:
    ' Enable screen updating and calculation
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    
    ' Release object references
    Set rngPartNumber = Nothing
    Set cell = Nothing
    Set dict = Nothing
    Set duplicates = Nothing
    Set ws = Nothing
End Sub

Sub CountCommercial_Errors()
    Dim rng As Range
    Dim cell As Range
    Dim commentCounts As Object
    Dim comment As Variant
    Dim msg As String
    Dim cellComments As Variant
    Dim i As Integer
    
    ' Initialize the dictionary to hold comment counts
    Set commentCounts = CreateObject("Scripting.Dictionary")
    
    ' List of possible comments
    Dim commentsList As Variant
    commentsList = Array("Key not found in clinical sheets", "Subcategory not found", _
                         "Item not found in clinical sheets", "Item found but only item no. matched", _
                         "No Description", "No Brand", "No Item Number", "No MFG", _
                         "No MPN", "Item may be tendered on multiple subcats")
    
    ' Initialize counts to zero
    For Each comment In commentsList
        commentCounts(comment) = 0
    Next comment
    
    ' Set the range to the selected range
    Set rng = Selection
    
    ' Count the occurrences of each comment
    For Each cell In rng
        If Not IsEmpty(cell.value) Then
            ' Clean up the cell content
            cell.value = Replace(cell.value, "_x000D_", "")
            cell.value = Replace(cell.value, vbCr, vbLf)
            cell.value = Replace(cell.value, vbLf & vbLf, vbLf)
            cell.value = Trim(cell.value)
            
            ' Split the cell content by line breaks
            cellComments = Split(cell.value, vbLf)
            For i = LBound(cellComments) To UBound(cellComments)
                ' Trim leading and trailing spaces
                cellComments(i) = Trim(cellComments(i))
                For Each comment In commentsList
                    If cellComments(i) = comment Then
                        commentCounts(comment) = commentCounts(comment) + 1
                    End If
                Next comment
            Next i
        End If
    Next cell
    
    ' Create the message for the pop-up
    msg = "Comments Breakdown" & vbCrLf & vbCrLf
    For Each comment In commentsList
        msg = msg & "- " & comment & " = " & commentCounts(comment) & vbCrLf
    Next comment
    
    ' Display the message
    MsgBox msg, vbInformation, "Comments Breakdown"
    
    ' Add filters to the header row if not already applied
    If Not rng.Worksheet.AutoFilterMode Then
        rng.Worksheet.Rows(1).AutoFilter
    End If
    
    ' Freeze the first row
    rng.Worksheet.Activate
    ActiveWindow.FreezePanes = False
    rng.Worksheet.Rows("2:2").Select
    ActiveWindow.FreezePanes = True
End Sub



