Attribute VB_Name = "Descriptions_EK_Main"
Public wbc As Workbook 'CLINICAL FILE
Public wscl As Worksheet 'clinical sheets
Public wsco As Worksheet 'commercial sheets
Public wscat As Worksheet 'Cat_Subcat_Major_Minor sheets'

Public str As String

Public scin() As String 'Subcat in index sheet
Public sccat() As String 'subcat in Cat_Subcat_Major_Minor sheet

Public catempty As Boolean
Public mandatory As Range  'mandatory columns
Public mancol() As Integer 'mandatory columns numbers array
Public shouldExit As Boolean 'exit query if this is true

Sub setsheets()
shouldExit = False
Application.ScreenUpdating = False
Application.Calculation = xlManual
Dim StrFlNm As String 'commercial - clinical file name

i = 0

Set wbc = ActiveWorkbook


If InStr(1, ActiveSheet.Name, "(Clinical)") > 0 And InStr(1, ActiveSheet.Name, "Category") > 0 Then

    Set wscl = ActiveSheet
    StrFlNm = "Commercial Data (All)"
    If Evaluate("ISREF('" & StrFlNm & "'!A1)") Then
    
        Set wsco = wbc.Sheets(StrFlNm)
    Else
        shouldExit = True
        MsgBox "Commercial Data (All) File Not Found"
        Exit Sub
    End If

Else
        MsgBox "This function only works in TRW Clinical sheet. Sheet name must be - Category# (Clinial) -"
        shouldExit = True
        Exit Sub

End If

Set wscat = wbc.Sheets("Cat_Subcat_Major_Minor")

Application.Calculation = xlCalculationAutomatic

End Sub


Sub TRWCheckUniqueID()

Application.ScreenUpdating = False
Application.Calculation = xlManual

Call setsheets
If shouldExit Then
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
Else
    Call checkuniqid
End If

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub

Sub checkuniqid()

Dim i As Integer
Dim lstrowcl As Integer 'clinical last row
Dim lstrowco As Integer 'commercial last row
Dim rscl As Integer 'clinical first row
Dim rsco As Integer 'commercial first row

Dim uniqidcl As Integer 'clinical MDT KEY column
Dim uniqidco As Integer 'commercial MDT KEY column
Dim counterr As Integer
counterr = 0
rscl = findincol("MDT Issue Flag", wscl.Columns(1)) 'Assuming is the first column is MDT Issue Flag in Clinical Sheets. Find the First Row where table starts
uniqidcl = findinrow("MDT KEY", wscl.Rows(rscl))
Dim testcolumn As Integer 'column to find last row in clinical file. Unique Identification will be used to define
testcolumn = findinrow("Unique Identification", wscl.Rows(rscl))
lstrowcl = findlrw(wscl.Range(wscl.Cells(rscl + 1, testcolumn), wscl.Cells(rscl + 99999, testcolumn))) - 1

If testcolumn = 0 Or lstrowcl = 0 Then

    MsgBox "Last row of clinical file wasn't found" & vbCrLf & _
        "Ensure column Unique Identification appears" & vbCrLf & _
        "Check if the last rows are consistent across the sheet"
    Exit Sub
End If

' check if all rows has MDT KEY against uniqueID


'find MDT KEY in Commercial sheet

rsco = findincol("MDT Issue Flag", wsco.Columns(1)) 'Assuming is the first column is MDT Issue Flag in Commercial Sheets
uniqidco = findinrow("MDT KEY", wsco.Rows(rsco))
lstrowco = findlrw(wsco.Range(wsco.Cells(rsco + 1, uniqidco), wsco.Cells(rsco + 99999, uniqidco))) - 1

For i = rscl + 1 To lstrowcl
    
    ' Check if the unique Key Missing
    If Trim(CStr(wscl.Cells(i, uniqidcl).value)) = "" Or IsEmpty(wscl.Cells(i, uniqidcl).value) Or wscl.Cells(i, uniqidcl).value = " |  |  | " Or Trim(CStr(wscl.Cells(i, uniqidcl).value)) = "#N/A" Then
    
       wscl.Cells(i, 1).value = wscl.Cells(i, 1).value & vbLf & "MDT KEY missing in Clinical"
       If Left(wscl.Cells(i, 1).value, 1) = vbLf Then wscl.Cells(i, 1).value = Right(wscl.Cells(i, 1).value, Len(wscl.Cells(i, 1).value) - 1)
       wscl.Cells(i, uniqidcl).Interior.ColorIndex = 31
       
    counterr = counterr + 1 ' error count
    
    Else
        If findincol(wscl.Cells(i, uniqidcl).value, wsco.Range(wsco.Cells(rsco, uniqidco), wsco.Cells(lstrowco, uniqidco))) = 0 Then
        ' Check if MDT Key Missing in Commercial
           wscl.Cells(i, 1).value = wscl.Cells(i, 1).value & vbLf & "MDT KEY missing in Commercial"
           If Left(wscl.Cells(i, 1).value, 1) = vbLf Then wscl.Cells(i, 1).value = Right(wscl.Cells(i, 1).value, Len(wscl.Cells(i, 1).value) - 1)
           wscl.Cells(i, uniqidcl).Interior.ColorIndex = 31
           counterr = counterr + 1 ' error count
        End If
        
        If i <> lstrowcl Then 'if not in last row, check if MDT KEY is duplicated in Clinical file
        
            If findincol(wscl.Cells(i, uniqidcl).value, wscl.Range(wscl.Cells(i + 1, uniqidcl), wscl.Cells(lstrowcl, uniqidcl))) > 1 Then
            
               wscl.Cells(i, 1).value = wscl.Cells(i, 1).value & vbLf & "MDT KEY duplicated"
               'remove VBLF if it is the first char
               If Left(wscl.Cells(i, 1).value, 1) = vbLf Then wscl.Cells(i, 1).value = Right(wscl.Cells(i, 1).value, Len(wscl.Cells(i, 1).value) - 1)
               counterr = counterr + 1 ' error count
            End If
        End If
    End If

Next i

    MsgBox "Checked Rows: " & lstrowcl & vbCrLf & _
           "Issue Count: " & counterr & vbCrLf




Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub

Sub TRWCheckMandatoryColumns()




Dim i As Integer
Dim lstrowcl As Integer 'clinical last row
Dim lstrowco As Integer 'commercial last row
Dim rscl As Integer 'clinical first row
Dim rsco As Integer 'commercial first row
Dim uniqidcl As Integer 'clinical MDT KEY column
Dim uniqidco As Integer 'commercial MDT KEY column
Dim counterr As Integer
counterr = 0

Call setsheets
If shouldExit Then
    Exit Sub
Else
    Call getmancol
End If
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

rscl = findincol("MDT Issue Flag", wscl.Columns(1)) 'Assuming is the first column is MDT Issue Flag in Clinical Sheets. Find the First Row where table starts
uniqidcl = findinrow("MDT KEY", wscl.Rows(rscl))
Dim testcolumn As Integer 'column to find last row in clinical file. Unique Identification will be used to define
testcolumn = findinrow("Unique Identification", wscl.Rows(rscl))
lstrowcl = findlrw(wscl.Range(wscl.Cells(rscl + 1, testcolumn), wscl.Cells(rscl + 99999, testcolumn))) - 1

If testcolumn = 0 Or lstrowcl = 0 Then

    MsgBox "Last row of clinical file wasn't found" & vbCrLf & _
            "Ensure column Unique Identification appears" & vbCrLf & _
            "Check if the last rows are consistent across the sheet"
    Exit Sub
End If


For i = rscl + 1 To lstrowcl 'loop through each row

    For j = 1 To UBound(mancol)
    
        If Trim(CStr(wscl.Cells(i, mancol(j)).value)) = "" Or UCase(Trim(CStr(wscl.Cells(i, mancol(j)).value))) = "NULL" Or UCase(Trim(CStr(wscl.Cells(i, mancol(j)).value))) = "EXCEMPT" _
        Or UCase(Trim(CStr(wscl.Cells(i, mancol(j)).value))) = "N/A" Or UCase(Trim(CStr(wscl.Cells(i, mancol(j)).value))) = "#N/A" _
        Or UCase(Trim(CStr(wscl.Cells(i, mancol(j)).value))) = "NOT AVALIABLE" Then
        
           wscl.Cells(i, 1).value = wscl.Cells(i, 1).value & vbLf & "Mandatory info missing"
            'remove VBLF if it is the first char
           If Left(wscl.Cells(i, 1).value, 1) = vbLf Then wscl.Cells(i, 1).value = Right(wscl.Cells(i, 1).value, Len(wscl.Cells(i, 1).value) - 1)
           wscl.Cells(i, mancol(j)).Interior.ColorIndex = 44
           counterr = counterr + 1 ' error count
        
        End If
    Next j

Next i

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    MsgBox "Checked Rows: " & lstrowcl & vbCrLf & _
           "Issue Count: " & counterr & vbCrLf




End Sub

Sub TRWCheckSubcat()

Application.ScreenUpdating = False
Application.Calculation = xlManual

Dim i As Integer
Dim lstrowcl As Integer 'clinical last row
Dim rscl As Integer 'clinical first row
Dim uniqidcl As Integer 'clinical MDT KEY column
Dim catcl As Integer 'clinical category column
Dim counterr As Integer
counterr = 0

Call setsheets

If shouldExit Then
    Exit Sub
Else
    Call getsubcatmj
End If


rscl = findincol("MDT Issue Flag", wscl.Columns(1)) 'Assuming is the first column is MDT Issue Flag in Clinical Sheets. Find the First Row where table starts
uniqidcl = findinrow("MDT KEY", wscl.Rows(rscl))
Dim testcolumn As Integer 'column to find last row in clinical file. Unique Identification will be used to define
testcolumn = findinrow("Unique Identification", wscl.Rows(rscl))
lstrowcl = findlrw(wscl.Range(wscl.Cells(rscl + 1, testcolumn), wscl.Cells(rscl + 99999, testcolumn))) - 1

catcl = findinrow("Product New Subcategory Number & Description", wscl.Rows(rscl))

If catcl = 0 Then

    MsgBox "Product New Subcategory Number & Description"
    Exit Sub

End If

For i = rscl + 1 To lstrowcl

    If Not (isinarray(wscl.Cells(i, catcl), sccat())) Then 'check if subcat exist in Cat_Subcat_Major_Minor Sheet
        
        wscl.Cells(i, 1).value = wscl.Cells(i, 1).value & vbLf & "Subcategory not in Cat_Subcat_Major_Minor sheet"
        wscl.Cells(i, catcl).Interior.ColorIndex = 16
        'remove VBLF if it is the first char
        If Left(wscl.Cells(i, 1).value, 1) = vbLf Then wscl.Cells(i, 1).value = Right(wscl.Cells(i, 1).value, Len(wscl.Cells(i, 1).value) - 1)
        counterr = counterr + 1 ' error count
    End If

Next i


    MsgBox "Checked Rows: " & lstrowcl & vbCrLf & _
           "Issue Count: " & counterr & vbCrLf

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub


Sub TRWCheckLatex()

Application.ScreenUpdating = False
Application.Calculation = xlManual

Dim i As Integer
Dim lstrowcl As Integer 'clinical last row
Dim rscl As Integer 'clinical first row
Dim uniqidcl As Integer 'clinical MDT KEY column
Dim ltxc As Integer 'clinical latex column
Call setsheets
Dim counterr As Integer
counterr = 0

If shouldExit Then
    Exit Sub
End If

rscl = findincol("MDT Issue Flag", wscl.Columns(1)) 'Assuming is the first column is MDT Issue Flag in Clinical Sheets. Find the First Row where table starts
uniqidcl = findinrow("MDT KEY", wscl.Rows(rscl))
Dim testcolumn As Integer 'column to find last row in clinical file. Unique Identification will be used to define
testcolumn = findinrow("Unique Identification", wscl.Rows(rscl))
lstrowcl = findlrw(wscl.Range(wscl.Cells(rscl + 1, testcolumn), wscl.Cells(rscl + 99999, testcolumn))) - 1

ltxc = findinrow("Does the product Contain Latex?", wscl.Range(wscl.Cells(rscl, 1), wscl.Cells(rscl, 100))) 'search in first 100 column then exit

If testcolumn = 0 Or lstrowcl = 0 Then

    MsgBox "Last row of clinical file wasn't found" & vbCrLf & _
        "Ensure column Unique Identification appears" & vbCrLf & _
        "Check if the last rows are consistent across the sheet"
        
    Exit Sub

End If

If ltxc = 0 Then

MsgBox "Does the product Contain Latex? column doesn't appear in clinical sheet!!"
    Exit Sub
End If

For i = rscl + 1 To lstrowcl

    If checklatexstatus(wscl.Cells(i, ltxc), wscl.Cells(i, ltxc + 1)) <> "" Then 'check if subcat exist in Index Sheet
        
        wscl.Cells(i, 1).value = wscl.Cells(i, 1).value & vbLf & "Latex Inconsistency"
        wscl.Cells(i, ltxc).Interior.ColorIndex = 43
        'remove VBLF if it is the first char
        If Left(wscl.Cells(i, 1).value, 1) = vbLf Then wscl.Cells(i, 1).value = Right(wscl.Cells(i, 1).value, Len(wscl.Cells(i, 1).value) - 1)
        counterr = counterr + 1 ' error count
        
    End If
    
Next i


    MsgBox "Checked Rows: " & lstrowcl & vbCrLf & _
           "Issue Count: " & counterr & vbCrLf

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub


Sub getmancol()

ReDim mancol(99) '99 doesn't mean anything. just to initialize
Application.ScreenUpdating = True

On Error Resume Next ' Handle cancel action
    Set mandatory = Application.InputBox("Please select a mandatory columns range:", Type:=8)
On Error GoTo 0 ' Re-enable normal error handling

ReDim mancol(1 To mandatory.Columns.Count)

For i = 1 To mandatory.Columns.Count 'assign mandatory column numbers to
    
    mancol(i) = mandatory.Columns(i).Column

Next i

End Sub

Sub getsubcatmj()  ' module gets the subcategories  Cat_Subcat_Major_Minor

Dim subcatmjc As Integer   'subcat column in major minor noun sheet
Dim subcatmjr As Integer   'subcat row in major minor noun sheet'ublic sccat() As String
'Subcategory Number Subcategory Description
Dim lstrow As Integer 'lastrow
Dim asize As Integer 'array size

asize = 0

subcatmjc = 5 'Trang confirmed that Subcategory Number Subcategory Description column will always be column E
subcatmjr = 1 ' assuming that subcat names will start from first row
lstrow = findlrw(wscat.Range(wscat.Cells(subcatmjr, subcatmjc), wscat.Cells(subcatmjr + 250, subcatmjc))) 'if more than 250 subcat increase the constant.

ReDim sccat(lstrow)

For i = subcatmjr + 1 To lstrow

    sccat(asize) = wscat.Cells(i, subcatmjc).value
    asize = asize + 1
Next i

ReDim Preserve sccat(asize - 1)

End Sub

Sub TRWInsertErrorColumn()

Dim countadded As Integer
countadded = 0

shouldExit = False
Application.ScreenUpdating = False
Application.Calculation = xlManual
Dim ws As Worksheet


For Each ws In ActiveWorkbook.Worksheets 'check if the sheet name has (Clinical) or (Commercial)
    If InStr(1, ws.Name, "(Clinical)") > 0 Then
        
       If findincol("MDT Issue Flag", ws.Range(ws.Range("A1"), ws.Range("A30"))) = 0 Then
       
            ws.Range("A:A").EntireColumn.Insert Shift:=xlToRight
            ws.Range("A1").value = "MDT Issue Flag"
            Columns("B:B").Select
            Selection.Copy
            Columns("A:A").Select
            Selection.PasteSpecial Paste:=xlPasteFormats
            Selection.WrapText = True
            Selection.EntireColumn.AutoFit
            Application.CutCopyMode = False
            countadded = countadded + 1
       End If
        
    

    
    ElseIf InStr(1, ws.Name, "Commercial Data (All)") > 0 Then
    
        If findincol("MDT Issue Flag", ws.Range(ws.Range("A1"), ws.Range("A30"))) = 0 Then
       
            ws.Range("A:A").EntireColumn.Insert Shift:=xlToRight
            ws.Range("A1").value = "MDT Issue Flag"
            Columns("B:B").Select
            Selection.Copy
            Columns("A:A").Select
            Selection.PasteSpecial Paste:=xlPasteFormats
            Selection.WrapText = True
            Selection.EntireColumn.AutoFit
            Application.CutCopyMode = False
            countadded = countadded + 1
       End If
    
    End If
    
    

Next

    MsgBox "Columns Added : " & countadded

End Sub



