Attribute VB_Name = "Descriptions_RC"
'https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/cc296089(v=office.12)?redirectedfrom=MSDN


Public Sub Description_Column_Add()
'Template setup
' find the last used column
'add description columns to the right of this
'add formula for building the description
'Columns to add
'Major Noun  Minor Noun 1    Minor Noun 2    Attribute 1 Attribute 2 Attribute 3 Attribute 4 Attribute 5 Attribute 6 Attribute 7 Attribute 8 Attribute 9 Attribute 10    Additional Freetext Brand   Item Number on Packaging    Single-Use  Latex Status    Sterile HSV PROPOSED DESCRIPTION    Length
'copy from hpv_macros(Descriptions) a1:U2 sheet to last column plus 1 at same row as existing headers
'copy formula in t2 and u2 down through the used rows
Dim wsrc As Workbook
Dim wdest As Workbook
Dim ssrc As Worksheet
Dim sdest As Worksheet
Dim dpath As String
Dim dFile As String
Dim dCol As Integer
Dim dRow As Integer
Dim lastRow As Integer
Dim lastCol As Integer
Dim dadd As Variant
Dim subCatCol As Variant
Dim subCatColLetter As String
Dim repText As String
Dim shNew As Worksheet

'set wsrc
Set wdest = ActiveWorkbook
'set wsrc =

dFile = "HPV_Macros.xlam"
    Set wsrc = Application.Workbooks.Item(dFile)
    Set sdest = wdest.ActiveSheet
    Set ssrc = wsrc.Worksheets("Descriptions")
    Set sdest = ActiveSheet
    
'add sheet Cat_Subcat_Major_Minor if not existing
If Not sheetExists("Cat_Subcat_Major_Minor") Then
    Set shNew = wdest.Worksheets.Add()
    shNew.Name = "Cat_Subcat_Major_Minor"
End If
    
    sdest.Activate
    dCol = sdest.Cells.SpecialCells(xlCellTypeLastCell).Column + 1
    'drow = 2
    dadd = Application.InputBox("Select row to add the header", Type:=8).Address
    dRow = Range(dadd).Row
    subCatCol = Application.InputBox("Select the column containing the sub category name", Type:=8).Address
    subCatColLetter = subCatCol
    subCatColLetter = "$" & TextFilter(subCatColLetter)
    
    ssrc.Range(ssrc.Cells(1, 1), ssrc.Cells(2, 21)).Copy 'sdest.Cells(drow, dcol)
    sdest.Cells(dRow, dCol).PasteSpecial xlPasteAll
    sdest.Cells(dRow, dCol).PasteSpecial xlPasteColumnWidths

    lastRow = sdest.Cells.SpecialCells(xlCellTypeLastCell).Row
    lastCol = sdest.Cells.SpecialCells(xlCellTypeLastCell).Column
    
    sdest.Range(Cells(dRow + 1, lastCol - 1), Cells(dRow + 1, lastCol)).Copy
    sdest.Range(Cells(dRow + 1, lastCol - 1), Cells(lastRow, lastCol)).PasteSpecial xlPasteFormulas
    
    'update the major/minor noun lookup formula. Replace $Z2 with column letter selected with subCatCol
    sdest.Range(Cells(dRow + 1, dCol), Cells(dRow + 1, dCol + 5)).Replace What:="$Z", Replacement:=subCatColLetter, LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        
    sdest.Range(Cells(dRow + 1, dCol), Cells(dRow + 1, dCol + 5)).Replace What:="[HPV_Macros.xlam]", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    
    repText = "$" & dRow & ",Cat_Subcat_Major_Minor!"
    sdest.Range(Cells(dRow + 1, dCol), Cells(dRow + 1, dCol + 5)).Replace What:="$1,Cat_Subcat_Major_Minor!", Replacement:=repText, LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        
    sdest.Range(Cells(dRow + 1, dCol), Cells(dRow + 1, dCol + 5)).Copy
    sdest.Range(Cells(dRow + 1, dCol), Cells(lastRow, dCol + 5)).PasteSpecial xlPasteFormulas
    Application.CutCopyMode = False
    sdest.Cells(dRow, dCol).Select
End Sub


Public Sub DescLength()
'Identify descriptions over 140 characters
'highlight in red any descriptions longer than 140 characters
'remove highligh where length <=140

Dim descLen As Integer
Dim inputLength As String
Dim selected As Integer
Dim highlighted As Integer
Dim c As Range
Const favColor = 50 '34

    inputLength = InputBox("Specify the max length of the field", "Max Length", 140)
    If IsNumeric(inputLength) Then
        descLen = CInt(inputLength)
    Else
        MsgBox "You need to specify a whole number"
        Exit Sub
    End If

If ActiveWorkbook Is Nothing Then
    Exit Sub
End If

On Error GoTo ErrorExit

    selected = 0
    highlighted = 0

    datasrc = Selection.Address
    If Range(datasrc).Count > 1 Then
        Range(datasrc).SpecialCells(xlCellTypeVisible).Select
    Else
        Range(datasrc).Select
    End If
    
SuspendAutoCalc
    For Each c In Selection.Cells
        selected = selected + 1
            If Len(c.value) > descLen Then
                'c.Interior.ColorIndex = favColor
                With c.Characters(Start:=inputLength + 1, Length:=Len(c)).Font
                    '.FontStyle = "Bold Italic"
                    .FontStyle = "Bold"
                    .ColorIndex = favColor
                End With
                With c.Characters(Start:=1, Length:=inputLength).Font
                    .FontStyle = "Normal"
                    .ColorIndex = 1
                End With
                highlighted = highlighted + 1
            Else
                c.Interior.Pattern = xlNone
                c.Font.FontStyle = "Normal"
                c.Font.ColorIndex = 1
            End If
    Next c
 
 'restore the user's calculation setting
    ResumeAutoCalc
    MsgBox "Cells Tested: " & selected & vbCr & "Cells highlighted: " & highlighted
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

Public Sub DescContainsMissing()
'Where there are  ***MISSING*** in part number, brand name
'highligh any cell containng ***MISSING***
'Remove highlight for other cells
Dim selected As Integer
Dim highlighted As Integer
Const favColor = 19

If ActiveWorkbook Is Nothing Then
    Exit Sub
End If

On Error GoTo errorHandler

    Dim c As Range
    selected = 0
    highlighted = 0

    datasrc = Selection.Address
    If Range(datasrc).Count > 1 Then
        Range(datasrc).SpecialCells(xlCellTypeVisible).Select
    Else
        Range(datasrc).Select
    End If
    
SuspendAutoCalc
    For Each c In Selection.Cells
        selected = selected + 1
            If InStr(1, c.value, "***MISSING***") > 0 Then
                c.Interior.ColorIndex = favColor
                highlighted = highlighted + 1
            Else
                c.Interior.Pattern = xlNone
            End If
    Next c
    MsgBox "Cells Tested: " & selected & vbCr & "Cells highlighted: " & highlighted
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

Public Sub FindDuplicates()
'Duplicate checking of description minus part number
'find duplicate values excluding the part number at the end of the description
'for each cell drop the text after the last space
'use array of descriptions less the part number with the address
'populate array from selected cells
'search through array and create array of description counts
'apply cell highlight to cells with description counts >1

Dim selected As Integer
Dim highlighted As Integer
Const favColor = 35

selected = 0
highlighted = 0

If ActiveWorkbook Is Nothing Then
    Exit Sub
End If

On Error GoTo ErrorExit

Dim desc() As String
Dim cellAdd() As String
Dim descCount() As Integer
Dim i As Integer    'loop counter
Dim cellCount As Integer

cellCount = Selection.Cells.Count
i = 0
'ReDim desc(cellCount)
'ReDim cellAdd(cellCount)
'ReDim descCount(cellCount)
ReDim desc(i)
ReDim cellAdd(i)
ReDim descCount(i)

'populate description, address and desc count arrays from selected cells
For Each c In Selection.Cells
    If c.value <> "" Then
        desc(i) = Left(c.value, Len(c.value) - InStr(1, StrReverse(c.value), " "))
        cellAdd(i) = c.Address
        descCount(i) = 1
        i = i + 1
        ReDim Preserve desc(i)
        ReDim Preserve cellAdd(i)
        ReDim Preserve descCount(i)
   End If
Next c

'compare each desc with the rest of the array looking for duplicates. add to the descCount
For i = 0 To UBound(cellAdd()) 'cellCount
    For j = (1 + i) To UBound(cellAdd()) 'cellCount
        If descCount(j) = 1 Then
            If desc(i) = desc(j) Then
                descCount(i) = descCount(i) + 1
                descCount(j) = descCount(j) + 1
            End If
        End If
    Next j
Next i

'highlight any cells where the descCount is >1
For i = 0 To UBound(cellAdd()) 'cellCount
    selected = selected + 1
    If cellAdd(i) <> "" Then
        If descCount(i) > 1 Then
            Range(cellAdd(i)).Interior.ColorIndex = favColor ' set highlight for cell address
            highlighted = highlighted + 1
        Else
            Range(cellAdd(i)).Interior.Pattern = xlNone
        End If
    Else: Exit For
    End If
Next i


MsgBox "Cells Tested: " & selected & vbCr & "Cells highlighted: " & highlighted

Exit Sub
ErrorExit:
        Debug.Print Err.Number & vbLf & Err.Description
End Sub


Sub TrimRange()
'Remove leading and trailing spaces
'Identify where there are double and triple spacings in all SCIT cleansed columns
'Identify non-standard characters (tab, carriage return…)
    Dim rng As Range
    Set rng = Application.Selection
    rng.value = Application.Trim(rng)

End Sub


Public Sub Trim_Extra()

'Remove leading and trailing spaces
'Identify where there are double and triple spacings in all SCIT cleansed columns
'Identify non-standard characters (tab, carriage return…)

'Highlight a section of cells
'step through cell by cell trimming the contents
'included in removal of HTML non-breaking space ascii 160 as this was showing up in some supplier data
'included removal of double spaces
'

Dim cellCount As Integer
Dim cellCorrected As Integer
Dim valueBeforeCorrection As String

cellCount = 0
cellCorrected = 0

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
        cellCount = Selection.SpecialCells(xlCellTypeVisible).Count
    End If
    For Each currCell In Selection.Cells
        If Not Application.IsNA(currCell.value) Then
            valueBeforeCorrection = currCell.value 'currCell.Formula
            'currCell.Formula = Application.WorksheetFunction.Trim(Replace(Replace(Replace(Trim(currCell.value), Chr(160), ""), Chr(10), ""), Chr(13), ""))
            currCell.value = Application.WorksheetFunction.Trim(Replace(Replace(Replace(Replace(Replace(Replace(Trim(currCell.value), Chr(9), " "), Chr(160), ""), Chr(10), ""), Chr(13), ""), ChrW(8195), " "), "  ", " "))
            'If currCell.Formula <> valueBeforeCorrection Then
            If currCell.value <> valueBeforeCorrection Then
                cellCorrected = cellCorrected + 1
            End If
        Else
            NACount = NACount + 1
        End If
    Next
    'restore the user's calculation setting
    ResumeAutoCalc
    If NACount > 0 Then
        MsgBox "Your data contains " & NACount & " cells with #N/A"
    End If
    
    MsgBox "Selected: " & cellCount & " cells" & vbCr & "Corrections made: " & cellCorrected
    Exit Sub
    
ErrorExit:
    
    'restore the user's calculation setting
    ResumeAutoCalc
    'the ErrorHandler code should only be executed if there is an error
    Exit Sub
errorHandler:
        Debug.Print Err.Number & vbLf & Err.Description
        MsgBox currCell.Address & " : " & Err.Number & vbLf & Err.Description
        Resume Next
        Resume ErrorExit

End Sub

Public Sub repeatWords()
'Identify where there are repeated or more than 2 iteration of word in string
'for each cell in range, use wordcount to get repeated words
'for each repeated word, find position in sting where the word is located
'use location of word with to set the text to bold and/or a colour
'With ActiveCell.Characters(Start:=88, Length:=18).Font
'        .Name = "Segoe UI"
'        .FontStyle = "Bold Italic"
'        .Size = 11
'        .Strikethrough = False
'        .Superscript = False
'        .Subscript = False
'        .OutlineFont = False
'        .Shadow = False
'        .Underline = xlUnderlineStyleSingle
'        .ThemeColor = xlThemeColorAccent2
'        .TintAndShade = 0
'        .ThemeFont = xlThemeFontNone
'    End With
' call WordCountDesc will return arrary of words found multiple times in the cell
Dim repeatWords() As String
Dim wordList() As String
Dim selected As Integer
Dim highlighted As Integer
Const favColor = 29

   
selected = 0
highlighted = 0

'repeatWords = ""
    If Selection.Count > 1 Then
        Selection.SpecialCells(xlCellTypeVisible).Select
        cellCount = Selection.SpecialCells(xlCellTypeVisible).Count
    End If
    For Each c In Selection.Cells
        selected = selected + 1
        
        If c.value <> "" Then
            c.Font.Bold = False
            'repeatWords = WordCountDesc(c.value, 2) 'use wheen wordCountDesc is set to return a string of repeated words
            repeatWords() = WordCountDesc(c.value, 2) 'use wheen wordCountDesc is set to return an array of repeated words
                If Not Not repeatWords Then
                    highlighted = highlighted + 1
                    ReDim Preserve repeatWords(UBound(repeatWords))
                    'wordList = Split(repeatWords, "|")
                    'ReDim Preserve wordList(UBound(wordList))
                    'go looking for the words in the cell and set them to bold
                    'MsgBox c.Address & ": " & Join(repeatWords, "|") & " : " & UBound(repeatWords)
                    'highlight words that have been repeated
                    'find locations of each instance of the repeated word, or walk through the cell contents looking for the word and setting the format
                    For Each w In repeatWords()
                        i = 1
                        wstart = 1
                        Do While i < Len(c.value) And wstart > 0
                            'MsgBox c.Address & ": " & w
                            wstart = InStr(i, c.value, w, vbTextCompare)
                            If wstart > 0 Then
                                With c.Characters(Start:=wstart, Length:=Len(w)).Font
                                    '.FontStyle = "Bold Italic"
                                    .FontStyle = "Bold"
                                    .ColorIndex = favColor
                                End With
                                i = wstart + Len(w)
                            End If
                            'MsgBox c.Address & ": " & w & " : " & wstart
                        Loop
                    Next w
                End If
            ReDim repeatWords(1)
        End If
    Next c
    
    MsgBox "Cells Tested: " & selected & vbCr & "Descriptions Effected: " & highlighted
End Sub

Function WordCountDesc(strText As String, Optional minOccurance As Integer, Optional strExclList As Range, Optional removeNumeric As Boolean) As String()
'Read list of words from exclusion list into array
'read text into array, broken by words
'remove any words from text array where they match with exclusion list
'calc frequency of remaining words
'output list of words in order of frequency  top 3 to start with


Dim arrExclList() As String
Dim arrStrText() As String
Dim arrOut() As String
Dim arrWordCount() As Integer

'If minOccurance Is Nothing Then
'    minOccurance = 2
'End If

arrStrText = Split(Trim(Replace(strText, " ", ",")), ",")
ReDim Preserve arrStrText(UBound(arrStrText))
ReDim Preserve arrWordCount(UBound(arrStrText))
'ReDim arrOut(0) ', 2)

If strExclList Is Nothing Then
Else
    'arrExclList = Split(strExclList & strExclList2 & strExclList3 & strExclList4, ",")
    'arrExclList = Split(strExclList, ",")
    ReDim arrExclList(strExclList.Rows.Count)
    For i = 1 To strExclList.Count
        arrExclList(i) = strExclList(i) 'strExclList(1, i)
    Next i
    
    'remove words from the exclusion list
    For i = 0 To UBound(arrStrText)
        For j = 1 To UBound(arrExclList)
            If arrStrText(i) = arrExclList(j) Then
                arrStrText(i) = ""
                Exit For
            End If
        Next
    Next
End If
    
''remove number and words with 1 or 2 letters then numbers
If removeNumeric Then
    For i = 0 To UBound(arrStrText)
        If arrStrText(i) <> "" Then
            If IsNumeric(arrStrText(i)) Then
                arrStrText(i) = ""
            End If
            For j = 1 To Len(arrStrText(i)) 'IIf(Len(arrStrText(i)) <= 3, Len(arrStrText(i)) - 1, 3)
                If IsNumeric(Mid(arrStrText(i), j, 1)) Then
                    arrStrText(i) = ""
                    Exit For
                End If
            Next
        End If
    Next
End If

'count occurance rate of words
For i = 0 To UBound(arrStrText)
    If arrStrText(i) <> "" Then
        arrWordCount(i) = 1
        For j = i + 1 To UBound(arrStrText)
            If i <> j And UCase(arrStrText(i)) = UCase(arrStrText(j)) Then
                arrWordCount(i) = arrWordCount(i) + 1
                arrStrText(j) = ""
            End If
        Next
    End If
Next

'Order the output array by most frequent
''find highest frequency
intMaxScore = 0
For i = 0 To UBound(arrWordCount)
    If arrWordCount(i) > intMaxScore Then
        intMaxScore = arrWordCount(i)
    End If
Next

''list words with that score
''subtract 1
''repeat til 1

'make output arrary with remaining keywords only
outCount = 0
For mc = intMaxScore To minOccurance Step -1
    For i = 0 To UBound(arrStrText)
        If arrStrText(i) <> "" And arrWordCount(i) = mc Then
            ReDim Preserve arrOut(outCount)
            'arrOut(outCount) = arrStrText(i) ' & ": " & arrWordCount(i)
            arrOut(outCount) = arrStrText(i) 'arrOut(outCount) = arrWordCount(i)
'            arrOut(outCount, 0) = arrStrText(i) '
'            arrOut(outCount, 1) = arrWordCount(i)
            outCount = outCount + 1
        End If
    Next
Next mc
'convert output array to comma separated list
'WordCountDesc = Join(arrOut, "|")
WordCountDesc = arrOut

End Function

'All should be capital case except for dimensions
'Function to create a static copy of the final commercial void of formula
'Function to add description construction columns (major minor nouns, attributes, brand, item number, constructed descriptions, length) towards the right of the worksheet columns

Sub ClearFormat()
'Remove leading and trailing spaces
'Identify where there are double and triple spacings in all SCIT cleansed columns
'Identify non-standard characters (tab, carriage return…)
    Dim rng As Range
    Set rng = Application.Selection
    rng.Style = "Normal"

End Sub


Public Sub StaticCommercial()
'Function to create a static copy of the final commercial void of formula
'identify commercial sheet
'save copy of sheet
'remove formulas
'rename to source name + 'static'
'    Sheets("comm").Select
'    Sheets("comm").Copy Before:=Sheets(2)
'    Sheets("comm (2)").Select
'    Sheets("comm (2)").Name = "comm-static"

Dim shName As String
Dim newShName As String
Dim wsSrc As Worksheet
Dim wsDest As Worksheet
Dim rng As Range

Set wsSrc = ActiveSheet
    shName = wsSrc.Name
    newShName = Left(shName & " - Static", 31)
    wsSrc.Copy , wsSrc
    With ActiveSheet
        If Not sheetExists(newShName) Then
            .Name = newShName
        End If
        Set rng = .Range(Cells(1, 1), Cells.SpecialCells(xlCellTypeLastCell))
        rng.Copy
        rng.PasteSpecial xlPasteValues
        .Cells(1, 1).Select
    End With
    Application.CutCopyMode = False
    
End Sub



Sub addMM()
    addText ("mm")
End Sub

Sub addCM()
    addText ("cm")
End Sub

Sub addUserSpec()
Dim textToAdd As String
textToAdd = ""
textToAdd = InputBox("Text to Add", "Text", "")
    If textToAdd = "" Then
        Exit Sub
    Else
        addText textToAdd
    End If

End Sub

Sub addPrefix()
'same as addUserSpec but adds text to the start of the cell

Dim textToAdd As String
textToAdd = ""
textToAdd = InputBox("Text to Add", "Text", "")
    If textToAdd = "" Then
        Exit Sub
    Else
        addText textToAdd, True ' second arg is prefix boolean
    End If

End Sub

Private Sub unitConvert(unitType As Integer)
'1 --> mm to cm
'2 --> cm to m
'3 --> inch to mm
'4 --> inch to cm

Dim unitVale As Single
If ActiveWorkbook Is Nothing Then
    Exit Sub
End If

If Selection.Count > 1 Then
    SuspendAutoCalc
    'load values into an array
    For Each cell In Selection.SpecialCells(xlCellTypeVisible).Cells
    'get number only, keep decimal point
    'apply conversion
    unitValue = NumericFilter(cell.value, ".")
    If Len(NumericFilter(cell.value, ".")) > 0 Then
        unitValue = NumericFilter(cell.value, ".") + 0
        Select Case unitType
            Case 1: 'mm to cm
                cell.value = Round(unitValue / 10, 2) & "cm"
            Case 2: 'cm to m
                cell.value = Round(unitValue / 100, 2) & "m"
            Case 3: 'inch to mm
                cell.value = Round(unitValue * 25.4, 2) & "mm"
            Case 4: 'inch to cm
                cell.value = Round(unitValue * 2.54, 2) & "cm"
        End Select
    End If
    Next
    ResumeAutoCalc
Else
   
    unitValue = NumericFilter(ActiveCell.value, ".") + 0
        Select Case unitType
            Case 1: 'mm to cm
                ActiveCell.value = Round(unitValue / 10, 2) & "cm"
            Case 2: 'cm to m
                ActiveCell.value = Round(unitValue / 100, 2) & "m"
            Case 3: 'inch to mm
                ActiveCell.value = Round(unitValue * 25.4, 2) & "mm"
            Case 4: 'inch to cm
                ActiveCell.value = Round(unitValue * 2.54, 2) & "cm"
        End Select
End If


End Sub

Sub ConvertMM2CM()
    unitConvert (1)
End Sub

Sub ConvertCM2m()
    unitConvert (2)
End Sub

Sub ConvertInch2MM()
    unitConvert (3)
End Sub

Sub ConvertInch2CM()
    unitConvert (4)
End Sub

Sub Combine2cells()
    If ActiveWorkbook Is Nothing Then
    Exit Sub
End If

If Selection.Count > 1 Then
    SuspendAutoCalc
    For Each cell In Selection.SpecialCells(xlCellTypeVisible).Cells
        With cell
            .value = .value & .Offset(0, 1).value
            .Offset(0, 1).value = ""
        End With
    Next
    ResumeAutoCalc
Else
    With ActiveCell
        .value = .value & .Offset(0, 1).value
        .Offset(0, 1).value = ""
    End With
End If

End Sub

Sub measureX10()

Dim unitVale As Single
Dim unit As String
If ActiveWorkbook Is Nothing Then
    Exit Sub
End If

If Selection.Count > 1 Then
    SuspendAutoCalc
    'load values into an array
    For Each cell In Selection.SpecialCells(xlCellTypeVisible).Cells
        unitValue = NumericFilter(cell.value, ".") + 0
        unit = TextFilter(cell.value)
        cell.value = (unitValue * 10) & unit
    Next
    ResumeAutoCalc
Else
   
    unitValue = NumericFilter(ActiveCell.value, ".") + 0
    unit = TextFilter(ActiveCell.value)
    ActiveCell.value = (unitValue * 10) & unit
End If

End Sub


Private Sub addText(strAdditionalText As String, Optional prefix As Boolean)

If ActiveWorkbook Is Nothing Then
    Exit Sub
End If

If Selection.Count > 1 Then
    'cellToAdd = Selection.SpecialCells(xlCellTypeVisible).cells
    
    SuspendAutoCalc
    'load values into an array
    For Each cell In Selection.SpecialCells(xlCellTypeVisible)
    'For Each cell In cellToAdd
        If Len(cell.value) > 0 Then
            If Len(cell.value) = Len(Replace(cell.value, strAdditionalText, "")) Then
               If prefix Then
                cell.value = strAdditionalText & cell.value
               Else
                cell.value = cell.value & strAdditionalText
            End If
            End If
        'Else
            'cell.value = strAdditionalText
        End If
    Next
    ResumeAutoCalc
Else
'    cellToAdd = Selection.ActiveCell.Range
    ActiveCell.value = ActiveCell.value & strAdditionalText
End If



End Sub
Function tmark2(dsc As String) As String
 
dsc = Trim(dsc)
Dim x As Integer
Dim j As Integer
Dim i As Integer
Dim k As Integer
Dim tmark As String
Dim blocation As Integer
 
x = 0
Dim brandlogo(2) As String
Dim tmdescription() As String
 
brandlogo(0) = Chr(169)
 
brandlogo(1) = Chr(174)
 
brandlogo(2) = Chr(153)
 
ReDim tmdescription(UBound(brandlogo))
 
For j = LBound(brandlogo) To UBound(brandlogo)
 
    If InStr(1, dsc, brandlogo(j)) > 0 Then

        tmdescription(x) = brandlogo(j)
        x = x + 1
    End If
Next j
 
If x = 0 Then
 
    tmark = ""

Else
 
    ReDim Preserve tmdescription(x - 1)
    ReDim altbrands(x - 1)
    Dim spt() As String

    spt = Split(dsc, " ")
 
    For k = LBound(tmdescription) To UBound(tmdescription)
 
        For i = LBound(spt) To UBound(spt)

            If InStr(1, spt(i), tmdescription(k)) > 0 Then
                blocation = InStr(1, spt(i), tmdescription(k))

                If Mid(dsc, InStr(1, dsc, tmdescription(k)) - 1, 1) = " " Then

                    tmark = tmark & " " & spt(i - 1)
                Else
                    
                    tmark = tmark & " " & Mid(spt(i), 1, blocation - 1)

                End If
 
            End If
        Next i
     Next k
End If
tmark2 = Trim(tmark)
 
End Function




Sub worksheet_Rename()
Dim changes As Integer
Dim testCount As Integer
Dim shFound As Boolean

changes = 0
testCount = 0

    For Each cell In Range(Cells(1, 1), Cells(Cells.SpecialCells(xlCellTypeLastCell).Row, 1))
        If cell.value <> "" And Cells(cell.Row, 2).value <> "" And cell.value <> Cells(cell.Row, 2).value Then
            'rename the sheet from the name in first column to the name in the second column
            'search sheets for sheet name in cell
            shFound = False
            For Each sheet In Sheets
                If sheet.Name = cell.value Then
                    shFound = True
                    sheet.Name = Cells(cell.Row, 2).value
                    changes = changes + 1
                    Exit For
                End If
            Next
        End If
       testCount = testCount + 1
    Next

MsgBox "Cells checked: " & testCount & " sheet name changes: " & changes

End Sub

