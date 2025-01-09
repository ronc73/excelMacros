Attribute VB_Name = "WorksheetFunctions"
Option Compare Text
Option Base 1


Sub loadOptions()
 Application.MacroOptions _
        Macro:="HarshTrim", _
        Description:="The result will be the contents of the reference cell with all characters stripped out except 0-9, A-Z and a-z.", _
        Category:="Custom Worksheet Functions", _
        ArgumentDescriptions:=Array( _
            "inputStr is the first argument.  The string to be operated on, usually a cell reference", _
            "OtherChr is an optional argument.  Allows you to specify additional characters to keep")

Application.MacroOptions _
        Macro:="FuzzyCompare", _
        Description:="Uses HarshTrim to do a comparison between the contents of two cells ignoring any characters other than numbers and letters", _
        Category:="Custom Worksheet Functions", _
        ArgumentDescriptions:=Array( _
            "str1 is the first striing to compare. ", _
            "str1 is the comparison argument.")

Application.MacroOptions _
        Macro:="NumericFilter", _
        Description:="The result will be the contents of the reference cell with all characters stripped out except 0-9", _
        Category:="Custom Worksheet Functions", _
        ArgumentDescriptions:=Array( _
            "inputStr is the first argument.  The string to be operated on, usually a cell reference", _
            "OtherChr is an optional argument.  Allows you to specify additional characters to keep")
            
Application.MacroOptions _
        Macro:="TextFilter", _
        Description:="The result will be the contents of the reference cell with all characters stripped out except A-Z and a-z and any other characters listed.", _
        Category:="Custom Worksheet Functions", _
        ArgumentDescriptions:=Array( _
            "inputStr is the first argument.  The string to be operated on, usually a cell reference", _
            "OtherChr is an optional argument.  Allows you to specify additional characters to keep")
            
 Application.MacroOptions _
        Macro:="PriceFilter", _
        Description:="The result will be the contents of the reference cell with all characters stripped out except 0-9 and . to import prices without additional text.", _
        Category:="Custom Worksheet Functions", _
        ArgumentDescriptions:=Array( _
            "inputStr is the first argument.  The string to be operated on, usually a cell reference", _
            "OtherChr is an optional argument.  Allows you to specify additional characters to keep")
            
End Sub
Public Function HarshTrim(inputStr As String, Optional OtherChr As String) As String
Attribute HarshTrim.VB_Description = "The result will be the contents of the reference cell with all characters stripped out except 0-9, A-Z and a-z."
Attribute HarshTrim.VB_ProcData.VB_Invoke_Func = " \n19"
'   Created 11/09/2012
'   Ron Campbell for HPV

'   The result will be the contents of the reference cell with all characters stripped out except 0-9, A-Z and a-z.
'   Use optional arguement OtherChr to keep other characters. Enclose otherChr in double quotes

'   Usage:
'   In the worksheet use in syntax HarshTrim(<Cell reference>)
'   Example: harshTrim(A2)
'   or harshtrim(A2,"./ ][") to remove
' Application.MacroOptions _
'        Macro:="HarshTrim", _
'        Description:="The result will be the contents of the reference cell with all characters stripped out except 0-9, A-Z and a-z.", _
'        Category:="Custom Worksheet Functions", _
'        ArgumentDescriptions:=Array( _
'            "inputStr is the first argument.  The string to be operated on, usually a cell reference", _
'            "OtherChr is an optional argument.  Allows you to specify additional characters to keep")


Dim output As String
Dim i As Integer
Dim s() As Byte

s = StrConv(inputStr, vbFromUnicode)
output = ""
For i = 0 To UBound(s)
    If (s(i) >= 48 And s(i) <= 57) Or _
    (s(i) >= 65 And s(i) <= 90) Or _
    (s(i) >= 97 And s(i) <= 122) Or _
    (InStr(1, OtherChr, Chr(s(i)), vbTextCompare) > 0) Then
        output = output & Chr(s(i))
'    ElseIf Right(output, 1) <> " " Then
'        output = output & " "
    End If
Next
HarshTrim = output
End Function

Function FuzzyCompare(str1 As String, str2 As String) As Boolean
Attribute FuzzyCompare.VB_Description = "Uses HarshTrim to do a comparison between the contents of two cells ignoring any characters other than numbers and letters"
Attribute FuzzyCompare.VB_ProcData.VB_Invoke_Func = " \n19"
'   Created 18/09/2012
'   Ron Campbell for HPV

'   In the worksheet use in syntax FuzzyCompare(<Cell reference 1>,<Cell reference 2>)
'   Example: FuzzyCompare(A2,B2)
'   The result will be 0 if they match or +/- 1 if they are different
'   Uses HarshTrim to do a comparison between the contents of two cells ignoring any characters other than numbers and letters
' Application.MacroOptions _
'        Macro:="FuzzyCompare", _
'        Description:="Uses HarshTrim to do a comparison between the contents of two cells ignoring any characters other than numbers and letters", _
'        Category:="Custom Worksheet Functions", _
'        ArgumentDescriptions:=Array( _
'            "str1 is the first striing to compare. ", _
'            "str1 is the comparison argument.")


FuzzyCompare = (StrComp(HarshTrim(str1), HarshTrim(str2), vbTextCompare) = 0)
End Function


Function NumericFilter(strCode As String, Optional OtherChr As String) As String
Attribute NumericFilter.VB_Description = "The result will be the contents of the reference cell with all characters stripped out except 0-9"
Attribute NumericFilter.VB_ProcData.VB_Invoke_Func = " \n19"
'   Created 22/11/2012
'   Ron Campbell for HPV

'   The result will be the contents of the reference cell with all characters stripped out except 0-9
'   In the worksheet use in syntax NumericFilter(<Cell reference>)
'   Example: NumericFilter(A2)
'Application.MacroOptions _
'        Macro:="NumericFilter", _
'        Description:="The result will be the contents of the reference cell with all characters stripped out except 0-9", _
'        Category:="Custom Worksheet Functions", _
'        ArgumentDescriptions:=Array( _
'            "inputStr is the first argument.  The string to be operated on, usually a cell reference", _
'            "OtherChr is an optional argument.  Allows you to specify additional characters to keep")

Dim output As String
Dim i As Integer
Dim s() As Byte

s = StrConv(strCode, vbFromUnicode)
output = ""
For i = 0 To UBound(s)
    If (s(i) >= 48 And s(i) <= 57 Or InStr(1, OtherChr, Chr(s(i)), vbTextCompare) > 0) Then
        output = output & Chr(s(i))
    End If
Next
If output = "/" Then
    output = ""
End If

NumericFilter = output
End Function

Function TextFilter(strCode As String, Optional OtherChr As String) As String
Attribute TextFilter.VB_Description = "The result will be the contents of the reference cell with all characters stripped out except A-Z and a-z and any other characters listed."
Attribute TextFilter.VB_ProcData.VB_Invoke_Func = " \n19"
'   Created 19/04/2016
'   Ron Campbell for HPV

'   The result will be the contents of the reference cell with all characters stripped out except A-z

'   Usage: Copy this code/module into the worksheet you wish to use the function in
'   In the worksheet use in syntax NumericFilter(<Cell reference>)
'   Example: NumericFilter(A2)

' Application.MacroOptions _
'        Macro:="TextFilter", _
'        Description:="The result will be the contents of the reference cell with all characters stripped out except A-Z and a-z and any other characters listed.", _
'        Category:="Custom Worksheet Functions", _
'        ArgumentDescriptions:=Array( _
'            "inputStr is the first argument.  The string to be operated on, usually a cell reference", _
'            "OtherChr is an optional argument.  Allows you to specify additional characters to keep")

Dim output As String
Dim i As Integer
Dim s() As Byte

s = StrConv(strCode, vbFromUnicode)
output = ""
For i = 0 To UBound(s)
    If (s(i) >= Asc("A") And s(i) <= Asc("z") Or InStr(1, OtherChr, Chr(s(i)), vbTextCompare) > 0) Then
        output = output & Chr(s(i))
    End If
Next
If output = "/" Then
    output = ""
End If

TextFilter = Trim(output)
End Function

Function PriceFilter(strCode As String) As String
Attribute PriceFilter.VB_Description = "The result will be the contents of the reference cell with all characters stripped out except 0-9 and . to import prices without additional text."
Attribute PriceFilter.VB_ProcData.VB_Invoke_Func = " \n19"
'   Created 26/03/2013
'   Ron Campbell for HPV

'   The result will be the contents of the reference cell with all characters stripped out except 0-9 and "." to import prices without additional text

'   Usage:
'   In the worksheet use in syntax NumericFilter(<Cell reference>)
'   Example: NumericFilter(A2)

' Application.MacroOptions _
'        Macro:="PriceFilter", _
'        Description:="The result will be the contents of the reference cell with all characters stripped out except 0-9 and "." to import prices without additional text.", _
'        Category:="Custom Worksheet Functions", _
'        ArgumentDescriptions:=Array( _
'            "inputStr is the first argument.  The string to be operated on, usually a cell reference", _
'            "OtherChr is an optional argument.  Allows you to specify additional characters to keep")

Dim output As String
Dim i As Integer
Dim s() As Byte

s = StrConv(strCode, vbFromUnicode)
output = ""
For i = 0 To UBound(s)
    If ((s(i) >= 48 And s(i) <= 57) Or s(i) = 46) Then
        output = output & Chr(s(i))
    End If
Next
PriceFilter = output
End Function


Function IsStrikeThrough(cellRef As Range) As Boolean
Application.Volatile
IsStrikeThrough = cellRef.Font.Strikethrough

End Function


Function getFilename() As String
    getFilename = ThisWorkbook.FullName
End Function

Function Quarter(refDate As Date) As Integer
'return the calendar quarter of a year
    Quarter = Fix(((Month(refDate) - 1) / 3)) + 1
End Function

Function QuarterYear(refDate As Date) As String
'return the calendar quarter and year
If refDate > 0 Then
    QuarterYear = Year(refDate) & "-" & "Q" & Quarter(refDate)
Else
    QuarterYear = ""
End If
End Function

Function FinYear(refDate As Date, Optional firstMonth As Integer) As String
'Return the financial year in the form YYYY/YYYY for a given date
Dim outYear As Integer

If firstMonth = 0 Then
    firstMonth = 7
End If

outYear = Year(refDate)
    If Month(refDate) >= firstMonth Then
        FinYear = outYear & "/" & outYear + 1
    Else
        FinYear = outYear - 1 & "/" & outYear
    End If
    
End Function
Function GLNCheck(gln As String) As String
Dim i As Integer
Dim sumEven As Integer
Dim sumOdd As Integer
Dim sumAll As Integer
Dim GLNCheckDigit As Integer

    If Len(gln) = 13 And IsNumeric(gln) Then
        sumEven = 0
        sumOdd = 0
        For i = 2 To 12 Step 2
            sumEven = sumEven + Mid(gln, i, 1)
        Next i
            sumEven = 3 * sumEven
        For i = 1 To 11 Step 2
            sumOdd = sumOdd + Mid(gln, i, 1)
        Next i
        sumAll = sumEven + sumOdd
        If sumAll Mod 10 = 0 Then
            GLNCheckDigit = 0
        Else
            GLNCheckDigit = 10 - sumAll Mod 10
        End If
        If GLNCheckDigit = Right(gln, 1) Then
            GLNCheck = "Valid GLN"
        Else
            GLNCheck = "Invalid GLN. Check digit should be : " & GLNCheckDigit
        End If
    Else
        GLNCheck = "Not a GLN: Must be 13 digits"
    End If
        
End Function

Function indexMatch(val As Variant, searchRange As Range, ReturnRange As Range, Optional tol As Double) As Variant
' look for a value within a range and return the contents of the corresponding cell in the second range
' tol is tolerence. standard +/- 0.01 to allow for rounding of price

Dim cellCount As Integer
Dim searchCell As Integer
Dim found As Boolean

cellCount = 0
For Each cell In searchRange
    If IsNumeric(val) Then
        If IsNumeric(cell.Value2) Then
            If Abs(cell.value - val) <= tol Then
                'found matching value
                found = True
            End If
        End If
    Else
        If cell.Value2 = val Then
            found = True
        End If
    End If
    If Not found Then
        cellCount = cellCount + 1
    Else
        Exit For
    End If
Next cell

If found Then
    searchCell = 0
    For Each cell In ReturnRange
        If searchCell = cellCount Then
            indexMatch = cell.value
            Exit For
        Else
            searchCell = searchCell + 1
        End If
    Next cell
Else
    For Each cell In searchRange
        indexMatch = Format(cell.value, "$#.00")
        Exit For
    Next cell
End If
End Function

Function ExtractTextBetween(strIn As String, ch As String, StartChInstance As Integer, EndChInstance As Integer) As String
'Loop through input string looking for character ch
'output is made up of the input string between the start instance of ch and the end instance of ch
Dim i As Integer
Dim chCounter
Dim strOut As String
Dim continue As Boolean

strIn = strIn & ch
On Error GoTo errExtractTextBetween
If Len(strIn) > 0 Then
    i = 1
    strOut = ""
    chCounter = 0
    continue = True
    
    Do
        If chCounter >= StartChInstance And chCounter <= EndChInstance Then
            strOut = strOut & Mid(strIn, i, 1)
        End If
        If InStr(1, ch, Mid(strIn, i, 1), vbTextCompare) > 0 Then
            chCounter = chCounter + 1
        End If
        If chCounter >= EndChInstance Then
            continue = False
        End If
        i = i + 1
    Loop While continue Or i < Len(strIn)
    
    ExtractTextBetween = Left(strOut, Len(strOut) - 1)
End If
Exit Function

errExtractTextBetween:
    ExtractTextBetween = ""

End Function


Function PivotDataString() As String
'use to create a string to use as a dynamic named range for use with pivot tables
'=OFFSET('Raw GL Data'!$A$1,0,0,COUNTA('Raw GL Data'!$A:$A),COUNTA('Raw GL Data'!$1:$1))
PivotDataString = "=OFFSET('" & Application.ActiveSheet.Name & "'!$A$1,0,0,COUNTA('" & Application.ActiveSheet.Name & "'!$A:$A),COUNTA('" & Application.ActiveSheet.Name & "'!$1:$1))"

End Function

Function UNSPSC_GetLevel(code As String) As Integer
'Return the UNSPSC level of a given UNSPSC code
'
If Len(code) = 8 Then
    If Right(code, 6) = "000000" Then
        UNSPSC_GetLevel = 1
    ElseIf Right(code, 4) = "0000" Then
        UNSPSC_GetLevel = 2
    ElseIf Right(code, 2) = "00" Then
        UNSPSC_GetLevel = 3
    Else
        UNSPSC_GetLevel = 4
    End If
Else
    UNSPSC_GetLevel = 0
    Exit Function
End If
End Function

Function UNSPSC_SetLevel(code As String, level As Integer) As String
UNSPSC_SetLevel = Left(code, level * 2) & String(2 * (4 - level), "0")
End Function

Function stringtest(anum As Integer) As String
stringtest = String(anum, "X")
End Function

Function ConcatRange(rng As Range, sep As String, all As Boolean) As String

Dim output As String
Dim firstCell As Boolean
output = ""
firstCell = True

For Each curcell In rng
    If firstCell Or (Len(output) = 0 And Not all) Then
        output = curcell
    Else
        If Len(curcell) > 0 Or all Then
            output = output & sep & curcell
        End If
    End If
    firstCell = False
Next

ConcatRange = output

End Function
Function SplitStr(str As String, sep As String, elementNo As Integer) As String
Dim output As String
Dim elements() As String
output = ""
elementNo = elementNo - 1
    elements() = Split(str, sep)
    If elementNo <= UBound(elements) Then
        output = elements(elementNo)
    End If
SplitStr = output

End Function

Function PatternFind(strText As String, strExclList As String) As String
Dim arrExclList() As String
Dim strOut As String
Dim wordStart As Integer
Dim arrOut() As String
Dim charFound As String ' placeholder character
Dim keepChar As Boolean ' charcter to keep before or after found word
Dim foundCount As Integer ' count of excl list words found in strtext

On Error GoTo PatternFind_Err
If strExclList <> "" Then
    arrExclList = Split(strExclList, ",")
    ReDim Preserve arrExclList(UBound(arrExclList))
Else
    Exit Function
End If
PatternFind = ""
strOut = UCase(strText)
foundCount = 0

For j = 0 To UBound(arrExclList)
    'look for number with word
    wordStart = InStr(1, strText, UCase(arrExclList(j)), vbTextCompare)
    'count numbers before word start or after wordstart + word length
    'before - read chars before word until non space and non numeric
    If wordStart > 0 Then
        foundCount = foundCount + 1
        keepChar = True
        ReDim Preserve arrOut(foundCount)
        arrOut(foundCount) = arrExclList(j)
        'Keep numbers before word
        i = 1
        If wordStart > 1 Then
            Do
                charFound = Mid(strText, wordStart - i, 1)
                keepChar = (Asc(charFound) >= 48 And Asc(charFound) <= 57) Or Asc(charFound) = 32
                If keepChar Then
                    ReDim Preserve arrOut(foundCount)
                    arrOut(foundCount) = charFound & arrOut(foundCount)
                End If 'keepChar
                i = i + 1
            Loop While keepChar And (wordStart - i) >= 1
            arrOut(foundCount) = Trim(arrOut(foundCount))
        End If 'wordStart > 1
        'Keep numbers after word
        i = 0
        If (wordStart + Len(arrExclList(j))) < Len(strText) Then
            Do
                charFound = Mid(strText, wordStart + Len(arrExclList(j)) + i, 1)
                keepChar = (Asc(charFound) >= 48 And Asc(charFound) <= 57) Or Asc(charFound) = 32
                If keepChar Then
                    ReDim Preserve arrOut(foundCount)
                    arrOut(foundCount) = arrOut(foundCount) & charFound
                End If 'keepChar
                i = i + 1
            Loop While keepChar And (i + wordStart + Len(arrExclList(j))) <= Len(strText)
            arrOut(foundCount) = Trim(arrOut(foundCount))
        End If ' (wordStart + Len(arrExclList(j)))< len(strText)
    End If 'wordStart >0
Next

If foundCount > 0 Then
    PatternFind = Join(arrOut, ",")
    PatternFind = Right(PatternFind, Len(PatternFind) - 1)
Else
    PatternFind = ""
End If

Exit Function

PatternFind_Err:
PatternFind = "Err" & Err.Number & " " & Err.Description

End Function
Function ReplaceMultiple(strText As String, strExclList As String) As String
Dim arrExclList() As String
Dim strOut As String

If strExclList <> "" Then
    arrExclList = Split(strExclList, ",")
    ReDim Preserve arrExclList(UBound(arrExclList))
Else
    Exit Function
End If

strOut = UCase(strText)
For j = 0 To UBound(arrExclList)
    strOut = Replace(strOut, UCase(arrExclList(j)), "", , , vbTextCompare)
Next

ReplaceMultiple = strOut

End Function

Function ExcludeWords(strText As String, Optional strExclList As String, Optional strExclList2 As String, Optional strExclList3 As String, Optional strExclList4 As String, Optional removeNumeric As Boolean) As String
'Read list of words from exclusion list into array
'read text into array, broken by words
'remove any words from text array where they match with exclusion list
Dim arrExclList() As String
Dim arrStrText() As String
Dim arrOut() As String

arrStrText = Split(Trim(strText), ",")
ReDim Preserve arrStrText(UBound(arrStrText))

If strExclList <> "" Then
    arrExclList = Split(strExclList & strExclList2 & strExclList3 & strExclList4, ",")
    ReDim Preserve arrExclList(UBound(arrExclList))
End If


'Remove Duplicate from input string
'For i = 0 To UBound(arrStrText)
'    For j = i To UBound(arrStrText)
'        If i <> j And arrStrText(i) = arrStrText(j) Then
'            arrStrText(j) = ""
'        End If
'    Next
'Next

'remove number and words with 1 or 2 letters then numbers
'For i = 0 To UBound(arrStrText)
'    If arrStrText(i) <> "" Then
'        If IsNumeric(arrStrText(i)) Then
'            arrStrText(i) = ""
'        End If
'        For j = 1 To Len(arrStrText(i)) 'IIf(Len(arrStrText(i)) <= 3, Len(arrStrText(i)) - 1, 3)
'            If IsNumeric(Mid(arrStrText(i), j, 1)) Then
'                arrStrText(i) = ""
'                Exit For
'            End If
'        Next
'    End If
'Next

'remove words from the exclusion list
If strExclList <> "" Then
    For i = 0 To UBound(arrStrText)
        For j = 0 To UBound(arrExclList)
            If arrStrText(i) = arrExclList(j) Then
                arrStrText(i) = ""
                Exit For
            End If
        Next
    Next
End If

'make output arrary with remaining keywords only
outCount = 0
For i = 0 To UBound(arrStrText)
    If arrStrText(i) <> "" Then
        ReDim Preserve arrOut(outCount)
        arrOut(outCount) = arrStrText(i)
        outCount = outCount + 1
    End If
Next


'convert output array to comma separated list
ExcludeWords = Join(arrOut, ",")

End Function

Function KeywordGenerator(str As String) As String
Dim output As String
Dim elements() As String
Dim elementNo As Integer
Dim sep As String

output = ""
sep = " "
    elements() = Split(str, sep)
elementNo = 0
While elementNo <= UBound(elements)
        If Len(Trim(elements(elementNo))) > 0 Then
            output = output & elements(elementNo) & ","
        End If
    elementNo = elementNo + 1
Wend

KeywordGenerator = output

End Function

Function KeywordMatch(strSrcWords As String, rngCompare As Range, strMajorNouns, UNSPSCoffset As Integer) As String
'take a set of keywords from the cell and compare it against the keywords in each cell of the comparison range
'return the UNSPSC code with word matches
'return the number of matches found
'return the number of words in each set
'08/08/17 -
'Add Major Noun List. Major nouns score more highly as a word match. List provided by MH

Dim arrSrcWords() As String
Dim arrCompare() As String ' store the UNSPSC keywords in an array for comparison
Dim arrMajorNoun() As String
Dim matchScore As Integer
Dim matchUNSPSC As String
Dim matchWords As String
Dim matchWordCount As Integer
Dim output As String
Dim ResultCount As Integer

Dim ResultUNSPSC() As String
Dim ResultUNSPSCWords() As Integer
Dim ResultScore() As Integer
Dim ResultWordMatchCount() As Integer
Dim ResultWordMatch() As String
Dim ResultMajorNoun() As String

Dim majorNounScore As Integer
Dim majorNounFactor As Single 'value placed on major noun starting point 2.0
Dim foundWord As String
Dim MaxScore As Integer

'Dim runTimerStart As Double
'Dim runTimerEnd As Double
arrSrcWords = Split(strSrcWords, ",")
'ReDim Preserve arrSrcWords(UBound(arrSrcWords))

arrMajorNoun = Split(strMajorNouns, ",")
'ReDim Preserve arrMajorNoun(UBound(arrMajorNoun))

MaxScore = 0
output = ""
ResultCount = 0
majorNounFactor = 2

ReDim Preserve ResultScore(ResultCount)
ReDim Preserve ResultUNSPSCWords(ResultCount)
ReDim Preserve ResultUNSPSC(ResultCount)
ReDim Preserve ResultWordMatch(ResultCount)
ReDim Preserve ResultWordMatchCount(ResultCount)
ReDim Preserve ResultMajorNoun(ResultCount)
    
For Each cell In rngCompare
'runTimerStart = Now()
matchWords = ""
matchScore = 0
majorNounScore = 0
matchWordCount = 0
arrCompare = Split(cell.value, ",")
ReDim Preserve arrCompare(UBound(arrCompare))

    'for each UNSPSC code count how many words are in the code description and in the invoice description
        For i = 0 To UBound(arrCompare)
            For j = 0 To UBound(arrSrcWords)
                If UCase(arrSrcWords(j)) = UCase(arrCompare(i)) Then
                    foundWord = arrCompare(i)
                    matchWords = matchWords & foundWord & ","
                    'is this word in the major nouns list
                        For k = 0 To UBound(arrMajorNoun)
                            If UCase(arrMajorNoun(k)) = UCase(foundWord) Then
                                majorNounScore = majorNounScore + 1
                                Exit For
                            End If
                        Next k
                    matchWordCount = matchWordCount + 1
                End If
            Next
        Next
    matchScore = (matchWordCount - majorNounScore) + majorNounScore * majorNounFactor
        If matchScore >= 1 Then
            'track highest match count for sorting
            If matchScore > MaxScore Then
               MaxScore = matchScore
            End If
                matchUNSPSC = cell.Offset(0, UNSPSCoffset).value
                'matchUNSPSC = Cells(Cell.Row, 1).Value
                ResultUNSPSC(ResultCount) = matchUNSPSC
                ResultScore(ResultCount) = matchScore
                ResultUNSPSCWords(ResultCount) = UBound(arrCompare)
                ResultWordMatch(ResultCount) = matchWords
                ResultWordMatchCount(ResultCount) = matchWordCount
                ResultMajorNoun(ResultCount) = majorNounScore
                ResultCount = ResultCount + 1
            If ResultCount >= 1 Then
                ReDim Preserve ResultScore(ResultCount)
                ReDim Preserve ResultUNSPSCWords(ResultCount)
                ReDim Preserve ResultUNSPSC(ResultCount)
                ReDim Preserve ResultWordMatch(ResultCount)
                ReDim Preserve ResultWordMatchCount(ResultCount)
                ReDim Preserve ResultMajorNoun(ResultCount)
            End If
        End If
Next

'form output in order of most matches to least matches
For k = MaxScore To 1 Step -1
    For m = 0 To ResultCount
        If ResultScore(m) = k Then
            output = output & ResultUNSPSC(m) & "|" & ResultScore(m) & "|" & ResultWordMatchCount(m) & "|" & ResultUNSPSCWords(m) + 1 & "|" & UBound(arrSrcWords) + 1 & "|" & ResultMajorNoun(m) & "|" & ResultWordMatch(m) & Chr(13) & Chr(10)
        End If
    Next
Next

If output <> "" Then
    output = "UNSPSC | Score |  Matched Count | UNSPSC Count | Description Count | Major Nouns | Matched Words" & Chr(13) & Chr(10) & output
Else
    output = "No Matches"
End If
'runTimerEnd = Now()
'output = "Runtime: " & (runTimerEnd - runTimerStart) & Chr(13) & Chr(10) & output
KeywordMatch = Left(output, 32000)

End Function

Function KeywordMatchCell(strSrcWords As String, strSrcWords2 As String) As String
'take a set of keywords from the cell and compare it against the keywords in another cell
'return the number of matches found
'Return matched words

Dim arrSrcWords() As String
Dim arrSrcWords2() As String
Dim matchScore As Integer
Dim matchWords As String
Dim matchWordCount As Integer
Dim output As String
Dim ResultCount As Integer


Dim foundWord As String

'Dim runTimerStart As Double
'Dim runTimerEnd As Double
arrSrcWords = Split(strSrcWords, ",")
arrSrcWords2 = Split(strSrcWords2, ",")


output = ""
ResultCount = 0

   
'runTimerStart = Now()
matchWords = ""
matchScore = 0
majorNounScore = 0
matchWordCount = 0

        For i = 0 To UBound(arrSrcWords)
            For j = 0 To UBound(arrSrcWords2)
                If UCase(arrSrcWords2(j)) = UCase(arrSrcWords(i)) And Len(UCase(arrSrcWords(i))) > 0 Then
                    foundWord = arrSrcWords(i)
                    matchWords = matchWords & foundWord & ","
                    matchWordCount = matchWordCount + 1
                End If
            Next
        Next


output = matchWordCount & ": " & matchWords

'runTimerEnd = Now()
'output = "Runtime: " & (runTimerEnd - runTimerStart) & Chr(13) & Chr(10) & output
KeywordMatchCell = Left(output, 32000)

End Function


Function ascii_codes(str As String) As String
'Dim str As String
Dim aa As String


' Stop
 
For i = 1 To Len(str)
    aa = aa & Asc(Mid(str, i, 1)) & ", "
Next
ascii_codes = aa
End Function

Function ascii_codes_Unique(str As String, Optional intCharMin As Integer, Optional intCharmax As Integer)

Dim currChar As Integer

If intCharMin = 0 Then
    intCharMin = 1
End If
If intCharmax = 0 Then
    intCharmax = 255
End If
aa = ""
For i = 1 To Len(str)
    currChar = Asc(Mid(str, i, 1))
    If currChar >= intCharMin And currChar <= intCharmax Then
        If InStr(1, aa, currChar, vbTextCompare) = 0 Then
            aa = aa & currChar & ", "
        End If
    End If
Next
ascii_codes_Unique = aa
End Function

Function ASCIICode(Character As String) As Long
    ASCIICode = AscW(Character)
End Function

Function ascii_codesw(str As String) As String
'Dim str As String
Dim aa As String


' Stop
 
For i = 1 To Len(str)
    aa = aa & AscW(Mid(str, i, 1)) & ", "
Next
ascii_codesw = aa
End Function

Function getSheetName(Optional ws As Range) As String
If ws Is Nothing Then
    'no reference given
    getSheetName = ActiveSheet.Name
Else
    getSheetName = ws.Worksheet.Name
End If

If InStr(1, getSheetName, " ") Then
    getSheetName = "'" & getSheetName & "'"
End If

End Function

Function getFileFromPath(source As String) As String
'extract the filename from a file and path
'created 14/12/16

 getFileFromPath = Trim(Right(Replace(source, "\", Application.WorksheetFunction.Rept(" ", Len(source))), Len(source)))
    
End Function


Function multiReplace(SrcRng As Range, replaceRange As Range) As String
'does series of search and replace functions. can give undesirable results is search word is part of another word.

Dim output As String
If SrcRng.Count > 1 Then
    Exit Function
End If
    
'If Left(SrcRng.Value, 1) <> "," Then
'    output = "," & SrcRng.Value
'Else
    output = " " & SrcRng.value & " "
'End If
    
'For Each cell In SrcRng
    For Each replacecell In replaceRange
        'replace/substitute cell.value find replaceCell.value replace with offset(replaceCell(0,1))
        output = Replace(output, replacecell.value, "" & replacecell.Offset(0, 1).value & "", , , vbTextCompare)
    Next replacecell
'Next cell
'multiReplace = Mid(output, 2, Len(output))
multiReplace = Trim(output)
End Function


Function multiReplace_WW(SrcRng As Range, replaceRange As Range, Optional repLimit As Integer) As String
'whole word search and replace. does replace per word in a description
'made to get around issue with multireplace where parts of words were being replaced

Dim output As String
Dim outputArr() As String
Dim repCount As Long
Dim replaceList() As String
Dim i, x, maxX As Long
Dim newValue As String

On Error GoTo mr_err

If SrcRng.Count > 1 Then
    Exit Function
End If
    
'If Left(SrcRng.Value, 1) <> "," Then
'    output = "," & SrcRng.Value
'Else
    output = SrcRng.value
'End If

x = 1
ReDim replaceList(replaceRange.Count, 2)
For Each c In replaceRange
    newValue = Trim(c.Offset(0, 1).value)
'    ReDim Preserve replaceList(x , 2)
    If Trim(c.value) <> newValue And newValue <> "" Then
        replaceList(x, 1) = c.value
        replaceList(x, 2) = newValue
        x = x + 1
    End If
Next c
maxX = x
'ReDim Preserve replaceList(maxX - 1, 2) As String


outputArr = Split(output, " ")
repCount = 0

For i = 0 To UBound(outputArr)
'    For Each replacecell In replaceRange
        'replace/substitute cell.value find replaceCell.value replace with offset(replaceCell(0,1))
'        If UCase(outputArr(i)) = UCase(replacecell.value) Then
'            outputArr(i) = replacecell.offset(0, 1).value
'            repCount = repCount + 1
'        End If
    For x = 1 To maxX ' UBound(replaceList)
        If UCase(outputArr(i)) = UCase(replaceList(x, 1)) And replaceList(x, 1) <> "" And replaceList(x, 2) <> "" Then
            outputArr(i) = replaceList(x, 2)
            repCount = repCount + 1
            If repLimit > 0 Then
                If repCount >= repLimit Then
                    GoTo doneNow
                End If
            End If
            GoTo next_i
        End If
        
    'Next replacecell
    Next x
next_i:
Next i
doneNow:

multiReplace_WW = Join(outputArr(), " ")
Exit Function

mr_err:
    MsgBox "MultiReplace Error: " & Err.Number & ":" & Err.Description
End Function


Function gtin_cal(strIn As String) As String
'calculate the check digit for a GTIN
'check digit replaces last digit of given GTIN

Dim n
Dim n1, n2, n3, n4, n5, n6, n7, n8, n9, n10, n11, n12, n13
Dim normal As String

If Len(strIn) = 14 Then
    normal = Left(strIn, 13)
ElseIf Len(strIn) = 13 Then
    normal = Left("0" & strIn, 13)
End If

n1 = CInt(Mid(normal, 1, 1))
n2 = CInt(Mid(normal, 2, 1))
n3 = CInt(Mid(normal, 3, 1))
n4 = CInt(Mid(normal, 4, 1))
n5 = CInt(Mid(normal, 5, 1))
n6 = CInt(Mid(normal, 6, 1))
n7 = CInt(Mid(normal, 7, 1))
n8 = CInt(Mid(normal, 8, 1))
n9 = CInt(Mid(normal, 9, 1))
n10 = CInt(Mid(normal, 10, 1))
n11 = CInt(Mid(normal, 11, 1))
n12 = CInt(Mid(normal, 12, 1))
n13 = CInt(Mid(normal, 13, 1))

n = (n1 + n3 + n5 + n7 + n9 + n11 + n13) * 3 + (n2 + n4 + n6 + n8 + n10 + n12) * 1

q = Int(n / 10)

r = n Mod 10
If (r <> 0) Then
append_value = (10 * (q + 1)) - n
Else
append_value = 0
End If
'gtin_cal = "0" & normal & CStr(append_value)
gtin_cal = normal & CStr(append_value)

End Function

Function gtin_cal_rt(strIn As String) As String
'calculate the check digit for a GTIN
'check digit replaces last digit of given GTIN
'same as gtin_cal but uses running total rather than discrete variables

Dim n
Dim normal As String

strIn = NumericFilter(strIn)

If Len(strIn) >= 14 Then
    normal = Left(strIn, 13)
ElseIf Len(strIn) = 13 Then
    normal = Left("0" & strIn, 13)
End If

'n = (n1 + n3 + n5 + n7 + n9 + n11 + n13) * 3 + (n2 + n4 + n6 + n8 + n10 + n12) * 1
n = 0
For i = 1 To 13
    DigitValue = Mid(normal, i, 1)
        If WorksheetFunction.IsOdd(i) Then
            n = n + DigitValue * 3
        Else
            n = n + DigitValue
        End If
Next i

q = Int(n / 10)
r = n Mod 10
If (r <> 0) Then
    append_value = (10 * (q + 1)) - n
Else
    append_value = 0
End If

gtin_cal_rt = normal & CStr(append_value)

End Function

Function GTIN_Valid(gtin As String) As String
'take given string
'is string length 14?
'is value given numeric
'compare check digit of provided number with calculated check digit
Dim error As String
Dim output As String

error = ""

If Len(gtin) = 14 Then
    If NumericFilter(gtin) = gtin Then
        If gtin = gtin_cal_rt(gtin) Then
            error = "OK"
        Else
            error = "Check digit fail"
        End If
    Else
        error = "Non numeric"
    End If
Else
    error = "Length <> 14"
End If

GTIN_Valid = error
End Function

Function getMaxCell() As Long
'find the last used row in the active sheet
'as opposed to lastcell from excel def that may be from previously use now empty cells

Dim maxRow As Long
Dim rowNo As Long
maxRow = 1

For Each c In Range("a1", ActiveSheet.Cells.Cells(1, ActiveSheet.Cells.SpecialCells(xlLastCell).Column)).Cells
    rowNo = ActiveSheet.Cells(ActiveSheet.Rows.Count, c.Column).End(xlUp).Row
    If maxRow < rowNo Then
        maxRow = rowNo
    End If
    'MsgBox "Column: " & c.Column & " Last Row: " & rowNo & " Max Row: " & maxRow
Next c
getMaxCell = maxRow
'getMaxCell = ActiveSheet.Cells.SpecialCells(xlLastCell).Row
'getMaxCell = ActiveCell.SpecialCells(xlLastCell).Column
End Function

Function calc(strIn As String) As String
'read the value of a cell and evaluate it if it's a numeric expression. replace the contents with the result

On Error GoTo calc_err
calc = Evaluate(strIn)
Exit Function
calc_err:
calc = strIn
End Function

Function WordCount(strText As String, Optional minOccurance As Integer, Optional strExclList As Range, Optional removeNumeric As Boolean) As String
'Function WordCount(strText As String, Optional strExclList As String, Optional removeNumeric As Boolean) As String
'11/09/2019
'
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
'    minOccurance = 1
'End If

arrStrText = Split(Trim(Replace(strText, " ", ",")), ",")
ReDim Preserve arrStrText(UBound(arrStrText))
ReDim Preserve arrWordCount(UBound(arrStrText))

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
            arrOut(outCount) = arrStrText(i) & ": " & arrWordCount(i)
            outCount = outCount + 1
        End If
    Next
Next mc

'convert output array to comma separated list
WordCount = Join(arrOut, ",")

End Function

Function fixMeasure(strIn As String) As String
Dim output As String
Dim unit As String
Dim strValue As String
Dim value As Single

output = strIn
strValue = NumericFilter(strIn, "./-")
unit = LCase(TextFilter(strIn, """"))
On Error GoTo exitNow
'does string contain a number
If strValue <> "" And unit <> "" Then
    'value = val(strValue)
    value = strValue + 0
    'is numeric then find units
    unit = LCase(TextFilter(strIn, """"))
    'is the unit metric
    Select Case unit
        Case "cm":
            'is value >100 should the unit be converted
            If value >= 100 Then 'conver to m
                value = value / 100
                unit = "m"
            Else
                unit = "cm"
            End If
        Case "mm":
            If value >= 1000 Then 'conver to m
                value = value / 1000
                unit = "m"
            ElseIf value >= 100 Then 'conver to cm
                value = value / 10
                unit = "cm"
            End If
        Case "m", "mtr", "meter", "metre", "meters", "metres":
                unit = "m"
        Case "g", "gr", "gms", "gram", "gm":
            If value >= 1000 Then 'conver to m
                value = value / 1000
                unit = "kg"
            Else
                unit = "g"
            End If
        Case "kg", "kgr", "kgms", "kgram", "kgm", "kilo", "kgs":
                unit = "kg"
        Case "ml", "mls", "cc":
            If value > 1000 Then 'conver to m
                value = value / 1000
                unit = "L"
            Else
                unit = "mL"
            End If
        Case "l", "ltr", "liter", "lter", "lts", "lt", "litre", "litres", "liters":
            If value < 1 Then
                value = value * 1000
                unit = "mL"
            Else
                unit = "L"
            End If
        Case "um", "micron":
                unit = "mcm"
        Case "ul":
                unit = "mcL"
        Case "fg", "fr", "f", "french":
                unit = "FR"
        Case "inch", "inches", "in", """", "'", Chr(34), Chr(34) & Chr(34):
            value = value * 25.4 'convert inch to mm
            If value > 1000 Then 'convert to m
                value = value / 1000
                unit = "m"
            ElseIf value >= 100 Then 'convert to cm
                value = value / 10
                unit = "cm"
            Else
                unit = "mm"
            End If
        Case Else:
            GoTo exitNow
        End Select
            output = Round(value, 2) & unit
End If

exitNow:
fixMeasure = output
End Function



Function getDimension(strIn As String, strDiv As String, intLen As Integer) As String

Dim strStart As Integer
Dim strEnd As Integer
Dim strLen As Integer


If WorksheetFunction.Find(strDiv, strIn) > intLen Then
    strStart = WorksheetFunction.Find(" ", strIn, WorksheetFunction.Max(1, WorksheetFunction.Find(strDiv, strIn) - intLen))
Else: strStart = 1
End If

On Error GoTo strEndError
strEnd = WorksheetFunction.Find(" ", strIn, WorksheetFunction.Find(strDiv, strIn) + Len(strDiv))
GoTo Finish

strEndError:
strEnd = Len(strIn) + 1

Finish:
strLen = strEnd - strStart

getDimension = Trim(Mid(strIn, strStart, strLen))

End Function

'Function getNextVisibleValue(addr As Address, rowCol As String, offset As Integer)
'Dim rng As Range
'
'With ActiveWorkbook
'    If rowCol = "Row" Then
'    'send up or down arrow offset number of times
'    ElseIf rowCol = "Col" Then
'    'send left or right offset number of times
'    End If
'End With
'
'End Function

Function multiRemove(strInput As String, replaceRange As Range, Optional replaceRange2 As Range) As String
'removes words from a cell where they are found in a given replaceRange

Dim output As String
    
    output = " " & strInput & " "
    If replaceRange2 Is Nothing Then
    Else
       Set replaceRange = Union(replaceRange, replaceRange2)
    End If
    
    For Each replacecell In replaceRange
        'replace/substitute cell.value find replaceCell.value replace with offset(replaceCell(0,1))
        If replacecell.value <> "" Then
            output = Replace(output, replacecell.value & " ", "", , , vbTextCompare)
        End If
    Next replacecell

'multiReplace = Mid(output, 2, Len(output))
multiRemove = Trim(output)
End Function


Function FilterCriteria(rng As Range) As String
'Failed attempt at applying a filter
'too many variations based on the value/value2 etc of a cell depending on the format of the cell

Dim FCol As Integer
    Dim FDate As Boolean
    Dim Filter As String
    Dim x As Integer
    Dim Filter2 As String
    FCol = 0
    FDate = False
    Filter = ""
    x = 0
    Filter2 = ""
    On Error GoTo Finish
    With rng.Parent.AutoFilter
        If Intersect(rng, .Range) Is Nothing Then GoTo Finish
        FCol = rng.Column - .Range.Column + 1
        If IsDate(.Range.Cells(1, FCol).text) Or IsDate(.Range.Cells(2, FCol).text) Then
            FDate = True
        End If
        With .Filters(FCol)
            If Not .On Then GoTo Finish
            If FDate = False Then
                Filter = .Criteria1
            Else
                For x = 1 To Len(.Criteria1)
                    If IsNumeric(Mid(.Criteria1, x, 1)) Then
                        Exit For
                    End If
                Next x
                Filter = Left(.Criteria1, x - 1) & Format(Mid(.Criteria1, x, 255), "dd/mm/yy")
            End If
            If .Operator <> 0 Then
                If FDate = False Then
                    Filter2 = .Criteria2
                Else
                    For x = 1 To Len(.Criteria2)
                        If IsNumeric(Mid(.Criteria2, x, 1)) Then
                            Exit For
                        End If
                    Next x
                    Filter2 = Left(.Criteria2, x - 1) & Format(Mid(.Criteria2, x, 255), "dd/mm/yy")
                End If
            End If
            Select Case .Operator
                Case xlAnd
                    Filter = Filter & " AND " & Filter2
                Case xlOr
                    Filter = Filter & " OR " & Filter2
            End Select
        End With
    End With
Finish:
    setFilterCriteria = Filter
End Function

Function getFilterCriteria(rng As Range, Optional incCriteria As Boolean) As String
'Failed attempt at applying a filter
'too many variations based on the value/value2 etc of a cell depending on the format of the cell

    Dim output As String
On Error GoTo Finish
    
    With rng.Parent.AutoFilter.Filters
        For c = 1 To .Count
            With .Item(c)
                If .On Then
                    If incCriteria Then
                        output = output & Replace(Replace(Cells(1, c).Address, "$1", ""), "$", "") & ":" & .Criteria1 & ", "
                    Else
                        output = output & Replace(Replace(Cells(1, c).Address, "$1", ""), "$", "") & ", "
                    End If
                End If
            End With
        Next c
    End With

Finish:
    If output = "" Then
        output = "None"
    End If
    getFilterCriteria = output
    'Resume Next
End Function

'
'Function getCharCount(strIn As String, chToCount As String) As Integer
'Dim i As Integer
'Dim chCounter
'Dim continue As Boolean
'
'On Error GoTo errGetCharCount
'
'If Len(strIn) > 0 Then
'    i = 1
'    chCounter = 0
'    continue = True
'
'    Do
'        If InStr(1, chToCount, Mid(strIn, i, 1), vbTextCompare) > 0 Then
'            chCounter = chCounter + 1
'        End If
'        i = i + 1
'    Loop While i < Len(strIn)
'
'    getCharCount = chCounter
'End If
'Exit Function
'
'errGetCharCount:
'    getCharCount = ""
'End Function

Function getCharCount(strIn As String, chToCount As String) As Integer

On Error GoTo errGetCharCount
If Len(strIn) > 0 Then
    getCharCount = Len(strIn) - Len(Replace(strIn, chToCount, ""))
End If
Exit Function

errGetCharCount:
    getCharCount = 0
End Function

Function GetLastxWords(strIn As String, chSeparator As String, intWordCount As Integer) As String
Dim intSepStart As Integer
Dim intSepEnd As Integer

intSepEnd = getCharCount(strIn, chSeparator) + 1
intSepStart = intSepEnd - intWordCount

GetLastxWords = ExtractTextBetween(strIn, chSeparator, intSepStart, intSepEnd)

End Function

Public Function getComments(ref As Range) As String
    If ref.CommentThreaded Is Nothing Then
        getComments = ""
    Else
        getComments = ref.CommentThreaded.text
    End If
End Function

Function WORDDIF(strA As String, strB As String) As String
'01/05/2023 added with permission Sophia Javier

    Dim WordsA As Variant, WordsB As Variant
    Dim ndxA As Long, ndxB As Long, strTemp As String
        
    WordsA = Split(strA, " ")
    WordsB = Split(strB, " ")
    
    For ndxB = LBound(WordsB) To UBound(WordsB)
        For ndxA = LBound(WordsA) To UBound(WordsA)
            If StrComp(WordsA(ndxA), WordsB(ndxB), vbTextCompare) = 0 Then
                WordsA(ndxA) = vbNullString
                Exit For
            End If
        Next ndxA
    Next ndxB
    
    For ndxA = LBound(WordsA) To UBound(WordsA)
        If WordsA(ndxA) <> vbNullString Then strTemp = strTemp & WordsA(ndxA) & " "
    Next ndxA
    
    WORDDIF = Trim(strTemp)

End Function

Function StringClean(strIn As String, codeList As Range) As String
Dim output As String
Dim i As Integer
Dim s() As Byte

s = StrConv(strIn, vbFromUnicode)
output = ""
 For i = 0 To UBound(s)
    For Each c In codeList.Cells
        If c.Column = 1 Then
            If (s(i) = c.value) Then
                If c.Next <> "" Then
                    output = output & c.Next
                Else
                    output = output & Chr(s(i))
                End If
                Exit For
            End If
        End If
    Next
Next
If output = "/" Then
    output = ""
End If

StringClean = Trim(output)
End Function
