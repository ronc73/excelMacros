Attribute VB_Name = "Macros_Internet"
Sub highlightFound()
'Added by Trang Tran 17/06/2019
'Highlight words in one list based on a secod list

Dim WorkRng1 As Range, WorkRng2 As Range, rng1 As Range, Rng2 As Range
xTitleId = "Duplicate Highlight"
Set WorkRng1 = Application.InputBox("Data to highlight if duplicate found:", xTitleId, "", Type:=8)
Set WorkRng2 = Application.InputBox("Data source comparing against:", xTitleId, Type:=8)
For Each rng1 In WorkRng1
    rng1Value = rng1.value
    For Each Rng2 In WorkRng2
        If rng1Value = Rng2.value Then
            rng1.Interior.Color = VBA.RGB(0, 204, 204)
            Exit For
        End If
    Next
Next

End Sub

Public Function Levenshtein(s1 As String, s2 As String)
'Added 9/01/17
'From Euan Boulton
'Levenshtein Distance presents a count of the number of changes required to convert one text string into another
'
Dim i As Integer
Dim j As Integer
Dim l1 As Integer
Dim l2 As Integer
Dim d() As Integer
Dim min1 As Integer
Dim min2 As Integer

l1 = Len(s1)
l2 = Len(s2)
ReDim d(l1, l2)
For i = 0 To l1
    d(i, 0) = i
Next
For j = 0 To l2
    d(0, j) = j
Next
For i = 1 To l1
    For j = 1 To l2
        If Mid(s1, i, 1) = Mid(s2, j, 1) Then
            d(i, j) = d(i - 1, j - 1)
        Else
            min1 = d(i - 1, j) + 1
            min2 = d(i, j - 1) + 1
            If min2 < min1 Then
                min1 = min2
            End If
            min2 = d(i - 1, j - 1) + 1
            If min2 < min1 Then
                min1 = min2
            End If
            d(i, j) = min1
        End If
    Next
Next
Levenshtein = d(l1, l2)
End Function


Public Sub HighlightDupesCaseInsensitive()
    Dim cell As Range
    Dim delimiter As String
 
    delimiter = InputBox("Enter the delimiter that separates values in a cell", "Delimiter", ", ")
 
    For Each cell In Application.Selection
        Call HighlightDupeWordsInCell(cell, delimiter, False)
    Next
End Sub
 
Sub HighlightDupeWordsInCell(cell As Range, Optional delimiter As String = " ", Optional CaseSensitive As Boolean = True)
    Dim text As String
    Dim words() As String
    Dim word As String
    Dim wordIndex, matchCount, positionInText As Integer
 
    If CaseSensitive Then
        words = Split(cell.value, delimiter)
    Else
        words = Split(LCase(cell.value), delimiter)
    End If
 
    For wordIndex = LBound(words) To UBound(words) - 1
        word = words(wordIndex)
        matchCount = 0
 
        For nextWordIndex = wordIndex + 1 To UBound(words)
            If word = words(nextWordIndex) Then
                matchCount = matchCount + 1
            End If
 
        Next nextWordIndex
 
        If matchCount > 0 Then
            text = ""
 
            For Index = LBound(words) To UBound(words)
                text = text & words(Index)
 
                If (words(Index) = word) Then
                    cell.Characters(Len(text) - Len(word) + 1, Len(word)).Font.Color = vbRed
                End If
 
                text = text & delimiter
            Next
        End If
    Next wordIndex
End Sub

Function RemoveDupeWords(text As String, Optional delimiter As String = " ") As String
    Dim dictionary As Object
    Dim x, part
 
    Set dictionary = CreateObject("Scripting.Dictionary")
    dictionary.CompareMode = vbTextCompare
    For Each x In Split(text, delimiter)
        part = Trim(x)
        If part <> "" And Not dictionary.exists(part) Then
            dictionary.Add part, Nothing
        End If
    Next
 
    If dictionary.Count > 0 Then
        RemoveDupeWords = Join(dictionary.keys, delimiter)
    Else
        RemoveDupeWords = ""
    End If
 
    Set dictionary = Nothing
End Function

