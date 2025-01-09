Attribute VB_Name = "Descriptions_EK_Funcrtions"



Sub getfile(StrFlNm As String)

With Application.FileDialog(msoFileDialogOpen)
  .AllowMultiSelect = False
  
  If .Show = -1 Then
    StrFlNm = .SelectedItems(1)
  Else
    Exit Sub
  End If
End With


End Sub

Function findinrow(s As String, rng As Range) As Integer

Dim c As Integer
c = 0
Dim hcr As Range
For Each hcr In rng.Cells
    If Trim(CStr(hcr.value)) = s Then
        c = hcr.Column
        Exit For
    End If
Next hcr

findinrow = c

End Function

Function findincol(s As String, rng As Range) As Integer

Dim r As Integer
r = 0
Dim hcr As Range
For Each hcr In rng.Cells
    If Trim(CStr(hcr.value)) = s Then
        r = hcr.Row
        Exit For
    End If
Next hcr

findincol = r


End Function

Function findlrw(rng As Range) As Integer

Dim r As Integer
Dim hcr As Range

For Each hcr In rng.Cells
    If Trim(CStr(hcr.value)) = "END" Or IsEmpty(hcr.value) Or hcr.value = " |  |  | " Then
        r = hcr.Row
        Exit For
    End If
Next hcr

findlrw = r


End Function

Sub fndrng(lstrw As Integer, frw As Integer, c As Integer, title As String, frng As Range, ws As Worksheet, src As Range)

c = findinrow(title, src)
lstrw = ws.UsedRange.Rows.Count

Set frng = ws.Range(ws.Cells(frw, c), ws.Cells(lstrw, c))


End Sub

Function checklatexstatus(ltxst As String, ltxop As String) As String

Dim b1 As String
Dim b2 As String

If UCase(ltxst) = "TRUE" Or UCase(ltxst) = "YES" Or UCase(ltxst) = "LATEX" Then

    b1 = "TRUE"

ElseIf UCase(ltxst) = "FALSE" Or UCase(ltxst) = "NO" Or UCase(ltxst) = "LATEX-FREE" Or UCase(ltxst) = "LATEX FREE" Or UCase(ltxst) = "N/A" Then

    b1 = "FALSE"
            
Else

    b1 = "check"

End If

If UCase(ltxop) = "TRUE" Or UCase(ltxop) = "YES" Or UCase(ltxop) = "LATEX (CONTAINS)" Then

    b2 = "TRUE"
    
ElseIf UCase(ltxop) = "LATEX FREE" Or UCase(ltxop) = "NO" Or UCase(ltxop) = "FALSE" Or UCase(ltxop) = "LATEX-FREE" Then

     b2 = "FALSE"
     
ElseIf UCase(ltxop) = "NO REFERENCE TO LATEX" Then
    
    b2 = ""
    
Else

    b2 = "check"

End If


If (b1 = "TRUE" And b2 = "FALSE") Or (b1 = "FALSE" And b2 = "TRUE") Or b1 = "check" Or b2 = "check" Or (b1 = "TRUE" And b2 = "") Then

    latexstatus = "Latex Inconsistency"

Else

    latexstatus = ""

End If

checklatexstatus = latexstatus

End Function

Function isinarray(str As String, inarray() As String) As Boolean
Dim b As Boolean

b = False

i = 0

For i = LBound(inarray) To UBound(inarray)

    If inarray(i) = str Then
    
        b = True
        Exit For
    End If
    

Next i

isinarray = b

End Function




