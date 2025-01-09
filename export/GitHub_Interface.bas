Attribute VB_Name = "GitHub_Interface"
Private Sub exportVBA_All()
    exportvba "export"
End Sub
Private Function exportvba(Optional Path As String)
Dim objVbComp As Variant ' VBComponent
Dim strPath As String
Dim varItem As Variant
Dim fso As String
Dim fileName As String

If Path <> "" Then
    fileName = Path
Else
    fileName = "export"
End If

strPath = Environ("USERPROFILE") & "\Macro\" & fileName & "\"

On Error Resume Next
    MkDir (strPath)
On Error GoTo 0

'Change the path to suit the users needs

  For Each varItem In ThisWorkbook.VBProject.VBComponents
  Set objVbComp = varItem

  Select Case objVbComp.Type
     Case 1 ' vbext_ct_StdModule
        objVbComp.Export strPath & "\" & objVbComp.Name & ".bas"
     Case vbext_ct_Document, vbext_ct_ClassModule
        ' ThisDocument and class modules
        objVbComp.Export strPath & "\" & objVbComp.Name & ".cls"
     Case 3 ' vbext_ct_MSForm
        objVbComp.Export strPath & "\" & objVbComp.Name & ".frm"
     Case 100 ' regular worksheet. no need to export
     Case Else
        objVbComp.Export strPath & "\" & objVbComp.Name & "." & objVbComp.Type
  End Select
Next varItem
End Function

Sub import_All_Code()
'MsgBox Application.UserName
'MsgBox Environ("APPDATA")
'MsgBox Environ(41)
MsgBox Environ("USERPROFILE")

End Sub

Sub AllEnvironVariables()
    Dim strEnviron As String
    Dim VarSplit As Variant
    Dim i As Long
    For i = 1 To 255
        strEnviron = Environ$(i)
        If LenB(strEnviron) = 0& Then GoTo TryNext:
        VarSplit = Split(strEnviron, "=")
        If UBound(VarSplit) > 1 Then Stop
        Range("A" & Range("A" & Rows.Count).End(xlUp).Row + 1).value = i
        Range("B" & Range("B" & Rows.Count).End(xlUp).Row + 1).value = VarSplit(0)
        Range("C" & Range("C" & Rows.Count).End(xlUp).Row + 1).value = VarSplit(1)
TryNext:
    Next
End Sub
