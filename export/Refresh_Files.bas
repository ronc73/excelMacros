Attribute VB_Name = "Refresh_Files"
'open a file, refresh the data then save it after

Sub VPCS_Report_Files_Refresh()

Dim filePath  As String
Dim FileList() As Variant
Dim fileName As Variant

    filePath = "H:\Shared\Operational\DataSystems\SCIT\CommonCatalogue\Data\"
    FileList = Array("VPCS_Catalogue_CE_Extract.xlsx", "VPCS_Catalogue_Eligible_Services.xlsx", "VPCS_Catalogue_Pharma.xlsx", "VPCS_Catalogue_RebateCodes.xlsx", "VPCS_Catalogue_Full_Extract.xlsx", "VPCS_Catalogue_Bulk Pricing.xlsx")
    'FileList = Array("VPCS_Catalogue_Pharma.xlsx", "VPCS_Catalogue_RebateCodes.xlsx", "VPCS_Catalogue_Bulk Pricing.xlsx")

    'For i = 1 To UBound(FileList)
    For Each fileName In FileList
        If doRefresh(filePath, fileName) Then
            result = result + fileName + " Success" & vbCr
        Else
            result = result + fileName + " Fail"
        End If
    Next fileName
    MsgBox result

End Sub

Function doRefresh(Path As String, fileName As Variant) As Boolean
Dim workingFile As Workbook

On Error GoTo doRefreshErr
    Set workingFile = Application.Workbooks.Open(Path & fileName)
    Call DBRefresh(True)
    'workingFile.Save
    workingFile.Close True
    doRefresh = True
    Exit Function
doRefreshErr:
    doRefresh = False
End Function
