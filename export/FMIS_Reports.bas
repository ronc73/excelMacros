Attribute VB_Name = "FMIS_Reports"
Sub ProfileReport()
Selection.TextToColumns Destination:=Range("A1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(31, 1), Array(65, 1), Array(81, 1)), _
        TrailingMinusNumbers:=True
End Sub

Sub LoadReportPrep()
LoadReportPrep_Action (1)
End Sub

Sub LoadReportPrep2()
LoadReportPrep_Action (2)
End Sub

Sub LoadReportPrep3()
LoadReportPrep_Action (3)
End Sub

Sub LoadReportPrep_Action(opt As Integer)
Dim wbk As Workbook
Dim wbkVPCS As Workbook

Dim shError As Worksheet
Dim shSuccess As Worksheet
Dim shReportData As Worksheet
Dim shParameters As Worksheet

Dim rngPastedData As Range

Dim lastRow As Long
Dim endDataRow As Long

Dim outFolder As String
Dim saveFolder As String
Dim saveFile As String
Dim saveFilePrefix As String
Dim saveFileSufix As String
Dim saveFileDateRange As String


Set wbk = ActiveWorkbook
Set shError = wbk.Sheets("LoadErrors")
Set shSuccess = wbk.Sheets("LoadSuccess")

If opt = 2 Then
    Set shReportData = wbk.Sheets("ReportData")
End If

If opt = 2 Or opt = 3 Then
    Set shParameters = wbk.Sheets("Parameters")
'save template file out as "Consolidated Output File - 2022-03-26_2022-04-01_ReRun.xlsx"
'output folder = h:\Shared\Operational\DataSystems\SCIT\CommonCatalogue\PROD\02-Output
    saveFolder = "h:\Shared\Operational\DataSystems\SCIT\CommonCatalogue\PROD\02-Output\"
    saveFilePrefix = "Consolidated Output File - "
    saveFileDateRange = Format(shParameters.Range("d2").value, "YYYY-MM-DD") & "_" & Format(shParameters.Range("d3").value, "YYYY-MM-DD")
    If Len(shParameters.Range("b4").value) = 2 Then
        saveFileSufix = ".xlsx"
    Else
        saveFileSufix = "_ReRun.xlsx"
    End If

    saveFile = saveFilePrefix & saveFileDateRange & saveFileSufix
    wbk.SaveAs saveFolder & saveFile, xlWorkbookDefault
End If

'open if not yet open H:\Shared\Operational\DataSystems\SCIT\CommonCatalogue\Data\VPCS_Catalogue.xlsx
    Call SuspendAutoCalc

If opt = 1 Or opt = 2 Then
    Set wbkVPCS = Application.Workbooks.Open("H:\Shared\Operational\DataSystems\SCIT\CommonCatalogue\Data\VPCS_Catalogue.xlsx")
End If
wbk.Activate
'With ActiveSheet

    If opt = 1 Then
    ''select range AH1:AH<last row>
        With shError
            lastRow = .Cells.SpecialCells(xlCellTypeLastCell).Row
            Set rngPastedData = .Range(Cells(1, 34), Cells(.Cells.SpecialCells(xlCellTypeLastCell).Row, 34))
            rngPastedData.Select
        '    split col AH by |
            'call split col by...
            Application.DisplayAlerts = False
            Call TextToCols(2, "|")
            Application.DisplayAlerts = True
        '    copy formulas from A2:AF2 through A:AF to last row
        End With
    ElseIf opt = 2 Then
        'refresh oracle data
        wbk.RefreshAll
        'copy report data to Load errors tab
        'table name p01649_prod_services_10010_EHVPFMS
        shReportData.ListObjects("p01649_prod_services_10010_EHVPFMS").Range.Copy
        shError.Range("AH1").PasteSpecial xlPasteValuesAndNumberFormats
    ElseIf opt = 3 Then
        'nothing specific to do for opt 3
        'wbk.RefreshAll
    End If

If opt = 2 Or opt = 1 Then
    shError.Activate
    lastRow = shError.Cells.SpecialCells(xlCellTypeLastCell).Row
    Set rngPastedData = shError.Range(Cells(1, 1), Cells(lastRow, 32))
    shError.ListObjects("Table1").Resize rngPastedData
    Set rngPastedData = shError.Range(Cells(2, 1), Cells(lastRow, 32))
    shError.Range(Cells(2, 1), Cells(2, 32)).Copy
    rngPastedData.PasteSpecial xlPasteFormulas
    
    rngPastedData.Calculate
    
    'replace formulas with values A:AF
    Application.CutCopyMode = False
    rngPastedData.Copy
    rngPastedData.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'move success rows to "LoadSuccess" sheet
    ''filter status for success
    
    ''copy lines to LoadSuccess
    ''remove lines from LoadErrors
    shError.ListObjects("Table1").Range.AutoFilter Field:=11, Criteria1:="Success"
        shError.ListObjects("Table1").DataBodyRange.Select
        
        Selection.Copy
        
        With shSuccess
            .Activate
            .Range("Table6[FILE_NAME]").Select
            .Paste
            Application.CutCopyMode = treu
        End With
        
        shError.Activate
        Range("Table1[FILE_NAME]").EntireRow.Delete
        shError.ShowAllData
        Range("AG:BK").EntireColumn.Delete xlLeft
        shError.Cells.SpecialCells(xlCellTypeLastCell).Select
        endDataRow = shError.Cells.SpecialCells(xlCellTypeLastCell).End(xlUp).Row
        Cells(shError.Cells.SpecialCells(xlCellTypeLastCell).Row, 1).Activate
    
        endDataRow = Cells(shError.Cells.SpecialCells(xlCellTypeLastCell).Row, 1).End(xlUp).Row
        Range(Cells(endDataRow + 1, 1), Cells(lastRow, 1)).EntireRow.Delete
        Range("A1").Select
    
End If

If opt = 2 Or opt = 3 Then
    'delete sheets parameters and Report data
    Application.DisplayAlerts = False
    shParameters.Delete
    If opt = 2 Then
        shReportData.Delete
    End If
    Application.DisplayAlerts = True
End If

Call PivotRefresh
'refresh pivots

Call ResumeAutoCalc

If opt = 2 Or opt = 3 Then
    wbk.Save
End If

End Sub

