Attribute VB_Name = "Auto_Calc"
Global AutoCalcFlag As Integer

Function SuspendAutoCalc()
'6/12/12 - Ron Campbell
'use at the start of a macro to turn off auto calculate
'if auto calculation is currently turned on then turn it off while the macro is running then turn it back on after.
'if auto calculate is turned off then leave it off
If ActiveWorkbook Is Nothing Then
    Exit Function
End If

    'set flag to the current setting of auto calc, manual, auto or semiauto
On Error GoTo errorHandler
    'turn off screen updating, speed improvement
    Application.ScreenUpdating = False
    If Not IsNull(Application.Calculation) Then
        AutoCalcFlag = Application.Calculation
    'Set to manual
        Application.Calculation = xlCalculationManual
    End If
    
ErrorExit:
    
    'the ErrorHandler code should only be executed if there is an error
    Exit Function
errorHandler:
        Debug.Print Err.Number & vbLf & Err.Description
        Resume ErrorExit

End Function

Function ResumeAutoCalc()
'6/12/12 - Ron Campbell
'used at the end of a macro to restore the setting for autocalculate
'Reverse of SuspendAutoCalc

If ActiveWorkbook Is Nothing Then
    Exit Function
End If

If AutoCalcFlag <> 0 Then
        Application.Calculation = AutoCalcFlag
    Else
        Application.Calculation = xlCalculationAutomatic
    End If
    Application.StatusBar = ""
    'turn on screen updating
    Application.ScreenUpdating = True

End Function


