Attribute VB_Name = "친욥5"
Sub 브4()
Attribute 브4.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 브4 브
'

'
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="-10000", Formula2:="50000"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .IMEMode = xlIMEModeOff
        .ShowInput = True
        .ShowError = True
    End With
End Sub
