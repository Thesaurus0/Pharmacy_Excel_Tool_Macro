Attribute VB_Name = "ģ��5"
Sub ��4()
Attribute ��4.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ��4 ��
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
