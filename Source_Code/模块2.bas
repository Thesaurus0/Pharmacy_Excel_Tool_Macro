Attribute VB_Name = "친욥2"
Sub 브2()
Attribute 브2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 브2 브
'

'
    Columns("F:F").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator _
        :=xlBetween, Formula1:="0", Formula2:="1"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
End Sub
Sub 브3()
Attribute 브3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 브3 브
'

'
    Columns("E:E").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="1/1/2001", Formula2:="12/31/2099"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
    Range("E6").Select
End Sub
