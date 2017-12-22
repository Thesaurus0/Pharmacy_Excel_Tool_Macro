Attribute VB_Name = "模块2"
Sub 宏2()
Attribute 宏2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 宏2 宏
'

'
    Sheets("商业公司配送费").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=rngStaticSalesCompanyNames"
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
    Range("A191").Select
End Sub
