Attribute VB_Name = "ģ��2"
Sub ��2()
Attribute ��2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ��2 ��
'

'
    Sheets("��ҵ��˾���ͷ�").Select
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
