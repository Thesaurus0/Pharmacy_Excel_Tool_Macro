Attribute VB_Name = "ģ��2"
Sub ��3()
Attribute ��3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ��3 ��
'
 
    Application.AutomationSecurity = msoAutomationSecurityForceDisable  '???

 

    Dim wbSource As Workbook
    Set wbSource = Workbooks.Open(Filename:="F:\Github_Local_Repository\Pharmacy_Excel_Tool_Macro\Pharmacy_Excel_Tool_Macro_V0.5.xlsm", ReadOnly:=True)

    Application.AutomationSecurity = msoAutomationSecurityByUI  '????
End Sub

Sub cccaa()
    Dim sht As Worksheet
    
    Set sht = shtMenu
    If sht.FilterMode Then  'advanced filter
        sht.ShowAllData
    End If
    
    If sht.AutoFilterMode Then  'auto filter
        If fZero(asDegree) Or asDegree = "SHOW_ALL_DATA" Then
            sht.AutoFilter.ShowAllData
        Else
            sht.AutoFilterMode = False
        End If
    End If
End Sub

Sub caaddfa()
    MsgBox format(1 / 10, "Ϊ�������ó�ʼ����")
End Sub
