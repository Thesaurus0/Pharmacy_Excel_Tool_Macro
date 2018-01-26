Attribute VB_Name = "模块2"
Sub 宏3()
Attribute 宏3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 宏3 宏
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
    MsgBox format(1 / 10, "为画面设置初始数据")
End Sub
