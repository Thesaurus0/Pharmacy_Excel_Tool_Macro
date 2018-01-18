Attribute VB_Name = "친욥2"
Sub 브3()
Attribute 브3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 브3 브
'
 
    Application.AutomationSecurity = msoAutomationSecurityForceDisable  '???

 

    Dim wbSource As Workbook
    Set wbSource = Workbooks.Open(Filename:="F:\Github_Local_Repository\Pharmacy_Excel_Tool_Macro\Pharmacy_Excel_Tool_Macro_V0.5.xlsm", ReadOnly:=True)

    Application.AutomationSecurity = msoAutomationSecurityByUI  '????
End Sub
