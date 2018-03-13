Attribute VB_Name = "친욥1"
Sub 브1()
Attribute 브1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 브1 브
'

'
   ' shtMenu.cbbCompanyList.Select
    shtMenu.cbbCompanyList.Activate
    shtMenu.cbbCompanyList.SelStart = 0
    shtMenu.cbbCompanyList.SelLength = Len(shtMenu.cbbCompanyList.Value)
    
End Sub
