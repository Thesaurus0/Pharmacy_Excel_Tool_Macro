VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtSysConf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Worksheet_Change(ByVal Target As Range)

    Application.EnableEvents = False

    On Error GoTo exit_function
    Call fShtSysConf_SheetChange_DevProdChange(Target)
    
exit_function:
    Application.EnableEvents = True
End Sub
 
 
