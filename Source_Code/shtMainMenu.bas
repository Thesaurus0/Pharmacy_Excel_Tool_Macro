VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub btnCloseAllSheet2_Click()
'    If shtHospital.Visible = xlSheetVisible Then
'        subMain_InvisibleHideAllBusinessSheets
'    Else
'        subMain_ShowAllBusinessSheets
'    End If
    subMain_InvisibleHideAllBusinessSheets
End Sub

Private Sub btnCloseAllSheets_Click()
    subMain_InvisibleHideAllBusinessSheets
End Sub

Private Sub btnImportCZLSalesToSaleComp_Click()
    
    fActiveVisibleSwitchSheet shtImportCZL2SalesCompSales, "AK15", False
End Sub

Private Sub btnNewRuleProducts_Click()
    subMain_NewRuleProducts
End Sub

Private Sub btnReplaceCZLSales2Comp_Click()
    subMain_ReplaceCZLSales2Comp
End Sub
