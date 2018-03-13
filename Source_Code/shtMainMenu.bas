VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub btnCalculateCZLInventory_Click()
    subMain_CZLInventory
End Sub

Private Sub btnCompareCZLInv_Click()
    subMain_CompareCZLInventory
End Sub

Private Sub btnCZLPurchaseOrder_Click()
    Dim arrSelf()
    fClearContentLeaveHeader shtCZLPurchaseOrder
    Call fCopyReadWholeSheetData2Array(shtSelfSalesOrder, arrSelf, , , fLetter2Num("H"))
    Call fWriteArray2Sheet(shtCZLPurchaseOrder, arrSelf)
    Erase arrSelf
    fActiveVisibleSwitchSheet shtCZLPurchaseOrder
End Sub

Private Sub btnCZLSalesOrder_Click()
    fActiveVisibleSwitchSheet shtCZLSales2Companies
End Sub

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

Private Sub btnCompanyNameReplaceConf_Click()
    fActiveVisibleSwitchSheet shtCompanyNameReplace, , False
End Sub



Private Sub btnCZLSalesToCompany_Click()
    fActiveVisibleSwitchSheet shtCZLSales2Companies, , False
End Sub

Private Sub btnCZLSalesToCompRawData_Click()
    fActiveVisibleSwitchSheet shtCZLSales2CompRawData, , False
End Sub

Private Sub btnImportCompInventory_Click()
    subMain_ImportSalesCompanyInventory
End Sub

Private Sub btnImportCZLSalesToSaleComp_Click()
    
    fActiveVisibleSwitchSheet shtImportCZL2SalesCompSales, "AK15", False
End Sub

Private Sub btnImportSalesCompSalesFile_Click()
    subMain_Ribbon_ImportSalesInfoFiles
End Sub

Private Sub btnNewRuleProducts_Click()
    subMain_NewRuleProducts
End Sub

Private Sub btnProductMaster_Click()
    subMain_ProductMaster
End Sub

Private Sub btnProfit_Click()
subMain_Profit
End Sub

Private Sub btnPromotionProducts_Click()
    fActiveVisibleSwitchSheet shtPromotionProduct, , False
End Sub

Private Sub btnPvTables_Click()
    subMain_RefreshAllPvTables
End Sub

Private Sub btnRawSalesInfo_Click()
subMain_RawSalesInfos
End Sub

Private Sub btnReplaceCompInv_Click()
    subMain_ReplaceInventory
End Sub

Private Sub btnReplaceCZLSales2Comp_Click()
    subMain_ReplaceCZLSales2Comp
End Sub
 
Private Sub btnReplaceSalesInfo_Click()
    subMain_ReplaceSalesInfos
End Sub

Private Sub btnSelfInventory_Click()
    subMain_SelfInventory
End Sub

Private Sub btnSelfPurchaseOrder_Click()
    subMain_SelfPurchaseOrder
End Sub

Private Sub btnSelfSalesOrder_Click()
    subMain_SelfSalesOrder
End Sub

Private Sub btnTrialCalProfit_Click()
    subMain_CalculateProfit_PreCal
End Sub

Private Sub btnUnifiedSalesInfo_Click()
subMain_SalesInfos
End Sub

Private Sub btnValidateAllSheet_Click()
    subMain_ValidateAllSheetsData
End Sub
