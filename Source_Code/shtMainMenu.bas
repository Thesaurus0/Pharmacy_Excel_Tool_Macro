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
    subMain_CalCZLInventory
End Sub

Private Sub btnCalRefund_Click()
    subMain_CalculateRefundOrSupplement
End Sub

Private Sub btnCompareChange_Click()
    subMain_CompareChangeWithPrevVersion
End Sub

Private Sub btnCompareCZLInv_Click()
    subMain_CompareCZLInventory
End Sub

Private Sub btnCreateMEFile_CZLSales2SComp_Click()
    subMain_CreateMonthEndFile_CZLSales2SComp
End Sub

Private Sub btnCreateMEFile_Profit_Click()
    subMain_CreateMonthEndFile_Profit
End Sub

Private Sub btnCZL_MEInvRollover_Click()
    subMain_CZLMonthEndInventoryRollOver
End Sub

Private Sub btnCZLCommConfig_Click()
subMain_FirstLevelCommission
End Sub

'Private Sub btnCZLInformedInvInput_Click()
'    fActiveVisibleSwitchSheet shtCZLInformedInvInput, , False
'End Sub

Private Sub btnCZLInvDiffSheet_Click()
    fActiveVisibleSwitchSheet shtCZLInvDiff, , False
End Sub

Private Sub btnCZLInventorySheet_Click()
    fActiveVisibleSwitchSheet shtCZLInventory, , False
End Sub

Private Sub btnCZLInventorySheet2_Click()

    fActiveVisibleSwitchSheet shtCZLInventory, , False
End Sub

Private Sub btnCZLInvImported_Click()
    Dim sCZLCompName As String
    sCZLCompName = fGetCompanyNameByID_Common("CZL")
    
    fRemoveFilterForSheet shtSalesCompInvUnified
    Call fSetFilterForSheet(shtSalesCompInvUnified, Array(SCompUnifiedInv.SalesCompany), Array(sCZLCompName))
        
    fActiveVisibleSwitchSheet shtSalesCompInvUnified, , False
End Sub

Private Sub btnCZLInvImported2_Click()
    fActiveVisibleSwitchSheet shtSalesCompInvCalcd, , False
End Sub

Private Sub btnCZLPurchaseOrder_Click()
    Call fPrepareCZLPurchaseFromSelfSales
'    Dim arrSelf()
'    fClearContentLeaveHeader shtCZLPurchaseOrder
'    Call fCopyReadWholeSheetData2Array(shtSelfSalesOrder, arrSelf, , , fLetter2Num("H"))
'    Call fWriteArray2Sheet(shtCZLPurchaseOrder, arrSelf)
'    Erase arrSelf
'    fActiveVisibleSwitchSheet shtCZLPurchaseOrder
End Sub


Private Sub btnCZLRolloverInv_Click()
    fActiveVisibleSwitchSheet shtCZLRolloverInv
End Sub

Private Sub btnCZLSalesMEAllForReview_Click()
    fActiveVisibleSwitchSheet shtCZLSales2SCompAll
End Sub

Private Sub btnCZLSalesMEAllForReview1_Click()
    fActiveVisibleSwitchSheet shtCZLSales2SCompAll
End Sub

Private Sub btnCZLSalesToCompanies_Click()
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

Private Sub btnCZLSalesToHospital_Click()
    Call fPrepareCZLSales2HospitalByFiltering
End Sub

Private Sub btnHospitalMaster_Click()
    subMain_Hospital
End Sub

Private Sub btnHospitalReplace_Click()
    subMain_HospitalReplacement
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


Private Sub btnMECalProfit_Click()
    subMain_CalculateProfit_MonthEnd
End Sub

Private Sub btnMECZLSales2SComp_Click()
    subMain_SaveCZLSales2SCompTableToHistory
End Sub

Private Sub btnMEProfitToHist_Click()
    subMain_SaveProfitTableToHistory
End Sub

Private Sub btnNewRuleProducts_Click()
    subMain_NewRuleProducts
End Sub

Private Sub btnOpenHistCZLSales2SComp_Click()
    subMain_OpenHistFile_CZLSales2SComp
End Sub

Private Sub btnOpenHistProfitFile_Click()
    subMain_OpenHistProfitFile
End Sub

Private Sub btnProducerMaster_Click()
    subMain_ProducerMaster
End Sub

Private Sub btnProductMaster_Click()
    subMain_ProductMaster
End Sub

Private Sub btnProductNameMaster_Click()
    subMain_ProductNameMaster
End Sub

Private Sub btnProductNameReplace_Click()
    subMain_ProductNameReplace
End Sub

Private Sub btnProductProducerReplace_Click()
    subMain_ProductProducerReplace
End Sub

Private Sub btnProductSeriesReplace_Click()
    subMain_ProductSeriesReplace
End Sub

Private Sub btnProductTaxRate_Click()
    fActiveVisibleSwitchSheet shtProductTaxRate, , False
End Sub

Private Sub btnProductUnitRatio_Click()
    subMain_ProductUnitRatio
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

Private Sub btnRefundSheet_Click()
    fActiveVisibleSwitchSheet shtRefund, , False
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

Private Sub btnSalesCompCommConf_Click()
    subMain_SecondLevelCommission
End Sub

Private Sub btnSalesCompInvCal_Click()
    subMain_CalculateSalesCompInventory
End Sub

Private Sub btnSalesCompInvDiffCal_Click()
    subMain_CompareSalesCompanyInventory
End Sub

Private Sub btnSalesCompInvDiffSheet_Click()
    fActiveVisibleSwitchSheet shtSalesCompInvDiff, , False
End Sub

Private Sub btnSalesCompInvCald_Click()
    fActiveVisibleSwitchSheet shtSalesCompInvCalcd, , False
End Sub

Private Sub btnSalesCompInvUnfied_Click()
    fActiveVisibleSwitchSheet shtSalesCompInvUnified, , False
    fRemoveFilterForSheet shtSalesCompInvUnified
End Sub

Private Sub btnSalesCompPurchase_Click()
    fActiveVisibleSwitchSheet shtCZLSales2Companies, , False
 '   fRemoveFilterForSheet shtCZLSales2Companies
End Sub

Private Sub btnSalesCompRolloverInv_Click()
    fActiveVisibleSwitchSheet shtSalesCompRolloverInv, , False
  '  fRemoveFilterForSheet shtSalesCompRolloverInv
End Sub

Private Sub btnSalesCompSales_Click()
    fActiveVisibleSwitchSheet shtSalesInfos, , False
 '   fRemoveFilterForSheet shtSalesInfos
End Sub

Private Sub btnSalesManCommConfig_Click()
    subMain_SalesManCommissionConfig
End Sub

Private Sub btnSComp_MEInvRollover_Click()
    subMain_SalesCompanyMonthEndInventoryRollOver
End Sub

Private Sub btnSCompInvRawData_Click()
    fActiveVisibleSwitchSheet shtInventoryRawDataRpt
End Sub

Private Sub btnSCompInvUnified_Click()
    fActiveVisibleSwitchSheet shtSalesCompInvUnified
End Sub

Private Sub btnSelfInventory_Click()
    subMain_CalculateSelfInventory
End Sub

Private Sub btnSelfInventorySheet_Click()
    fActiveVisibleSwitchSheet shtSelfInventory, , False
End Sub

Private Sub btnSelfPurchaseOrder_Click()
    subMain_SelfPurchaseOrder
End Sub

Private Sub btnSelfSalesOrder_Click()
    subMain_SelfSalesOrder
End Sub

Private Sub btnSelfSalesOrderPrededuct_Click()
    fActiveVisibleSwitchSheet shtSelfSalesPreDeduct, , False
End Sub

Private Sub btnSellPriceInAdv_Click()
    fActiveVisibleSwitchSheet shtSellPriceInAdv, , False
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

Private Sub CommandButton1_Click()
    Sub_DataMigration
End Sub
