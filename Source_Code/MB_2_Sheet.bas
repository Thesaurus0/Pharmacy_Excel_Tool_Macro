Attribute VB_Name = "MB_2_Sheet"
Option Explicit
 
Sub subMain_GetCurrentRowBusinessInfo()
    Call sub_PresstgGetSearchBy
    Call subActivateRibbonTab
End Sub

Function fGetColIndexs(shtParam As Worksheet, Optional ByRef iColIndex_Hospital As Integer _
            , Optional ByRef iColIndex_SalesCompany As Integer, Optional ByRef iColIndex_ProductProducer As Integer _
            , Optional ByRef iColIndex_ProductName As Integer, Optional ByRef iColIndex_ProductSeries As Integer _
            , Optional ByRef iColIndex_LotNum As Integer) As Boolean
    iColIndex_LotNum = 0
    iColIndex_ProductSeries = 0
    iColIndex_ProductName = 0
    iColIndex_ProductProducer = 0
    iColIndex_SalesCompany = 0
     
    Dim bOut As Boolean
    bOut = True
    
    Select Case UCase(shtParam.CodeName)
        Case UCase("shtFirstLevelCommission")
            iColIndex_ProductSeries = FirstLevelComm.ProductSeries
            iColIndex_ProductName = FirstLevelComm.ProductName
            iColIndex_ProductProducer = FirstLevelComm.ProductProducer
            iColIndex_SalesCompany = FirstLevelComm.SalesCompany
        Case UCase("shtHospital")
            iColIndex_Hospital = enHospital.HospitalName
        Case UCase("shtHospitalReplace")
            iColIndex_Hospital = enHospitalReplace.ToHospital
        Case UCase("shtSalesInfos")
            iColIndex_Hospital = Sales2Hospital.Hospital
            iColIndex_LotNum = Sales2Hospital.LotNum
            iColIndex_ProductSeries = Sales2Hospital.ProductSeries
            iColIndex_ProductName = Sales2Hospital.ProductName
            iColIndex_ProductProducer = Sales2Hospital.ProductProducer
            iColIndex_SalesCompany = Sales2Hospital.SalesCompany
        Case UCase("shtSalesCompInvUnified")
            iColIndex_LotNum = SCompUnifiedInv.LotNum
            iColIndex_ProductSeries = SCompUnifiedInv.ProductSeries
            iColIndex_ProductName = SCompUnifiedInv.ProductName
            iColIndex_ProductProducer = SCompUnifiedInv.ProductProducer
            iColIndex_SalesCompany = SCompUnifiedInv.SalesCompany
        Case UCase("shtCZLInvDiff")
            iColIndex_LotNum = CZLInvDiff.LotNum
            iColIndex_ProductSeries = CZLInvDiff.ProductSeries
            iColIndex_ProductName = CZLInvDiff.ProductName
            iColIndex_ProductProducer = CZLInvDiff.ProductProducer
        Case UCase("shtSalesCompInvDiff")
            iColIndex_LotNum = SCompInvDiff.LotNum
            iColIndex_ProductSeries = SCompInvDiff.ProductSeries
            iColIndex_ProductName = SCompInvDiff.ProductName
            iColIndex_ProductProducer = SCompInvDiff.ProductProducer
            iColIndex_SalesCompany = SCompInvDiff.SalesCompany
        Case UCase("shtSalesCompInvCalcd")
            iColIndex_LotNum = SCompInvCalcd.LotNum
            iColIndex_ProductSeries = SCompInvCalcd.ProductSeries
            iColIndex_ProductName = SCompInvCalcd.ProductName
            iColIndex_ProductProducer = SCompInvCalcd.ProductProducer
            iColIndex_SalesCompany = SCompInvCalcd.SalesCompany
        Case UCase("shtProfit")
            iColIndex_Hospital = Profit.Hospital
            iColIndex_LotNum = Profit.LotNum
            iColIndex_ProductSeries = Profit.ProductSeries
            iColIndex_ProductName = Profit.ProductName
            iColIndex_ProductProducer = Profit.ProductProducer
            iColIndex_SalesCompany = Profit.SalesCompany
        Case UCase("shtCZLInventory")
            iColIndex_LotNum = CZLInv.LotNum
            iColIndex_ProductSeries = CZLInv.ProductSeries
            iColIndex_ProductName = CZLInv.ProductName
            iColIndex_ProductProducer = CZLInv.ProductProducer
        Case UCase("shtSelfInventory")
            iColIndex_LotNum = SelfInv.LotNum
            iColIndex_ProductSeries = SelfInv.ProductSeries
            iColIndex_ProductName = SelfInv.ProductName
            iColIndex_ProductProducer = SelfInv.ProductProducer
        Case UCase("shtSelfSalesOrder")
            iColIndex_LotNum = SelfSales.LotNum
            iColIndex_ProductSeries = SelfSales.ProductSeries
            iColIndex_ProductName = SelfSales.ProductName
            iColIndex_ProductProducer = SelfSales.ProductProducer
        Case UCase("shtSelfPurchaseOrder")
            iColIndex_LotNum = SelfPurchase.LotNum
            iColIndex_ProductSeries = SelfPurchase.ProductSeries
            iColIndex_ProductName = SelfPurchase.ProductName
            iColIndex_ProductProducer = SelfPurchase.ProductProducer
        Case UCase("shtRefund")
            iColIndex_Hospital = Refund.Hospital
            iColIndex_LotNum = Refund.LotNum
            iColIndex_ProductSeries = Refund.ProductSeries
            iColIndex_ProductName = Refund.ProductName
            iColIndex_ProductProducer = Refund.ProductProducer
            iColIndex_SalesCompany = Refund.SalesCompany
        Case UCase("shtSecondLevelCommission")
            iColIndex_Hospital = SecondLevelComm.Hospital
            iColIndex_ProductSeries = SecondLevelComm.ProductSeries
            iColIndex_ProductName = SecondLevelComm.ProductName
            iColIndex_ProductProducer = SecondLevelComm.ProductProducer
            iColIndex_SalesCompany = SecondLevelComm.SalesCompany
        Case UCase("shtCZLSales2Companies")
            iColIndex_LotNum = CZLSales2Comp.LotNum
            iColIndex_ProductSeries = CZLSales2Comp.ProductSeries
            iColIndex_ProductName = CZLSales2Comp.ProductName
            iColIndex_ProductProducer = CZLSales2Comp.ProductProducer
            iColIndex_SalesCompany = CZLSales2Comp.SalesCompany
        Case UCase("shtPromotionProduct")
            iColIndex_Hospital = PromoteProduct.Hospital
            iColIndex_ProductSeries = PromoteProduct.ProductSeries
            iColIndex_ProductName = PromoteProduct.ProductName
            iColIndex_ProductProducer = PromoteProduct.ProductProducer
            iColIndex_SalesCompany = PromoteProduct.SalesCompany
        Case UCase("shtProductSeriesReplace")
            iColIndex_ProductSeries = ProdSerReplace.ProductSeries
            iColIndex_ProductName = ProdSerReplace.ProductName
            iColIndex_ProductProducer = ProdSerReplace.ProductProducer
        Case UCase("shtProductMaster")
            iColIndex_ProductSeries = ProductMst.ProductSeries
            iColIndex_ProductName = ProductMst.ProductName
            iColIndex_ProductProducer = ProductMst.ProductProducer
        Case UCase("shtSellPriceInAdv")
            iColIndex_ProductSeries = SellPriceInAdv.ProductSeries
            iColIndex_ProductName = SellPriceInAdv.ProductName
            iColIndex_ProductProducer = SellPriceInAdv.ProductProducer
            iColIndex_SalesCompany = SellPriceInAdv.SalesCompany
        Case UCase("shtCZLSales2SCompAll")
            iColIndex_ProductSeries = CZLSales2CompHist.ProductSeries
            iColIndex_ProductName = CZLSales2CompHist.ProductName
            iColIndex_ProductProducer = CZLSales2CompHist.ProductProducer
            iColIndex_SalesCompany = CZLSales2CompHist.SalesCompany
        Case UCase("shtProductNameReplace")
            iColIndex_ProductName = ProdNameReplace.ProductName
            iColIndex_ProductProducer = ProdNameReplace.ProductProducer
        Case Else
            bOut = False
    End Select
    
    fGetColIndexs = bOut
End Function

Private Sub btnSelfInventory_Click(control As IRibbonControl)
    Call Sub_SearchOnCurrentSheetBySearchBy(shtSelfInventory)
'    'fPresstgGetSearchBy
'    Call fRefreshEditBoxFromShtDataStage
'
'    If Len(ebProductProducer_val) > 0 Then
''        Call fSetFilterForSheet(shtSelfInventory, Array(SelfInv.ProductProducer, SelfInv.ProductName, SelfInv.ProductSeries, SelfInv.LotNum) _
'                , Array(ebProductProducer_val, ebProductName_val, ebProductSeries_val, ebLotNum_val))
'        Call fSetFilterForSheet(shtSelfInventory, Array(SelfInv.ProductProducer, SelfInv.ProductName, SelfInv.ProductSeries) _
'                , Array(ebProductProducer_val, ebProductName_val, ebProductSeries_val))
'        Call fGotoCell(shtSelfInventory.Range("A2"), True)
'    Else
'        fRemoveFilterForSheet shtSelfInventory
'    End If
'
'    fShowActivateSheet shtSelfInventory
End Sub
Private Sub btnSelfPurchase_Click(control As IRibbonControl)
    Call Sub_SearchOnCurrentSheetBySearchBy(shtSelfPurchaseOrder)
'    Call fRefreshEditBoxFromShtDataStage
'    If Len(ebProductProducer_val) > 0 Then
'        'Call fSetFilterForSheet(shtSelfPurchaseOrder, Array(SelfPurchase.ProductProducer, SelfPurchase.ProductName, SelfPurchase.ProductSeries, SelfPurchase.LotNum) _
'                , Array(ebProductProducer_val, ebProductName_val, ebProductSeries_val, ebLotNum_val))
'        Call fSetFilterForSheet(shtSelfPurchaseOrder, Array(SelfPurchase.ProductProducer, SelfPurchase.ProductName, SelfPurchase.ProductSeries) _
'                , Array(ebProductProducer_val, ebProductName_val, ebProductSeries_val))
'        Call fGotoCell(shtSelfPurchaseOrder.Range("A2"), True)
'    Else
'        fRemoveFilterForSheet shtSelfPurchaseOrder
'    End If
'
'    fShowActivateSheet shtSelfPurchaseOrder
End Sub
Private Sub btnSelfSales_Click(control As IRibbonControl)
    Call Sub_SearchOnCurrentSheetBySearchBy(shtSelfSalesOrder)
'    Call fRefreshEditBoxFromShtDataStage
'    If Len(ebProductProducer_val) > 0 Then
''        Call fSetFilterForSheet(shtSelfSalesOrder, Array(SelfSales.ProductProducer, SelfSales.ProductName, SelfSales.ProductSeries, SelfSales.LotNum) _
'                , Array(ebProductProducer_val, ebProductName_val, ebProductSeries_val, ebLotNum_val))
'        Call fSetFilterForSheet(shtSelfSalesOrder, Array(SelfSales.ProductProducer, SelfSales.ProductName, SelfSales.ProductSeries) _
'                , Array(ebProductProducer_val, ebProductName_val, ebProductSeries_val))
'        Call fGotoCell(shtSelfSalesOrder.Range("A2"), True)
'    Else
'        fRemoveFilterForSheet shtSelfSalesOrder
'    End If
'
'    fShowActivateSheet shtSelfSalesOrder
End Sub
Private Sub btnSCompInvImported_Click()
    Call Sub_SearchOnCurrentSheetBySearchBy(shtSalesCompInvUnified)
'    Call fRefreshEditBoxFromShtDataStage
'    If Len(ebProductProducer_val) > 0 Then
''        Call fSetFilterForSheet(shtSalesCompInvUnified, Array(SCompUnifiedInv.SalesCompany, SCompUnifiedInv.ProductProducer, SCompUnifiedInv.ProductName, SCompUnifiedInv.ProductSeries, SCompUnifiedInv.LotNum) _
'                , Array(ebSalesCompany_val, ebProductProducer_val, ebProductName_val, ebProductSeries_val, ebLotNum_val))
'        Call fSetFilterForSheet(shtSalesCompInvUnified, Array(SCompUnifiedInv.SalesCompany, SCompUnifiedInv.ProductProducer, SCompUnifiedInv.ProductName, SCompUnifiedInv.ProductSeries) _
'                , Array(ebSalesCompany_val, ebProductProducer_val, ebProductName_val, ebProductSeries_val))
'        Call fGotoCell(shtSalesCompInvUnified.Range("A2"), True)
'    Else
'        fRemoveFilterForSheet shtSalesCompInvUnified
'    End If
'
'    fShowActivateSheet shtSalesCompInvUnified
End Sub
Private Sub btnPromotionProduct_Click()
    Call Sub_SearchOnCurrentSheetBySearchBy(shtPromotionProduct)
'    Call fRefreshEditBoxFromShtDataStage
'    If Len(ebProductProducer_val) > 0 Then
'        Call fSetFilterForSheet(shtPromotionProduct, Array(PromoteProduct.SalesCompany, SecondLevelComm.Hospital, PromoteProduct.ProductProducer, PromoteProduct.ProductName, PromoteProduct.ProductSeries) _
'                , Array(ebSalesCompany_val, ebHospital_val, ebProductProducer_val, ebProductName_val, ebProductSeries_val))
'        Call fGotoCell(shtPromotionProduct.Range("A2"), True)
'    Else
'        fRemoveFilterForSheet shtPromotionProduct
'    End If
'
'    fShowActivateSheet shtPromotionProduct
End Sub
Private Sub btnFirstLevelComm_Click()
    Call Sub_SearchOnCurrentSheetBySearchBy(shtFirstLevelCommission)
'    Call fRefreshEditBoxFromShtDataStage
'    If Len(ebProductProducer_val) > 0 Then
'        Call fSetFilterForSheet(shtFirstLevelCommission, Array(FirstLevelComm.SalesCompany, FirstLevelComm.ProductProducer, FirstLevelComm.ProductName, FirstLevelComm.ProductSeries) _
'                , Array(ebSalesCompany_val, ebProductProducer_val, ebProductName_val, ebProductSeries_val))
'        Call fGotoCell(shtFirstLevelCommission.Range("A2"), True)
'    Else
'        fRemoveFilterForSheet shtFirstLevelCommission
'    End If
'
'    fShowActivateSheet shtFirstLevelCommission
End Sub
Private Sub btnSecondLevelComm_Click()
    Call Sub_SearchOnCurrentSheetBySearchBy(shtSecondLevelCommission)
    
'    Call fRefreshEditBoxFromShtDataStage
'    If Len(ebProductProducer_val) > 0 Then
'        Call fSetFilterForSheet(shtSecondLevelCommission, Array(SecondLevelComm.SalesCompany, SecondLevelComm.Hospital, SecondLevelComm.ProductProducer, SecondLevelComm.ProductName, SecondLevelComm.ProductSeries) _
'                , Array(ebSalesCompany_val, ebHospital_val, ebProductProducer_val, ebProductName_val, ebProductSeries_val))
'
'        Call fGotoCell(shtSecondLevelCommission.Range("A2"), True)
'    Else
'        fRemoveFilterForSheet shtSecondLevelCommission
'    End If
'
'    fShowActivateSheet shtSecondLevelCommission
End Sub
Private Sub btnSalePriceInAdv_Click()
    Call Sub_SearchOnCurrentSheetBySearchBy(shtSellPriceInAdv)
'    'todo
'    Call fRefreshEditBoxFromShtDataStage
'    If Len(ebProductProducer_val) > 0 Then
'        Call fSetFilterForSheet(shtSellPriceInAdv, Array(SellPriceInAdv.SalesCompany, SellPriceInAdv.ProductProducer, SellPriceInAdv.ProductName, SellPriceInAdv.ProductSeries) _
'                , Array(ebSalesCompany_val, ebProductProducer_val, ebProductName_val, ebProductSeries_val))
'
'        Call fGotoCell(shtSellPriceInAdv.Range("A2"), True)
'    Else
'        fRemoveFilterForSheet shtSellPriceInAdv
'    End If
'
'    fShowActivateSheet shtSellPriceInAdv
End Sub

Private Sub btnSalesInfo_Click()
    Call Sub_SearchOnCurrentSheetBySearchBy(shtSalesInfos)
'    Call fRefreshEditBoxFromShtDataStage
'    If Len(ebProductProducer_val) > 0 Then
'        Call fSetFilterForSheet(shtSalesInfos, Array(Sales2Hospital.SalesCompany, Sales2Hospital.Hospital, Sales2Hospital.ProductProducer, Sales2Hospital.ProductName, Sales2Hospital.ProductSeries, Sales2Hospital.LotNum) _
'                , Array(ebSalesCompany_val, ebHospital_val, ebProductProducer_val, ebProductName_val, ebProductSeries_val, ebLotNum_val))
'
'        Call fGotoCell(shtSalesInfos.Range("A2"), True)
'    Else
'        fRemoveFilterForSheet shtSalesInfos
'    End If
'
'    fShowActivateSheet shtSalesInfos
End Sub

Private Sub btnSCompInvDiff_Click(control As IRibbonControl)
    Call Sub_SearchOnCurrentSheetBySearchBy(shtSalesCompInvDiff)
'    Call fRefreshEditBoxFromShtDataStage
'    If Len(ebProductProducer_val) > 0 Then
'        Call fSetFilterForSheet(shtSalesCompInvDiff, Array(SCompInvDiff.SalesCompany, SCompInvDiff.ProductProducer, SCompInvDiff.ProductName, SCompInvDiff.ProductSeries) _
'                , Array(ebSalesCompany_val, ebProductProducer_val, ebProductName_val, ebProductSeries_val))
'        Call fGotoCell(shtSalesCompInvDiff.Range("A2"), True)
'    Else
'        fRemoveFilterForSheet shtSalesCompInvDiff
'       ' Call fSetFilterForSheet(shtSalesCompInvDiff, Array(CZLInvDiff.SalesCompany), Array(sCZLCompName))
'    End If
'
'    fShowActivateSheet shtSalesCompInvDiff
End Sub
Private Sub tbSCompInvCalcd_Click(control As IRibbonControl)
    Call Sub_SearchOnCurrentSheetBySearchBy(shtSalesCompInvCalcd)
'
'    Call fRefreshEditBoxFromShtDataStage
'    If Len(ebProductProducer_val) > 0 Then
'        Call fSetFilterForSheet(shtSalesCompInvCalcd, Array(SCompInvCalcd.SalesCompany, SCompInvCalcd.ProductProducer, SCompInvCalcd.ProductName, SCompInvCalcd.ProductSeries) _
'                , Array(ebSalesCompany_val, ebProductProducer_val, ebProductName_val, ebProductSeries_val))
'        Call fGotoCell(shtSalesCompInvCalcd.Range("A2"), True)
'    Else
'        fRemoveFilterForSheet shtSalesCompInvCalcd
'    End If
'
'    fShowActivateSheet shtSalesCompInvCalcd
End Sub

Private Sub btnSCompRolloverInv_Click(control As IRibbonControl)
    
    Call Sub_SearchOnCurrentSheetBySearchBy(shtSalesCompRolloverInv)
    
'    Call fRefreshEditBoxFromShtDataStage
'    If Len(ebProductProducer_val) > 0 Then
'        Call fSetFilterForSheet(shtSalesCompRolloverInv, Array(SCompRollover.SalesCompany, SCompRollover.ProductProducer, SCompRollover.ProductName, SCompRollover.ProductSeries) _
'                , Array(ebSalesCompany_val, ebProductProducer_val, ebProductName_val, ebProductSeries_val))
'        Call fGotoCell(shtSalesCompRolloverInv.Range("A2"), True)
'    Else
'        fRemoveFilterForSheet shtSalesCompRolloverInv
'    End If
'
'    fShowActivateSheet shtSalesCompRolloverInv
End Sub
Private Sub btnSCompPurchase_Click(control As IRibbonControl)
    Call Sub_SearchOnCurrentSheetBySearchBy(shtCZLSales2Companies)
'    Call fRefreshEditBoxFromShtDataStage
'    If Len(ebProductProducer_val) > 0 Then
'        Call fSetFilterForSheet(shtCZLSales2Companies, Array(CZLSales2Comp.SalesCompany, CZLSales2Comp.ProductProducer, CZLSales2Comp.ProductName, CZLSales2Comp.ProductSeries) _
'                , Array(ebSalesCompany_val, ebProductProducer_val, ebProductName_val, ebProductSeries_val))
'        Call fGotoCell(shtCZLSales2Companies.Range("A2"), True)
'    Else
'        fRemoveFilterForSheet shtCZLSales2Companies
'    End If
'
'    fShowActivateSheet shtCZLSales2Companies
End Sub
Private Sub btnSCompSalesToHospital_Click(control As IRibbonControl)
    Call Sub_SearchOnCurrentSheetBySearchBy(shtSalesInfos)
'    Call fRefreshEditBoxFromShtDataStage
'    If Len(ebProductProducer_val) > 0 Then
'        Call fSetFilterForSheet(shtSalesInfos, Array(Sales2Hospital.SalesCompany, Sales2Hospital.ProductProducer, Sales2Hospital.ProductName, Sales2Hospital.ProductSeries, Sales2Hospital.Hospital) _
'                , Array(ebSalesCompany_val, ebProductProducer_val, ebProductName_val, ebProductSeries_val, ebHospital_val))
'        Call fGotoCell(shtSalesInfos.Range("A2"), True)
'    Else
'        fRemoveFilterForSheet shtSalesInfos
'    End If
'
'    fShowActivateSheet shtSalesInfos
End Sub

Private Sub btnCZLInvDiff_Click(control As IRibbonControl)
    Call Sub_SearchOnCurrentSheetBySearchBy(shtCZLInvDiff)
'    Call fRefreshEditBoxFromShtDataStage
'    Dim sCZLCompName As String
'    sCZLCompName = fGetCompanyNameByID_Common("CZL")
'
'    If Len(ebProductProducer_val) > 0 Then
'        Call fSetFilterForSheet(shtCZLInvDiff, Array(CZLInvDiff.ProductProducer, CZLInvDiff.ProductName, CZLInvDiff.ProductSeries) _
'                , Array(ebProductProducer_val, ebProductName_val, ebProductSeries_val))
'        Call fGotoCell(shtCZLInvDiff.Range("A2"), True)
'    Else
'        fRemoveFilterForSheet shtCZLInvDiff
'       ' Call fSetFilterForSheet(shtCZLInvDiff, Array(CZLInvDiff.SalesCompany), Array(sCZLCompName))
'    End If
'
'    fShowActivateSheet shtCZLInvDiff
End Sub
Private Sub btnCZLInvImported_Click(control As IRibbonControl)
    Call Sub_SearchOnCurrentSheetBySearchBy(shtSalesCompInvUnified)
'    Call fRefreshEditBoxFromShtDataStage
'    Dim sCZLCompName As String
'    sCZLCompName = fGetCompanyNameByID_Common("CZL")
'
'    If Len(ebProductProducer_val) > 0 Then
'        Call fSetFilterForSheet(shtSalesCompInvUnified, Array(SCompUnifiedInv.SalesCompany, SCompUnifiedInv.ProductProducer, SCompUnifiedInv.ProductName, SCompUnifiedInv.ProductSeries, SCompUnifiedInv.LotNum) _
'                , Array(sCZLCompName, ebProductProducer_val, ebProductName_val, ebProductSeries_val, ebLotNum_val))
'        Call fGotoCell(shtSalesCompInvUnified.Range("A2"), True)
'    Else
'        fRemoveFilterForSheet shtSalesCompInvUnified
'        Call fSetFilterForSheet(shtSalesCompInvUnified, Array(SCompUnifiedInv.SalesCompany), Array(sCZLCompName))
'    End If
'
'    fShowActivateSheet shtSalesCompInvUnified
End Sub
 
Private Sub tbCZLInvCalcd_Click(control As IRibbonControl)
    Call Sub_SearchOnCurrentSheetBySearchBy(shtCZLInventory)
'    Call fRefreshEditBoxFromShtDataStage
'    If Len(ebProductProducer_val) > 0 Then
'        Call fSetFilterForSheet(shtCZLInventory, Array(CZLInv.ProductProducer, CZLInv.ProductName, CZLInv.ProductSeries) _
'                , Array(ebProductProducer_val, ebProductName_val, ebProductSeries_val))
'        Call fGotoCell(shtCZLInventory.Range("A2"), True)
'    Else
'        fRemoveFilterForSheet shtCZLInventory
'    End If
'
'    fShowActivateSheet shtCZLInventory
End Sub


Private Sub btnCZLSalesToSCompAll_Click(control As IRibbonControl)
    Call Sub_SearchOnCurrentSheetBySearchBy(shtCZLSales2SCompAll)
'    Call fRefreshEditBoxFromShtDataStage
'    If Len(ebProductProducer_val) > 0 Then
'        Call fSetFilterForSheet(shtCZLSales2SCompAll, Array(CZLSales2CompHist.SalesCompany, CZLSales2CompHist.ProductProducer, CZLSales2CompHist.ProductName, CZLSales2CompHist.ProductSeries) _
'                , Array(ebSalesCompany_val, ebProductProducer_val, ebProductName_val, ebProductSeries_val))
'        Call fGotoCell(shtCZLSales2SCompAll.Range("A2"), True)
'    Else
'        fRemoveFilterForSheet shtCZLSales2SCompAll
'    End If
'
'    fShowActivateSheet shtCZLSales2SCompAll
End Sub
Private Sub btnCZLSalesToSComp_Click(control As IRibbonControl)
    Call Sub_SearchOnCurrentSheetBySearchBy(shtCZLSales2Companies)
'    Call fRefreshEditBoxFromShtDataStage
'    If Len(ebProductProducer_val) > 0 Then
'        Call fSetFilterForSheet(shtCZLSales2Companies, Array(CZLSales2Comp.SalesCompany, CZLSales2Comp.ProductProducer, CZLSales2Comp.ProductName, CZLSales2Comp.ProductSeries) _
'                , Array(ebSalesCompany_val, ebProductProducer_val, ebProductName_val, ebProductSeries_val))
'        Call fGotoCell(shtCZLSales2Companies.Range("A2"), True)
'    Else
'        fRemoveFilterForSheet shtCZLSales2Companies
'    End If
'
'    fShowActivateSheet shtCZLSales2Companies
End Sub
Private Sub btnCZLSalesToHospital_Click(control As IRibbonControl)
    Call Sub_SearchOnCurrentSheetBySearchBy(shtSalesInfos)
'    Call fRefreshEditBoxFromShtDataStage
'    Dim sCZLName As String
'
'    If Len(ebProductProducer_val) > 0 Then
'        sCZLName = fGetCompanyNameByID_Common("CZL")
'
'        Call fSetFilterForSheet(shtSalesInfos, Array(Sales2Hospital.SalesCompany, Sales2Hospital.ProductProducer, Sales2Hospital.ProductName, Sales2Hospital.ProductSeries) _
'                , Array(sCZLName, ebProductProducer_val, ebProductName_val, ebProductSeries_val))
'        Call fGotoCell(shtSalesInfos.Range("A2"), True)
'    Else
'        fRemoveFilterForSheet shtSalesInfos
'    End If
'
'    fShowActivateSheet shtSalesInfos
End Sub

Private Sub btnCZLRolloverInv_Click(control As IRibbonControl)
    Call Sub_SearchOnCurrentSheetBySearchBy(shtCZLRolloverInv)
'    Call fRefreshEditBoxFromShtDataStage
'    If Len(ebProductProducer_val) > 0 Then
'        Call fSetFilterForSheet(shtCZLRolloverInv, Array(CZLRollover.ProductProducer, CZLRollover.ProductName, CZLRollover.ProductSeries) _
'                , Array(ebProductProducer_val, ebProductName_val, ebProductSeries_val))
'        Call fGotoCell(shtCZLRolloverInv.Range("A2"), True)
'    Else
'        fRemoveFilterForSheet shtCZLRolloverInv
'    End If
'
'    fShowActivateSheet shtCZLRolloverInv
End Sub

Private Sub btnCZLPurchase_Click(control As IRibbonControl)
    Call Sub_SearchOnCurrentSheetBySearchBy(shtCZLPurchaseOrder)
'    Call fRefreshEditBoxFromShtDataStage
'    fPrepareCZLPurchaseFromSelfSales
'
'    If Len(ebProductProducer_val) > 0 Then
'        Call fSetFilterForSheet(shtCZLPurchaseOrder, Array(SelfSales.ProductProducer, SelfSales.ProductName, SelfSales.ProductSeries) _
'                , Array(ebProductProducer_val, ebProductName_val, ebProductSeries_val))
'        Call fGotoCell(shtCZLPurchaseOrder.Range("A2"), True)
'    Else
'        fRemoveFilterForSheet shtCZLPurchaseOrder
'    End If
'
'    fShowActivateSheet shtCZLPurchaseOrder
End Sub

Private Sub btnProfit_Click(control As IRibbonControl)
    Call Sub_SearchOnCurrentSheetBySearchBy(shtProfit)
'    Call fRefreshEditBoxFromShtDataStage
'    If Len(ebProductProducer_val) > 0 Then
'        Call fSetFilterForSheet(shtProfit, Array(Profit.SalesCompany, Profit.Hospital, Profit.ProductProducer, Profit.ProductName, Profit.ProductSeries) _
'                , Array(ebSalesCompany_val, ebHospital_val, ebProductProducer_val, ebProductName_val, ebProductSeries_val))
'        Call fGotoCell(shtProfit.Range("A2"), True)
'    Else
'        fRemoveFilterForSheet shtProfit
'    End If
'
'    fShowActivateSheet shtProfit
End Sub
Private Sub btnRefund_Click(control As IRibbonControl)
    Call Sub_SearchOnCurrentSheetBySearchBy(shtRefund)
'    Call fRefreshEditBoxFromShtDataStage
'    If Len(ebProductProducer_val) > 0 Then
'        Call fSetFilterForSheet(shtRefund, Array(Refund.SalesCompany, Refund.Hospital, Refund.ProductProducer, Refund.ProductName, Refund.ProductSeries) _
'                , Array(ebSalesCompany_val, ebHospital_val, ebProductProducer_val, ebProductName_val, ebProductSeries_val))
'        Call fGotoCell(shtRefund.Range("A2"), True)
'    Else
'        fRemoveFilterForSheet shtRefund
'    End If
'
'    fShowActivateSheet shtRefund
End Sub
Private Sub btnProductNameReplace_Click(control As IRibbonControl)
    Call Sub_SearchOnCurrentSheetBySearchBy(shtProductNameReplace)
'    Call fRefreshEditBoxFromShtDataStage
'    If Len(ebProductProducer_val) > 0 Then
'        Call fSetFilterForSheet(shtProductNameReplace, Array(ProdNameReplace.ProductProducer, ProdNameReplace.ProductName) _
'                , Array(ebProductProducer_val, ebProductName_val))
'        Call fGotoCell(shtProductNameReplace.Range("A2"), True)
'    Else
'        fRemoveFilterForSheet shtProductNameReplace
'    End If
'
'    fShowActivateSheet shtProductNameReplace
End Sub
Private Sub btnProductSeriesReplace_Click(control As IRibbonControl)
    Call Sub_SearchOnCurrentSheetBySearchBy(shtProductSeriesReplace)
'    Call fRefreshEditBoxFromShtDataStage
'    If Len(ebProductProducer_val) > 0 Then
'        Call fSetFilterForSheet(shtProductSeriesReplace, Array(ProdSerReplace.ProductProducer, ProdSerReplace.ProductName, ProdSerReplace.ProductSeries) _
'                , Array(ebProductProducer_val, ebProductName_val, ebProductSeries_val))
'        Call fGotoCell(shtProductSeriesReplace.Range("A2"), True)
'    Else
'        fRemoveFilterForSheet shtProductSeriesReplace
'    End If
'
'    fShowActivateSheet shtProductSeriesReplace
End Sub
Private Sub btnProductMaster_Click(control As IRibbonControl)
    Call Sub_SearchOnCurrentSheetBySearchBy(shtProductMaster)
'    Call fRefreshEditBoxFromShtDataStage
'    If Len(ebProductProducer_val) > 0 Then
'        Call fSetFilterForSheet(shtProductMaster, Array(ProductMst.ProductProducer, ProductMst.ProductName, ProductMst.ProductSeries) _
'                , Array(ebProductProducer_val, ebProductName_val, ebProductSeries_val))
'        Call fGotoCell(shtProductMaster.Range("A2"), True)
'    Else
'        fRemoveFilterForSheet shtProductMaster
'    End If
'
'    fShowActivateSheet shtProductMaster
End Sub
