Attribute VB_Name = "MB_1_RibbonUI"
Option Explicit

#If VBA7 And Win64 Then  'Win64
    Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
#Else
    Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As Long)
#End If

Public mRibbonObj As IRibbonUI
Private ebSalesCompany_val As String
Private ebProductProducer_val As String
Private ebProductName_val As String
Private ebProductSeries_val As String
Private ebLotNum_val As String
Private ebHospital_val As String

'=============================================================
Sub subRefreshRibbon()
    fGetRibbonReference.Invalidate
End Sub
Sub ERP_UI_Onload(ribbon As IRibbonUI)
  Set mRibbonObj = ribbon
  
  fCreateAddNameUpdateNameWhenExists "nmRibbonPointer", ObjPtr(ribbon)
  'Names("nmRibbonPointer").RefersTo = ObjPtr(ribbon)
  
  mRibbonObj.ActivateTab "ERP_2010"
  ThisWorkbook.Saved = True
End Sub
Function fGetRibbonReference() As IRibbonUI
    If Not mRibbonObj Is Nothing Then Set fGetRibbonReference = mRibbonObj: Exit Function
    
    Dim objRibbon As Object
    Dim lRibPointer As LongPtr
    
    lRibPointer = [nmRibbonPointer]
    
    CopyMemory objRibbon, lRibPointer, LenB(lRibPointer)
    
    Set fGetRibbonReference = objRibbon
    Set mRibbonObj = objRibbon
    Set objRibbon = Nothing
End Function
'---------------------------------------------------------------------
Sub Button_onAction(control As IRibbonControl)
    Call fGetControlAttributes(control, "ACTION")
End Sub
Sub Button_getImage(control As IRibbonControl, ByRef imageMso)
    Call fGetControlAttributes(control, "IMAGE", imageMso)
End Sub
Sub Button_getLabel(control As IRibbonControl, ByRef label)
    Call fGetControlAttributes(control, "LABEL", label)
End Sub
Sub Button_getSize(control As IRibbonControl, ByRef Size)
    Call fGetControlAttributes(control, "SIZE", Size)
End Sub

'================== toggle button common function===========================================
Sub ToggleButtonToSwitchSheet_onAction(control As IRibbonControl, pressed As Boolean)
    Dim sht As Worksheet
    Set sht = fGetSheetByUIRibbonTag(control.Tag)
    
    If Not sht Is Nothing Then
        fToggleSheetVisibleFromUIRibbonControl pressed, sht, control
    End If
    Set sht = Nothing
End Sub

Sub ToggleButtonToSwitchSheet_getPressed(control As IRibbonControl, ByRef returnedVal)
    Dim sht As Worksheet
    Set sht = fGetSheetByUIRibbonTag(control.Tag)
    
    If sht Is Nothing Then
        returnedVal = False
    Else
        returnedVal = (sht.Visible = xlSheetVisible And ActiveSheet Is sht)
    End If
End Sub
Function fGetSheetByUIRibbonTag(ByVal asButtonTag As String) As Worksheet
    Dim sht As Worksheet
    
    If fSheetExistsByCodeName(asButtonTag, sht) Then
        Set fGetSheetByUIRibbonTag = sht
    Else
        MsgBox "The button's Tag is not corresponding to any worksheet in this workbook, please check the customUI.xml you prepared," _
            & " The design thought is that the button's tag is the name of a sheet, so that the common function ToggleButtonToSwitchSheet_onAction/getPressed can get a worksheet."
    End If
    Set sht = Nothing
End Function
Function fToggleSheetVisibleFromUIRibbonControl(ByVal pressed As Boolean, sht As Worksheet, control As IRibbonControl)
    If pressed Then
        If ActiveSheet.CodeName <> sht.CodeName Then
            fActiveVisibleSwitchSheet sht
        End If
    Else
        If ActiveSheet.CodeName <> sht.CodeName Then
            fActiveVisibleSwitchSheet sht
        Else
            fVeryHideSheet sht
        End If
    End If
    
    'fGetRibbonReference.InvalidateControl (control.id)
    fGetRibbonReference.Invalidate
End Function

'---------------------------------------------------------------------

'==========================dev prod switch===================================
Sub btnSwitchDevProd_onAction(control As IRibbonControl, pressed As Boolean)
    sub_SwitchDevProdMode
End Sub

Sub btnSwitchDevProd_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = fIsDev()
End Sub
Sub btnSwitchDevProd_getVisible(control As IRibbonControl, ByRef returnedVal)
    'returnedVal = fIsDev()
    returnedVal = True
End Sub
Sub grpDevFacilities_getVisible(control As IRibbonControl, ByRef returnedVal)
    returnedVal = fIsDev()
End Sub
'---------------------------------------------------------------------

'================ dev facilities ==============================================
Sub btnListAllFunctions_onAction(control As IRibbonControl)
    sub_ListAllFunctionsOfThisWorkbook
End Sub
Sub btnExportSourceCode_onAction(control As IRibbonControl)
    sub_ExportModulesSourceCodeToFolder
End Sub
Sub btnGenNumberList_onAction(control As IRibbonControl)
    sub_GenNumberList
End Sub
Sub btnGenAlphabetList_onAction(control As IRibbonControl)
    sub_GenAlpabetList
End Sub
Sub btnListAllActiveXOnCurrSheet_onAction(control As IRibbonControl)
    Sub_ListActiveXControlOnActiveSheet
End Sub
Sub btnResetOnError_onAction(control As IRibbonControl)
    sub_ResetOnError_Initialize
End Sub
'------------------------------------------------------------------------------

Sub EditBox_onChange(control As IRibbonControl, text As String)
    Select Case control.id
        Case "ebSalesCompany"
            ebSalesCompany_val = text
            shtDataStage.Range("K1").Value = ebSalesCompany_val
        Case "ebProductProducer"
            ebProductProducer_val = text
            shtDataStage.Range("K2").Value = ebProductProducer_val
        Case "ebProductName"
            ebProductName_val = text
            shtDataStage.Range("K3").Value = ebProductName_val
        Case "ebProductSeries"
            ebProductSeries_val = text
            shtDataStage.Range("K4").Value = ebProductSeries_val
        Case "ebLotnum"
            ebLotNum_val = text
            shtDataStage.Range("K5").Value = ebLotNum_val
        Case "ebHospital"
            ebHospital_val = text
            shtDataStage.Range("K6").Value = ebHospital_val
        Case Else
    End Select
    
'    MsgBox "商业公司 ebSalesCompany： " & vbTab & ebSalesCompany_val _
'        & vbCr & "生产厂家 ebProductProducer： " & vbTab & ebProductProducer_val _
'        & vbCr & "药品名称 ebProductName： " & vbTab & ebProductName_val _
'        & vbCr & "药品规格 ebProductSeries： " & vbTab & ebProductSeries_val _
'        & vbCr & "批号 ebLotNum_val： " & vbTab & ebLotNum_val _
'        & vbCr & "医院ebHospital_val： " & vbTab & ebHospital_val
End Sub

Private Function fRefreshEditBoxFromShtDataStage()
    If Len(ebProductProducer_val) <= 0 And Len(ebProductName_val) <= 0 And Len(ebProductSeries_val) <= 0 Then
        ebSalesCompany_val = Trim(shtDataStage.Range("K1").Value)
        ebProductProducer_val = Trim(shtDataStage.Range("K2").Value)
        ebProductName_val = Trim(shtDataStage.Range("K3").Value)
        ebProductSeries_val = Trim(shtDataStage.Range("K4").Value)
        ebLotNum_val = Trim(shtDataStage.Range("K5").Value)
        ebHospital_val = Trim(shtDataStage.Range("K6").Value)
    End If
End Function

Private Sub EditBox_getText(control As IRibbonControl, ByRef returnedVal)
    Select Case control.id
        Case "ebSalesCompany"
            Call fGetCurrSheetCurrRowSalesCompany(ebSalesCompany_val)
            returnedVal = ebSalesCompany_val
            shtDataStage.Range("K1").Value = ebSalesCompany_val
        Case "ebProductProducer"
            Call fGetCurrSheetCurrRowProductProducer(ebProductProducer_val)
            returnedVal = ebProductProducer_val
            shtDataStage.Range("K2").Value = ebProductProducer_val
        Case "ebProductName"
            Call fGetCurrSheetCurrRowProductName(ebProductName_val)
            returnedVal = ebProductName_val
            shtDataStage.Range("K3").Value = ebProductName_val
        Case "ebProductSeries"
            Call fGetCurrSheetCurrRowProductSeries(ebProductSeries_val)
            returnedVal = ebProductSeries_val
            shtDataStage.Range("K4").Value = ebProductSeries_val
        Case "ebLotnum"
            Call fGetCurrSheetCurrRowLotNum(ebLotNum_val)
            returnedVal = ebLotNum_val
            shtDataStage.Range("K5").Value = ebLotNum_val
        Case "ebHospital"
            Call fGetCurrSheetCurrRowHospital(ebHospital_val)
            returnedVal = ebHospital_val
            shtDataStage.Range("K6").Value = ebHospital_val
        Case Else
    End Select
End Sub

'Function fPresstgSearchBy(bPressed As Boolean)
Function fPresstgSearchBy()
    'If bPressed Then
        fGetRibbonReference.InvalidateControl "ebSalesCompany"
        fGetRibbonReference.InvalidateControl "ebProductProducer"
        fGetRibbonReference.InvalidateControl "ebProductName"
        fGetRibbonReference.InvalidateControl "ebProductSeries"
        fGetRibbonReference.InvalidateControl "ebLotnum"
        fGetRibbonReference.InvalidateControl "ebHospital"
'    Else
'    End If
End Function
Private Sub UIbtnHome(control As IRibbonControl)
    Call Sub_ToHomeSheet
End Sub
Private Sub btnSelfInventory_Click(control As IRibbonControl)
    'fPresstgSearchBy
    Call fRefreshEditBoxFromShtDataStage
    
    If Len(ebProductProducer_val) > 0 Then
'        Call fSetFilterForSheet(shtSelfInventory, Array(SelfInv.ProductProducer, SelfInv.ProductName, SelfInv.ProductSeries, SelfInv.LotNum) _
                , Array(ebProductProducer_val, ebProductName_val, ebProductSeries_val, ebLotNum_val))
        Call fSetFilterForSheet(shtSelfInventory, Array(SelfInv.ProductProducer, SelfInv.ProductName, SelfInv.ProductSeries) _
                , Array(ebProductProducer_val, ebProductName_val, ebProductSeries_val))
        Call fGotoCell(shtSelfInventory.Range("A2"), True)
    Else
        fRemoveFilterForSheet shtSelfInventory
    End If
    
    fShowActivateSheet shtSelfInventory
End Sub
Private Sub btnSelfPurchase_Click(control As IRibbonControl)
    Call fRefreshEditBoxFromShtDataStage
    If Len(ebProductProducer_val) > 0 Then
        'Call fSetFilterForSheet(shtSelfPurchaseOrder, Array(SelfPurchase.ProductProducer, SelfPurchase.ProductName, SelfPurchase.ProductSeries, SelfPurchase.LotNum) _
                , Array(ebProductProducer_val, ebProductName_val, ebProductSeries_val, ebLotNum_val))
        Call fSetFilterForSheet(shtSelfPurchaseOrder, Array(SelfPurchase.ProductProducer, SelfPurchase.ProductName, SelfPurchase.ProductSeries) _
                , Array(ebProductProducer_val, ebProductName_val, ebProductSeries_val))
        Call fGotoCell(shtSelfPurchaseOrder.Range("A2"), True)
    Else
        fRemoveFilterForSheet shtSelfPurchaseOrder
    End If
    
    fShowActivateSheet shtSelfPurchaseOrder
End Sub
Private Sub btnSelfSales_Click(control As IRibbonControl)
    Call fRefreshEditBoxFromShtDataStage
    If Len(ebProductProducer_val) > 0 Then
'        Call fSetFilterForSheet(shtSelfSalesOrder, Array(SelfSales.ProductProducer, SelfSales.ProductName, SelfSales.ProductSeries, SelfSales.LotNum) _
                , Array(ebProductProducer_val, ebProductName_val, ebProductSeries_val, ebLotNum_val))
        Call fSetFilterForSheet(shtSelfSalesOrder, Array(SelfSales.ProductProducer, SelfSales.ProductName, SelfSales.ProductSeries) _
                , Array(ebProductProducer_val, ebProductName_val, ebProductSeries_val))
        Call fGotoCell(shtSelfSalesOrder.Range("A2"), True)
    Else
        fRemoveFilterForSheet shtSelfSalesOrder
    End If
    
    fShowActivateSheet shtSelfSalesOrder
End Sub
Private Sub btnSCompInvImported_Click()
    Call fRefreshEditBoxFromShtDataStage
    If Len(ebProductProducer_val) > 0 Then
'        Call fSetFilterForSheet(shtSalesCompInvUnified, Array(SCompUnifiedInv.SalesCompany, SCompUnifiedInv.ProductProducer, SCompUnifiedInv.ProductName, SCompUnifiedInv.ProductSeries, SCompUnifiedInv.LotNum) _
                , Array(ebSalesCompany_val, ebProductProducer_val, ebProductName_val, ebProductSeries_val, ebLotNum_val))
        Call fSetFilterForSheet(shtSalesCompInvUnified, Array(SCompUnifiedInv.SalesCompany, SCompUnifiedInv.ProductProducer, SCompUnifiedInv.ProductName, SCompUnifiedInv.ProductSeries) _
                , Array(ebSalesCompany_val, ebProductProducer_val, ebProductName_val, ebProductSeries_val))
        Call fGotoCell(shtSalesCompInvUnified.Range("A2"), True)
    Else
        fRemoveFilterForSheet shtSalesCompInvUnified
    End If
    
    fShowActivateSheet shtSalesCompInvUnified
End Sub
Private Sub btnPromotionProduct_Click()
    Call fRefreshEditBoxFromShtDataStage
    If Len(ebProductProducer_val) > 0 Then
        Call fSetFilterForSheet(shtPromotionProduct, Array(PromoteProduct.SalesCompany, SecondLevelComm.Hospital, PromoteProduct.ProductProducer, PromoteProduct.ProductName, PromoteProduct.ProductSeries) _
                , Array(ebSalesCompany_val, ebHospital_val, ebProductProducer_val, ebProductName_val, ebProductSeries_val))
        Call fGotoCell(shtPromotionProduct.Range("A2"), True)
    Else
        fRemoveFilterForSheet shtPromotionProduct
    End If
    
    fShowActivateSheet shtPromotionProduct
End Sub
Private Sub btnFirstLevelComm_Click()
    Call fRefreshEditBoxFromShtDataStage
    If Len(ebProductProducer_val) > 0 Then
        Call fSetFilterForSheet(shtFirstLevelCommission, Array(FirstLevelComm.SalesCompany, FirstLevelComm.ProductProducer, FirstLevelComm.ProductName, FirstLevelComm.ProductSeries) _
                , Array(ebSalesCompany_val, ebProductProducer_val, ebProductName_val, ebProductSeries_val))
        Call fGotoCell(shtFirstLevelCommission.Range("A2"), True)
    Else
        fRemoveFilterForSheet shtFirstLevelCommission
    End If
    
    fShowActivateSheet shtFirstLevelCommission
End Sub
Private Sub btnSecondLevelComm_Click()
    Call fRefreshEditBoxFromShtDataStage
    If Len(ebProductProducer_val) > 0 Then
        Call fSetFilterForSheet(shtSecondLevelCommission, Array(SecondLevelComm.SalesCompany, SecondLevelComm.Hospital, SecondLevelComm.ProductProducer, SecondLevelComm.ProductName, SecondLevelComm.ProductSeries) _
                , Array(ebSalesCompany_val, ebHospital_val, ebProductProducer_val, ebProductName_val, ebProductSeries_val))
    
        Call fGotoCell(shtSecondLevelCommission.Range("A2"), True)
    Else
        fRemoveFilterForSheet shtSecondLevelCommission
    End If
    
    fShowActivateSheet shtSecondLevelCommission
End Sub
Private Sub btnSalePriceInAdv_Click()
    'todo
    Call fRefreshEditBoxFromShtDataStage
    If Len(ebProductProducer_val) > 0 Then
        Call fSetFilterForSheet(shtSellPriceInAdv, Array(SellPriceInAdv.SalesCompany, SellPriceInAdv.ProductProducer, SellPriceInAdv.ProductName, SellPriceInAdv.ProductSeries) _
                , Array(ebSalesCompany_val, ebProductProducer_val, ebProductName_val, ebProductSeries_val))
    
        Call fGotoCell(shtSellPriceInAdv.Range("A2"), True)
    Else
        fRemoveFilterForSheet shtSellPriceInAdv
    End If
    
    fShowActivateSheet shtSellPriceInAdv
End Sub

Private Sub btnSalesInfo_Click()
    Call fRefreshEditBoxFromShtDataStage
    If Len(ebProductProducer_val) > 0 Then
        Call fSetFilterForSheet(shtSalesInfos, Array(Sales2Hospital.SalesCompany, Sales2Hospital.Hospital, Sales2Hospital.ProductProducer, Sales2Hospital.ProductName, Sales2Hospital.ProductSeries, Sales2Hospital.LotNum) _
                , Array(ebSalesCompany_val, ebHospital_val, ebProductProducer_val, ebProductName_val, ebProductSeries_val, ebLotNum_val))
    
        Call fGotoCell(shtSalesInfos.Range("A2"), True)
    Else
        fRemoveFilterForSheet shtSalesInfos
    End If
    
    fShowActivateSheet shtSalesInfos
End Sub

Private Sub btnSCompInvDiff_Click(control As IRibbonControl)
    Call fRefreshEditBoxFromShtDataStage
    If Len(ebProductProducer_val) > 0 Then
        Call fSetFilterForSheet(shtSalesCompInvDiff, Array(SCompInvDiff.SalesCompany, SCompInvDiff.ProductProducer, SCompInvDiff.ProductName, SCompInvDiff.ProductSeries) _
                , Array(ebSalesCompany_val, ebProductProducer_val, ebProductName_val, ebProductSeries_val))
        Call fGotoCell(shtSalesCompInvDiff.Range("A2"), True)
    Else
        fRemoveFilterForSheet shtSalesCompInvDiff
       ' Call fSetFilterForSheet(shtSalesCompInvDiff, Array(CZLInvDiff.SalesCompany), Array(sCZLCompName))
    End If
    
    fShowActivateSheet shtSalesCompInvDiff
End Sub
Private Sub tbSCompInvCalcd_Click(control As IRibbonControl)
    Call fRefreshEditBoxFromShtDataStage
    If Len(ebProductProducer_val) > 0 Then
        Call fSetFilterForSheet(shtSalesCompInvCalcd, Array(SCompInvCalcd.SalesCompany, SCompInvCalcd.ProductProducer, SCompInvCalcd.ProductName, SCompInvCalcd.ProductSeries) _
                , Array(ebSalesCompany_val, ebProductProducer_val, ebProductName_val, ebProductSeries_val))
        Call fGotoCell(shtSalesCompInvCalcd.Range("A2"), True)
    Else
        fRemoveFilterForSheet shtSalesCompInvCalcd
    End If
    
    fShowActivateSheet shtSalesCompInvCalcd
End Sub

Private Sub btnSCompRolloverInv_Click(control As IRibbonControl)
    Call fRefreshEditBoxFromShtDataStage
    If Len(ebProductProducer_val) > 0 Then
        Call fSetFilterForSheet(shtSalesCompRolloverInv, Array(SCompRollover.SalesCompany, SCompRollover.ProductProducer, SCompRollover.ProductName, SCompRollover.ProductSeries) _
                , Array(ebSalesCompany_val, ebProductProducer_val, ebProductName_val, ebProductSeries_val))
        Call fGotoCell(shtSalesCompRolloverInv.Range("A2"), True)
    Else
        fRemoveFilterForSheet shtSalesCompRolloverInv
    End If
    
    fShowActivateSheet shtSalesCompRolloverInv
End Sub
Private Sub btnSCompPurchase_Click(control As IRibbonControl)
    Call fRefreshEditBoxFromShtDataStage
    If Len(ebProductProducer_val) > 0 Then
        Call fSetFilterForSheet(shtCZLSales2Companies, Array(CZLSales2Comp.SalesCompany, CZLSales2Comp.ProductProducer, CZLSales2Comp.ProductName, CZLSales2Comp.ProductSeries) _
                , Array(ebSalesCompany_val, ebProductProducer_val, ebProductName_val, ebProductSeries_val))
        Call fGotoCell(shtCZLSales2Companies.Range("A2"), True)
    Else
        fRemoveFilterForSheet shtCZLSales2Companies
    End If
    
    fShowActivateSheet shtCZLSales2Companies
End Sub
Private Sub btnSCompSalesToHospital_Click(control As IRibbonControl)
    Call fRefreshEditBoxFromShtDataStage
    If Len(ebProductProducer_val) > 0 Then
        Call fSetFilterForSheet(shtSalesInfos, Array(Sales2Hospital.SalesCompany, Sales2Hospital.ProductProducer, Sales2Hospital.ProductName, Sales2Hospital.ProductSeries, Sales2Hospital.Hospital) _
                , Array(ebSalesCompany_val, ebProductProducer_val, ebProductName_val, ebProductSeries_val, ebHospital_val))
        Call fGotoCell(shtSalesInfos.Range("A2"), True)
    Else
        fRemoveFilterForSheet shtSalesInfos
    End If
    
    fShowActivateSheet shtSalesInfos
End Sub

Private Sub btnCZLInvDiff_Click(control As IRibbonControl)
    Call fRefreshEditBoxFromShtDataStage
    Dim sCZLCompName As String
    sCZLCompName = fGetCompanyNameByID_Common("CZL")
    
    If Len(ebProductProducer_val) > 0 Then
        Call fSetFilterForSheet(shtCZLInvDiff, Array(CZLInvDiff.ProductProducer, CZLInvDiff.ProductName, CZLInvDiff.ProductSeries) _
                , Array(ebProductProducer_val, ebProductName_val, ebProductSeries_val))
        Call fGotoCell(shtCZLInvDiff.Range("A2"), True)
    Else
        fRemoveFilterForSheet shtCZLInvDiff
       ' Call fSetFilterForSheet(shtCZLInvDiff, Array(CZLInvDiff.SalesCompany), Array(sCZLCompName))
    End If
    
    fShowActivateSheet shtCZLInvDiff
End Sub
Private Sub btnCZLInvImported_Click(control As IRibbonControl)
    Call fRefreshEditBoxFromShtDataStage
    Dim sCZLCompName As String
    sCZLCompName = fGetCompanyNameByID_Common("CZL")
    
    If Len(ebProductProducer_val) > 0 Then
        Call fSetFilterForSheet(shtSalesCompInvUnified, Array(SCompUnifiedInv.SalesCompany, SCompUnifiedInv.ProductProducer, SCompUnifiedInv.ProductName, SCompUnifiedInv.ProductSeries, SCompUnifiedInv.LotNum) _
                , Array(sCZLCompName, ebProductProducer_val, ebProductName_val, ebProductSeries_val, ebLotNum_val))
        Call fGotoCell(shtSalesCompInvUnified.Range("A2"), True)
    Else
        fRemoveFilterForSheet shtSalesCompInvUnified
        Call fSetFilterForSheet(shtSalesCompInvUnified, Array(SCompUnifiedInv.SalesCompany), Array(sCZLCompName))
    End If
    
    fShowActivateSheet shtSalesCompInvUnified
End Sub
 
Private Sub tbCZLInvCalcd_Click(control As IRibbonControl)
    Call fRefreshEditBoxFromShtDataStage
    If Len(ebProductProducer_val) > 0 Then
        Call fSetFilterForSheet(shtCZLInventory, Array(CZLInv.ProductProducer, CZLInv.ProductName, CZLInv.ProductSeries) _
                , Array(ebProductProducer_val, ebProductName_val, ebProductSeries_val))
        Call fGotoCell(shtCZLInventory.Range("A2"), True)
    Else
        fRemoveFilterForSheet shtCZLInventory
    End If
    
    fShowActivateSheet shtCZLInventory
End Sub


Private Sub btnCZLSalesToSCompAll_Click(control As IRibbonControl)
    Call fRefreshEditBoxFromShtDataStage
    If Len(ebProductProducer_val) > 0 Then
        Call fSetFilterForSheet(shtCZLSales2SCompAll, Array(CZLSales2CompHist.SalesCompany, CZLSales2CompHist.ProductProducer, CZLSales2CompHist.ProductName, CZLSales2CompHist.ProductSeries) _
                , Array(ebSalesCompany_val, ebProductProducer_val, ebProductName_val, ebProductSeries_val))
        Call fGotoCell(shtCZLSales2SCompAll.Range("A2"), True)
    Else
        fRemoveFilterForSheet shtCZLSales2SCompAll
    End If
    
    fShowActivateSheet shtCZLSales2SCompAll
End Sub
Private Sub btnCZLSalesToSComp_Click(control As IRibbonControl)
    Call fRefreshEditBoxFromShtDataStage
    If Len(ebProductProducer_val) > 0 Then
        Call fSetFilterForSheet(shtCZLSales2Companies, Array(CZLSales2Comp.SalesCompany, CZLSales2Comp.ProductProducer, CZLSales2Comp.ProductName, CZLSales2Comp.ProductSeries) _
                , Array(ebSalesCompany_val, ebProductProducer_val, ebProductName_val, ebProductSeries_val))
        Call fGotoCell(shtCZLSales2Companies.Range("A2"), True)
    Else
        fRemoveFilterForSheet shtCZLSales2Companies
    End If
    
    fShowActivateSheet shtCZLSales2Companies
End Sub
Private Sub btnCZLSalesToHospital_Click(control As IRibbonControl)
    Call fRefreshEditBoxFromShtDataStage
    Dim sCZLName As String
    
    If Len(ebProductProducer_val) > 0 Then
        sCZLName = fGetCompanyNameByID_Common("CZL")
    
        Call fSetFilterForSheet(shtSalesInfos, Array(Sales2Hospital.SalesCompany, Sales2Hospital.ProductProducer, Sales2Hospital.ProductName, Sales2Hospital.ProductSeries) _
                , Array(sCZLName, ebProductProducer_val, ebProductName_val, ebProductSeries_val))
        Call fGotoCell(shtSalesInfos.Range("A2"), True)
    Else
        fRemoveFilterForSheet shtSalesInfos
    End If
    
    fShowActivateSheet shtSalesInfos
End Sub

Private Sub btnCZLRolloverInv_Click(control As IRibbonControl)
    Call fRefreshEditBoxFromShtDataStage
    If Len(ebProductProducer_val) > 0 Then
        Call fSetFilterForSheet(shtCZLRolloverInv, Array(CZLRollover.ProductProducer, CZLRollover.ProductName, CZLRollover.ProductSeries) _
                , Array(ebProductProducer_val, ebProductName_val, ebProductSeries_val))
        Call fGotoCell(shtCZLRolloverInv.Range("A2"), True)
    Else
        fRemoveFilterForSheet shtCZLRolloverInv
    End If
    
    fShowActivateSheet shtCZLRolloverInv
End Sub

Private Sub btnCZLPurchase_Click(control As IRibbonControl)
    Call fRefreshEditBoxFromShtDataStage
    fPrepareCZLPurchaseFromSelfSales
    
    If Len(ebProductProducer_val) > 0 Then
        Call fSetFilterForSheet(shtCZLPurchaseOrder, Array(SelfSales.ProductProducer, SelfSales.ProductName, SelfSales.ProductSeries) _
                , Array(ebProductProducer_val, ebProductName_val, ebProductSeries_val))
        Call fGotoCell(shtCZLPurchaseOrder.Range("A2"), True)
    Else
        fRemoveFilterForSheet shtCZLPurchaseOrder
    End If
    
    fShowActivateSheet shtCZLPurchaseOrder
End Sub

Private Sub btnProfit_Click(control As IRibbonControl)
    Call fRefreshEditBoxFromShtDataStage
    If Len(ebProductProducer_val) > 0 Then
        Call fSetFilterForSheet(shtProfit, Array(Profit.SalesCompany, Profit.Hospital, Profit.ProductProducer, Profit.ProductName, Profit.ProductSeries) _
                , Array(ebSalesCompany_val, ebHospital_val, ebProductProducer_val, ebProductName_val, ebProductSeries_val))
        Call fGotoCell(shtProfit.Range("A2"), True)
    Else
        fRemoveFilterForSheet shtProfit
    End If
    
    fShowActivateSheet shtProfit
End Sub
Private Sub btnRefund_Click(control As IRibbonControl)
    Call fRefreshEditBoxFromShtDataStage
    If Len(ebProductProducer_val) > 0 Then
        Call fSetFilterForSheet(shtRefund, Array(Refund.SalesCompany, Refund.Hospital, Refund.ProductProducer, Refund.ProductName, Refund.ProductSeries) _
                , Array(ebSalesCompany_val, ebHospital_val, ebProductProducer_val, ebProductName_val, ebProductSeries_val))
        Call fGotoCell(shtRefund.Range("A2"), True)
    Else
        fRemoveFilterForSheet shtRefund
    End If
    
    fShowActivateSheet shtRefund
End Sub
Private Sub btnProductNameReplace_Click(control As IRibbonControl)
    Call fRefreshEditBoxFromShtDataStage
    If Len(ebProductProducer_val) > 0 Then
        Call fSetFilterForSheet(shtProductNameReplace, Array(ProdNameReplace.ProductProducer, ProdNameReplace.ProductName) _
                , Array(ebProductProducer_val, ebProductName_val))
        Call fGotoCell(shtProductNameReplace.Range("A2"), True)
    Else
        fRemoveFilterForSheet shtProductNameReplace
    End If
    
    fShowActivateSheet shtProductNameReplace
End Sub
Private Sub btnProductSeriesReplace_Click(control As IRibbonControl)
    Call fRefreshEditBoxFromShtDataStage
    If Len(ebProductProducer_val) > 0 Then
        Call fSetFilterForSheet(shtProductSeriesReplace, Array(ProdSerReplace.ProductProducer, ProdSerReplace.ProductName, ProdSerReplace.ProductSeries) _
                , Array(ebProductProducer_val, ebProductName_val, ebProductSeries_val))
        Call fGotoCell(shtProductSeriesReplace.Range("A2"), True)
    Else
        fRemoveFilterForSheet shtProductSeriesReplace
    End If
    
    fShowActivateSheet shtProductSeriesReplace
End Sub
Private Sub btnProductMaster_Click(control As IRibbonControl)
    Call fRefreshEditBoxFromShtDataStage
    If Len(ebProductProducer_val) > 0 Then
        Call fSetFilterForSheet(shtProductMaster, Array(ProductMst.ProductProducer, ProductMst.ProductName, ProductMst.ProductSeries) _
                , Array(ebProductProducer_val, ebProductName_val, ebProductSeries_val))
        Call fGotoCell(shtProductMaster.Range("A2"), True)
    Else
        fRemoveFilterForSheet shtProductMaster
    End If
    
    fShowActivateSheet shtProductMaster
End Sub
''Private Sub tgSearchBy_Click(control As IRibbonControl, pressed As Boolean)
'Private Sub tgSearchBy_Click(control As IRibbonControl)
''    tgSearchBy_Val = pressed
'    'Call fPresstgSearchBy(pressed)
'    Call fPresstgSearchBy
'End Sub

Private Sub tgSearchBy_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = tgSearchBy_Val
End Sub

'Private Sub btnRemoveFilter_Click(control As IRibbonControl)
'    Sub_RemoveFilterForAcitveSheet
'End Sub
 

Function fGetCurrSheetCurrRowSalesCompany(sOut As String) As Boolean
    Call fGetValueFromCurrSheetCurrRow("SALES_COMPANY", sOut)
End Function
Function fGetCurrSheetCurrRowProductProducer(sOut As String) As Boolean
    Call fGetValueFromCurrSheetCurrRow("PRODUCT_PRODUCER", sOut)
End Function
Function fGetCurrSheetCurrRowProductName(sOut As String) As Boolean
    Call fGetValueFromCurrSheetCurrRow("PRODUCT_NAME", sOut)
End Function
Function fGetCurrSheetCurrRowProductSeries(sOut As String) As Boolean
    Call fGetValueFromCurrSheetCurrRow("PRODUCT_SERIES", sOut)
End Function
Function fGetCurrSheetCurrRowLotNum(sOut As String) As Boolean
    Call fGetValueFromCurrSheetCurrRow("LOT_NUM", sOut)
End Function
Function fGetCurrSheetCurrRowHospital(sOut As String) As Boolean
    Call fGetValueFromCurrSheetCurrRow("HOSPITAL", sOut)
End Function

Function fGetValueFromCurrSheetCurrRow(sColType As String, sOut As String) 'As Boolean
    Dim iColIndex As Integer
    
    sOut = ""
    If ActiveCell.Row <= 1 Then Exit Function
    
    Select Case sColType
        Case "HOSPITAL"
            Call fGetHospitalColIndex(ActiveSheet, iColIndex)
        Case "SALES_COMPANY"
            Call fGetSalesCompanyColIndex(ActiveSheet, iColIndex)
        Case "PRODUCT_PRODUCER"
            Call fGetProductProcuderColIndex(ActiveSheet, iColIndex)
        Case "PRODUCT_NAME"
            Call fGetProductProductNameColIndex(ActiveSheet, iColIndex)
        Case "PRODUCT_SERIES"
            Call fGetProductSeriesColIndex(ActiveSheet, iColIndex)
        Case "LOT_NUM"
            Call fGetLotNumColIndex(ActiveSheet, iColIndex)
        Case Else
            MsgBox "wrong param sColType to fGetValueFromCurrSheetCurrRow: " & sColType, vbCritical
    End Select
    
    If iColIndex <= 0 Then
        If ActiveSheet Is shtCZLInvDiff Then
            sOut = fGetCompanyNameByID_Common("CZL")
        End If
        
        Exit Function
    End If
        
    sOut = Trim(ActiveSheet.Cells(ActiveCell.Row, iColIndex).Value)
End Function


Function fReGetValue_tgSearchBy() As Boolean
    fGetRibbonReference.InvalidateControl "tgSearchBy"
    'fGetRibbonReference.Invalidate
    fReGetValue_tgSearchBy = tgSearchBy_Val
End Function

'Private Sub dwSearchTables_getItemCount(control As IRibbonControl, ByRef returnedVal)
'    returnedVal = 8
'End Sub
'
'Private Sub dwSearchTables_Click(control As IRibbonControl, id As String, index As Integer)
'    Dim sTable
'   ' Call dwSearchTables_getItemLabel(control, index, sTable)
'    MsgBox id & vbCr & sTable
'End Sub
'
'Private Sub dwSearchTables_getItemID(control As IRibbonControl, index As Integer, ByRef id)
'    Select Case index
'        Case 0
'            id = "Table_" & index
'        Case 1
'            id = "tbSCompInvInformed"
'        Case 2
'            id = "tbSCompInvCalcd"
'    End Select
'End Sub
'
'Private Sub dwSearchTables_getItemLabel(control As IRibbonControl, index As Integer, ByRef returnedVal)
'    Select Case index
'        Case 0
'            returnedVal = "药品"
'        Case 1
'            returnedVal = "(商业公司)库存表(导入的)"
'        Case 2
'            returnedVal = "(商业公司)库存表(计算的)"
'        Case Else
'            returnedVal = "药品90"
'    End Select
'End Sub

Function fGetSalesCompanyColIndex(shtParam As Worksheet, ByRef iColIndex As Integer) 'As Boolean
    Select Case shtParam.CodeName
        Case "shtSalesInfos"
            iColIndex = Sales2Hospital.SalesCompany
        Case "shtSalesCompInvUnified"
            iColIndex = SCompUnifiedInv.SalesCompany
        Case "shtPromotionProduct"
            iColIndex = PromoteProduct.SalesCompany
'        Case "shtCZLInvDiff"
'            iColIndex = CZLInvDiff.SalesCompany
        Case "shtSalesCompInvDiff"
            iColIndex = SCompInvDiff.SalesCompany
        Case "shtSalesCompInvCalcd"
            iColIndex = SCompInvCalcd.SalesCompany
        Case "shtProfit"
            iColIndex = Profit.SalesCompany
        Case "shtRefund"
            iColIndex = Refund.SalesCompany
        Case "shtSecondLevelCommission"
            iColIndex = SecondLevelComm.SalesCompany
        Case "shtSellPriceInAdv"
            iColIndex = SellPriceInAdv.SalesCompany
        Case "shtCZLSales2Companies"
            iColIndex = CZLSales2Comp.SalesCompany
        Case "shtCZLSales2SCompAll"
            iColIndex = CZLSales2CompHist.SalesCompany
        Case Else
            iColIndex = 0: Exit Function
    End Select
    
    'fGetSalesCompanyColIndex = True
End Function
Function fGetProductProcuderColIndex(shtParam As Worksheet, ByRef iColIndex As Integer) 'As Boolean
    Select Case shtParam.CodeName
        Case "shtSalesInfos"
            iColIndex = Sales2Hospital.ProductProducer
        Case "shtPromotionProduct"
            iColIndex = PromoteProduct.ProductProducer
        Case "shtSalesCompInvUnified"
            iColIndex = SCompUnifiedInv.ProductProducer
        Case "shtCZLInvDiff"
            iColIndex = CZLInvDiff.ProductProducer
        Case "shtSalesCompInvDiff"
            iColIndex = SCompInvDiff.ProductProducer
        Case "shtSalesCompInvCalcd"
            iColIndex = SCompInvCalcd.ProductProducer
        Case "shtProductNameReplace"
            iColIndex = ProdNameReplace.ProductProducer
        Case "shtProductSeriesReplace"
            iColIndex = ProdSerReplace.ProductProducer
        Case "shtProfit"
            iColIndex = Profit.ProductProducer
        Case "shtCZLInventory"
            iColIndex = CZLInv.ProductProducer
        Case "shtSelfInventory"
            iColIndex = SelfInv.ProductProducer
        Case "shtSelfSalesOrder"
            iColIndex = SelfSales.ProductProducer
        Case "shtSelfPurchaseOrder"
            iColIndex = SelfPurchase.ProductProducer
        Case "shtProductMaster"
            iColIndex = ProductMst.ProductProducer
        Case "shtRefund"
            iColIndex = Refund.ProductProducer
        Case "shtSecondLevelCommission"
            iColIndex = SecondLevelComm.ProductProducer
        Case "shtSellPriceInAdv"
            iColIndex = SellPriceInAdv.ProductProducer
        Case "shtCZLSales2Companies"
            iColIndex = CZLSales2Comp.ProductProducer
        Case "shtCZLSales2SCompAll"
            iColIndex = CZLSales2CompHist.ProductProducer
        Case Else
            iColIndex = 0: Exit Function
    End Select
    
    'fGetProductProcuderColIndex = True
End Function

Function fGetProductProductNameColIndex(shtParam As Worksheet, ByRef iColIndex As Integer) 'As Boolean
    Select Case shtParam.CodeName
        Case "shtSalesInfos"
            iColIndex = Sales2Hospital.ProductName
        Case "shtPromotionProduct"
            iColIndex = PromoteProduct.ProductName
        Case "shtSalesCompInvUnified"
            iColIndex = SCompUnifiedInv.ProductName
        Case "shtCZLInvDiff"
            iColIndex = CZLInvDiff.ProductName
        Case "shtSalesCompInvDiff"
            iColIndex = SCompInvDiff.ProductName
        Case "shtSalesCompInvCalcd"
            iColIndex = SCompInvCalcd.ProductName
        Case "shtProductNameReplace"
            iColIndex = ProdNameReplace.ProductName
        Case "shtProductSeriesReplace"
            iColIndex = ProdSerReplace.ProductName
        Case "shtProfit"
            iColIndex = Profit.ProductName
        Case "shtCZLInventory"
            iColIndex = CZLInv.ProductName
        Case "shtSelfInventory"
            iColIndex = SelfInv.ProductName
        Case "shtSelfSalesOrder"
            iColIndex = SelfSales.ProductName
        Case "shtSelfPurchaseOrder"
            iColIndex = SelfPurchase.ProductName
        Case "shtProductMaster"
            iColIndex = ProductMst.ProductName
        Case "shtRefund"
            iColIndex = Refund.ProductName
        Case "shtSecondLevelCommission"
            iColIndex = SecondLevelComm.ProductName
        Case "shtSellPriceInAdv"
            iColIndex = SellPriceInAdv.ProductName
        Case "shtCZLSales2Companies"
            iColIndex = CZLSales2Comp.ProductName
        Case "shtCZLSales2SCompAll"
            iColIndex = CZLSales2CompHist.ProductName
        Case Else
            iColIndex = 0: Exit Function
    End Select
    
    'fGetProductProductNameColIndex = True
End Function

Function fGetProductSeriesColIndex(shtParam As Worksheet, ByRef iColIndex As Integer) 'As Boolean
    Select Case shtParam.CodeName
        Case "shtSalesInfos"
            iColIndex = Sales2Hospital.ProductSeries
        Case "shtPromotionProduct"
            iColIndex = PromoteProduct.ProductSeries
        Case "shtSalesCompInvUnified"
            iColIndex = SCompUnifiedInv.ProductSeries
        Case "shtCZLInvDiff"
            iColIndex = CZLInvDiff.ProductSeries
        Case "shtSalesCompInvDiff"
            iColIndex = SCompInvDiff.ProductSeries
        Case "shtSalesCompInvCalcd"
            iColIndex = SCompInvCalcd.ProductSeries
        Case "shtProductSeriesReplace"
            iColIndex = ProdSerReplace.ProductSeries
        Case "shtProfit"
            iColIndex = Profit.ProductSeries
        Case "shtCZLInventory"
            iColIndex = CZLInv.ProductSeries
        Case "shtSelfInventory"
            iColIndex = SelfInv.ProductSeries
        Case "shtSelfSalesOrder"
            iColIndex = SelfSales.ProductSeries
        Case "shtSelfPurchaseOrder"
            iColIndex = SelfPurchase.ProductSeries
        Case "shtProductMaster"
            iColIndex = ProductMst.ProductSeries
        Case "shtRefund"
            iColIndex = Refund.ProductSeries
        Case "shtSecondLevelCommission"
            iColIndex = SecondLevelComm.ProductSeries
        Case "shtSellPriceInAdv"
            iColIndex = SellPriceInAdv.ProductSeries
        Case "shtCZLSales2Companies"
            iColIndex = CZLSales2Comp.ProductSeries
        Case "shtCZLSales2SCompAll"
            iColIndex = CZLSales2CompHist.ProductSeries
        Case Else
            iColIndex = 0: Exit Function
    End Select
    
    'fGetProductSeriesColIndex = True
End Function
Function fGetLotNumColIndex(shtParam As Worksheet, ByRef iColIndex As Integer) ' As Boolean
    Select Case shtParam.CodeName
        Case "shtSalesInfos"
            iColIndex = Sales2Hospital.LotNum
        Case "shtSalesCompInvUnified"
            iColIndex = SCompUnifiedInv.LotNum
        Case "shtCZLInvDiff"
            iColIndex = CZLInvDiff.LotNum
        Case "shtSalesCompInvDiff"
            iColIndex = SCompInvDiff.LotNum
        Case "shtSalesCompInvCalcd"
            iColIndex = SCompInvCalcd.LotNum
        Case "shtProfit"
            iColIndex = Profit.LotNum
        Case "shtCZLInventory"
            iColIndex = CZLInv.LotNum
        Case "shtSelfInventory"
            iColIndex = SelfInv.LotNum
        Case "shtSelfSalesOrder"
            iColIndex = SelfSales.LotNum
        Case "shtSelfPurchaseOrder"
            iColIndex = SelfPurchase.LotNum
        Case "shtRefund"
            iColIndex = Refund.LotNum
        Case "shtSecondLevelCommission"
            iColIndex = 0
        Case "shtCZLSales2Companies"
            iColIndex = CZLSales2Comp.LotNum
        Case Else
            iColIndex = 0: Exit Function
    End Select
End Function

Function fGetHospitalColIndex(shtParam As Worksheet, ByRef iColIndex As Integer)
    Select Case shtParam.CodeName
        Case "shtSalesInfos"
            iColIndex = Sales2Hospital.Hospital
        Case "shtPromotionProduct"
            iColIndex = PromoteProduct.Hospital
        Case "shtProfit"
            iColIndex = Profit.Hospital
        Case "shtRefund"
            iColIndex = Refund.Hospital
        Case "shtSecondLevelCommission"
            iColIndex = SecondLevelComm.Hospital
        Case Else
            iColIndex = 0: Exit Function
    End Select
End Function
Function fGetControlAttributes(control As IRibbonControl, sType As String, Optional ByRef val)
    If Not (sType = "LABEL" Or sType = "IMAGE" Or sType = "SIZE" Or sType = "ACTION") Then
        fErr "wrong param to fGetControlAttributes: " & vbCr & "sType=" & sType & vbCr & "control=" & control.id
    End If
    
    Select Case control.id
        Case "btnHome"
            Select Case sType
                Case "LABEL":   val = "主菜单"
                Case "IMAGE":   val = "OpenStartPage"
                Case "SIZE":    val = "true"
                Case "ACTION":  Call Sub_ToHomeSheet
            End Select
        Case "btnCloseAllSheets"
            Select Case sType
                Case "LABEL":   val = "关闭所有"
                Case "IMAGE":   val = "DeclineInvitation"
                Case "SIZE":    val = "true"
                Case "ACTION":  Call subMain_InvisibleHideAllBusinessSheets
            End Select
        Case "btnFilterBySelected"
            Select Case sType
                Case "LABEL":   val = "以所选过滤"
                Case "IMAGE":   val = "FilterBySelection"
                Case "SIZE":    val = "true"
                Case "ACTION":  Call Sub_FilterBySelectedCells
            End Select
        Case "btnSortBySelected"
            Select Case sType
                Case "LABEL":   val = "以所选排序"
                Case "IMAGE":   val = "SortUp"
                Case "SIZE":    val = "true"
                Case "ACTION":  Call sub_SortBySelectedCells
            End Select
            
        Case "btnRemoveFilter"
            Select Case sType
                Case "LABEL":   val = "清除过滤"
                Case "IMAGE":   val = "FilterClearAllFilters"
                Case "SIZE":    val = "true"
                Case "ACTION":  Call Sub_RemoveFilterForAcitveSheet
            End Select
        Case "tgSearchBy"
            Select Case sType
                Case "LABEL":   val = "获取当前行"
                Case "IMAGE":   val = "RecurrenceEdit"
                Case "SIZE":    val = "false"
                Case "ACTION":  Call fPresstgSearchBy
            End Select
        Case "btnSCompInvImported"
            Select Case sType
                Case "LABEL":   val = "(商业公司)库存表(导入的)"
                Case "IMAGE":   val = "PageSetupSheetDialog"
                Case "SIZE":    val = "false"
                Case "ACTION":  Call btnSCompInvImported_Click
            End Select
        Case "btnPromotionProduct"
            Select Case sType
                Case "LABEL":   val = "推广品"
                Case "IMAGE":   val = "PageSetupSheetDialog"
                Case "SIZE":    val = "false"
                Case "ACTION":  Call btnPromotionProduct_Click
            End Select
        Case "btnFirstLevelComm"
            Select Case sType
                Case "LABEL":   val = "采芝林配送费"
                Case "IMAGE":   val = "PageSetupSheetDialog"
                Case "SIZE":    val = "false"
                Case "ACTION":  Call btnFirstLevelComm_Click
            End Select
        Case "btnSecondLevelComm"
            Select Case sType
                Case "LABEL":   val = "商业公司配送费"
                Case "IMAGE":   val = "PageSetupSheetDialog"
                Case "SIZE":    val = "false"
                Case "ACTION":  Call btnSecondLevelComm_Click
            End Select
        Case "btnSalePriceInAdv"
            Select Case sType
                Case "LABEL":   val = "药品预收供货价"
                Case "IMAGE":   val = "PageSetupSheetDialog"
                Case "SIZE":    val = "false"
                Case "ACTION":  Call btnSalePriceInAdv_Click
            End Select
        Case "btnSalesInfo"
            Select Case sType
                Case "LABEL":   val = "替换统一的销售流向"
                Case "IMAGE":   val = "PageSetupSheetDialog"
                Case "SIZE":    val = "false"
                Case "ACTION":  Call btnSalesInfo_Click
            End Select
            
        Case "tbtnProfitSheet"
            Select Case sType
                Case "LABEL":   val = "利润表"
                Case "IMAGE":   val = "ControlSubFormReport"
                'Case "ACTION":  Call ToggleButtonToSwitchSheet_onAction(control)
            End Select
        Case "tbtnRefundSheet"
            Select Case sType
                Case "LABEL":   val = "补差表"
                Case "IMAGE":   val = "ControlSubFormReport"
                'Case "ACTION":  Call ToggleButtonToSwitchSheet_onAction(control)
            End Select
        Case "tbtnSelfInv"
            Select Case sType
                Case "LABEL":   val = "(本公司)库存表"
                Case "IMAGE":   val = "ControlSubFormReport"
                'Case "ACTION":  Call ToggleButtonToSwitchSheet_onAction(control)
            End Select
        Case "tbtnCZLInvCalcd"
            Select Case sType
                Case "LABEL":   val = "(采芝林)库存表(本公司计算)"
                Case "IMAGE":   val = "ControlSubFormReport"
                'Case "ACTION":  Call ToggleButtonToSwitchSheet_onAction(control)
            End Select

    End Select
    
End Function

Sub testaaaaa()
    Dim control As IRibbonControl
    Dim sht As Worksheet
    Set sht = fGetSheetByUIRibbonTag("shtProductMaster")
    
    If Not sht Is Nothing Then
        fToggleSheetVisibleFromUIRibbonControl True, sht, control
    End If
    Set sht = Nothing
End Sub
