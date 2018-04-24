Attribute VB_Name = "MB_1_RibbonUI"
Option Explicit

#If VBA7 And Win64 Then  'Win64
    Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As Long)
#Else
    Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As Long)
#End If

Private mRibbonObj As IRibbonUI
Private ebSalesCompany_val As String
Private ebProductProducer_val As String
Private ebProductName_val As String
Private ebProductSeries_val As String
Private ebLotNum_val As String

'=============================================================
Sub ERP_UI_Onload(ribbon As IRibbonUI)
  Set mRibbonObj = ribbon
  
  fCreateAddNameUpdateNameWhenExists "nmRibbonPointer", ObjPtr(ribbon)
  'Names("nmRibbonPointer").RefersTo = ObjPtr(ribbon)
  
  mRibbonObj.ActivateTab "ERP_2010"
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

Sub subUIebSalesCompany_onChange(control As IRibbonControl, text As String)
    ebSalesCompany_val = text
End Sub
Sub subUIebProductProducer_onChange(control As IRibbonControl, text As String)
    ebProductProducer_val = text
End Sub
Sub subUIebProductName_onChange(control As IRibbonControl, text As String)
    ebProductName_val = text
End Sub
Sub subUIebProductSeries_onChange(control As IRibbonControl, text As String)
    ebProductSeries_val = text
End Sub
Sub subUIebLotnum_onChange(control As IRibbonControl, text As String)
    ebLotNum_val = text
End Sub

Private Sub subUIebSalesCompany_getText(control As IRibbonControl, ByRef returnedVal)
    If fGetCurrSheetCurrRowSalesCompany(ebSalesCompany_val) Then
        returnedVal = ebSalesCompany_val
    Else
        returnedVal = ""
    End If
End Sub
Private Sub subUIebProductProducer_getText(control As IRibbonControl, ByRef returnedVal)
    If fGetCurrSheetCurrRowProductProducer(ebProductProducer_val) Then
        returnedVal = ebProductProducer_val
    Else
        returnedVal = ""
    End If
End Sub

Private Sub subUIebProductName_getText(control As IRibbonControl, ByRef returnedVal)
    If fGetCurrSheetCurrRowProductName(ebProductName_val) Then
        returnedVal = ebProductName_val
    Else
        returnedVal = ""
    End If
End Sub
Private Sub subUIebProductSeries_getText(control As IRibbonControl, ByRef returnedVal)
    If fGetCurrSheetCurrRowProductSeries(ebProductSeries_val) Then
        returnedVal = ebProductSeries_val
    Else
        returnedVal = ""
    End If
End Sub
Private Sub subUIebLotnum_getText(control As IRibbonControl, ByRef returnedVal)
    If fGetCurrSheetCurrRowLotNum(ebLotNum_val) Then
        returnedVal = ebLotNum_val
    Else
        returnedVal = ""
    End If
End Sub

Private Sub UIbtnHome(control As IRibbonControl)
    Call Sub_ToHomeSheet
End Sub
Private Sub btnSelfInventory_Click(control As IRibbonControl)
    If Len(ebProductProducer_val) > 0 Then
        Call fSetFilterForSheet(shtSelfInventory, Array(SelfInv.ProductProducer, SelfInv.ProductName, SelfInv.ProductSeries, SelfInv.LotNum) _
                , Array(ebProductProducer_val, ebProductName_val, ebProductSeries_val, ebLotNum_val))
    Else
        fRemoveFilterForSheet shtSelfInventory
    End If
    
    fShowActivateSheet shtSelfInventory
End Sub
Private Sub btnSelfPurchase_Click(control As IRibbonControl)
    If Len(ebProductProducer_val) > 0 Then
        Call fSetFilterForSheet(shtSelfPurchaseOrder, Array(SelfPurchase.ProductProducer, SelfPurchase.ProductName, SelfPurchase.ProductSeries, SelfPurchase.LotNum) _
                , Array(ebProductProducer_val, ebProductName_val, ebProductSeries_val, ebLotNum_val))
    Else
        fRemoveFilterForSheet shtSelfPurchaseOrder
    End If
    
    fShowActivateSheet shtSelfPurchaseOrder
End Sub
Private Sub btnSelfSales_Click(control As IRibbonControl)
    If Len(ebProductProducer_val) > 0 Then
        Call fSetFilterForSheet(shtSelfSalesOrder, Array(SelfSales.ProductProducer, SelfSales.ProductName, SelfSales.ProductSeries, SelfSales.LotNum) _
                , Array(ebProductProducer_val, ebProductName_val, ebProductSeries_val, ebLotNum_val))
    Else
        fRemoveFilterForSheet shtSelfSalesOrder
    End If
    
    fShowActivateSheet shtSelfSalesOrder
End Sub
Private Sub btnSCompInvImported_Click(control As IRibbonControl)
    If Len(ebProductProducer_val) > 0 Then
        Call fSetFilterForSheet(shtSalesCompInvUnified, Array(SCompUnifiedInv.SalesCompany, SCompUnifiedInv.ProductProducer, SCompUnifiedInv.ProductName, SCompUnifiedInv.ProductSeries, SCompUnifiedInv.LotNum) _
                , Array(ebSalesCompany_val, ebProductProducer_val, ebProductName_val, ebProductSeries_val, ebLotNum_val))
    Else
        fRemoveFilterForSheet shtSalesCompInvUnified
    End If
    
    fShowActivateSheet shtSalesCompInvUnified
End Sub

Private Sub tbSCompInvCalcd_Click(control As IRibbonControl)
    If Len(ebProductProducer_val) > 0 Then
        Call fSetFilterForSheet(shtSalesCompInvCalcd, Array(SCompInvCalcd.SalesCompany, SCompInvCalcd.ProductProducer, SCompInvCalcd.ProductName, SCompInvCalcd.ProductSeries, SCompInvCalcd.LotNum) _
                , Array(ebSalesCompany_val, ebProductProducer_val, ebProductName_val, ebProductSeries_val, ebLotNum_val))
    Else
        fRemoveFilterForSheet shtSalesCompInvCalcd
    End If
    
    fShowActivateSheet shtSalesCompInvCalcd
End Sub

Private Sub btnCZLInvImported_Click(control As IRibbonControl)
    Dim sCZLCompName As String
    sCZLCompName = fGetCompanyNameByID_Common("CZL")
    
    If Len(ebProductProducer_val) > 0 Then
        Call fSetFilterForSheet(shtSalesCompInvUnified, Array(SCompUnifiedInv.SalesCompany, SCompUnifiedInv.ProductProducer, SCompUnifiedInv.ProductName, SCompUnifiedInv.ProductSeries, SCompUnifiedInv.LotNum) _
                , Array(sCZLCompName, ebProductProducer_val, ebProductName_val, ebProductSeries_val, ebLotNum_val))
    Else
        fRemoveFilterForSheet shtSalesCompInvUnified
        Call fSetFilterForSheet(shtSalesCompInvUnified, Array(SCompUnifiedInv.SalesCompany), Array(sCZLCompName))
    End If
    
    fShowActivateSheet shtSalesCompInvUnified
End Sub
 
Private Sub tbCZLInvCalcd_Click(control As IRibbonControl)
    If Len(ebProductProducer_val) > 0 Then
        Call fSetFilterForSheet(shtCZLInventory, Array(CZLInv.ProductProducer, CZLInv.ProductName, CZLInv.ProductSeries, CZLInv.LotNum) _
                , Array(ebProductProducer_val, ebProductName_val, ebProductSeries_val, ebLotNum_val))
    Else
        fRemoveFilterForSheet shtCZLInventory
    End If
    
    fShowActivateSheet shtCZLInventory
End Sub
Private Sub btnCZLSalesToSComp_Click(control As IRibbonControl)
    If Len(ebProductProducer_val) > 0 Then
        Call fSetFilterForSheet(shtCZLSales2Companies, Array(CZLSales2Comp.SalesCompany, CZLSales2Comp.ProductProducer, CZLSales2Comp.ProductName, CZLSales2Comp.ProductSeries, CZLSales2Comp.LotNum) _
                , Array(ebSalesCompany_val, ebProductProducer_val, ebProductName_val, ebProductSeries_val, ebLotNum_val))
    Else
        fRemoveFilterForSheet shtCZLSales2Companies
    End If
    
    fShowActivateSheet shtCZLSales2Companies
End Sub

Private Sub btnProfit_Click(control As IRibbonControl)
    If Len(ebProductProducer_val) > 0 Then
        Call fSetFilterForSheet(shtProfit, Array(Profit.SalesCompany, Profit.ProductProducer, Profit.ProductName, Profit.ProductSeries, Profit.LotNum) _
                , Array(ebSalesCompany_val, ebProductProducer_val, ebProductName_val, ebProductSeries_val, ebLotNum_val))
    Else
        fRemoveFilterForSheet shtProfit
    End If
    
    fShowActivateSheet shtProfit
End Sub
Private Sub btnProductNameReplace_Click(control As IRibbonControl)
    If Len(ebProductProducer_val) > 0 Then
        Call fSetFilterForSheet(shtProductNameReplace, Array(ProdNameReplace.ProductProducer, ProdNameReplace.ProductName) _
                , Array(ebProductProducer_val, ebProductName_val))
    Else
        fRemoveFilterForSheet shtProductNameReplace
    End If
    
    fShowActivateSheet shtProductNameReplace
End Sub
Private Sub btnProductSeriesReplace_Click(control As IRibbonControl)
    If Len(ebProductProducer_val) > 0 Then
        Call fSetFilterForSheet(shtProductSeriesReplace, Array(ProdSerReplace.ProductProducer, ProdSerReplace.ProductName, ProdSerReplace.ProductSeries) _
                , Array(ebProductProducer_val, ebProductName_val, ebProductSeries_val))
    Else
        fRemoveFilterForSheet shtProductSeriesReplace
    End If
    
    fShowActivateSheet shtProductSeriesReplace
End Sub
Private Sub btnProductMaster_Click(control As IRibbonControl)
    If Len(ebProductProducer_val) > 0 Then
        Call fSetFilterForSheet(shtProductMaster, Array(ProductMst.ProductProducer, ProductMst.ProductName, ProductMst.ProductSeries) _
                , Array(ebProductProducer_val, ebProductName_val, ebProductSeries_val))
    Else
        fRemoveFilterForSheet shtProductMaster
    End If
    
    fShowActivateSheet shtProductMaster
End Sub
'Private Sub tgSearchBy_Click(control As IRibbonControl, pressed As Boolean)
Private Sub tgSearchBy_Click(control As IRibbonControl)
'    tgSearchBy_Val = pressed
    'Call fPresstgSearchBy(pressed)
    Call fPresstgSearchBy
End Sub

Private Sub tgSearchBy_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = tgSearchBy_Val
End Sub

Private Sub btnRemoveFilter_Click(control As IRibbonControl)
    Sub_RemoveFilterForAcitveSheet
End Sub

Private Sub btnCloseAllSheets_Click(control As IRibbonControl)
    subMain_InvisibleHideAllBusinessSheets
End Sub

'Function fPresstgSearchBy(bPressed As Boolean)
Function fPresstgSearchBy()
    'If bPressed Then
        fGetRibbonReference.InvalidateControl "ebSalesCompany"
        fGetRibbonReference.InvalidateControl "ebProductProducer"
        fGetRibbonReference.InvalidateControl "ebProductName"
        fGetRibbonReference.InvalidateControl "ebProductSeries"
        fGetRibbonReference.InvalidateControl "ebLotnum"
'    Else
'    End If
End Function

Function fGetCurrSheetCurrRowSalesCompany(sOut As String) As Boolean
    Dim iColIndex As Integer
    
    If ActiveCell.Row <= 1 Then Exit Function
    If Not fGetSalesCompanyColIndex(ActiveSheet, iColIndex) Then Exit Function
        
    sOut = Trim(ActiveSheet.Cells(ActiveCell.Row, iColIndex).Value)
    fGetCurrSheetCurrRowSalesCompany = True
End Function
Function fGetCurrSheetCurrRowProductProducer(sOut As String) As Boolean
    Dim iColIndex As Integer
    
    If ActiveCell.Row <= 1 Then Exit Function
    If Not fGetProductProcuderColIndex(ActiveSheet, iColIndex) Then Exit Function
        
    sOut = Trim(ActiveSheet.Cells(ActiveCell.Row, iColIndex).Value)
    
    fGetCurrSheetCurrRowProductProducer = True
End Function
Function fGetCurrSheetCurrRowProductName(sOut As String) As Boolean
    Dim iColIndex As Integer
    
    If ActiveCell.Row <= 1 Then Exit Function
    If Not fGetProductProductNameColIndex(ActiveSheet, iColIndex) Then Exit Function
        
    sOut = Trim(ActiveSheet.Cells(ActiveCell.Row, iColIndex).Value)
    
    fGetCurrSheetCurrRowProductName = True
End Function
Function fGetCurrSheetCurrRowProductSeries(sOut As String) As Boolean
    Dim iColIndex As Integer
    
    If ActiveCell.Row <= 1 Then Exit Function
    If Not fGetProductSeriesColIndex(ActiveSheet, iColIndex) Then Exit Function
        
    sOut = Trim(ActiveSheet.Cells(ActiveCell.Row, iColIndex).Value)
    
    fGetCurrSheetCurrRowProductSeries = True
End Function
Function fGetCurrSheetCurrRowLotNum(sOut As String) As Boolean
    Dim iColIndex As Integer
    
    If ActiveCell.Row <= 1 Then Exit Function
    If Not fGetLogNumColIndex(ActiveSheet, iColIndex) Then Exit Function
        
    sOut = Trim(ActiveSheet.Cells(ActiveCell.Row, iColIndex).Value)
    
    fGetCurrSheetCurrRowLotNum = True
End Function

Function fReGetValue_tgSearchBy() As Boolean
    fGetRibbonReference.InvalidateControl "tgSearchBy"
    'fGetRibbonReference.Invalidate
    fReGetValue_tgSearchBy = tgSearchBy_Val
End Function

'Function fGetProdProducerCol(shtParam As Worksheet) As Integer
'    Select Case shtParam.CodeName
'        Case "shtSalesCompInvDiff"
'            fGetProdProducerCol = SCompInvDiff.ProductProducer
'        Case Else
'            fGetProdProducerCol = 9
'    End Select
'End Function

Sub ddd()
    Debug.Print tgSearchBy_Val
End Sub


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

Function fGetSalesCompanyColIndex(shtParam As Worksheet, ByRef iColIndex As Integer) As Boolean
    Select Case shtParam.CodeName
        Case "shtSalesCompInvUnified"
            iColIndex = SCompUnifiedInv.SalesCompany
'        Case "shtCZLInvDiff"
'            iColIndex = CZLInvDiff.SalesCompany
        Case "shtSalesCompInvDiff"
            iColIndex = SCompInvDiff.SalesCompany
        Case "shtSalesCompInvCalcd"
            iColIndex = SCompInvCalcd.SalesCompany
        Case "shtProfit"
            iColIndex = Profit.SalesCompany
        Case Else
            iColIndex = 0: Exit Function
    End Select
    
    fGetSalesCompanyColIndex = True
End Function
Function fGetProductProcuderColIndex(shtParam As Worksheet, ByRef iColIndex As Integer) As Boolean
    Select Case shtParam.CodeName
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
        Case Else
            iColIndex = 0: Exit Function
    End Select
    
    fGetProductProcuderColIndex = True
End Function

Function fGetProductProductNameColIndex(shtParam As Worksheet, ByRef iColIndex As Integer) As Boolean
    Select Case shtParam.CodeName
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
        Case Else
            iColIndex = 0: Exit Function
    End Select
    
    fGetProductProductNameColIndex = True
End Function

Function fGetProductSeriesColIndex(shtParam As Worksheet, ByRef iColIndex As Integer) As Boolean
    Select Case shtParam.CodeName
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
        Case Else
            iColIndex = 0: Exit Function
    End Select
    
    fGetProductSeriesColIndex = True
End Function
Function fGetLogNumColIndex(shtParam As Worksheet, ByRef iColIndex As Integer) As Boolean
    Select Case shtParam.CodeName
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
        Case Else
            iColIndex = 0: Exit Function
    End Select
    
    fGetLogNumColIndex = True
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
