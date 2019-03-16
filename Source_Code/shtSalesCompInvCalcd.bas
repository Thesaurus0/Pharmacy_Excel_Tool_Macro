VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtSalesCompInvCalcd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Option Base 1

Enum SCompInvCalcd
    SalesCompany = 1
    ProductProducer = 2
    ProductName = 3
    ProductSeries = 4
    ProductUnit = 5
    LotNum = 6
    InventoryQty = 7
    PurchasePrice = 8
End Enum

Private Sub btnRolloverInv_Click()
    Dim sSalesComp As String
    Dim sProducer As String
    Dim sProductName As String
    Dim ssProductSeries As String
    Dim lCurrRow As Long
    
    lCurrRow = ActiveCell.Row
    sSalesComp = shtSalesCompInvCalcd.Cells(lCurrRow, SCompInvCalcd.SalesCompany)
    sProducer = shtSalesCompInvCalcd.Cells(lCurrRow, SCompInvCalcd.ProductProducer)
    sProductName = shtSalesCompInvCalcd.Cells(lCurrRow, SCompInvCalcd.ProductName)
    ssProductSeries = shtSalesCompInvCalcd.Cells(lCurrRow, SCompInvCalcd.ProductSeries)
    
    fActiveVisibleSwitchSheet shtSalesCompRolloverInv
    
    If Len(sProductName) > 0 And lCurrRow > 1 Then
        'SelfSales = CZLPurchase
        Call fSetFilterForSheet(shtSalesCompRolloverInv, Array(SCompRollover.SalesCompany, SCompRollover.ProductProducer, SCompRollover.ProductName, SCompRollover.ProductSeries) _
                , Array(sSalesComp, sProducer, sProductName, ssProductSeries))
    Else
        fRemoveFilterForSheet shtSalesCompRolloverInv
    End If
End Sub

Private Sub btnSCompPurchase_Click()
    Dim sSalesComp As String
    Dim sProducer As String
    Dim sProductName As String
    Dim ssProductSeries As String
    Dim lCurrRow As Long
    
    lCurrRow = ActiveCell.Row
    sSalesComp = shtSalesCompInvCalcd.Cells(lCurrRow, SCompInvCalcd.SalesCompany)
    sProducer = shtSalesCompInvCalcd.Cells(lCurrRow, SCompInvCalcd.ProductProducer)
    sProductName = shtSalesCompInvCalcd.Cells(lCurrRow, SCompInvCalcd.ProductName)
    ssProductSeries = shtSalesCompInvCalcd.Cells(lCurrRow, SCompInvCalcd.ProductSeries)
    
    If Len(sProductName) > 0 And lCurrRow > 1 Then
        'SelfSales = CZLPurchase
        Call fSetFilterForSheet(shtCZLSales2Companies, Array(CZLSales2Comp.SalesCompany, CZLSales2Comp.ProductProducer, CZLSales2Comp.ProductName, CZLSales2Comp.ProductSeries) _
                , Array(sSalesComp, sProducer, sProductName, ssProductSeries))
    Else
        fRemoveFilterForSheet shtCZLSales2Companies
    End If
    
    fActiveVisibleSwitchSheet shtCZLSales2Companies
End Sub

Private Sub btnSCompSales_Click()
    Dim sSalesComp As String
    Dim sProducer As String
    Dim sProductName As String
    Dim ssProductSeries As String
    Dim lCurrRow As Long
    
    lCurrRow = ActiveCell.Row
    
    sSalesComp = shtSalesCompInvCalcd.Cells(lCurrRow, SCompInvCalcd.SalesCompany)
    sProducer = shtSalesCompInvCalcd.Cells(lCurrRow, SCompInvCalcd.ProductProducer)
    sProductName = shtSalesCompInvCalcd.Cells(lCurrRow, SCompInvCalcd.ProductName)
    ssProductSeries = shtSalesCompInvCalcd.Cells(lCurrRow, SCompInvCalcd.ProductSeries)
    
    If Len(sProductName) > 0 And lCurrRow > 1 Then
        'SelfSales = CZLPurchase
        Call fSetFilterForSheet(shtSalesInfos, Array(Sales2Hospital.SalesCompany, Sales2Hospital.ProductProducer, Sales2Hospital.ProductName, Sales2Hospital.ProductSeries) _
                , Array(sSalesComp, sProducer, sProductName, ssProductSeries))
    Else
        fRemoveFilterForSheet shtSalesInfos
    End If
    
    fActiveVisibleSwitchSheet shtSalesInfos
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error GoTo exit_sub
    Application.ScreenUpdating = False
    
    Const ProducerCol = SCompInvCalcd.ProductProducer
    Const ProductNameCol = SCompInvCalcd.ProductName
    Const ProductSeriesCol = SCompInvCalcd.ProductSeries
    Const ProductUnitCol = SCompInvCalcd.ProductUnit
'    Const SellQuantityCol = 6
    Const SellPriceCol = SCompInvCalcd.PurchasePrice
    Const LotNumCol = SCompInvCalcd.LotNum

    Dim sLotNum As String

    Dim rgIntersect As Range
    Set rgIntersect = Intersect(Target, Me.Columns(ProductNameCol))

    'product name
    If Not rgIntersect Is Nothing Then
        If rgIntersect.Areas.Count > 1 Then GoTo exit_sub    'fErr "不能选多个"
        If rgIntersect.Rows.Count <> 1 Then GoTo exit_sub

        Dim sProducer As String
        Dim sValidationListAddr As String

        sProducer = rgIntersect.Offset(0, ProducerCol - ProductNameCol).Value
        Call fGetProductNameValidationListAndSetToCell(rgIntersect, sProducer)

'        If fNzero(sProducer) Then
'            Call fSetFilterForSheet(shtProductNameMaster, ProductNameMst.ProdProducer, sProducer)
'            Call fCopyFilteredDataToRange(shtProductNameMaster, 2)
'
'            sValidationListAddr = "=" & shtDataStage.Columns("A").Address(external:=True)
'            'Call fSetValidationListForshtProductNameReplace_ProductName(sValidationListAddr, 3)
'            Call fSetValidationListForRange(rgIntersect, sValidationListAddr)
'        End If
    Else
        'product SeriesCol
        Set rgIntersect = Intersect(Target, Me.Columns(ProductSeriesCol))

        If Not rgIntersect Is Nothing Then
            If rgIntersect.Areas.Count > 1 Then GoTo exit_sub    'fErr "不能选多个"
            If rgIntersect.Rows.Count <> 1 Then GoTo exit_sub

            sProducer = rgIntersect.Offset(0, ProducerCol - ProductSeriesCol).Value
            sProductName = rgIntersect.Offset(0, ProductNameCol - ProductSeriesCol).Value
            Call fGetProductSeriesValidationListAndSetToCell(rgIntersect, sProducer, sProductName)

'            If fNzero(sProducer) And fNzero(sProductName) Then
'                Call fSetFilterForSheet(shtProductMaster, Array(1, 2), Array(sProducer, sProductName))
'                Call fCopyFilteredDataToRange(shtProductMaster, 3)
'
'                sValidationListAddr = "=" & shtDataStage.Columns("A").Address(external:=True)
'                'Call fSetValidationListForshtProductNameReplace_ProductName(sValidationListAddr, 3)
'                Call fSetValidationListForRange(rgIntersect, sValidationListAddr)
'            End If
        Else
            'product Unit
            Set rgIntersect = Intersect(Target, Me.Columns(ProductUnitCol))

            If Not rgIntersect Is Nothing Then
                If rgIntersect.Areas.Count > 1 Then GoTo exit_sub    'fErr "不能选多个"
                If rgIntersect.Rows.Count <> 1 Then GoTo exit_sub

                sProducer = rgIntersect.Offset(0, ProducerCol - ProductUnitCol).Value
                sProductName = rgIntersect.Offset(0, ProductNameCol - ProductUnitCol).Value
                sProductSeries = rgIntersect.Offset(0, ProductSeriesCol - ProductUnitCol).Value

                If fNzero(sProducer) And fNzero(sProductName) Then
                    Call fSetFilterForSheet(shtProductMaster, Array(1, 2, 3), Array(sProducer, sProductName, sProductSeries))
                    Call fCopyFilteredDataToRange(shtProductMaster, 4)

                    sValidationListAddr = "=" & shtDataStage.Columns("A").Address(external:=True)
                    'Call fSetValidationListForshtProductNameReplace_ProductName(sValidationListAddr, 3)
                    Call fSetValidationListForRange(rgIntersect, sValidationListAddr)
                End If
            Else
                'Sell Price
                Set rgIntersect = Intersect(Target, Me.Columns(SellPriceCol))

                If Not rgIntersect Is Nothing Then
'                    If rgIntersect.Areas.Count > 1 Then GoTo exit_sub    'fErr "不能选多个"
'                    If rgIntersect.Rows.Count <> 1 Then GoTo exit_sub
'
'                    sProducer = rgIntersect.Offset(0, ProducerCol - SellPriceCol).Value
'                    sProductName = rgIntersect.Offset(0, ProductNameCol - SellPriceCol).Value
'                    sProductSeries = rgIntersect.Offset(0, ProductSeriesCol - SellPriceCol).Value
'                    sLotNum = rgIntersect.Offset(0, LotNumCol - SellPriceCol).Value
'
'                    If fNzero(sProducer) And fNzero(sProductName) Then
'                        Call fSetFilterForSheet(shtSelfSalesOrder, Array(1, 2, 3, 8), Array(sProducer, sProductName, sProductSeries, sLotNum))
'                        Call fCopyFilteredDataToRange(shtSelfSalesOrder, 7)
'
'                        sValidationListAddr = "=" & shtDataStage.Columns("A").Address(external:=True)
'                        'Call fSetValidationListForshtProductNameReplace_ProductName(sValidationListAddr, 3)
'                        Call fSetValidationListForRange(rgIntersect, sValidationListAddr)
'                    End If
                Else
                    'Lot Number
                    Set rgIntersect = Intersect(Target, Me.Columns(LotNumCol))

                    If Not rgIntersect Is Nothing Then
                        If rgIntersect.Areas.Count > 1 Then GoTo exit_sub    'fErr "不能选多个"
                        If rgIntersect.Rows.Count <> 1 Then GoTo exit_sub

                        sProducer = rgIntersect.Offset(0, ProducerCol - LotNumCol).Value
                        sProductName = rgIntersect.Offset(0, ProductNameCol - LotNumCol).Value
                        sProductSeries = rgIntersect.Offset(0, ProductSeriesCol - LotNumCol).Value

                        If fNzero(sProducer) And fNzero(sProductName) And fNzero(sProductSeries) Then
                            Call fSetFilterForSheet(shtSelfSalesOrder, Array(SelfSales.ProductSeries, SelfSales.ProductName, SelfSales.ProductSeries), Array(sProducer, sProductName, sProductSeries))
                            Call fCopyFilteredDataToRange(shtSelfSalesOrder, SelfSales.LotNum)

                            sValidationListAddr = "=" & shtDataStage.Columns("A").Address(external:=True)
                            'Call fSetValidationListForshtProductNameReplace_ProductName(sValidationListAddr, 3)
                            Call fSetValidationListForRange(rgIntersect, sValidationListAddr)
                        End If
                    Else
'                        'Sell Quantity
'                        Set rgIntersect = Intersect(Target, Me.Columns(SellQuantityCol))
'
'                        If Not rgIntersect Is Nothing Then
'                            If rgIntersect.Areas.Count > 1 Then GoTo exit_sub    'fErr "不能选多个"
'                            If rgIntersect.Rows.Count <> 1 Then GoTo exit_sub
'
'                            sProducer = rgIntersect.Offset(0, ProducerCol - SellQuantityCol).Value
'                            sProductName = rgIntersect.Offset(0, ProductNameCol - SellQuantityCol).Value
'                            sProductSeries = rgIntersect.Offset(0, ProductSeriesCol - SellQuantityCol).Value
'                            sLotNum = rgIntersect.Offset(0, LotNumCol - SellQuantityCol).Value
'
'                            If fNzero(sProducer) And fNzero(sProductName) Then
'                                Call fSetFilterForSheet(shtSelfSalesOrder, Array(1, 2, 3, 8), Array(sProducer, sProductName, sProductSeries, sLotNum))
'                                Call fCopyFilteredDataToRange(shtSelfSalesOrder, 6)
'
'                                sValidationListAddr = "=" & shtDataStage.Columns("A").Address(external:=True)
'                                'Call fSetValidationListForshtProductNameReplace_ProductName(sValidationListAddr, 3)
'                                Call fSetValidationListForRange(rgIntersect, sValidationListAddr)
'                            End If
'                        Else
'
'                        End If
                    End If
                End If
            End If
        End If
    End If

'    Dim lCurrRow As Long
'    Dim lCurrCol As Long
'    lCurrRow = ActiveCell.Row
'    lCurrCol = ActiveCell.Column
'    If lCurrCol = LotNumCol Then
'        sProducer = Me.Cells(lCurrRow, ProducerCol).Value
'        sProductName = Me.Cells(lCurrRow, ProductNameCol).Value
'        sProductSeries = Me.Cells(lCurrRow, ProductSeriesCol).Value
'        sLotNum = Me.Cells(lCurrRow, LotNumCol).Value
'
'        Call fSetFilterForSheet(shtSelfSalesOrder, Array(1, 2, 3, 8), Array(sProducer, sProductName, sProductSeries, sLotNum))
'        Call fSetFilterForSheet(shtSelfSalesOrder, Array(1, 2, 3, 8), Array(sProducer, sProductName, sProductSeries, sLotNum))
'    End If
    
exit_sub:
    fEnableExcelOptionsAll
    Application.ScreenUpdating = True
    
    If Err.Number <> 0 Then fMsgBox Err.Description
End Sub
