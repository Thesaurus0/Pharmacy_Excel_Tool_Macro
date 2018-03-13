VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtSelfInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error GoTo Exit_Sub
    Application.ScreenUpdating = False
'
'    Const ProducerCol = 1
'    Const ProductNameCol = 2
'    Const ProductSeriesCol = 3
'    Const ProductUnitCol = 4
''    Const SellQuantityCol = 6
''    Const SellPriceCol = 7
'    Const LotNumCol = 5
'
'    Dim sLotNum As String
'
'    Dim rgIntersect As Range
'    Set rgIntersect = Intersect(Target, Me.Columns(ProductNameCol))
'
'    'product name
'    If Not rgIntersect Is Nothing Then
'        If rgIntersect.Areas.Count > 1 Then GoTo Exit_Sub    'fErr "不能选多个"
'        If rgIntersect.Rows.Count <> 1 Then GoTo Exit_Sub
'
'        Dim sProducer As String
'        Dim sValidationListAddr As String
'
'        sProducer = rgIntersect.Offset(0, ProducerCol - ProductNameCol).Value
'
'        If fNzero(sProducer) Then
'            Call fSetFilterForSheet(shtProductNameMaster, 1, sProducer)
'            Call fCopyFilteredDataToRange(shtProductNameMaster, 2)
'
'            sValidationListAddr = "=" & shtDataStage.Columns("A").Address(external:=True)
'            'Call fSetValidationListForshtProductNameReplace_ProductName(sValidationListAddr, 3)
'            Call fSetValidationListForRange(rgIntersect, sValidationListAddr)
'        End If
'    Else
'        'product SeriesCol
'        Set rgIntersect = Intersect(Target, Me.Columns(ProductSeriesCol))
'
'        If Not rgIntersect Is Nothing Then
'            If rgIntersect.Areas.Count > 1 Then GoTo Exit_Sub    'fErr "不能选多个"
'            If rgIntersect.Rows.Count <> 1 Then GoTo Exit_Sub
'
'            sProducer = rgIntersect.Offset(0, ProducerCol - ProductSeriesCol).Value
'            sProductName = rgIntersect.Offset(0, ProductNameCol - ProductSeriesCol).Value
'
'            If fNzero(sProducer) And fNzero(sProductName) Then
'                Call fSetFilterForSheet(shtProductMaster, Array(1, 2), Array(sProducer, sProductName))
'                Call fCopyFilteredDataToRange(shtProductMaster, 3)
'
'                sValidationListAddr = "=" & shtDataStage.Columns("A").Address(external:=True)
'                'Call fSetValidationListForshtProductNameReplace_ProductName(sValidationListAddr, 3)
'                Call fSetValidationListForRange(rgIntersect, sValidationListAddr)
'            End If
'        Else
'            'product Unit
'            Set rgIntersect = Intersect(Target, Me.Columns(ProductUnitCol))
'
'            If Not rgIntersect Is Nothing Then
'                If rgIntersect.Areas.Count > 1 Then GoTo Exit_Sub    'fErr "不能选多个"
'                If rgIntersect.Rows.Count <> 1 Then GoTo Exit_Sub
'
'                sProducer = rgIntersect.Offset(0, ProducerCol - ProductUnitCol).Value
'                sProductName = rgIntersect.Offset(0, ProductNameCol - ProductUnitCol).Value
'                sProductSeries = rgIntersect.Offset(0, ProductSeriesCol - ProductUnitCol).Value
'
'                If fNzero(sProducer) And fNzero(sProductName) Then
'                    Call fSetFilterForSheet(shtProductMaster, Array(1, 2, 3), Array(sProducer, sProductName, sProductSeries))
'                    Call fCopyFilteredDataToRange(shtProductMaster, 4)
'
'                    sValidationListAddr = "=" & shtDataStage.Columns("A").Address(external:=True)
'                    'Call fSetValidationListForshtProductNameReplace_ProductName(sValidationListAddr, 3)
'                    Call fSetValidationListForRange(rgIntersect, sValidationListAddr)
'                End If
'            Else
'                'Lot Number
'                Set rgIntersect = Intersect(Target, Me.Columns(LotNumCol))
'
'                If Not rgIntersect Is Nothing Then
'                    If rgIntersect.Areas.Count > 1 Then GoTo Exit_Sub    'fErr "不能选多个"
'                    If rgIntersect.Rows.Count <> 1 Then GoTo Exit_Sub
'
'                    sProducer = rgIntersect.Offset(0, ProducerCol - LotNumCol).Value
'                    sProductName = rgIntersect.Offset(0, ProductNameCol - LotNumCol).Value
'                    sProductSeries = rgIntersect.Offset(0, ProductSeriesCol - LotNumCol).Value
'
'                    If fNzero(sProducer) And fNzero(sProductName) Then
'                        Call fSetFilterForSheet(shtSelfPurchaseOrder, Array(1, 2, 3), Array(sProducer, sProductName, sProductSeries))
'                        Call fCopyFilteredDataToRange(shtSelfPurchaseOrder, 8)
'
'                        sValidationListAddr = "=" & shtDataStage.Columns("A").Address(external:=True)
'                        'Call fSetValidationListForshtProductNameReplace_ProductName(sValidationListAddr, 3)
'                        Call fSetValidationListForRange(rgIntersect, sValidationListAddr)
'                    End If
'                Else
''                    'Sell Price
''                    Set rgIntersect = Intersect(Target, Me.Columns(SellPriceCol))
''
''                    If Not rgIntersect Is Nothing Then
''                        If rgIntersect.Areas.Count > 1 Then GoTo Exit_Sub    'fErr "不能选多个"
''                        If rgIntersect.Rows.Count <> 1 Then GoTo Exit_Sub
''
''                        sProducer = rgIntersect.Offset(0, ProducerCol - SellPriceCol).Value
''                        sProductName = rgIntersect.Offset(0, ProductNameCol - SellPriceCol).Value
''                        sProductSeries = rgIntersect.Offset(0, ProductSeriesCol - SellPriceCol).Value
''                        sLotNum = rgIntersect.Offset(0, LotNumCol - SellPriceCol).Value
''
''                        If fNzero(sProducer) And fNzero(sProductName) Then
''                            Call fSetFilterForSheet(shtSelfPurchaseOrder, Array(1, 2, 3, 8), Array(sProducer, sProductName, sProductSeries, sLotNum))
''                            Call fCopyFilteredDataToRange(shtSelfPurchaseOrder, 7)
''
''                            sValidationListAddr = "=" & shtDataStage.Columns("A").Address(external:=True)
''                            'Call fSetValidationListForshtProductNameReplace_ProductName(sValidationListAddr, 3)
''                            Call fSetValidationListForRange(rgIntersect, sValidationListAddr)
''                        End If
''                    Else
''                        'Sell Quantity
''                        Set rgIntersect = Intersect(Target, Me.Columns(SellQuantityCol))
''
''                        If Not rgIntersect Is Nothing Then
''                            If rgIntersect.Areas.Count > 1 Then GoTo exit_sub    'fErr "不能选多个"
''                            If rgIntersect.Rows.Count <> 1 Then GoTo exit_sub
''
''                            sProducer = rgIntersect.Offset(0, ProducerCol - SellQuantityCol).Value
''                            sProductName = rgIntersect.Offset(0, ProductNameCol - SellQuantityCol).Value
''                            sProductSeries = rgIntersect.Offset(0, ProductSeriesCol - SellQuantityCol).Value
''                            sLotNum = rgIntersect.Offset(0, LotNumCol - SellQuantityCol).Value
''
''                            If fNzero(sProducer) And fNzero(sProductName) Then
''                                Call fSetFilterForSheet(shtSelfPurchaseOrder, Array(1, 2, 3, 8), Array(sProducer, sProductName, sProductSeries, sLotNum))
''                                Call fCopyFilteredDataToRange(shtSelfPurchaseOrder, 6)
''
''                                sValidationListAddr = "=" & shtDataStage.Columns("A").Address(external:=True)
''                                'Call fSetValidationListForshtProductNameReplace_ProductName(sValidationListAddr, 3)
''                                Call fSetValidationListForRange(rgIntersect, sValidationListAddr)
''                            End If
''                        Else
''
''                        End If
''                    End If
'                End If
'            End If
'        End If
'    End If
    
    Dim lCurrRow As Long
    Dim lCurrCol As Long
    lCurrRow = ActiveCell.Row
    lCurrCol = ActiveCell.Column
    If lCurrCol = LotNumCol Then
        sProducer = Me.Cells(lCurrRow, ProducerCol).Value
        sProductName = Me.Cells(lCurrRow, ProductNameCol).Value
        sProductSeries = Me.Cells(lCurrRow, ProductSeriesCol).Value
        sLotNum = Me.Cells(lCurrRow, LotNumCol).Value

'        Call fSetFilterForSheet(shtSelfPurchaseOrder, Array(1, 2, 3, 8), Array(sProducer, sProductName, sProductSeries, sLotNum))
        Call fSetFilterForSheet(shtSelfSalesOrder, Array(1, 2, 3, 8), Array(sProducer, sProductName, sProductSeries, sLotNum))
    End If
    
Exit_Sub:
    fEnableExcelOptionsAll
    Application.ScreenUpdating = True
    
    If Err.Number <> 0 Then fMsgBox Err.Description
End Sub
