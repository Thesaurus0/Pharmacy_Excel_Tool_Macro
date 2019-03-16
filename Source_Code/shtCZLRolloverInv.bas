VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtCZLRolloverInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Option Base 1

Enum CZLRollover
    ProductProducer = 1
    ProductName = 2
    ProductSeries = 3
    ProductUnit = 4
    LotNum = 5
    RolloverQty = 6
End Enum

Private Sub btnValidate_Click()
    Call fValidateSheet
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error GoTo exit_sub
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
''        Call fSetFilterForSheet(shtSelfPurchaseOrder, Array(1, 2, 3, 8), Array(sProducer, sProductName, sProductSeries, sLotNum))
'        Call fSetFilterForSheet(shtSelfSalesOrder, Array(1, 2, 3, 8), Array(sProducer, sProductName, sProductSeries, sLotNum))
'    End If
    
exit_sub:
    fEnableExcelOptionsAll
    Application.ScreenUpdating = True
    
    If Err.Number <> 0 Then fMsgBox Err.Description
End Sub

Function fValidateSheet(Optional bErrMsgBox As Boolean = True) As Boolean
    On Error GoTo exit_sub
    
    Dim lErrRowNo As Long, lErrColNo As Long
    
    Call fTrimAllCellsForSheet(Me)
    
    Dim arrData()
'    Dim dictColIndex As Dictionary

    fInitialization
    gsRptID = "CALCULATE_PROFIT"
    Call fReadSysConfig_InputTxtSheetFile
'
'    Call fReadSheetDataByConfig("SELF_PURCHASE_ORDER", dictColIndex, arrData, , , , , Me)
    
'    Call fValidateDuplicateInArray(arrData, Array(dictColIndex("ProductProducer"), dictColIndex("ProductName")) _
'                    , False, me, 1, 1, "生产厂家+药品名称")
    
    Call fCopyReadWholeSheetData2Array(Me, arrData)
    Call fValidateDuplicateInArray(arrData, Array(CZLRollover.ProductProducer, CZLRollover.ProductName, CZLRollover.ProductSeries, CZLRollover.LotNum) _
                    , False, Me, 1, 1, " 生产厂家+药品名称+药品规格")
     
    Call fValidateBlankInArray(arrData, CZLRollover.ProductProducer, Me, 1, 1, "生产厂家")
    Call fValidateBlankInArray(arrData, CZLRollover.ProductName, Me, 1, 1, "药品名称")
    Call fValidateBlankInArray(arrData, CZLRollover.ProductSeries, Me, 1, 1, "药品规格")
    Call fValidateBlankInArray(arrData, CZLRollover.ProductUnit, Me, 1, 1, "药品单位")
    'Call fValidateBlankInArray(arrData, dictColIndex("PurchasePrice"), Me, 1, 1, "出货单价")
    'Call fValidateBlankInArray(arrData, dictColIndex("LotNum"), Me, 1, 1, "批号")
    
    Call fSortDataInSheetSortSheetData(Me, Array(CZLRollover.ProductProducer, CZLRollover.ProductName, CZLRollover.ProductSeries, CZLRollover.LotNum))

    Call fCheckIfProductExistsInProductMaster(arrData, CZLRollover.ProductProducer, CZLRollover.ProductName, CZLRollover.ProductSeries, lErrRowNo, lErrColNo)

    'Call fCheckIfCompanyNameExistsInrngStaticSalesCompanyNames(arrData, CZLRollover.SalesCompany, "[商业公司]", lErrRowNo, lErrColNo)
    
    If bErrMsgBox Then fMsgBox "[" & Me.Name & "]表 保存成功", vbInformation: ThisWorkbook.Save
exit_sub:
    fEnableExcelOptionsAll
    Set dictColIndex = Nothing
    Erase arrData
    
    If fCheckIfGotBusinessError Then
        fValidateSheet = False
    Else
        If fCheckIfUnCapturedExceptionAbnormalError Then
            fValidateSheet = False
        Else
            fShowAndActiveSheet Me
            fValidateSheet = True
        End If
    End If
    
    If lErrRowNo > 0 Then
        fShowAndActiveSheet Me
        Application.Goto Me.Cells(lErrRowNo, lErrColNo) ', True
    End If
'    If Err.Number <> 0 Then
'        fShowAndActiveSheet Me
'        fValidateSheet = False
'    Else
'        fValidateSheet = True
'    End If
End Function


