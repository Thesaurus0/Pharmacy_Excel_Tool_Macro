VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtCZLPurchaseOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub btnValidate_Click()
    Call fValidateSheet
End Sub

Function fValidateSheet(Optional bErrMsgBox As Boolean = True) As Boolean
    On Error GoTo exit_sub
    
    Call fTrimAllCellsForSheet(Me)
    
    Dim arrData()
    Dim dictColIndex As Dictionary
    Dim lErrRowNo As Long
    Dim lErrColNo As Long
    
    fInitialization
    gsRptID = "CALCULATE_PROFIT"
    Call fReadSysConfig_InputTxtSheetFile
    
    Call fReadSheetDataByConfig("SELF_SALES_ORDER", dictColIndex, arrData, , , , , Me)
    
'    Call fValidateDuplicateInArray(arrData, Array(dictColIndex("ProductProducer"), dictColIndex("ProductName")) _
'                    , False, me, 1, 1, "��������+ҩƷ����")
    
    Call fValidateBlankInArray(arrData, dictColIndex("ProductProducer"), Me, 1, 1, "��������")
    Call fValidateBlankInArray(arrData, dictColIndex("ProductName"), Me, 1, 1, "ҩƷ����")
    Call fValidateBlankInArray(arrData, dictColIndex("ProductSeries"), Me, 1, 1, "ҩƷ���")
    Call fValidateBlankInArray(arrData, dictColIndex("ProductUnit"), Me, 1, 1, "ҩƷ��λ")
    Call fValidateDateColInArray(arrData, dictColIndex("SalesDate"), Me, 1, 1, "���۳�������")
    'Call fValidateBlankInArray(arrData, dictColIndex("SellPrice"), Me, 1, 1, "��������")
   ' Call fValidateBlankInArray(arrData, dictColIndex("LotNum"), Me, 1, 1, "����")
    
    Call fSortDataInSheetSortSheetData(Me, Array(dictColIndex("SalesDate") _
                                                , dictColIndex("ProductProducer") _
                                                , dictColIndex("ProductName"), dictColIndex("ProductUnit")))

    Call fCheckIfProductExistsInProductMaster(arrData, dictColIndex("ProductProducer"), dictColIndex("ProductName"), dictColIndex("ProductSeries"), lErrRowNo, lErrColNo)

    Call fCheckIfLotNumExistsInSelfPurchaseOrder(arrData, dictColIndex("ProductProducer"), dictColIndex("ProductName"), dictColIndex("ProductSeries"), dictColIndex("LotNum"), lErrRowNo, lErrColNo)
    
    Call fCheckIfSelfSellAmountIsGreaterThanPurchaseByLotNumber(arrData, dictColIndex("ProductProducer"), dictColIndex("ProductName"), dictColIndex("ProductSeries"), dictColIndex("LotNum"), lErrRowNo, lErrColNo)

    If bErrMsgBox Then fMsgBox "[" & Me.Name & "]�� û�з��ִ���", vbInformation
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
End Function

Private Sub Worksheet_Change(ByVal Target As Range)
    fResetdictSelfSalesOD
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error GoTo exit_sub
    Application.ScreenUpdating = False
    
    Const ProducerCol = 1
    Const ProductNameCol = 2
    Const ProductSeriesCol = 3
    Const ProductUnitCol = 4
    Const SellQuantityCol = 6
    Const SellPriceCol = 7
    Const LotNumCol = 8
    
    Dim sLotNum As String
    
    Dim rgIntersect As Range
    Set rgIntersect = Intersect(Target, Me.Columns(ProductNameCol))
    
    If Not rgIntersect Is Nothing Then
        If rgIntersect.Areas.Count > 1 Then GoTo exit_sub    'fErr "����ѡ���"
        If rgIntersect.Rows.Count <> 1 Then GoTo exit_sub
            
        Dim sProducer As String
        Dim sValidationListAddr As String
        
        sProducer = rgIntersect.Offset(0, ProducerCol - ProductNameCol).Value
        
        If fNzero(sProducer) Then
            Call fSetFilterForSheet(shtProductNameMaster, 1, sProducer)
            Call fCopyFilteredDataToRange(shtProductNameMaster, 2)
            
            sValidationListAddr = "=" & shtDataStage.Columns("A").Address(external:=True)
            'Call fSetValidationListForshtProductNameReplace_ProductName(sValidationListAddr, 3)
            Call fSetValidationListForRange(rgIntersect, sValidationListAddr)
        End If
    Else
        'product SeriesCol
        Set rgIntersect = Intersect(Target, Me.Columns(ProductSeriesCol))
        
        If Not rgIntersect Is Nothing Then
            If rgIntersect.Areas.Count > 1 Then GoTo exit_sub    'fErr "����ѡ���"
            If rgIntersect.Rows.Count <> 1 Then GoTo exit_sub
            
            sProducer = rgIntersect.Offset(0, ProducerCol - ProductSeriesCol).Value
            sProductName = rgIntersect.Offset(0, ProductNameCol - ProductSeriesCol).Value
            
            If fNzero(sProducer) And fNzero(sProductName) Then
                Call fSetFilterForSheet(shtProductMaster, Array(1, 2), Array(sProducer, sProductName))
                Call fCopyFilteredDataToRange(shtProductMaster, 3)
                
                sValidationListAddr = "=" & shtDataStage.Columns("A").Address(external:=True)
                'Call fSetValidationListForshtProductNameReplace_ProductName(sValidationListAddr, 3)
                Call fSetValidationListForRange(rgIntersect, sValidationListAddr)
            End If
        Else
            'product Unit
            Set rgIntersect = Intersect(Target, Me.Columns(ProductUnitCol))
            
            If Not rgIntersect Is Nothing Then
                If rgIntersect.Areas.Count > 1 Then GoTo exit_sub    'fErr "����ѡ���"
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
                'Lot Number
                Set rgIntersect = Intersect(Target, Me.Columns(LotNumCol))

                If Not rgIntersect Is Nothing Then
                    If rgIntersect.Areas.Count > 1 Then GoTo exit_sub    'fErr "����ѡ���"
                    If rgIntersect.Rows.Count <> 1 Then GoTo exit_sub

                    sProducer = rgIntersect.Offset(0, ProducerCol - LotNumCol).Value
                    sProductName = rgIntersect.Offset(0, ProductNameCol - LotNumCol).Value
                    sProductSeries = rgIntersect.Offset(0, ProductSeriesCol - LotNumCol).Value

                    If fNzero(sProducer) And fNzero(sProductName) And fNzero(sProductSeries) Then
                        Call fSetFilterForSheet(shtSelfPurchaseOrder, Array(1, 2, 3), Array(sProducer, sProductName, sProductSeries))
                        Call fCopyFilteredDataToRange(shtSelfPurchaseOrder, 8)

                        sValidationListAddr = "=" & shtDataStage.Columns("A").Address(external:=True)
                        'Call fSetValidationListForshtProductNameReplace_ProductName(sValidationListAddr, 3)
                        Call fSetValidationListForRange(rgIntersect, sValidationListAddr)
                    End If
                Else
                    'Sell Price
                    Set rgIntersect = Intersect(Target, Me.Columns(SellPriceCol))
    
                    If Not rgIntersect Is Nothing Then
'                        If rgIntersect.Areas.Count > 1 Then GoTo Exit_Sub    'fErr "����ѡ���"
'                        If rgIntersect.Rows.Count <> 1 Then GoTo Exit_Sub
'
'                        sProducer = rgIntersect.Offset(0, ProducerCol - SellPriceCol).Value
'                        sProductName = rgIntersect.Offset(0, ProductNameCol - SellPriceCol).Value
'                        sProductSeries = rgIntersect.Offset(0, ProductSeriesCol - SellPriceCol).Value
'                        sLotNum = rgIntersect.Offset(0, LotNumCol - SellPriceCol).Value
'
'                        If fNzero(sProducer) And fNzero(sProductName) Then
'                            Call fSetFilterForSheet(shtSelfPurchaseOrder, Array(1, 2, 3, 8), Array(sProducer, sProductName, sProductSeries, sLotNum))
'                            Call fCopyFilteredDataToRange(shtSelfPurchaseOrder, 7)
'
'                            sValidationListAddr = "=" & shtDataStage.Columns("A").Address(external:=True)
'                            'Call fSetValidationListForshtProductNameReplace_ProductName(sValidationListAddr, 3)
'                            Call fSetValidationListForRange(rgIntersect, sValidationListAddr)
'                        End If
                    Else
'                        'Sell Quantity
'                        Set rgIntersect = Intersect(Target, Me.Columns(SellQuantityCol))
'
'                        If Not rgIntersect Is Nothing Then
'                            If rgIntersect.Areas.Count > 1 Then GoTo exit_sub    'fErr "����ѡ���"
'                            If rgIntersect.Rows.Count <> 1 Then GoTo exit_sub
'
'                            sProducer = rgIntersect.Offset(0, ProducerCol - SellQuantityCol).Value
'                            sProductName = rgIntersect.Offset(0, ProductNameCol - SellQuantityCol).Value
'                            sProductSeries = rgIntersect.Offset(0, ProductSeriesCol - SellQuantityCol).Value
'                            sLotNum = rgIntersect.Offset(0, LotNumCol - SellQuantityCol).Value
'
'                            If fNzero(sProducer) And fNzero(sProductName) Then
'                                Call fSetFilterForSheet(shtSelfPurchaseOrder, Array(1, 2, 3, 8), Array(sProducer, sProductName, sProductSeries, sLotNum))
'                                Call fCopyFilteredDataToRange(shtSelfPurchaseOrder, 6)
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
    
exit_sub:
    fEnableExcelOptionsAll
    Application.ScreenUpdating = True
    
    If Err.Number <> 0 Then fMsgBox Err.Description
End Sub
