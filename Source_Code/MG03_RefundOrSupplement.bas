Attribute VB_Name = "MG03_RefundOrSupplement"
Option Explicit
Option Base 1

Dim dictCZLSCompProdPrice As Dictionary
Dim dictCZLProductPrice As Dictionary
Dim arrCZLSales()
'Dim dictCZLSalesReverse As Dictionary
Dim dictCZLSalesDeduct As Dictionary
Dim dictCZLSalesMinus As Dictionary

Sub subMain_CalculateRefundOrSupplement()
    If Not fIsDev Then On Error GoTo error_handling
    
    fCheckIfErrCountNotZero_SCompSalesInfo
    fCheckIfErrCountNotZero_CZLSales2Comp
    
    fRemoveFilterForSheet shtSalesInfos
    fRemoveFilterForSheet shtRefund
    fRemoveFilterForSheet shtCZLSales2SCompAll
    fRemoveFilterForSheet shtCZLSales2Companies
    
    Call fShowSheet(shtSalesInfos)
    Call fHideSheet(shtException)
    If fIsDev Then Call fShowSheet(shtRefund):   Call fShowSheet(shtException)
    Call fShowSheet(shtCZLSales2SCompAll)
    
    Call fUnProtectSheet(shtRefund)
    Call fDeleteRowsFromSheetLeaveHeader(shtRefund)
    Call fDeleteRowsFromSheetLeaveHeader(shtCZLSales2SCompAll)
    Call fDeleteRowsFromSheetLeaveHeader(shtException)

    fInitialization

    gsRptID = "CALCULATE_PROFIT"
    Call fReadSysConfig_InputTxtSheetFile
    
    Call fCombineThisAndHistToSheetCZLSales2SCompAll
    
    'master file
    'sort should be same between shtSalesInfos and shtCZLSales2SCompAll
    Call fSortDataInSheetSortSheetData(shtSalesInfos, Array(Sales2Hospital.SalesDate, Sales2Hospital.SalesCompany _
                                                            , Sales2Hospital.ProductProducer, Sales2Hospital.ProductName, Sales2Hospital.ProductSeries))
    Call fCopyReadWholeSheetData2Array(shtSalesInfos, arrMaster)
    
    Call fProcessData
    
    Call fAppendArray2Sheet(shtRefund, arrOutput)
    
    Call fBasicCosmeticFormatSheet(shtRefund, Refund.[_last])
    If dictErrorRows.Count <= 0 And dictErrorRows.Count <= 0 Then
        Call fSetConditionFormatForOddEvenLine(shtRefund, Refund.[_last])
    Else
        Call fDeleteAllConditionFormatFromSheet(shtRefund)
    End If
    
    Call fSetBorderLineForSheet(shtRefund, Refund.[_last])
    
    shtRefund.Visible = xlSheetVisible
    shtRefund.Activate
    fGotoCell shtRefund.Range("A1")
error_handling:
    If fCheckIfUnCapturedExceptionAbnormalError Then GoTo reset_excel_options
    
    If dictErrorRows.Count > 0 Or dictErrorRows.Count > 0 Then
        shtException.Visible = xlSheetVisible
        
        Dim lExcepMaxCol As Long
        lExcepMaxCol = fGetValidMaxCol(shtException)
        
        Call fSetFormatBoldOrangeBorderForHeader(shtException, lExcepMaxCol)
        Call fSetBorderLineForSheet(shtException, lExcepMaxCol)
        Call fBasicCosmeticFormatSheet(shtException, lExcepMaxCol)
        Call fSetFormatForOddEvenLineByFixColor(shtException, lExcepMaxCol)
        
        If Not fFindInWorksheet(shtException.Cells, "找不到可扣的出货记录", False) Is Nothing Then
            'shtException.Columns(4).ColumnWidth = 100
            Call fFreezeSheet(shtException, , 2)
        End If
        
        shtException.Activate
    End If
    
    Call fSetFormatForExceptionCells(shtRefund, dictErrorRows, "REPORT_ERROR_COLOR")
    Call fSetFormatForExceptionCells(shtRefund, dictWarningRows, "REPORT_WARNING_COLOR")
    
    If Not fCheckIfGotBusinessError(False) Then
        fMsgBox "计算完成，请检查工作表：[" & shtRefund.Name & "] 中，请检查！", vbInformation
    End If
    
    If fCheckIfGotBusinessError Then GoTo reset_excel_options
    If fCheckIfUnCapturedExceptionAbnormalError Then GoTo reset_excel_options
reset_excel_options:
    Err.Clear
    fClearRefVariables
    Set dictCZLSalesDeduct = Nothing
    Set dictCZLSalesMinus = Nothing
    Set dictCZLSCompProdPrice = Nothing
    Set dictCZLProductPrice = Nothing
            
    'Set dictCZLSalesReverse = Nothing
    Erase arrCZLSales
    fEnableExcelOptionsAll
End Sub

Private Function fGetActualNetPriceByCZLSales(sCZLSalesKey As String _
            , dblQuantity As Double, ByRef dblActualNetPrice As Double) As Boolean
    dblActualNetPrice = 0
    
    If dblQuantity > 0 Then
        fGetActualNetPriceByCZLSales = fCalculateActualPriceFromCZLSalesAllNoraml(sCZLSalesKey, dblQuantity, dblActualNetPrice)
    ElseIf dblQuantity < 0 Then
        fGetActualNetPriceByCZLSales = fCalculateActualPriceFromCZLSalesAllMinus(sCZLSalesKey, dblQuantity, dblActualNetPrice)
'    ElseIf dblQuantity < 0 Then
'        fGetActualNetPriceByCZLSales = fCalculateActualPriceFromCZLSalesAllWithdraw(sCZLSalesKey, dblQuantity, dblActualNetPrice)
    Else
        'fErr "销售数量为0"
    End If
End Function


Private Function fSetBackToshtCZLSalesCalWithDeductedData()
    Call fDeleteRowsFromSheetLeaveHeader(shtCZLSales2SCompAll)
    
    If UBound(arrCZLSales, 1) >= LBound(arrCZLSales, 1) Then
        shtCZLSales2SCompAll.Range("A2").Resize(UBound(arrCZLSales, 1), UBound(arrCZLSales, 2)).Value2 = arrCZLSales
    End If
End Function
'====================== CZL Sales to SComp=================================================================
Private Function fReadCZLSalesAll2Dictionary()
    Dim sTmpKey As String
    Dim sSalesComp As String
    Dim sProducer As String, sProductName As String, sProductSeries As String
    Dim dblSellQuantity As Double
    Dim dblHospitalQuantity As Double
    Dim dictMinusTo As Dictionary
    Dim dictDeductTo As Dictionary
    'Dim dictReverseTo As Dictionary
    Dim lEachRow As Long
    
    Call fSortDataInSheetSortSheetData(shtCZLSales2SCompAll, Array(CZLSales2CompHist.SalesDate, CZLSales2CompHist.SalesCompany _
                                                                 , CZLSales2CompHist.ProductProducer, CZLSales2CompHist.ProductName, CZLSales2CompHist.ProductSeries))
    Call fCopyReadWholeSheetData2Array(shtCZLSales2SCompAll, arrCZLSales)
    
    Set dictDeductTo = New Dictionary
    'Set dictReverseTo = New Dictionary
    Set dictMinusTo = New Dictionary
    Set dictCZLSalesDeduct = New Dictionary
    'Set dictCZLSalesReverse = New Dictionary
    Set dictCZLSalesMinus = New Dictionary
    
    For lEachRow = LBound(arrCZLSales, 1) To UBound(arrCZLSales, 1)
        sSalesComp = arrCZLSales(lEachRow, CZLSales2CompHist.SalesCompany)
        sProducer = arrCZLSales(lEachRow, CZLSales2CompHist.ProductProducer)
        sProductName = arrCZLSales(lEachRow, CZLSales2CompHist.ProductName)
        sProductSeries = arrCZLSales(lEachRow, CZLSales2CompHist.ProductSeries)
        
        dblSellQuantity = arrCZLSales(lEachRow, CZLSales2CompHist.Quantity)
        dblHospitalQuantity = arrCZLSales(lEachRow, CZLSales2CompHist.DeductQty)
        
        If dblSellQuantity = 0 Then GoTo next_row
        
        sTmpKey = sSalesComp & DELIMITER & sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries
            
        If dblSellQuantity < 0 Then
            If dblSellQuantity > dblHospitalQuantity Then
                fActiveVisibleSwitchSheet shtCZLSales2SCompAll
                fErr "数据出错，退货的情况下，销售数量不应该大于医院抵扣数量" _
                            & vbCr & "工作表：" & shtCZLSales2SCompAll.Name _
                            & vbCr & "行号：" & lEachRow + 1 _
                            & vbCr & "商业公司：" & sSalesComp _
                            & vbCr & "药品：" & sProducer & " , " & sProductName & ", " & sProductSeries _
                            & vbCr & vbCr & "请检查【" & shtCZLSales2SCompAll.Name & "】表。"
            End If
            If dblHospitalQuantity > 0 Then
                fActiveVisibleSwitchSheet shtCZLSales2SCompAll
                fErr "数据出错，退货的情况下，医院销售数量不应该 > 0" _
                            & vbCr & "工作表：" & shtCZLSales2SCompAll.Name _
                            & vbCr & "行号：" & lEachRow + 1 _
                            & vbCr & "商业公司：" & sSalesComp _
                            & vbCr & "药品：" & sProducer & " , " & sProductName & ", " & sProductSeries
            End If
            
            If dblSellQuantity < dblHospitalQuantity Then
                If Not dictCZLSalesMinus.Exists(sTmpKey) Then
                    dictCZLSalesMinus.Add sTmpKey, lEachRow
                End If
                dictMinusTo(sTmpKey) = lEachRow
            End If
        Else
            If dblSellQuantity < dblHospitalQuantity Then
                fActiveVisibleSwitchSheet shtCZLSales2SCompAll
                fErr "数据出错，医院抵扣数量不应该大于销售数量" _
                            & vbCr & "工作表：" & shtCZLSales2SCompAll.Name _
                            & vbCr & "行号：" & lEachRow + 1 _
                            & vbCr & "商业公司：" & sSalesComp _
                            & vbCr & "药品：" & sProducer & " , " & sProductName & ", " & sProductSeries _
                            & vbCr & vbCr & "请检查【" & shtCZLSales2SCompAll.Name & "】表。"
            End If
    
            If dblHospitalQuantity < 0 Then
                fActiveVisibleSwitchSheet shtCZLSales2SCompAll
                fErr "数据出错，医院销售数量不应该 < 0" _
                            & vbCr & "工作表：" & shtCZLSales2SCompAll.Name _
                            & vbCr & "行号：" & lEachRow + 1 _
                            & vbCr & "商业公司：" & sSalesComp _
                            & vbCr & "药品：" & sProducer & " , " & sProductName & ", " & sProductSeries
            End If
            
'            If dblHospitalQuantity > 0 Then
'                If Not dictCZLSalesReverse.Exists(sTmpKey) Then
'                    dictCZLSalesReverse.Add sTmpKey, lEachRow
'                End If
'                dictReverseTo(sTmpKey) = lEachRow
'            End If
            
            If dblSellQuantity > dblHospitalQuantity Then
                If Not dictCZLSalesDeduct.Exists(sTmpKey) Then
                    dictCZLSalesDeduct.Add sTmpKey, lEachRow
                End If
                dictDeductTo(sTmpKey) = lEachRow
            End If
        End If
next_row:
    Next
    
'    For lEachRow = 0 To dictCZLSalesReverse.Count - 1
'        dictCZLSalesReverse(dictCZLSalesReverse.Keys(lEachRow)) = dictCZLSalesReverse.Items(lEachRow) _
'                    & DELIMITER & dictReverseTo.Items(lEachRow)
'    Next
    
    For lEachRow = 0 To dictCZLSalesDeduct.Count - 1
        dictCZLSalesDeduct(dictCZLSalesDeduct.Keys(lEachRow)) = dictCZLSalesDeduct.Items(lEachRow) _
                    & DELIMITER & dictDeductTo.Items(lEachRow)
    Next
    For lEachRow = 0 To dictCZLSalesMinus.Count - 1
        dictCZLSalesMinus(dictCZLSalesMinus.Keys(lEachRow)) = dictCZLSalesMinus.Items(lEachRow) _
                    & DELIMITER & dictMinusTo.Items(lEachRow)
    Next
    
    'Set dictReverseTo = Nothing
    Set dictDeductTo = Nothing
    Set dictMinusTo = Nothing
End Function
Private Function fCalculateActualPriceFromCZLSalesAllNoraml(sCZLSalesKey As String _
                    , ByVal dblSalesQuantity As Double, ByRef dblActualPrice As Double) As Boolean
    Dim bOut As Boolean
    Dim lDeductStartRow As Long
    Dim lDeductEndRow As Long
    Dim dblCZLSellQuantity As Double
    Dim dblHospitalQuantity As Double
    Dim dblToDeduct As Double
    Dim dblBalance As Double
    Dim dblCurrRowAvailable As Double
    Dim lEachRow As Long
    Dim dblAccAmt As Double
    Dim dblPrice As Double
    
    bOut = False
    
    If dictCZLSalesDeduct Is Nothing Then Call fReadCZLSalesAll2Dictionary
    
    If Not dictCZLSalesDeduct.Exists(sCZLSalesKey) Then GoTo exit_fun
    
    lDeductStartRow = Split(dictCZLSalesDeduct(sCZLSalesKey), DELIMITER)(0)
    lDeductEndRow = Split(dictCZLSalesDeduct(sCZLSalesKey), DELIMITER)(1)
    
    dblAccAmt = 0
    dblToDeduct = dblSalesQuantity
    For lEachRow = lDeductStartRow To lDeductEndRow
        If dblToDeduct <= 0 Then Exit For
        
        dblCZLSellQuantity = arrCZLSales(lEachRow, CZLSales2CompHist.Quantity)
        dblHospitalQuantity = arrCZLSales(lEachRow, CZLSales2CompHist.DeductQty)
        dblPrice = arrCZLSales(lEachRow, CZLSales2CompHist.Price)
        
'        If dblCZLSellQuantity <= dblHospitalQuantity Then fErr "这一行的日期晚，不应该出现完全抵扣" _
'                        & vbCr & "工作表：" & shtCZLSales2SCompAll.Name _
'                        & vbCr & "行号：" & lEachRow + 1
        
        dblCurrRowAvailable = dblCZLSellQuantity - dblHospitalQuantity
        dblBalance = dblToDeduct - dblCurrRowAvailable
        
        If dblBalance >= 0 Then  'still has to find next row to deduct
            If lEachRow < lDeductEndRow Then    'move the deduct dictionary to next row
                dictCZLSalesDeduct(sCZLSalesKey) = (lEachRow + 1) & DELIMITER & lDeductEndRow
            Else
                dictCZLSalesDeduct.Remove sCZLSalesKey
            End If
            
            dblAccAmt = dblAccAmt + dblCurrRowAvailable * dblPrice
            arrCZLSales(lEachRow, CZLSales2CompHist.DeductQty) = dblCZLSellQuantity
        Else
            dblAccAmt = dblAccAmt + dblToDeduct * dblPrice
            arrCZLSales(lEachRow, CZLSales2CompHist.DeductQty) = dblCZLSellQuantity + dblBalance
        End If
        
'        'to extend dictreverse to new row
'        If dictCZLSalesReverse.Exists(sCZLSalesKey) Then
'            'If CLng(Split(dictCZLSalesReverse(sCZLSalesKey), DELIMITER)(1)) < lEachRow Then
'            dictCZLSalesReverse(sCZLSalesKey) = Split(dictCZLSalesReverse(sCZLSalesKey), DELIMITER)(0) & DELIMITER & lEachRow
'            'End If
'        Else
'            dictCZLSalesReverse.Add sCZLSalesKey, lEachRow & DELIMITER & lEachRow
'        End If
        
        dblToDeduct = dblBalance
    Next
    
    If dblToDeduct <= 0 Then
        bOut = True
        dblActualPrice = dblAccAmt / dblSalesQuantity
    End If
    
exit_fun:
    fCalculateActualPriceFromCZLSalesAllNoraml = bOut
End Function

Private Function fCalculateActualPriceFromCZLSalesAllMinus(sCZLSalesKey As String _
                    , ByVal dblSalesQuantity As Double, ByRef dblActualPrice As Double) As Boolean
    Dim bOut As Boolean
    Dim lDeductStartRow As Long
    Dim lDeductEndRow As Long
    Dim dblCZLSellQuantity As Double
    Dim dblHospitalQuantity As Double
    Dim dblToDeduct As Double
    Dim dblBalance As Double
    Dim dblCurrRowAvailable As Double
    Dim lEachRow As Long
    Dim dblAccAmt As Double
    Dim dblPrice As Double
    
    bOut = False
    
    If dictCZLSalesMinus Is Nothing Then Call fReadCZLSalesAll2Dictionary
    
    If Not dictCZLSalesMinus.Exists(sCZLSalesKey) Then GoTo exit_fun
    
    lDeductStartRow = Split(dictCZLSalesMinus(sCZLSalesKey), DELIMITER)(0)
    lDeductEndRow = Split(dictCZLSalesMinus(sCZLSalesKey), DELIMITER)(1)
    
    dblAccAmt = 0
    dblToDeduct = dblSalesQuantity
    For lEachRow = lDeductStartRow To lDeductEndRow
        If dblToDeduct >= 0 Then Exit For
        
        dblCZLSellQuantity = arrCZLSales(lEachRow, CZLSales2CompHist.Quantity)
        dblHospitalQuantity = arrCZLSales(lEachRow, CZLSales2CompHist.DeductQty)
        dblPrice = arrCZLSales(lEachRow, CZLSales2CompHist.Price)
        
'        If dblCZLSellQuantity <= dblHospitalQuantity Then fErr "这一行的日期晚，不应该出现完全抵扣" _
'                        & vbCr & "工作表：" & shtCZLSales2SCompAll.Name _
'                        & vbCr & "行号：" & lEachRow + 1
        
        dblCurrRowAvailable = dblCZLSellQuantity - dblHospitalQuantity
        dblBalance = dblToDeduct - dblCurrRowAvailable
        
        If dblBalance <= 0 Then  'still has to find next row to deduct
            If lEachRow < lDeductEndRow Then    'move the deduct dictionary to next row
                dictCZLSalesMinus(sCZLSalesKey) = (lEachRow + 1) & DELIMITER & lDeductEndRow
            Else
                dictCZLSalesMinus.Remove sCZLSalesKey
            End If
            
            dblAccAmt = dblAccAmt + dblCurrRowAvailable * dblPrice
            arrCZLSales(lEachRow, CZLSales2CompHist.DeductQty) = dblCZLSellQuantity
        Else
            dblAccAmt = dblAccAmt + dblToDeduct * dblPrice
            arrCZLSales(lEachRow, CZLSales2CompHist.DeductQty) = dblCZLSellQuantity + dblBalance
        End If
        
'        'to extend dictreverse to new row
'        If dictCZLSalesReverse.Exists(sCZLSalesKey) Then
'            'If CLng(Split(dictCZLSalesReverse(sCZLSalesKey), DELIMITER)(1)) < lEachRow Then
'            dictCZLSalesReverse(sCZLSalesKey) = Split(dictCZLSalesReverse(sCZLSalesKey), DELIMITER)(0) & DELIMITER & lEachRow
'            'End If
'        Else
'            dictCZLSalesReverse.Add sCZLSalesKey, lEachRow & DELIMITER & lEachRow
'        End If
        
        dblToDeduct = dblBalance
    Next
    
    If dblToDeduct >= 0 Then
        bOut = True
        dblActualPrice = Abs(dblAccAmt / dblSalesQuantity)
    End If
    
exit_fun:
    fCalculateActualPriceFromCZLSalesAllMinus = bOut
End Function
Private Function fCombineThisAndHistToSheetCZLSales2SCompAll()
    Dim sHistFileFullPath As String
    Dim wbMonthly As Workbook
    Dim shtMonth As Worksheet
    Dim arrData()
    
    On Error GoTo err_h
    
    'history
    Call fGetLatestCreatedMEFileCZLSales2SCompAndUpdateConfig(sHistFileFullPath, wbMonthly, shtMonth)
    
    Call fRemoveFilterForSheet(shtMonth)
    Call fCopyReadWholeSheetData2Array(shtMonth, arrData)
    Call fAppendArray2Sheet(shtCZLSales2SCompAll, arrData)
    
    'this
    Call fCopyReadWholeSheetData2Array(shtCZLSales2Companies, arrData)
    
    Dim arrThis()
    Call fSubstractArray(arrData, arrThis)
    Erase arrData
    Call fAppendArray2Sheet(shtCZLSales2SCompAll, arrThis)
    Erase arrThis
    
    Call fSetBorderLineForSheet(shtCZLSales2SCompAll)
    
err_h:
    If Not wbMonthly Is Nothing Then fCloseWorkBookWithoutSave wbMonthly
    If gErrNum <> 0 Then fErr
End Function
Private Function fSubstractArray(ByRef arrFrom(), ByRef arrThis())
    Dim lEachRow As Long
    Dim i As Integer
    Dim arrSourceCols()
    Dim arrToCols()
    
    If UBound(arrFrom, 1) < LBound(arrFrom, 1) Then arrThis = Array(): Exit Function
    
    arrSourceCols = Array(CZLSales2Comp.SalesCompany, CZLSales2Comp.SalesDate _
                        , CZLSales2Comp.ProductProducer, CZLSales2Comp.ProductName, CZLSales2Comp.ProductSeries _
                        , CZLSales2Comp.ProductUnit, CZLSales2Comp.LotNum _
                        , CZLSales2Comp.ConvertedQuantity, CZLSales2Comp.ConvertedPrice, CZLSales2Comp.RecalAmount)
    arrToCols() = Array(CZLSales2CompHist.SalesCompany, CZLSales2CompHist.SalesDate _
                      , CZLSales2CompHist.ProductProducer, CZLSales2CompHist.ProductName, CZLSales2CompHist.ProductSeries _
                      , CZLSales2CompHist.ProductUnit, CZLSales2CompHist.LotNum _
                      , CZLSales2CompHist.Quantity, CZLSales2CompHist.Price, CZLSales2CompHist.Amount _
                      )
    ReDim arrThis(LBound(arrFrom, 1) To UBound(arrFrom, 1), CZLSales2CompHist.[_first] To CZLSales2CompHist.[_last])
    
    For lEachRow = LBound(arrFrom, 1) To UBound(arrFrom, 1)
        For i = LBound(arrSourceCols) To UBound(arrSourceCols)
            arrThis(lEachRow, arrToCols(i)) = arrFrom(lEachRow, arrSourceCols(i))
        Next
    Next
    
    Erase arrSourceCols
    Erase arrToCols
End Function
 
Private Function fProcessData()
    Dim lEachRow As Long
    Dim dictMissedSecondLComm As Dictionary
    Dim sHospital As String
    Dim sSalesCompName As String
    Dim sSalesCompNameID As String
    Dim sSalesCompID As String
    Dim sProducer As String
    Dim sProductName  As String
    Dim sProductSeries As String
    Dim sLotNum As String
    Dim sSecondLevelCommKey As String
    Dim sSecondLevelCommPasteKey As String
    Dim sProductKey As String
    Dim sCZLSalesKey As String
    Dim sMsg As String
    
    Dim dblPromPrdRebate As Double
    Dim dblSalesTaxRate As Double
    Dim dblPurchaseTaxRate As Double
    Dim bIsPromotionProduct As Boolean
    
    Dim dblQuantity As Double
    Dim dblSellPrice As Double
    Dim dblDueNetPrice As Double
    Dim dblSecondLevelComm As Double
    Dim dblActualNetPrice As Double
    Dim sLastActualNetPrice As String
    Dim sCZLName As String
    Dim dictNoValidCZLSales As Dictionary
    
    Set dictErrorRows = New Dictionary
    Set dictNoValidCZLSales = New Dictionary
 
    sCZLName = fGetCompanyNameByID_Common("CZL")
    Set dictMissedSecondLComm = New Dictionary

    ReDim arrOutput(1 To UBound(arrMaster, 1), 1 To Refund.[_last])
    
    For lEachRow = LBound(arrMaster, 1) To UBound(arrMaster, 1)
        sHospital = Trim(arrMaster(lEachRow, Sales2Hospital.Hospital))
        sSalesCompName = Trim(arrMaster(lEachRow, Sales2Hospital.SalesCompany))
        sProducer = Trim(arrMaster(lEachRow, Sales2Hospital.ProductProducer))
        sProductName = Trim(arrMaster(lEachRow, Sales2Hospital.ProductName))
        sProductSeries = Trim(arrMaster(lEachRow, Sales2Hospital.ProductSeries))
        dblQuantity = arrMaster(lEachRow, Sales2Hospital.ConvertedQuantity)
        dblSellPrice = arrMaster(lEachRow, Sales2Hospital.ConvertedPrice)
        sLotNum = arrMaster(lEachRow, Sales2Hospital.LotNum)
        
        sProductKey = sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries
    
        arrOutput(lEachRow, Refund.Hospital) = sHospital
        arrOutput(lEachRow, Refund.SalesCompany) = sSalesCompName
        arrOutput(lEachRow, Refund.ProductProducer) = sProducer
        arrOutput(lEachRow, Refund.ProductName) = sProductName
        arrOutput(lEachRow, Refund.ProductSeries) = sProductSeries
        arrOutput(lEachRow, Refund.SalesDate) = arrMaster(lEachRow, Sales2Hospital.SalesDate)
        arrOutput(lEachRow, Refund.ProductUnit) = arrMaster(lEachRow, Sales2Hospital.ProductUnit)
        arrOutput(lEachRow, Refund.Quantity) = dblQuantity
        arrOutput(lEachRow, Refund.BidPrice) = dblSellPrice
        arrOutput(lEachRow, Refund.LotNum) = "'" & sLotNum
        
        '==== second level commission ==========================================
        sSecondLevelCommKey = sSalesCompName & DELIMITER & sHospital & DELIMITER _
                            & sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries
    
        If Not fGetSecondLevelComm(sSecondLevelCommKey, dblSecondLevelComm) Then
            dblSecondLevelComm = fGetConfigSecondLevelDefaultComm(sSalesCompName)
            
            sSecondLevelCommPasteKey = fComposeSecondLevelColumnsStryByConfig(sSalesCompName, sHospital _
                                                        , sProducer, sProductName, sProductSeries, dblSecondLevelComm)
            If Not dictMissedSecondLComm.Exists(sSecondLevelCommPasteKey) Then
                dictMissedSecondLComm.Add sSecondLevelCommPasteKey, "'" & (lEachRow + 1)
            Else
                dictMissedSecondLComm(sSecondLevelCommPasteKey) = dictMissedSecondLComm(sSecondLevelCommPasteKey) & "," & (lEachRow + 1)
            End If
        End If
        '-----------------------------------------------------------------------------------------------
        sCZLSalesKey = sSalesCompName & DELIMITER & sProductKey
        
        sMsg = ""
        bIsPromotionProduct = fIsPromotionProduct(sHospital, sProductKey, dblSellPrice, sSalesCompName, dblPromPrdRebate, dblSalesTaxRate, dblPurchaseTaxRate, dblSecondLevelComm, 0)
        If bIsPromotionProduct Then
            dblDueNetPrice = 0
            dblActualNetPrice = 0
            arrOutput(lEachRow, Refund.DueNetPrice) = "推广品"
            arrOutput(lEachRow, Refund.ActualNetPrice) = "推广品"
            arrOutput(lEachRow, Refund.PriceDeviation) = "推广品"
            arrOutput(lEachRow, Refund.AmountDeviation) = "推广品"
        Else
            If sSalesCompName = sCZLName Then   '采芝林
                arrOutput(lEachRow, Refund.DueNetPrice) = sCZLName
                arrOutput(lEachRow, Refund.ActualNetPrice) = sCZLName
                arrOutput(lEachRow, Refund.PriceDeviation) = sCZLName
                arrOutput(lEachRow, Refund.AmountDeviation) = sCZLName
            Else
                dblDueNetPrice = dblSellPrice * (1 - dblSecondLevelComm)
                arrOutput(lEachRow, Refund.DueNetPrice) = dblDueNetPrice
                
                'Call fGetActualNetPrice(sSalesCompName, sProductKey, sLotNum, oActualNetPrice)
                If Not fGetActualNetPriceByCZLSales(sCZLSalesKey, dblQuantity, dblActualNetPrice) Then
                    sLastActualNetPrice = GetAvailableActualNetPrices(sCZLSalesKey, sProductKey)
                    
                    If Len(Trim(sLastActualNetPrice)) <= 0 Then
                        dblActualNetPrice = 0
                        sMsg = "该商业公司+药品在采芝林的销售记录中没有可扣数量，因此找不到其准确的供货价。并且该药品也没有采芝林历史销售记录，故找不到任何可用的实际供货价。"
                        Call fAddErrorColumnTodictErrorRows(lEachRow + 1, Refund.ActualNetPrice)
                    Else
                        dblActualNetPrice = Split(sLastActualNetPrice, "~")(0)
                        sMsg = "该商业公司+药品在采芝林的销售记录中没有可扣数量，因此找不到其准确的供货价，现找到了所有历史实际供货价，并按照第一个价格计算了补差，请核对。(第一个为最近一次的价格。)"
                        Call fAddWarningColumnTodictWarningRows(lEachRow + 1, Refund.ActualNetPrice)
                    End If
            
                    If Not dictNoValidCZLSales.Exists(sCZLSalesKey) Then
                        dictNoValidCZLSales.Add sCZLSalesKey, sLastActualNetPrice & DELIMITER & sMsg & DELIMITER & (lEachRow + 1)
                    Else
                        dictNoValidCZLSales(sCZLSalesKey) = dictNoValidCZLSales(sCZLSalesKey) & "," & (lEachRow + 1)
                    End If
                    
                    arrOutput(lEachRow, Refund.ActualNetPrice) = sLastActualNetPrice
                Else
                    arrOutput(lEachRow, Refund.ActualNetPrice) = dblActualNetPrice
                End If

                arrOutput(lEachRow, Refund.PriceDeviation) = arrOutput(lEachRow, Refund.DueNetPrice) - dblActualNetPrice
                arrOutput(lEachRow, Refund.AmountDeviation) = arrOutput(lEachRow, Refund.PriceDeviation) * arrOutput(lEachRow, Refund.Quantity)
            End If
        End If
next_sales:
    Next
    
    Call fSetBackToshtCZLSalesCalWithDeductedData
    Call fAddNoValidCZLSalesToSheetException(dictNoValidCZLSales)
'    Call fAddNoSalesManConfToSheetException(dictNoSalesManConf)
    Set dictNoValidCZLSales = Nothing
End Function

'=============================================================
Private Function fReadCZLSales2SCompForAllPrices()
    Dim i As Long
    Dim sFullKeyStr As String
    Dim sKeyStr As String
    Dim dblPrice As Double
    Dim lMaxRow As Long
    Dim lMaxCol As Long
    
    Call fSortDataInSheetSortSheetData(shtCZLSales2SCompAll, Array(CZLSales2CompHist.SalesDate, CZLSales2CompHist.SalesCompany _
                                                                , CZLSales2CompHist.ProductProducer, CZLSales2CompHist.ProductName, CZLSales2CompHist.ProductSeries))
    lMaxRow = fGetValidMaxRow(shtCZLSales2SCompAll)
    lMaxCol = fGetValidMaxCol(shtCZLSales2SCompAll)
        
    Dim arrData()
    Call fCopyReadWholeSheetData2Array(shtCZLSales2SCompAll, arrData)
    
    Set dictCZLSCompProdPrice = New Dictionary
    Set dictCZLProductPrice = New Dictionary
    
    For i = UBound(arrData, 1) To LBound(arrData, 1) Step -1
        sFullKeyStr = Trim(arrData(i, CZLSales2CompHist.SalesCompany)) & DELIMITER _
                & Trim(arrData(i, CZLSales2CompHist.ProductProducer)) & DELIMITER _
                & Trim(arrData(i, CZLSales2CompHist.ProductName)) & DELIMITER _
                & Trim(arrData(i, CZLSales2CompHist.ProductSeries))
        sKeyStr = Trim(arrData(i, CZLSales2CompHist.ProductProducer)) & DELIMITER _
                & Trim(arrData(i, CZLSales2CompHist.ProductName)) & DELIMITER _
                & Trim(arrData(i, CZLSales2CompHist.ProductSeries))
        
        If fZero(Replace(sKeyStr, DELIMITER, "")) Then GoTo next_row
        
        dblPrice = arrData(i, CZLSales2CompHist.Price)
        
        If dictCZLSCompProdPrice.Exists(sFullKeyStr) Then
            If InStr("~" & dictCZLSCompProdPrice(sFullKeyStr) & "~", "~" & dblPrice & "~") <= 0 Then
                dictCZLSCompProdPrice(sFullKeyStr) = dictCZLSCompProdPrice(sFullKeyStr) & "~" & CStr(dblPrice)
            End If
        Else
            dictCZLSCompProdPrice.Add sFullKeyStr, CStr(dblPrice)
        End If
        
        If dictCZLProductPrice.Exists(sKeyStr) Then
            If InStr("~" & dictCZLProductPrice(sKeyStr) & "~", "~" & dblPrice & "~") <= 0 Then
                dictCZLProductPrice(sKeyStr) = dictCZLProductPrice(sKeyStr) & "~" & dblPrice
            End If
        Else
            dictCZLProductPrice.Add sKeyStr, CStr(dblPrice)
        End If
next_row:
    Next
    Erase arrData
    
'    For i = 0 To dictCZLSCompProdPrice.Count - 1
'        sFullKeyStr = dictCZLSCompProdPrice.Keys(i)
'        sKeyStr = Right(sFullKeyStr, Len(sFullKeyStr) - InStr(sFullKeyStr, DELIMITER))
'
'        If dictCZLProductPrice.Exists(sKeyStr) Then
'            dictCZLProductPrice.Remove sKeyStr
'        End If
'    Next
End Function
Private Function GetAvailableActualNetPrices(sCZLSalesKey As String, sProductKey As String) As String
    Dim sPrices As String
    
    If dictCZLSCompProdPrice Is Nothing Then Call fReadCZLSales2SCompForAllPrices

    If dictCZLSCompProdPrice.Exists(sCZLSalesKey) Then
        sPrices = dictCZLSCompProdPrice(sCZLSalesKey)
    Else
        If dictCZLProductPrice.Exists(sProductKey) Then
            sPrices = dictCZLProductPrice(sProductKey)
        Else
            sPrices = ""
        End If
    End If
    GetAvailableActualNetPrices = sPrices
End Function
'-----------------------------------------------------------------------------------

Private Function fAddNoValidCZLSalesToSheetException(dictNoValidCZLSales As Dictionary)
    Dim arrLeftPart()
    Dim lUniqRecCnt As Long
    Dim lRecCount As Long
    Dim i As Integer
    Dim j As Integer
    Dim lStartRow As Long
    
    lUniqRecCnt = dictNoValidCZLSales.Count
    If lUniqRecCnt > 0 Then
        lStartRow = fGetshtExceptionNewRow
        arrLeftPart = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictNoValidCZLSales, , False)
        
        'shtException.Columns(4).ColumnWidth = 100
        shtException.Cells(lStartRow - 1, 1).Value = "找不到可扣的采芝林销售流向(若为退货，则是先找退货，找不到再找医院销售，去抵扣)"
        shtException.Cells(lStartRow - 1, 1).WrapText = False
        Call fPrepareHeaderToSheet(shtException, Array("商业公司", "药品厂家", "药品名称", "规格", "历史价格(供参考)", "借误信息", "行号"), lStartRow)
        shtException.Rows(lStartRow - 1 & ":" & lStartRow).Font.Color = RGB(255, 0, 0)
        shtException.Rows(lStartRow - 1 & ":" & lStartRow).Font.Bold = True
        Call fAppendArray2Sheet(shtException, arrLeftPart, , False)
        
        lRecCount = fGetDictionayDelimiteredItemsCount(dictNoValidCZLSales)
        
        shtException.Cells(lStartRow + 1, UBound(arrLeftPart, 2) + 1).Resize(dictNoValidCZLSales.Count, 3).Value = fConvertDictionaryDelimiteredItemsTo2DimenArrayForPaste(dictNoValidCZLSales, , True)
        Erase arrLeftPart
        If lStartRow = 2 Then Call fFreezeSheet(shtException, , 2)
        
        shtException.Visible = xlSheetVisible
        shtException.Activate
        
        If fNzero(gsBusinessErrorMsg) Then gsBusinessErrorMsg = gsBusinessErrorMsg & vbCr & vbCr & vbCr & "===============================" & vbCr & vbCr

        gsBusinessErrorMsg = gsBusinessErrorMsg & dictNoValidCZLSales.Count & "个销售流向的实际供货价找不到，因为采芝林没有发过这样的贷给那些商业公司，" _
                        & "或者所发过的货已被完全扣除，请确认您是否忘记导入了采芝林某些月份的销售流向，或者采芝林本身的销售流向中有遗漏数据。"
    End If
End Function


Sub subMain_SaveCZLSales2SCompTableToHistory()
    Dim sHistFileFullPath As String
    Dim wbMonthly As Workbook
    Dim shtMonth As Worksheet
    
    If Not fIsDev() Then On Error GoTo error_handling

    fInitialization
    
    'If Not fPromptToConfirmToContinue("您确定要把本软件中的【采芝林销售流向(到商业公司)表】添加到历史文件中去吗？") Then fErr
    If Not fPromptToConfirmToContinue("您确定要把本软件中的【采芝林销售流向(本次+历史)(试算表)】添加到历史文件中去吗？") Then fErr
    
    Call fGetLatestCreatedMEFileCZLSales2SCompAndUpdateConfig(sHistFileFullPath, wbMonthly, shtMonth)
    
    Dim arrData()
    Dim lPasteStartRow As Long
    Dim lMaxRow As Long
    
    lMaxRow = fGetValidMaxRow(shtCZLSales2SCompAll)     'shtCZLSales2Companies
    lPasteStartRow = fGetValidMaxRow(shtMonth) + 1
    Call fCopyReadWholeSheetData2Array(shtCZLSales2SCompAll, arrData)
    Call fDeleteRowsFromSheetLeaveHeader(shtMonth)
    Call fAppendArray2Sheet(shtMonth, arrData)
    Erase arrData
    
'    Dim arrSource()
'    Dim arrDest()
'    Dim i As Integer
'    Dim lColToCopy As Integer
'    Dim lDestCol As Integer
'
'    arrSource = Array(CZLSales2Comp.SalesCompany, CZLSales2Comp.SalesDate, CZLSales2Comp.ProductProducer _
'                    , CZLSales2Comp.ProductName, CZLSales2Comp.ProductSeries, CZLSales2Comp.ProductUnit _
'                    , CZLSales2Comp.LotNum, CZLSales2Comp.ConvertedQuantity, CZLSales2Comp.ConvertedPrice, CZLSales2Comp.RecalAmount)
'    arrDest = Array(CZLSales2CompHist.SalesCompany, CZLSales2CompHist.SalesDate, CZLSales2CompHist.ProductProducer _
'                    , CZLSales2CompHist.ProductName, CZLSales2CompHist.ProductSeries, CZLSales2CompHist.ProductUnit _
'                    , CZLSales2CompHist.LotNum, CZLSales2CompHist.Quantity, CZLSales2CompHist.Price, CZLSales2CompHist.RecalAmount)
'
'    For i = LBound(arrSource) To UBound(arrSource)
'        'lColToCopy = CZLSales2Comp.SalesCompany
'        lColToCopy = arrSource(i)
'        lDestCol = arrDest(i)
'        arrData = fReadRangeDatatoArrayByStartEndPos(shtCZLSales2Companies, 2, CLng(lColToCopy), lMaxRow, CLng(lColToCopy))
'
'        If Not fArrayIsEmpty(arrData) Then
'            If lColToCopy = CZLSales2Comp.LotNum Then fConvertArrayColToText arrData, 1
'            shtMonth.Cells(lPasteStartRow, lDestCol).Resize(UBound(arrData, 1), 1) = arrData
'        End If
'
'        Erase arrData
'    Next
    
    Call fBasicCosmeticFormatSheet(shtMonth)
    
    Call fSetConditionFormatForOddEvenLine(shtMonth)
    
    Call fSetBorderLineForSheet(shtMonth)

    Set shtMonth = Nothing
    fSaveAndCloseWorkBook wbMonthly
    
error_handling:
    Set shtMonth = Nothing
    If Not wbMonthly Is Nothing Then fCloseWorkBookWithoutSave wbMonthly
    If fCheckIfGotBusinessError Then GoTo reset_excel_options
    If fCheckIfUnCapturedExceptionAbnormalError Then GoTo reset_excel_options
    
    fMsgBox "计算补差后的采芝林销售流向(本次+历史)已经保存到历史文件中： " & vbCr & sHistFileFullPath, vbInformation
reset_excel_options:
    Err.Clear
    fClearRefVariables
    fEnableExcelOptionsAll
    'End
End Sub
