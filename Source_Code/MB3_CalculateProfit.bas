Attribute VB_Name = "MB3_CalculateProfit"
Option Explicit
Option Base 1
 
Dim dictFirstCommColIndex As Dictionary
Dim dictSecondCommColIndex As Dictionary

Public dictErrorRows As Dictionary
Public dictWarningRows As Dictionary

Sub subMain_CalculateProfit()
    'If Not fIsDev Then On Error GoTo error_handling
    fCheckIfErrCountNotZero_SCompSalesInfo
    
    shtSalesInfos.Visible = xlSheetVisible
    shtException.Visible = xlSheetVeryHidden
    Call fUnProtectSheet(shtProfit)
    'Call fCleanSheetOutputResetSheetOutput(shtProfit)
    Call fDeleteRowsFromSheetLeaveHeader(shtProfit)
   ' Call fCleanSheetOutputResetSheetOutput(shtException)
    Call fDeleteRowsFromSheetLeaveHeader(shtException)
    'shtException.Cells.NumberFormat = "@"
    'shtException.Cells.WrapText = True

    fInitialization

    gsRptID = "CALCULATE_PROFIT"

    Call fReadSysConfig_InputTxtSheetFile

    gsRptFilePath = fReadSysConfig_Output(, gsRptType)
    
    Call fLoadFilesAndRead2Variables

    Call fPrepareOutputSheetHeaderAndTextColumns(shtProfit)

'    ReDim arrExceptionRows(1 To UBound(arrMaster, 1) * 4)
'    mlExcepCnt = 0
    Set dictErrorRows = New Dictionary
    
    Call fProcessData
    
'    If mlExcepCnt > 0 Then
'        ReDim Preserve arrExceptionRows(1 To mlExcepCnt)
'    Else
'        arrExceptionRows = Array()
'    End If
    
    'If Not shtException.Visible = xlSheetVisible Then shtException.Visible = xlSheetVeryHidden
    
    If dictErrorRows.Count > 0 Then shtException.Visible = xlSheetVisible
     
    Call fAppendArray2Sheet(shtProfit, arrOutput)
    
    Call fBasicCosmeticFormatSheet(shtProfit)
    If dictErrorRows.Count <= 0 And dictErrorRows.Count <= 0 Then Call fSetConditionFormatForOddEvenLine(shtProfit)
    Call fSetBorderLineForSheet(shtProfit)
    
    shtProfit.Visible = xlSheetVisible
    shtProfit.Activate
    
    
       ' Call fProtectSheetAndAllowEdit(shtSalesRawDataRpt, shtSalesRawDataRpt.Columns(4), UBound(arrOutput, 1) + 1, UBound(arrOutput, 2), False)
        Call fPostProcess(shtProfit)
        
        shtProfit.Visible = xlSheetVisible
        shtProfit.Activate
    fGotoCell shtProfit.Range("A1")
    
error_handling:
    If fCheckIfUnCapturedExceptionAbnormalError Then GoTo reset_excel_options
    
    If dictErrorRows.Count > 0 Then
        Dim lExcepMaxCol As Long
        lExcepMaxCol = fGetValidMaxCol(shtException)
        
        Call fSetFormatBoldOrangeBorderForHeader(shtException, lExcepMaxCol)
        Call fSetBorderLineForSheet(shtException, lExcepMaxCol)
        Call fBasicCosmeticFormatSheet(shtException, lExcepMaxCol)
        Call fSetFormatForOddEvenLineByFixColor(shtException, lExcepMaxCol)
        
        If Not fFindInWorksheet(shtException.Cells, "找不到可扣的本公司出货记录", False) Is Nothing Then
            'shtException.Columns(4).ColumnWidth = 100
            Call fFreezeSheet(shtException, , 2)
        End If
        
        shtException.Activate
    End If
    
    Call fSetFormatForExceptionCells(shtProfit, dictErrorRows, "REPORT_ERROR_COLOR")
    Call fSetFormatForExceptionCells(shtProfit, dictWarningRows, "REPORT_WARNING_COLOR")
    
    If Not fCheckIfGotBusinessError(False) Then
        fMsgBox "计算完成，请检查工作表：[" & shtProfit.Name & "] 中，请检查！", vbInformation
    End If
     
    If fCheckIfGotBusinessError Then GoTo reset_excel_options
    If fCheckIfUnCapturedExceptionAbnormalError Then GoTo reset_excel_options
    
reset_excel_options:
    Err.Clear
    fClearRefVariables
    fEnableExcelOptionsAll
    'End
End Sub

Private Function fLoadFilesAndRead2Variables()
    'gsCompanyID
    Call fLoadFileByFileTag("UNIFIED_SALES_INFO")
    
    Call fSortDataInSheetSortSheetDataByFileSpec("UNIFIED_SALES_INFO", Array("MatchedProductProducer" _
                                    , "MatchedProductName" _
                                    , "MatchedProductSeries" _
                                    , "SalesDate"))
    
    Call fReadMasterSheetData("UNIFIED_SALES_INFO", , , True)
End Function


Private Function fProcessData()
    Dim lEachRow As Long
    Dim dictMissedFirstLComm As Dictionary
    Dim dictMissedSecondLComm As Dictionary
    Dim dictNoValidSelfSales As Dictionary
    Dim dictNoSalesManConf As Dictionary
    Dim dictNoPriceRecInAdv As Dictionary
    
    Dim sHospital As String
    Dim sSalesCompName As String
    Dim sSalesCompNameID As String
    Dim sSalesCompID As String
    Dim sProducer As String
    Dim sProductName  As String
    Dim sProductSeries As String
    
    Dim sSalesManKey As String
    Dim sFirstLevelCommKey As String
    Dim sFirstLevelCommPasteKey As String
    Dim sSecondLevelCommKey As String
    Dim sSecondLevelCommPasteKey As String
    Dim sProductKey As String
    
    Dim dblQuantity As Double
    Dim dblFirstLevelComm As Double
    Dim dblSecondLevelComm As Double
    'Dim dblCommForRefund As Double
    Dim dblGrossPrice2CZL As Double
    Dim dblPriceForRefund As Double
    Dim dblCostPrice As Double
    Dim sSalesMan_1 As String, sSalesMan_2 As String, sSalesMan_3 As String, sSalesManager As String
    Dim dblComm_1 As Double, dblComm_2 As Double, dblComm_3 As Double, dblSalesMgrComm As Double
    Dim sSalesMan_4 As String, sSalesMan_5 As String, sSalesMan_6 As String
    Dim dblComm_4 As Double, dblComm_5 As Double, dblComm_6 As Double
    Dim dblGrossProfitAmt As Double
    Dim dblPriceRecInAdv As Double
    Dim dblCommForRefund As Double
    'Dim dblProdProducerRefundRate As Double
'    Dim dblNewRSalesTaxRate As Double
'    Dim dblNewRPurchaseTaxRate As Double
    Dim dblPromPrdRebate As Double
    Dim dblSalesTaxRate As Double
    Dim dblPurchaseTaxRate As Double
    Dim sAllCostPrice As String
    Dim sMsg As String
    Dim bIsPromotionProduct As Boolean
    
    Dim dblSellPrice As Double
    
    Call fRedimArrOutputBaseArrMaster
    
    Set dictMissedFirstLComm = New Dictionary
    Set dictMissedSecondLComm = New Dictionary
    Set dictNoValidSelfSales = New Dictionary
    Set dictNoSalesManConf = New Dictionary
    Set dictNoPriceRecInAdv = New Dictionary
    
    For lEachRow = LBound(arrMaster, 1) To UBound(arrMaster, 1)
        If dictMstColIndex.Exists("OrigSalesInfoID") Then
            arrOutput(lEachRow, dictRptColIndex("OrigSalesInfoID")) = arrMaster(lEachRow, dictMstColIndex("OrigSalesInfoID"))
        End If
        
        If dictMstColIndex.Exists("SeqNo") Then
            arrOutput(lEachRow, dictRptColIndex("SeqNo")) = arrMaster(lEachRow, dictMstColIndex("SeqNo"))
        End If

        sHospital = Trim(arrMaster(lEachRow, dictMstColIndex("MatchedHospital")))
        sSalesCompName = Trim(arrMaster(lEachRow, dictMstColIndex("SalesCompanyName")))
        sProducer = Trim(arrMaster(lEachRow, dictMstColIndex("MatchedProductProducer")))
        sProductName = Trim(arrMaster(lEachRow, dictMstColIndex("MatchedProductName")))
        sProductSeries = Trim(arrMaster(lEachRow, dictMstColIndex("MatchedProductSeries")))
        dblQuantity = arrMaster(lEachRow, dictMstColIndex("ConvertQuantity"))
        dblSellPrice = arrMaster(lEachRow, dictMstColIndex("ConvertSellPrice"))
    
        sProductKey = sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries
    
        arrOutput(lEachRow, dictRptColIndex("ProductKey")) = sProductKey
        arrOutput(lEachRow, dictRptColIndex("Hospital")) = sHospital
        arrOutput(lEachRow, dictRptColIndex("SalesCompanyName")) = sSalesCompName
        arrOutput(lEachRow, dictRptColIndex("ProductProducer")) = sProducer
        arrOutput(lEachRow, dictRptColIndex("ProductName")) = sProductName
        arrOutput(lEachRow, dictRptColIndex("ProductSeries")) = sProductSeries
        arrOutput(lEachRow, dictRptColIndex("SalesDate")) = arrMaster(lEachRow, dictMstColIndex("SalesDate"))
        arrOutput(lEachRow, dictRptColIndex("ProductUnit")) = arrMaster(lEachRow, dictMstColIndex("MatchedProductUnit"))
        arrOutput(lEachRow, dictRptColIndex("Quantity")) = dblQuantity
        arrOutput(lEachRow, dictRptColIndex("SellPrice")) = dblSellPrice
        arrOutput(lEachRow, dictRptColIndex("SellAmount")) = arrMaster(lEachRow, dictMstColIndex("RecalSellAmount"))
        arrOutput(lEachRow, dictRptColIndex("LotNum")) = "'" & arrMaster(lEachRow, dictMstColIndex("LotNum"))
        
        arrOutput(lEachRow, dictRptColIndex("SalesRecordKey")) = sSalesCompName & sProducer & sProductName & sProductSeries _
                        & sHospital & Format(arrMaster(lEachRow, dictMstColIndex("SalesDate")), "yyyymmdd") & dblQuantity & arrMaster(lEachRow, dictMstColIndex("LotNum"))
        
        bIsPromotionProduct = fIsPromotionProduct(sHospital, sProductKey, dblSellPrice, sSalesCompName, dblPromPrdRebate, dblSalesTaxRate, dblPurchaseTaxRate, dblSecondLevelComm, dblCommForRefund) ', dblProdProducerRefundRate)
        
'        dblPriceRecInAdvance = fGetPriceRecInAdvance(sProducer, sProductName, sProductSeries)
        dblPriceRecInAdv = fGetSellPriceInAdv(sSalesCompName, sProductKey)
        
        If dblPriceRecInAdv <= 0 Then
            If Not dictNoPriceRecInAdv.Exists(sSalesCompName & DELIMITER & sProductKey) Then
                dictNoPriceRecInAdv.Add sSalesCompName & DELIMITER & sProductKey, "'" & lEachRow + 1
            Else
                dictNoPriceRecInAdv(sProductKey) = dictNoPriceRecInAdv(sSalesCompName & DELIMITER & sProductKey) & "," & (lEachRow + 1)
            End If
            Call fAddErrorColumnTodictErrorRows(lEachRow + 1, dictRptColIndex("PriceRecInAdvance"))
        End If
        
        If bIsPromotionProduct Then
            dblGrossPrice2CZL = dblSellPrice * dblPromPrdRebate
           'dblPriceForRefund = (dblSellPrice - dblSellPrice * dblSecondLevelComm) / (1 - dblProdProducerRefundRate) * dblPromPrdRebate '(中标价 - 中标价 * 配送费点数) / 0.92 * 0.53
            'dblPriceForRefund = (dblSellPrice - dblSellPrice * dblSecondLevelComm)
            dblPriceForRefund = (dblSellPrice - dblSellPrice * dblCommForRefund)
        Else
            '==== first level czl commission ==========================================
            sFirstLevelCommKey = sSalesCompName & DELIMITER & sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries
            
            If Not fGetFirstLevelComm(sFirstLevelCommKey, dblFirstLevelComm) Then
                dblFirstLevelComm = fGetConfigFirstLevelDefaultComm()
                
                sFirstLevelCommPasteKey = fComposeFirstLevelColumnsStryByConfig(sSalesCompName, sProducer, sProductName _
                                                        , sProductSeries, dblFirstLevelComm)
                If Not dictMissedFirstLComm.Exists(sFirstLevelCommPasteKey) Then
                    dictMissedFirstLComm.Add sFirstLevelCommPasteKey, "'" & (lEachRow + 1)
                Else
                    dictMissedFirstLComm(sFirstLevelCommPasteKey) = dictMissedFirstLComm(sFirstLevelCommPasteKey) & "," & (lEachRow + 1)
                End If
            End If
            '-----------------------------------------------------------------------------------------------
            
            '==== second level commission ==========================================
            sSecondLevelCommKey = sSalesCompName & DELIMITER & sHospital & DELIMITER _
                                & sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries
        
            If Not fGetSecondLevelComm(sSecondLevelCommKey, dblSecondLevelComm, dblCommForRefund) Then
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
            
            dblGrossPrice2CZL = dblSellPrice * (1 - dblFirstLevelComm) * (1 - dblSecondLevelComm)
            'dblPriceForRefund = dblGrossPrice2CZL
            dblPriceForRefund = (dblSellPrice - dblSellPrice * dblCommForRefund)
        End If
        
        arrOutput(lEachRow, dictRptColIndex("GrossPrice2CZL")) = dblGrossPrice2CZL
        arrOutput(lEachRow, dictRptColIndex("GrossAmount2CZL")) = dblGrossPrice2CZL * dblQuantity
        
        'If bIsPromotionProduct Then
            arrOutput(lEachRow, dictRptColIndex("PriceForRefund")) = dblPriceForRefund
        'End If
        
        '==== cost price ==========================================
        sMsg = ""
        If bIsPromotionProduct Then
            dblCostPrice = 0
            arrOutput(lEachRow, dictRptColIndex("CostPrice")) = dblCostPrice
        Else
            If Not fCalculateCostPriceFromSelfSalesOrder(sProductKey, dblQuantity, dblCostPrice) Then
                sAllCostPrice = GetAvailableSelfSalesPrices(sProductKey)
                
                If Len(Trim(sAllCostPrice)) <= 0 Then
                    dblCostPrice = 0
                    sMsg = "该药品在本公司的销售记录中没有可扣数量，因此找不到其准确的成本价。并且该药品也没有历史销售记录，故找不到任何可用的成本价。"
                    Call fAddErrorColumnTodictErrorRows(lEachRow + 1, dictRptColIndex("CostPrice"))
                Else
                    dblCostPrice = Split(sAllCostPrice, "~")(0)
                    sMsg = "该药品在本公司的销售记录中没有可扣数量，因此找不到其准确的成本价，现找到了所有历史成本价，并按照第一个价格计算了利润，请核对。(第一个为最近一次的价格。)"
                    Call fAddWarningColumnTodictWarningRows(lEachRow + 1, dictRptColIndex("CostPrice"))
                End If
        
                If Not dictNoValidSelfSales.Exists(sProductKey) Then
                    dictNoValidSelfSales.Add sProductKey, sAllCostPrice & DELIMITER & sMsg & DELIMITER & (lEachRow + 1)
                Else
                    dictNoValidSelfSales(sProductKey) = dictNoValidSelfSales(sProductKey) & "," & (lEachRow + 1)
                End If
                arrOutput(lEachRow, dictRptColIndex("CostPrice")) = sAllCostPrice
            Else
                arrOutput(lEachRow, dictRptColIndex("CostPrice")) = dblCostPrice
            End If
            'dblCostPrice = fGetLatestPriceFromProductMaster(sProductKey)
        End If

        arrOutput(lEachRow, dictRptColIndex("CostAmount")) = dblCostPrice * dblQuantity

        If bIsPromotionProduct Then
            arrOutput(lEachRow, dictRptColIndex("SalesTaxPerUnit")) = (dblGrossPrice2CZL - dblCostPrice) * dblSalesTaxRate
            arrOutput(lEachRow, dictRptColIndex("PurchaeTaxPerUnit")) = (dblGrossPrice2CZL - dblCostPrice - arrOutput(lEachRow, dictRptColIndex("SalesTaxPerUnit"))) _
                                                    * dblPurchaseTaxRate
        ElseIf Not fIsNewRuleProduct(sProductKey) Then
            arrOutput(lEachRow, dictRptColIndex("SalesTaxPerUnit")) = dblGrossPrice2CZL * fGetTaxRate(sProductKey)
            arrOutput(lEachRow, dictRptColIndex("PurchaeTaxPerUnit")) = 0
        Else
            Call fGetNewRuleProductTaxRate(sProductKey, dblSalesTaxRate, dblPurchaseTaxRate)
            
            arrOutput(lEachRow, dictRptColIndex("SalesTaxPerUnit")) = (dblGrossPrice2CZL - dblCostPrice) * dblSalesTaxRate
            arrOutput(lEachRow, dictRptColIndex("PurchaeTaxPerUnit")) = (dblGrossPrice2CZL - dblCostPrice - arrOutput(lEachRow, dictRptColIndex("SalesTaxPerUnit"))) _
                                                    * dblPurchaseTaxRate
        End If
        
        If bIsPromotionProduct Then
            arrOutput(lEachRow, dictRptColIndex("GrossProfitPerUnit")) = dblGrossPrice2CZL - dblCostPrice _
                                                            - arrOutput(lEachRow, dictRptColIndex("SalesTaxPerUnit")) _
                                                            - arrOutput(lEachRow, dictRptColIndex("PurchaeTaxPerUnit")) _
                                                            - (dblSellPrice * dblSecondLevelComm)
        Else
            arrOutput(lEachRow, dictRptColIndex("GrossProfitPerUnit")) = dblGrossPrice2CZL - dblCostPrice _
                                                            - arrOutput(lEachRow, dictRptColIndex("SalesTaxPerUnit")) _
                                                            - arrOutput(lEachRow, dictRptColIndex("PurchaeTaxPerUnit"))
        End If
        
        dblGrossProfitAmt = arrOutput(lEachRow, dictRptColIndex("GrossProfitPerUnit")) * dblQuantity
        arrOutput(lEachRow, dictRptColIndex("GrossProfitAmt")) = dblGrossProfitAmt
        '-----------------------------------------------------------------------------------------------
        
        '==== salesman commission ==========================================
        sSalesManKey = sSalesCompName & DELIMITER & sHospital & DELIMITER _
                    & sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries & DELIMITER & dblSellPrice
        If Not fCalculateSalesManCommissionFromshtSalesManCommConfig(sSalesManKey _
                                , sSalesMan_1, sSalesMan_2, sSalesMan_3, sSalesMan_4, sSalesMan_5, sSalesMan_6, dblComm_1, dblComm_2, dblComm_3, dblComm_4, dblComm_5, dblComm_6 _
                                , sSalesManager, dblSalesMgrComm) Then
            If Not dictNoSalesManConf.Exists(sSalesManKey) Then
                dictNoSalesManConf.Add sSalesManKey, "'" & lEachRow + 1
            Else
                dictNoSalesManConf(sSalesManKey) = dictNoSalesManConf(sSalesManKey) & "," & (lEachRow + 1)
            End If
            
'            mlExcepCnt = mlExcepCnt + 1:    arrExceptionRows(mlExcepCnt) = lEachRow + 1
'            mlExcepCnt = mlExcepCnt + 1:    arrExceptionRows(mlExcepCnt) = dictRptColIndex("SalesMan_1")
            Call fAddErrorColumnTodictErrorRows(lEachRow + 1, dictRptColIndex("SalesMan_1"))
        End If

        arrOutput(lEachRow, dictRptColIndex("SalesMan_1")) = sSalesMan_1
        arrOutput(lEachRow, dictRptColIndex("SalesMan_2")) = sSalesMan_2
        arrOutput(lEachRow, dictRptColIndex("SalesMan_3")) = sSalesMan_3
        arrOutput(lEachRow, dictRptColIndex("SalesMan_4")) = sSalesMan_4
        arrOutput(lEachRow, dictRptColIndex("SalesMan_5")) = sSalesMan_5
        arrOutput(lEachRow, dictRptColIndex("SalesMan_6")) = sSalesMan_6
        arrOutput(lEachRow, dictRptColIndex("SalesManList")) = sSalesMan_1 _
                                                             & IIf(Len(sSalesMan_2) > 0, ", " & sSalesMan_2, "") _
                                                             & IIf(Len(sSalesMan_3) > 0, ", " & sSalesMan_3, "") _
                                                             & IIf(Len(sSalesMan_4) > 0, ", " & sSalesMan_4, "") _
                                                             & IIf(Len(sSalesMan_5) > 0, ", " & sSalesMan_5, "") _
                                                             & IIf(Len(sSalesMan_6) > 0, ", " & sSalesMan_6, "")
        arrOutput(lEachRow, dictRptColIndex("SalesCommission_1")) = dblComm_1
        arrOutput(lEachRow, dictRptColIndex("SalesCommission_2")) = dblComm_2
        arrOutput(lEachRow, dictRptColIndex("SalesCommission_3")) = dblComm_3
        arrOutput(lEachRow, dictRptColIndex("SalesCommission_4")) = dblComm_4
        arrOutput(lEachRow, dictRptColIndex("SalesCommission_5")) = dblComm_5
        arrOutput(lEachRow, dictRptColIndex("SalesCommission_6")) = dblComm_6
        '-----------------------------------------------------------------------------------------------
        
        arrOutput(lEachRow, dictRptColIndex("NetProfitPerUnit")) = arrOutput(lEachRow, dictRptColIndex("GrossProfitPerUnit")) _
                                            - arrOutput(lEachRow, dictRptColIndex("SalesCommission_1")) _
                                            - arrOutput(lEachRow, dictRptColIndex("SalesCommission_2")) _
                                            - arrOutput(lEachRow, dictRptColIndex("SalesCommission_3")) _
                                            - arrOutput(lEachRow, dictRptColIndex("SalesCommission_4")) _
                                            - arrOutput(lEachRow, dictRptColIndex("SalesCommission_5")) _
                                            - arrOutput(lEachRow, dictRptColIndex("SalesCommission_6"))
        arrOutput(lEachRow, dictRptColIndex("NetProfitAmount")) = arrOutput(lEachRow, dictRptColIndex("NetProfitPerUnit")) _
                                                                * dblQuantity
        
        arrOutput(lEachRow, dictRptColIndex("SalesManagerCommissoin")) = arrOutput(lEachRow, dictRptColIndex("NetProfitAmount")) _
                                                                * dblSalesMgrComm

        arrOutput(lEachRow, dictRptColIndex("SalesManList")) = fComposeSalesManList(sSalesManager, sSalesMan_1, sSalesMan_2, sSalesMan_3)
        arrOutput(lEachRow, dictRptColIndex("SalesCommissionAmt_1")) = dblComm_1 * dblQuantity
        arrOutput(lEachRow, dictRptColIndex("SalesCommissionAmt_2")) = dblComm_2 * dblQuantity
        arrOutput(lEachRow, dictRptColIndex("SalesCommissionAmt_3")) = dblComm_3 * dblQuantity
        arrOutput(lEachRow, dictRptColIndex("SalesCommissionAmt_4")) = dblComm_4 * dblQuantity
        arrOutput(lEachRow, dictRptColIndex("SalesCommissionAmt_5")) = dblComm_5 * dblQuantity
        arrOutput(lEachRow, dictRptColIndex("SalesCommissionAmt_6")) = dblComm_6 * dblQuantity
        arrOutput(lEachRow, dictRptColIndex("SalesCommission_Total")) = arrOutput(lEachRow, dictRptColIndex("SalesCommissionAmt_1")) _
                                                                        + arrOutput(lEachRow, dictRptColIndex("SalesCommissionAmt_2")) _
                                                                        + arrOutput(lEachRow, dictRptColIndex("SalesCommissionAmt_3")) _
                                                                        + arrOutput(lEachRow, dictRptColIndex("SalesCommissionAmt_4")) _
                                                                        + arrOutput(lEachRow, dictRptColIndex("SalesCommissionAmt_5")) _
                                                                        + arrOutput(lEachRow, dictRptColIndex("SalesCommissionAmt_6"))
        
        If arrOutput(lEachRow, dictRptColIndex("SellAmount")) > 0 Then _
        arrOutput(lEachRow, dictRptColIndex("NetProfitRate")) = arrOutput(lEachRow, dictRptColIndex("NetProfitAmount")) _
                                                                / arrOutput(lEachRow, dictRptColIndex("SellAmount"))
                                                                
        arrOutput(lEachRow, dictRptColIndex("PriceRecInAdvance")) = dblPriceRecInAdv
        arrOutput(lEachRow, dictRptColIndex("RefundPerUnit")) = dblPriceForRefund - dblPriceRecInAdv
        arrOutput(lEachRow, dictRptColIndex("RefundAmount")) = arrOutput(lEachRow, dictRptColIndex("RefundPerUnit")) * dblQuantity
next_sales:
    Next
    
    Dim lStartRow As Long
    Dim lRecCount As Long
    Dim arrTmp()
    
    lRecCount = fGetDictionayDelimiteredItemsCount(dictMissedFirstLComm)
    If lRecCount > 0 Then
        arrTmp = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictMissedFirstLComm)
        Call fAppendArray2Sheet(shtFirstLevelCommission, arrTmp)
        
        fMsgBox lRecCount & "条销售流向记录的采芝林的配送费没有设置，系统已经自动把它们添加到了【" & shtFirstLevelCommission.Name & "】" _
            & vbCr & "您可以查看该表中最后面的数据"
    End If
    
    lRecCount = fGetDictionayDelimiteredItemsCount(dictMissedSecondLComm)
    If lRecCount > 0 Then
        arrTmp = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictMissedSecondLComm)
        Call fAppendArray2Sheet(shtSecondLevelCommission, arrTmp)
        
        fMsgBox lRecCount & "条销售流向记录的商业公司的配送费没有设置，系统已经自动把它们添加到了【" & shtSecondLevelCommission.Name & "】" _
            & vbCr & "您可以查看该表中最后面的数据"
    End If
     
    
    Call fSetBackToshtSelfSalesCalWithDeductedData
    Call fAddNoValidSelfSalesToSheetException(dictNoValidSelfSales)
    Call fAddNoSalesManConfToSheetException(dictNoSalesManConf)
    Call fAddNoPriceRecInAdvToSheetException(dictNoPriceRecInAdv)
    Set dictNoPriceRecInAdv = Nothing
    Set dictNoSalesManConf = Nothing
    Set dictNoValidSelfSales = Nothing
End Function

Function fAddNoValidSelfSalesToSheetException(dictNoValidSelfSales As Dictionary)
    Dim arrNewProductSeries()
    Dim lUniqRecCnt As Long
    Dim lRecCount As Long
    Dim i As Integer
    Dim j As Integer
        Dim lStartRow As Long
    'Dim arrTmp
    
    lUniqRecCnt = dictNoValidSelfSales.Count
    If lUniqRecCnt > 0 Then
        lStartRow = fGetshtExceptionNewRow
        arrNewProductSeries = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictNoValidSelfSales, , False)
        
        'shtException.Columns(4).ColumnWidth = 100
        shtException.Cells(lStartRow - 1, 1).Value = "找不到可扣的本公司出货记录(若为退货，则是找不到退货可抵扣)"
        shtException.Cells(lStartRow - 1, 1).WrapText = False
        Call fPrepareHeaderToSheet(shtException, Array("药品厂家", "药品名称", "规格", "历史价格(供参考)", "借误信息", "行号"), lStartRow)
        shtException.Rows(lStartRow - 1 & ":" & lStartRow).Font.Color = RGB(255, 0, 0)
        shtException.Rows(lStartRow - 1 & ":" & lStartRow).Font.Bold = True
        Call fAppendArray2Sheet(shtException, arrNewProductSeries)
        'sErr = fUbound(arrNewProductSeries)
        
        lRecCount = fGetDictionayDelimiteredItemsCount(dictNoValidSelfSales)
        
        shtException.Cells(lStartRow + 1, 4).Resize(dictNoValidSelfSales.Count, 3).Value = fConvertDictionaryDelimiteredItemsTo2DimenArrayForPaste(dictNoValidSelfSales, , False)
        Erase arrNewProductSeries
        If lStartRow = 2 Then Call fFreezeSheet(shtException, , 2)
        
        shtException.Visible = xlSheetVisible
        shtException.Activate
        
        If fNzero(gsBusinessErrorMsg) Then gsBusinessErrorMsg = gsBusinessErrorMsg & vbCr & vbCr & vbCr & "===============================" & vbCr & vbCr

        gsBusinessErrorMsg = gsBusinessErrorMsg & lUniqRecCnt & "个药品" & lRecCount & "条销售流向在本公司出货记录中无出库可扣除(若为退货，则是找不到医院销售抵扣)，您可能要：" & vbCr _
            & "(1). 在【本公司出货】中添加一条替换记录" & vbCr _
            & "计算这些销售流向进行到一半，没有可以扣的出货记录，所以把它们的成本价格标注0，"
            
            '& "(2). 在【药品主表】中修改其最新价格" & vbCr & vbCr
    End If
End Function

Function fAddNoPriceRecInAdvToSheetException(dictNoPriceRecInAdv As Dictionary)
    Dim arrData()
    Dim lUniqRecCnt As Long
    'Dim lRecCount As Long
    Dim i As Integer
    Dim j As Integer
    Dim lStartRow As Long
            
    lUniqRecCnt = dictNoPriceRecInAdv.Count
    If lUniqRecCnt > 0 Then
        lStartRow = fGetshtExceptionNewRow
        
        shtException.Cells(lStartRow - 1, 1).Value = "没有设置供货价的药品！"
        shtException.Cells(lStartRow - 1, 1).WrapText = False
        Call fPrepareHeaderToSheet(shtException, Array("药品厂家", "药品名称", "规格", "", "行号"), lStartRow)
        shtException.Rows(lStartRow - 1 & ":" & lStartRow).Font.Color = RGB(255, 0, 0)
        shtException.Rows(lStartRow - 1 & ":" & lStartRow).Font.Bold = True
        
        arrData = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictNoPriceRecInAdv, , False)
        Call fAppendArray2Sheet(shtException, arrData)
        
        shtException.Cells(lStartRow + 1, 5).Resize(lUniqRecCnt, 1).Value = fConvertDictionaryItemsTo2DimenArrayForPaste(dictNoPriceRecInAdv)

        If lStartRow = 2 Then Call fFreezeSheet(shtException, , 2)
        
        If fNzero(gsBusinessErrorMsg) Then gsBusinessErrorMsg = gsBusinessErrorMsg & vbCr & vbCr & vbCr & "===============================" & vbCr & vbCr
        
        gsBusinessErrorMsg = gsBusinessErrorMsg & lUniqRecCnt & "个药品没有设置供货价，您需要：" & vbCr _
            & "(1). 在【药品主表】中设置其供货价."
    End If
End Function

Function fAddNoSalesManConfToSheetException(dictNoSalesManConf As Dictionary)
    Dim arrNoSalesMan()
    Dim lUniqRecCnt As Long
    'Dim lRecCount As Long
    Dim i As Integer
    Dim j As Integer
    Dim lStartRow As Long
            
    lUniqRecCnt = dictNoSalesManConf.Count
    If lUniqRecCnt > 0 Then
        lStartRow = fGetshtExceptionNewRow
        
        shtException.Cells(lStartRow - 1, 1).Value = "找不到业务员的记录  --> 有可能只是中价标没有！"
        shtException.Cells(lStartRow - 1, 1).WrapText = False
        Call fPrepareHeaderToSheet(shtException, Array("商业公司", "医院", "药品厂家", "药品名称", "规格", "中标价", "行号"), lStartRow)
        shtException.Rows(lStartRow - 1 & ":" & lStartRow).Font.Color = RGB(255, 0, 0)
        shtException.Rows(lStartRow - 1 & ":" & lStartRow).Font.Bold = True
        
        arrNoSalesMan = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictNoSalesManConf, , False)
        Call fAppendArray2Sheet(shtException, arrNoSalesMan)
        
        shtException.Cells(lStartRow + 1, 7).Resize(lUniqRecCnt, 1).Value = fConvertDictionaryItemsTo2DimenArrayForPaste(dictNoSalesManConf)

        If lStartRow = 2 Then Call fFreezeSheet(shtException, , 2)
        
        If fNzero(gsBusinessErrorMsg) Then gsBusinessErrorMsg = gsBusinessErrorMsg & vbCr & vbCr & vbCr & "===============================" & vbCr & vbCr
        
        gsBusinessErrorMsg = gsBusinessErrorMsg & lUniqRecCnt & "条销售流向找不到业务员的记录，您需要：" & vbCr _
            & "(1). 在【业务员佣金设置】中添加业务员佣金设置"
    End If
End Function
'Function fCalculateCostPrice(lEachRow As Long, ByRef dictNoValidSelfSales As Dictionary) As Double
'    Dim dblCostPrice As Double
'
'    Dim sProducer As String
'    Dim sProductName  As String
'    Dim sProductSeries As String
'    Dim dblSalesQuantity As Double
'
'    Dim sTmpKey As String
'
'    sProducer = Trim(arrOutput(lEachRow, dictRptColIndex("ProductProducer")))
'    sProductName = Trim(arrOutput(lEachRow, dictRptColIndex("ProductName")))
'    sProductSeries = Trim(arrOutput(lEachRow, dictRptColIndex("ProductSeries")))
'
'    dblSalesQuantity = arrOutput(lEachRow, dictRptColIndex("Quantity"))
'
'    If Not fCalculateCostPriceFromSelfSalesOrder(sProducer, sProductName, sProductSeries, dblSalesQuantity, dblCostPrice) Then
'        sTmpKey = sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries
'
'        If Not dictNoValidSelfSales.Exists(sTmpKey) Then
'            dictNoValidSelfSales.Add sTmpKey, lEachRow + 1
'        Else
'            dictNoValidSelfSales(sTmpKey) = dictNoValidSelfSales(sTmpKey) & "," & (lEachRow + 1)
'        End If
'        dblCostPrice = fGetLatestPriceFromProductMaster(sProducer, sProductName, sProductSeries)
'    End If
'
'    fCalculateCostPrice = dblCostPrice
'End Function

'Function fCalculateGrossPrice2CZL(lEachRow As Long, ByRef dictMissedFirstLComm As Dictionary, ByRef dictMissedSecondLComm As Dictionary) As Double
'    Dim dblGrossPrice2CZL As Double
'
'    Dim dblFirstLevelComm As Double
'    Dim dblSecondLevelComm As Double
'
'    Dim sHospital As String
'    Dim sSalesCompName As String
'    Dim sSalesCompID As String
'    Dim sProducer As String
'    Dim sProductName  As String
'    Dim sProductSeries As String
'
'    Dim sTmpKey As String
'
'    sHospital = Trim(arrOutput(lEachRow, dictRptColIndex("Hospital")))
'    sSalesCompName = Trim(arrOutput(lEachRow, dictRptColIndex("SalesCompanyName")))
'    sProducer = Trim(arrOutput(lEachRow, dictRptColIndex("ProductProducer")))
'    sProductName = Trim(arrOutput(lEachRow, dictRptColIndex("ProductName")))
'    sProductSeries = Trim(arrOutput(lEachRow, dictRptColIndex("ProductSeries")))
'
'    'sSalesCompID = fGetSalesCompanyID(sSalesCompName)
'    If Not fGetFirstLevelComm(sSalesCompName, sProducer, sProductName, sProductSeries, dblFirstLevelComm) Then
'        dblFirstLevelComm = fGetConfigFirstLevelDefaultComm()
'
'        'sTmpKey = sSalesCompName & DELIMITER & sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries
'        sTmpKey = fComposeFirstLevelColumnsStryByConfig(sSalesCompName, sProducer, sProductName, sProductSeries, dblFirstLevelComm)
'        If Not dictMissedFirstLComm.Exists(sTmpKey) Then
'            dictMissedFirstLComm.Add sTmpKey, lEachRow + 1
'        End If
'    End If
'
'    If Not fGetSecondLevelComm(sSalesCompName, sHospital, sProducer, sProductName, sProductSeries, dblSecondLevelComm) Then
'        dblSecondLevelComm = fGetConfigSecondLevelDefaultComm(sSalesCompName)
'
'        'sTmpKey = sSalesCompName & DELIMITER & sHospital & DELIMITER & sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries
'        sTmpKey = fComposeSecondLevelColumnsStryByConfig(sSalesCompName, sHospital, sProducer, sProductName, sProductSeries, dblSecondLevelComm)
'        If Not dictMissedSecondLComm.Exists(sTmpKey) Then
'            dictMissedSecondLComm.Add sTmpKey, lEachRow + 1
'        End If
'    End If
'
'    Dim dblSellPrice As Double
'    dblSellPrice = arrOutput(lEachRow, dictRptColIndex("SellPrice"))
'    dblGrossPrice2CZL = dblSellPrice * (1 - dblFirstLevelComm) * (1 - dblSecondLevelComm)
'
'    fCalculateGrossPrice2CZL = dblGrossPrice2CZL
'End Function

Function fComposeFirstLevelColumnsStryByConfig(sSalesCompName As String, sProducer As String _
                    , sProductName As String, sProductSeries As String, dblComm As Double) As String
    If dictFirstCommColIndex Is Nothing Then Set dictFirstCommColIndex = fReadInputFileSpecConfigItem("FIRST_LEVEL_COMMISSION", "LETTER_INDEX", shtFirstLevelCommission)
    
    Dim i As Integer
    Dim arr() As String
    
    ReDim arr(1 To dictFirstCommColIndex.Count)
    arr(dictFirstCommColIndex("SalesCompany")) = sSalesCompName
    arr(dictFirstCommColIndex("ProductProducer")) = sProducer
    arr(dictFirstCommColIndex("ProductName")) = sProductName
    arr(dictFirstCommColIndex("ProductSeries")) = sProductSeries
    arr(dictFirstCommColIndex("Commission")) = dblComm

    fComposeFirstLevelColumnsStryByConfig = Join(arr, DELIMITER)
    Erase arr
End Function
 
Function fComposeSecondLevelColumnsStryByConfig(sSalesCompName As String, sHospital As String, sProducer As String _
                    , sProductName As String, sProductSeries As String, dblComm As Double) As String
    
    Dim i As Integer
    Dim arr() As String
    
    ReDim arr(SecondLevelComm.[_first] To SecondLevelComm.[_last])
    arr(SecondLevelComm.SalesCompany) = sSalesCompName
    arr(SecondLevelComm.Hospital) = sHospital
    arr(SecondLevelComm.ProductProducer) = sProducer
    arr(SecondLevelComm.ProductName) = sProductName
    arr(SecondLevelComm.ProductSeries) = sProductSeries
    arr(SecondLevelComm.Commission) = dblComm

    fComposeSecondLevelColumnsStryByConfig = Join(arr, DELIMITER)
    Erase arr
End Function

'Function fCalculateSalesManCommission(lEachRow As Long, ByRef sSalesMan_1 As String, ByRef sSalesMan_2 As String, ByRef sSalesMan_3 As String _
'                            , ByRef dblComm_1 As Double, ByRef dblComm_2 As Double, ByRef dblComm_3 As Double _
'                            , ByRef dictNoSalesManConf As Dictionary) As Boolean
'    Dim sHospital As String
'    Dim sSalesCompany As String
'    Dim sSalesCompID As String
'    Dim sProducer As String
'    Dim sProductName  As String
'    Dim sProductSeries As String
'
'    Dim sTmpKey As String
'
'    sHospital = Trim(arrOutput(lEachRow, dictRptColIndex("Hospital")))
'    sSalesCompany = Trim(arrOutput(lEachRow, dictRptColIndex("SalesCompanyName")))
'    sProducer = Trim(arrOutput(lEachRow, dictRptColIndex("ProductProducer")))
'    sProductName = Trim(arrOutput(lEachRow, dictRptColIndex("ProductName")))
'    sProductSeries = Trim(arrOutput(lEachRow, dictRptColIndex("ProductSeries")))
'
'    If Not fCalculateSalesManCommissionFromshtSalesManCommConfig(sSalesCompany, sHospital, sProducer, sProductName, sProductSeries _
'                                , sSalesMan_1, sSalesMan_2, sSalesMan_3, dblComm_1, dblComm_2, dblComm_3) Then
'        sTmpKey = sSalesCompany & DELIMITER & sHospital & DELIMITER _
'                    & sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries
'        If Not dictNoSalesManConf.Exists(sTmpKey) Then
'            dictNoSalesManConf.Add sTmpKey, lEachRow + 1
'        End If
'    End If
'End Function

'Function fAppendDictionaryKeys2Worksheet(dict As Dictionary, sht As Worksheet)
'    Dim arr()
'    Dim lStartRow As Long
'
'    If dict.Count > 0 Then
'        arr = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dict, , False)
'        lStartRow = fGetValidMaxRow(sht) + 1
'
'        Call fAppendArray2Sheet(sht, arr)
'    End If
'End Function

Function fGetshtExceptionNewRow()
    Dim lNewRow As Long
    lNewRow = fGetValidMaxRow(shtException)
    If lNewRow <= 1 Then
        lNewRow = lNewRow + 2
    Else
        lNewRow = lNewRow + 5
    End If
    
    fGetshtExceptionNewRow = lNewRow
End Function

Sub subMain_CalculateProfit_MonthEnd()
    If Not fIsDev() Then On Error GoTo error_handling

    fInitialization
    
    If Not fPromptToConfirmToContinue("您确定要把【本公司销售出货(试算)】写入正式的库存表中吗？这个操作无法撤消，您确定吗？") Then fErr
    
    Dim arrData()
    Dim lPasteStartRow As Long
    Dim lMaxRow As Long
    
    lMaxRow = fGetValidMaxRow(shtSelfSalesPreDeduct)
    Call fCopyReadWholeSheetData2Array(shtSelfSalesPreDeduct, arrData)
    Call fDeleteRowsFromSheetLeaveHeader(shtSelfSalesOrder)
    Call fAppendArray2Sheet(shtSelfSalesOrder, arrData)
    Erase arrData
    
    Call fBasicCosmeticFormatSheet(shtSelfSalesOrder)
    
    Call fSetConditionFormatForOddEvenLine(shtSelfSalesOrder)
    
    Call fSetBorderLineForSheet(shtSelfSalesOrder)
    
error_handling:
    If fCheckIfGotBusinessError Then GoTo reset_excel_options
    If fCheckIfUnCapturedExceptionAbnormalError Then GoTo reset_excel_options
    
    Call fShowActivateSheet(shtSelfSalesOrder)
    fMsgBox "【本公司销售出货(试算)】已经写入正式的库存表中了： " & vbCr & shtSelfSalesOrder.Name, vbInformation
reset_excel_options:
    Err.Clear
    fClearRefVariables
    fEnableExcelOptionsAll
    
'    Dim response As VbMsgBoxResult
'    response = MsgBox(Prompt:="该操作会扣减本公司出库，无法撤消，你确定要进行计算利润和佣金吗？" _
'                        & vbCr & "继续，请点【Yes】" & vbCr & "否则，请点【No】" _
'                        , Buttons:=vbCritical + vbYesNo + vbDefaultButton2)
'    If response <> vbYes Then Exit Sub
'
'    Set shtSelfSalesCal = shtSelfSalesOrder
'    Call subMain_CalculateProfit
End Sub

Sub subMain_CalculateProfit_PreCal()
    'If Not fIsDev() Then On Error GoTo error_handling
    
    'Set shtSelfSalesCal = shtSelfSalesPreDeduct
    
    Call fRemoveFilterForSheet(shtSelfSalesOrder)
    Call fRemoveFilterForSheet(shtSelfSalesPreDeduct)
    
    shtSelfSalesOrder.Cells.Copy shtSelfSalesPreDeduct.Cells
    Call subMain_CalculateProfit
'error_handling:
'    If fCheckIfGotBusinessError Then
'        GoTo reset_excel_options
'    End If
'
'    If fCheckIfUnCapturedExceptionAbnormalError Then
'        GoTo reset_excel_options
'    End If
'
'    fMsgBox "成功整合在工作表：[" & shtSalesRawDataRpt.Name & "] 中，请检查！", vbInformation
'
'    Application.Goto shtSalesRawDataRpt.Range("A" & fGetValidMaxRow(shtSalesRawDataRpt)), True
'reset_excel_options:
'    Err.Clear
'    fEnableExcelOptionsAll
'    End
End Sub

Private Function fComposeSalesManList(sSalesManager As String, sSalesMan_1 As String, sSalesMan_2 As String, sSalesMan_3 As String) As String
    Dim sOut As String
    Dim sSalesManList As String
    
    If Len(Trim(sSalesMan_1)) > 0 Then
        sSalesManList = sSalesMan_1
    End If
    If Len(Trim(sSalesMan_2)) > 0 Then
        sSalesManList = IIf(Len(Trim(sSalesManList)) > 0, sSalesManList & ",", "") & sSalesMan_2
    End If
    If Len(Trim(sSalesMan_3)) > 0 Then
        sSalesManList = IIf(Len(Trim(sSalesManList)) > 0, sSalesManList & ",", "") & sSalesMan_3
    End If
    
    If Len(Trim(sSalesManager)) > 0 Then
        If Len(Trim(sSalesManList)) > 0 Then
            sOut = sSalesManager '& "(" & sSalesManList & ")"
        Else
            sOut = sSalesManager
        End If
    Else
        If Len(Trim(sSalesManList)) > 0 Then
            'sOut = "(" & sSalesManList & ")"    'sSalesManList
            sOut = sSalesManList      'sSalesManList
        Else
            sOut = ""
        End If
    End If
    
    fComposeSalesManList = sOut
End Function

Function fAddErrorColumnTodictErrorRows(ByVal lRowNo As Long, lCol As Long)
    If dictErrorRows Is Nothing Then Set dictErrorRows = New Dictionary
    
    If dictErrorRows.Exists(lRowNo) Then
        dictErrorRows(lRowNo) = dictErrorRows(lRowNo) & DELIMITER & CStr(lCol)
    Else
        dictErrorRows.Add lRowNo, CStr(lCol)
    End If
End Function
Function fAddWarningColumnTodictWarningRows(ByVal lRowNo As Long, lCol As Long)
    If dictWarningRows Is Nothing Then Set dictWarningRows = New Dictionary
    
    If dictWarningRows.Exists(lRowNo) Then
        dictWarningRows(lRowNo) = dictWarningRows(lRowNo) & DELIMITER & CStr(lCol)
    Else
        dictWarningRows.Add lRowNo, CStr(lCol)
    End If
End Function

