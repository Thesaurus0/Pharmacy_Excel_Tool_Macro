Attribute VB_Name = "MB3_CalculateProfit"
Option Explicit
Option Base 1

'Dim arrMissed1stLevelComm()
'Dim arrMissed2ndLevelComm()
Dim dictFirstCommColIndex As Dictionary
Dim dictSecondCommColIndex As Dictionary

Dim arrExceptionRows()
Dim mlExcepCnt As Long

Sub subMain_CalculateProfit()
    'If Not fIsDev Then On Error GoTo error_handling
    'On Error GoTo error_handling
    shtSalesInfos.Visible = xlSheetVisible
    shtException.Visible = xlSheetVeryHidden
    Call fUnProtectSheet(shtProfit)
    Call fCleanSheetOutputResetSheetOutput(shtProfit)
    Call fCleanSheetOutputResetSheetOutput(shtException)

    fInitialization

    gsRptID = "CALCULATE_PROFIT"

    Call fReadSysConfig_InputTxtSheetFile

    gsRptFilePath = fReadSysConfig_Output(, gsRptType)
    
    Call fLoadFilesAndRead2Variables

    Call fPrepareOutputSheetHeaderAndTextColumns(shtProfit)

    ReDim arrExceptionRows(1 To UBound(arrMaster, 1) * 2)
    mlExcepCnt = 0
    
    Call fProcessData
    If mlExcepCnt > 0 Then
        ReDim Preserve arrExceptionRows(1 To mlExcepCnt)
    Else
        arrExceptionRows = Array()
    End If
    
    If Not shtException.Visible = xlSheetVisible Then shtException.Visible = xlSheetVeryHidden
    
    'If shtException.Visible = xlSheetVisible Then
        Call fAppendArray2Sheet(shtProfit, arrOutput)
    
    
        'Call fReSequenceSeqNo
    
    '    Call fSortDataInSheetSortSheetData(shtSalesRawDataRpt, Array(dictRptColIndex("SalesCompanyName") _
                                                                    , dictRptColIndex("Hospital") _
                                                                    , dictRptColIndex("SalesDate") _
                                                                    , dictRptColIndex("ProductProducer") _
                                                                    , dictRptColIndex("ProductName") _
                                                                    , dictRptColIndex("ProductUnit")))
        Call fFormatOutputSheet(shtProfit)
    
       ' Call fProtectSheetAndAllowEdit(shtSalesRawDataRpt, shtSalesRawDataRpt.Columns(4), UBound(arrOutput, 1) + 1, UBound(arrOutput, 2), False)
        Call fPostProcess(shtProfit)
    
        shtProfit.Visible = xlSheetVisible
        shtProfit.Activate
        shtProfit.Range("A1").Select
        
error_handling:
    If shtException.Visible = xlSheetVisible Then
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
    
    If mlExcepCnt > 0 Then Call fSetFormatForExceptionCells(shtProfit, arrExceptionRows, "REPORT_NO_VALID_SELF_SALES_EXCEPTION_COLOR")
    
    fMsgBox "计算完成，请检查工作表：[" & shtProfit.Name & "] 中，请检查！", vbInformation
    Call fSetReneratedReport(, shtProfit.Name)
    If fCheckIfGotBusinessError Then GoTo reset_excel_options
    If fCheckIfUnCapturedExceptionAbnormalError Then GoTo reset_excel_options
    

reset_excel_options:
    
    Err.Clear
    fEnableExcelOptionsAll
    End
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
    Call fRedimArrOutputBaseArrMaster
    
    Dim lEachRow As Long
    Dim dictMissedFirstLComm As Dictionary
    Dim dictMissedSecondLComm As Dictionary
    Dim dictNoValidSelfSales As Dictionary
    
    Dim sCompanyLongID As String
    Dim sCompanyName As String
    
    Dim dblGrossPrice As Double
    Dim dblCostPrice As Double
    
    Set dictMissedFirstLComm = New Dictionary
    Set dictMissedSecondLComm = New Dictionary
    Set dictNoValidSelfSales = New Dictionary
    For lEachRow = LBound(arrMaster, 1) To UBound(arrMaster, 1)
        If dictMstColIndex.Exists("OrigSalesInfoID") Then
            arrOutput(lEachRow, dictRptColIndex("OrigSalesInfoID")) = arrMaster(lEachRow, dictMstColIndex("OrigSalesInfoID"))
        End If
        
        If dictMstColIndex.Exists("SeqNo") Then
            arrOutput(lEachRow, dictRptColIndex("SeqNo")) = arrMaster(lEachRow, dictMstColIndex("SeqNo"))
        End If

        arrOutput(lEachRow, dictRptColIndex("SalesCompanyName")) = arrMaster(lEachRow, dictMstColIndex("SalesCompanyName"))
        arrOutput(lEachRow, dictRptColIndex("SalesDate")) = arrMaster(lEachRow, dictMstColIndex("SalesDate"))
        arrOutput(lEachRow, dictRptColIndex("ProductProducer")) = arrMaster(lEachRow, dictMstColIndex("MatchedProductProducer"))
        arrOutput(lEachRow, dictRptColIndex("ProductName")) = arrMaster(lEachRow, dictMstColIndex("MatchedProductName"))
        arrOutput(lEachRow, dictRptColIndex("ProductSeries")) = arrMaster(lEachRow, dictMstColIndex("MatchedProductSeries"))
        arrOutput(lEachRow, dictRptColIndex("ProductUnit")) = arrMaster(lEachRow, dictMstColIndex("MatchedProductUnit"))
        arrOutput(lEachRow, dictRptColIndex("Hospital")) = arrMaster(lEachRow, dictMstColIndex("MatchedHospital"))
        arrOutput(lEachRow, dictRptColIndex("Quantity")) = arrMaster(lEachRow, dictMstColIndex("ConvertQuantity"))
        arrOutput(lEachRow, dictRptColIndex("SellPrice")) = arrMaster(lEachRow, dictMstColIndex("ConvertSellPrice"))
        arrOutput(lEachRow, dictRptColIndex("SellAmount")) = arrMaster(lEachRow, dictMstColIndex("RecalSellAmount"))
        
        dblGrossPrice = fCalculateGrossPrice(lEachRow, dictMissedFirstLComm, dictMissedSecondLComm)
        arrOutput(lEachRow, dictRptColIndex("GrossPrice")) = dblGrossPrice
        
        dblCostPrice = fCalculateCostPrice(lEachRow, dictNoValidSelfSales)
        arrOutput(lEachRow, dictRptColIndex("CostPrice")) = dblCostPrice
        
        arrOutput(lEachRow, dictRptColIndex("GrossProfitPerUnit")) = dblGrossPrice - dblCostPrice
        arrOutput(lEachRow, dictRptColIndex("GrossProfitAmt")) = (dblGrossPrice - dblCostPrice) * arrOutput(lEachRow, dictRptColIndex("Quantity"))
        arrOutput(lEachRow, dictRptColIndex("TaxAmount")) = arrOutput(lEachRow, dictRptColIndex("GrossProfitAmt")) * fGetTaxRate
        
        Call fCalculateSalesManCommission(lEachRow, sSalesMan_1, sSalesMan_2, sSalesMan_3, dblComm_1, dblComm_2, dblComm_3)
        
        arrOutput(lEachRow, dictRptColIndex("SalesMan_1")) = ""
        arrOutput(lEachRow, dictRptColIndex("SalesMan_2")) = ""
        arrOutput(lEachRow, dictRptColIndex("SalesMan_3")) = ""
        arrOutput(lEachRow, dictRptColIndex("SalesManList")) = ""
        arrOutput(lEachRow, dictRptColIndex("SalesCommission_1")) = 0
        arrOutput(lEachRow, dictRptColIndex("SalesCommission_2")) = 0
        arrOutput(lEachRow, dictRptColIndex("SalesCommission_3")) = 0
next_sales:
    Next
    
    Dim arrMissedFistLComm()
    Dim sErr As String
    If dictMissedFirstLComm.Count > 0 Then
        arrMissedFistLComm = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictMissedFirstLComm, , False)
        Dim lStartRow As Long
        lStartRow = fGetValidMaxRow(shtFirstLevelCommission) + 1
        
        Call fAppendArray2Sheet(shtFirstLevelCommission, arrMissedFistLComm)
        sErr = fUbound(arrMissedFistLComm)
        Erase arrMissedFistLComm
        
        fMsgBox sErr & "条销售流向记录的采芝林的配送费没有设置，系统已经自动把它们添加到了【" & shtFirstLevelCommission.Name & "】" _
            & vbCr & "您可以查看该表中最后面的数据"
    End If
    
    Dim arrMissedSecondLComm()
    If dictMissedSecondLComm.Count > 0 Then
        arrMissedSecondLComm = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictMissedSecondLComm, , False)
        lStartRow = fGetValidMaxRow(shtSecondLevelCommission) + 1
        
        Call fAppendArray2Sheet(shtSecondLevelCommission, arrMissedSecondLComm)
        sErr = fUbound(arrMissedSecondLComm)
        Erase arrMissedSecondLComm
        
        fMsgBox sErr & "条销售流向记录的商业公司的配送费没有设置，系统已经自动把它们添加到了【" & shtSecondLevelCommission.Name & "】" _
            & vbCr & "您可以查看该表中最后面的数据"
    End If
    
    Call fSetBackToshtSelfSalesOrderWithDeductedData
    Call fAddNoValidSelfSalesToSheetException(dictNoValidSelfSales)
End Function

Function fAddNoValidSelfSalesToSheetException(dictNoValidSelfSales As Dictionary)
    Dim arrNewProductSeries()
    Dim sErr As String
    Dim lRecCount As Long
    Dim i As Integer
    Dim j As Integer
    Dim arrTmp
    
    If dictNoValidSelfSales.Count > 0 Then
        arrNewProductSeries = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictNoValidSelfSales, , False)
        Dim lStartRow As Long
        lStartRow = fGetValidMaxRow(shtException)
        If lStartRow = 0 Then
            lStartRow = lStartRow + 2
        Else
            lStartRow = lStartRow + 6
        End If
        
        shtException.Cells.NumberFormat = "@"
        shtException.Cells.WrapText = True
        shtException.Columns(4).ColumnWidth = 100
        shtException.Cells(lStartRow - 1, 1).Value = "找不到可扣的本公司出货记录"
        Call fPrepareHeaderToSheet(shtException, Array("药品厂家", "药品名称", "规格", "行号"), lStartRow)
        shtException.Rows(lStartRow - 1 & ":" & lStartRow).Font.Color = RGB(255, 0, 0)
        shtException.Rows(lStartRow - 1 & ":" & lStartRow).Font.Bold = True
        Call fAppendArray2Sheet(shtException, arrNewProductSeries)
        sErr = fUbound(arrNewProductSeries)
        
        lRecCount = 0
        For i = 0 To dictNoValidSelfSales.Count - 1
            'lRecCount = lRecCount + UBound(Split(dictNoValidSelfSales.Items(i), ",")) - 1
            arrTmp = Split(dictNoValidSelfSales.Items(i), ",")
            lRecCount = lRecCount + UBound(arrTmp)
            
            For j = 0 To UBound(arrTmp)
                mlExcepCnt = mlExcepCnt + 1
                arrExceptionRows(mlExcepCnt) = arrTmp(j)
                mlExcepCnt = mlExcepCnt + 1
                arrExceptionRows(mlExcepCnt) = dictRptColIndex("CostPrice")
            Next
            
            Erase arrTmp
        Next
        
        shtException.Cells(lStartRow + 1, 4).Resize(sErr, 1).Value = fConvertDictionaryItemsTo2DimenArrayForPaste(dictNoValidSelfSales)
        Erase arrNewProductSeries
        If lStartRow = 2 Then Call fFreezeSheet(shtException, , 2)
        
        shtException.Visible = xlSheetVisible
        shtException.Activate
        
        If fNzero(gsBusinessErrorMsg) Then gsBusinessErrorMsg = gsBusinessErrorMsg & vbCr & vbCr & vbCr & "===============================" & vbCr & vbCr

        gsBusinessErrorMsg = gsBusinessErrorMsg & sErr & "个药品" & lRecCount & "条销售流向在本公司出货记录中无出库可扣除，您可能要：" & vbCr _
            & "(1). 在【本公司出货】中添加一条替换记录" & vbCr _
            & "(2). 在【药品主表】中修改其最新价格" & vbCr & vbCr _
            & "计算这些销售流向进行到一半，没有可以扣的出货记录，所以把它们的成本价格标注0，"
    End If
End Function

Function fCalculateCostPrice(lEachRow As Long, ByRef dictNoValidSelfSales As Dictionary) As Double
    Dim dblCostPrice As Double
    
    Dim sProducer As String
    Dim sProductName  As String
    Dim sProductSeries As String
    Dim dblSalesQuantity As Double
    
    Dim sTmpKey As String
    
    sProducer = Trim(arrOutput(lEachRow, dictRptColIndex("ProductProducer")))
    sProductName = Trim(arrOutput(lEachRow, dictRptColIndex("ProductName")))
    sProductSeries = Trim(arrOutput(lEachRow, dictRptColIndex("ProductSeries")))
    
    dblSalesQuantity = arrOutput(lEachRow, dictRptColIndex("Quantity"))
    
    If Not fCalculateCostPriceFromSelfSalesOrder(sProducer, sProductName, sProductSeries, dblSalesQuantity, dblCostPrice) Then
        sTmpKey = sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries
        
        If Not dictNoValidSelfSales.Exists(sTmpKey) Then
            dictNoValidSelfSales.Add sTmpKey, lEachRow + 1
        Else
            dictNoValidSelfSales(sTmpKey) = dictNoValidSelfSales(sTmpKey) & "," & (lEachRow + 1)
        End If
        dblCostPrice = fGetLatestPriceFromProductMaster(sProducer, sProductName, sProductSeries)
    End If

    fCalculateCostPrice = dblCostPrice
End Function

Function fCalculateGrossPrice(lEachRow As Long, ByRef dictMissedFirstLComm As Dictionary, ByRef dictMissedSecondLComm As Dictionary) As Double
    Dim dblGrossPrice As Double
    
    Dim dblFirstLevelComm As Double
    Dim dblSecondLevelComm As Double
    
    Dim sHospital As String
    Dim sSalesCompName As String
    Dim sSalesCompID As String
    Dim sProducer As String
    Dim sProductName  As String
    Dim sProductSeries As String
    
    Dim sTmpKey As String
    
    sHospital = Trim(arrOutput(lEachRow, dictRptColIndex("Hospital")))
    sSalesCompName = Trim(arrOutput(lEachRow, dictRptColIndex("SalesCompanyName")))
    sProducer = Trim(arrOutput(lEachRow, dictRptColIndex("ProductProducer")))
    sProductName = Trim(arrOutput(lEachRow, dictRptColIndex("ProductName")))
    sProductSeries = Trim(arrOutput(lEachRow, dictRptColIndex("ProductSeries")))

    'sSalesCompID = fGetSalesCompanyID(sSalesCompName)
    If Not fGetFirstLevelComm(sSalesCompName, sProducer, sProductName, sProductSeries, dblFirstLevelComm) Then
        dblFirstLevelComm = fGetConfigFirstLevelDefaultComm()
        
        'sTmpKey = sSalesCompName & DELIMITER & sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries
        sTmpKey = fComposeFirstLevelColumnsStryByConfig(sSalesCompName, sProducer, sProductName, sProductSeries, dblFirstLevelComm)
        If Not dictMissedFirstLComm.Exists(sTmpKey) Then
            dictMissedFirstLComm.Add sTmpKey, lEachRow + 1
        End If
    End If
    
    If Not fGetSecondLevelComm(sSalesCompName, sHospital, sProducer, sProductName, sProductSeries, dblSecondLevelComm) Then
        dblSecondLevelComm = fGetConfigSecondLevelDefaultComm(sSalesCompName)
        
        'sTmpKey = sSalesCompName & DELIMITER & sHospital & DELIMITER & sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries
        sTmpKey = fComposeSecondLevelColumnsStryByConfig(sSalesCompName, sHospital, sProducer, sProductName, sProductSeries, dblSecondLevelComm)
        If Not dictMissedSecondLComm.Exists(sTmpKey) Then
            dictMissedSecondLComm.Add sTmpKey, lEachRow + 1
        End If
    End If
    
    Dim dblSellPrice As Double
    dblSellPrice = arrOutput(lEachRow, dictRptColIndex("SellPrice"))
    dblGrossPrice = dblSellPrice * (1 - dblFirstLevelComm) * (1 - dblSecondLevelComm)
    
    fCalculateGrossPrice = dblGrossPrice
End Function

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
    If dictSecondCommColIndex Is Nothing Then Set dictSecondCommColIndex = fReadInputFileSpecConfigItem("SECOND_LEVEL_COMMISSION", "LETTER_INDEX", shtSecondLevelCommission)
    
    Dim i As Integer
    Dim arr() As String
    
    ReDim arr(1 To dictSecondCommColIndex.Count)
    arr(dictSecondCommColIndex("SalesCompany")) = sSalesCompName
    arr(dictSecondCommColIndex("Hospital")) = sHospital
    arr(dictSecondCommColIndex("ProductProducer")) = sProducer
    arr(dictSecondCommColIndex("ProductName")) = sProductName
    arr(dictSecondCommColIndex("ProductSeries")) = sProductSeries
    arr(dictSecondCommColIndex("Commission")) = dblComm

    fComposeSecondLevelColumnsStryByConfig = Join(arr, DELIMITER)
    Erase arr
End Function

Function fCalculateSalesManCommission(lEachRow As Long, ByRef sSalesMan_1 As String, ByRef sSalesMan_2 As String, ByRef sSalesMan_3 As String _
                            , ByRef dblComm_1 As Double, ByRef dblComm_2 As Double, ByRef dblComm_3 As Double)
    sSalesMan_1 = ""
    sSalesMan_2 = ""
    sSalesMan_3 = ""
    dblComm_1 = 0
    dblComm_2 = 0
    dblComm_3 = 0
    
    
    Dim sHospital As String
    Dim sSalesCompName As String
    Dim sSalesCompID As String
    Dim sProducer As String
    Dim sProductName  As String
    Dim sProductSeries As String
    
    Dim sTmpKey As String
    
    sHospital = Trim(arrOutput(lEachRow, dictRptColIndex("Hospital")))
    sSalesCompName = Trim(arrOutput(lEachRow, dictRptColIndex("SalesCompanyName")))
    sProducer = Trim(arrOutput(lEachRow, dictRptColIndex("ProductProducer")))
    sProductName = Trim(arrOutput(lEachRow, dictRptColIndex("ProductName")))
    sProductSeries = Trim(arrOutput(lEachRow, dictRptColIndex("ProductSeries")))
    
    
    sTmpKey = fComposeFirstLevelColumnsStryByConfig(sSalesCompName, sProducer, sProductName, sProductSeries, dblFirstLevelComm)
    
    If Not fCalculateSalesManCommissionFromshtSalesManCommConfig(sSalesCompName, sHospital, sProducer, sProductName _
                                    , sProductSeries, dblSecondLevelComm) Then
        dblFirstLevelComm = fGetConfigFirstLevelDefaultComm()
        
        'sTmpKey = sSalesCompName & DELIMITER & sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries
        sTmpKey = fComposeFirstLevelColumnsStryByConfig(sSalesCompName, sProducer, sProductName, sProductSeries, dblFirstLevelComm)
        If Not dictMissedFirstLComm.Exists(sTmpKey) Then
            dictMissedFirstLComm.Add sTmpKey, lEachRow + 1
        End If
    End If
End Function

Sub subMain_CalculateProfit_MonthEnd()
    
End Sub
