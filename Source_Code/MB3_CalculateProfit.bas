Attribute VB_Name = "MB3_CalculateProfit"
Option Explicit
Option Base 1

Sub subMain_CalculateProfit()
    'If Not fIsDev Then On Error GoTo error_handling
    On Error GoTo error_handling
    shtSalesRawDataRpt.Visible = xlSheetVisible
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

    Call fProcessData
    
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
        shtException.Activate
    End If
    
    If fCheckIfGotBusinessError Then GoTo reset_excel_options
    If fCheckIfUnCapturedExceptionAbnormalError Then GoTo reset_excel_options

    Call fSetReneratedReport(, shtProfit.Name)
    fMsgBox "计算完成，请检查工作表：[" & shtProfit.Name & "] 中，请检查！", vbInformation

reset_excel_options:
    
    Err.Clear
    fEnableExcelOptionsAll
    End
End Sub


Private Function fLoadFilesAndRead2Variables()
    'gsCompanyID
    Call fLoadFileByFileTag("UNIFIED_SALES_INFO")
    Call fReadMasterSheetData("UNIFIED_SALES_INFO", , , True)
End Function


Private Function fProcessData()
    Call fRedimArrOutputBaseArrMaster
    
    Dim lEachRow As Long
    Dim sCompanyLongID As String
    Dim sCompanyName As String
    
    Dim dblGrossPrice As Double
    
    Dim sTmpKey As String
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
        
        dblGrossPrice = fCalculateGrossPrice(lEachRow)
        
        arrOutput(lEachRow, dictRptColIndex("GrossPrice")) = dblGrossPrice
        arrOutput(lEachRow, dictRptColIndex("CostPrice")) = 0
        arrOutput(lEachRow, dictRptColIndex("GrossProfitPerUnit")) = 0
        arrOutput(lEachRow, dictRptColIndex("GrossProfitAmt")) = 0
        arrOutput(lEachRow, dictRptColIndex("SalesMan_1")) = ""
        arrOutput(lEachRow, dictRptColIndex("SalesMan_2")) = ""
        arrOutput(lEachRow, dictRptColIndex("SalesMan_3")) = ""
        arrOutput(lEachRow, dictRptColIndex("SalesManList")) = ""
        arrOutput(lEachRow, dictRptColIndex("SalesCommission_1")) = 0
        arrOutput(lEachRow, dictRptColIndex("SalesCommission_2")) = 0
        arrOutput(lEachRow, dictRptColIndex("SalesCommission_3")) = 0

next_sales:
    Next
     
End Function


Function fCalculateGrossPrice(lEachRow As Long)
    Dim dblFirstLevelComm As Double
    
    Dim sSalesCompName As String
    Dim sSalesCompID As String
    Dim sProducer As String
    Dim sProductName  As String
    Dim sProductSeries As String
    
    sSalesCompName = arrOutput(lEachRow, dictRptColIndex("SalesCompanyName"))
    sProducer = arrOutput(lEachRow, dictRptColIndex("ProductProducer"))
    sProductName = arrOutput(lEachRow, dictRptColIndex("ProductName"))
    sProductSeries = arrOutput(lEachRow, dictRptColIndex("ProductSeries"))

    'sSalesCompID = fGetSalesCompanyID(sSalesCompName)
    dblFirstLevelComm = fGetFirstLevelCommi(sSalesCompName, sProducer, sProductName, sProductSeries)
    
    
End Function
