Attribute VB_Name = "MB3_CalculateProfit"
Option Explicit
Option Base 1

'Dim arrMissed1stLevelComm()
'Dim arrMissed2ndLevelComm()
Dim dictFirstCommColIndex As Dictionary
Dim dictSecondCommColIndex As Dictionary

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
    Dim dictMissedFirstLComm As Dictionary
    Dim dictMissedSecondLComm As Dictionary
    
    Dim sCompanyLongID As String
    Dim sCompanyName As String
    
    Dim dblGrossPrice As Double
    
    Set dictMissedFirstLComm = New Dictionary
    Set dictMissedSecondLComm = New Dictionary
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
    
    Dim arrMissedFistLComm()
    Dim sErr As String
    If dictMissedFirstLComm.Count > 0 Then
        arrMissedFistLComm = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictMissedFirstLComm, , False)
        Dim lStartRow As Long
        lStartRow = fGetValidMaxRow(shtFirstLevelCommission) + 1
        
        Call fAppendArray2Sheet(shtFirstLevelCommission, dictMissedFirstLComm)
        sErr = fUbound(dictMissedFirstLComm)
        Erase arrMissedFistLComm
        
        fMsgBox sErr & "条销售流向记录的采芝林的配送费没有设置，系统已经自动把它们添加到了【" & shtFirstLevelCommission.Name & "】" _
            & vbCr & "您可以查看该表中最后面的数据"
    End If
    
    Dim arrMissedSecondLComm()
    Dim sErr As String
    If dictMissedSecondLComm.Count > 0 Then
        arrMissedFistLComm = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictMissedSecondLComm, , False)
        Dim lStartRow As Long
        lStartRow = fGetValidMaxRow(shtSecondLevelCommission) + 1
        
        Call fAppendArray2Sheet(shtSecondLevelCommission, dictMissedSecondLComm)
        sErr = fUbound(dictMissedSecondLComm)
        Erase arrMissedFistLComm
        
        fMsgBox sErr & "条销售流向记录的商业公司的配送费没有设置，系统已经自动把它们添加到了【" & shtSecondLevelCommission.Name & "】" _
            & vbCr & "您可以查看该表中最后面的数据"
    End If
     
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
    If dictFirstCommColIndex Is Nothing Then dictFirstCommColIndex = fReadInputFileSpecConfigItem("FIRST_LEVEL_COMMISSION", "LETTER_INDEX", shtFirstLevelCommission)
    
    Dim i As Integer
    Dim arr() As String
    
    ReDim arr(0 To dictFirstCommColIndex.Count - 1)
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
    If dictSecondCommColIndex Is Nothing Then dictSecondCommColIndex = fReadInputFileSpecConfigItem("SECOND_LEVEL_COMMISSION", "LETTER_INDEX", shtFirstLevelCommission)
    
    Dim i As Integer
    Dim arr() As String
    
    ReDim arr(0 To dictSecondCommColIndex.Count - 1)
    arr(dictSecondCommColIndex("SalesCompany")) = sSalesCompName
    arr(dictSecondCommColIndex("Hospital")) = sHospital
    arr(dictSecondCommColIndex("ProductProducer")) = sProducer
    arr(dictSecondCommColIndex("ProductName")) = sProductName
    arr(dictSecondCommColIndex("ProductSeries")) = sProductSeries
    arr(dictSecondCommColIndex("Commission")) = dblComm

    fComposeSecondLevelColumnsStryByConfig = Join(arr, DELIMITER)
    Erase arr
End Function


