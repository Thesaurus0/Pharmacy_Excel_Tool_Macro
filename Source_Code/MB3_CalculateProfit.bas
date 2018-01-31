Attribute VB_Name = "MB3_CalculateProfit"
Option Explicit
Option Base 1

Public shtSelfSalesCal As Worksheet

'Dim arrMissed1stLevelComm()
'Dim arrMissed2ndLevelComm()
Dim dictFirstCommColIndex As Dictionary
Dim dictSecondCommColIndex As Dictionary

Dim arrExceptionRows()
Dim mlExcepCnt As Long

Sub subMain_CalculateProfit()
    'If Not fIsDev Then On Error GoTo error_handling
    'On Error GoTo error_handling
    
    If fGetReplaceUnifyErrorRowCount > 0 Then
        fMsgBox "����������������ҩƷ��ϵͳ���Ҳ������޷����������Ӷ�����ȴ�����Щ����"
        shtSalesInfos.Visible = xlSheetVisible
        shtException.Visible = xlSheetVisible:         shtException.Activate
        End
    End If
    
    shtSalesInfos.Visible = xlSheetVisible
    shtException.Visible = xlSheetVeryHidden
    Call fUnProtectSheet(shtProfit)
    Call fCleanSheetOutputResetSheetOutput(shtProfit)
    Call fCleanSheetOutputResetSheetOutput(shtException)
    'shtException.Cells.NumberFormat = "@"
    'shtException.Cells.WrapText = True

    fInitialization

    gsRptID = "CALCULATE_PROFIT"

    Call fReadSysConfig_InputTxtSheetFile

    gsRptFilePath = fReadSysConfig_Output(, gsRptType)
    
    Call fLoadFilesAndRead2Variables

    Call fPrepareOutputSheetHeaderAndTextColumns(shtProfit)

    ReDim arrExceptionRows(1 To UBound(arrMaster, 1) * 4)
    mlExcepCnt = 0
    
    Call fProcessData
    
    If mlExcepCnt > 0 Then
        ReDim Preserve arrExceptionRows(1 To mlExcepCnt)
    Else
        arrExceptionRows = Array()
    End If
    
    If Not shtException.Visible = xlSheetVisible Then shtException.Visible = xlSheetVeryHidden
    
    If mlExcepCnt > 0 Then
        shtException.Visible = xlSheetVisible
    End If
    
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
        
        If Not fFindInWorksheet(shtException.Cells, "�Ҳ����ɿ۵ı���˾������¼", False) Is Nothing Then
            'shtException.Columns(4).ColumnWidth = 100
            Call fFreezeSheet(shtException, , 2)
        End If
        
        shtException.Activate
    End If
    
    If mlExcepCnt > 0 Then Call fSetFormatForExceptionCells(shtProfit, arrExceptionRows, "REPORT_ERROR_COLOR")
    
    fMsgBox "������ɣ����鹤����[" & shtProfit.Name & "] �У����飡", vbInformation
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
    Dim dictNoSalesManConf As Dictionary
    
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
    Dim dblGrossPrice As Double
    Dim dblCostPrice As Double
    Dim sSalesMan_1 As String, sSalesMan_2 As String, sSalesMan_3 As String, sSalesManager As String
    Dim dblComm_1 As Double, dblComm_2 As Double, dblComm_3 As Double, dblSalesMgrComm As Double
    Dim dblGrossProfitAmt As Double
    Dim dblNewRSalesTaxRate As Double
    Dim dblNewRPurchaseTaxRate As Double
    
    Dim dblSellPrice As Double
    
    Set dictMissedFirstLComm = New Dictionary
    Set dictMissedSecondLComm = New Dictionary
    Set dictNoValidSelfSales = New Dictionary
    Set dictNoSalesManConf = New Dictionary
    
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
        
        dblGrossPrice = dblSellPrice * (1 - dblFirstLevelComm) * (1 - dblSecondLevelComm)
        arrOutput(lEachRow, dictRptColIndex("GrossPrice")) = dblGrossPrice
        
        '==== cost price ==========================================
        sProductKey = sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries
        
        If Not fCalculateCostPriceFromSelfSalesOrder(sProductKey, dblQuantity, dblCostPrice) Then
            mlExcepCnt = mlExcepCnt + 1
            arrExceptionRows(mlExcepCnt) = (lEachRow + 1)
            mlExcepCnt = mlExcepCnt + 1
            arrExceptionRows(mlExcepCnt) = dictRptColIndex("CostPrice")

            If Not dictNoValidSelfSales.Exists(sProductKey) Then
                dictNoValidSelfSales.Add sProductKey, "'" & (lEachRow + 1)
            Else
                dictNoValidSelfSales(sProductKey) = dictNoValidSelfSales(sProductKey) & "," & (lEachRow + 1)
            End If
            dblCostPrice = fGetLatestPriceFromProductMaster(sProductKey)
        End If
        
        arrOutput(lEachRow, dictRptColIndex("CostPrice")) = dblCostPrice
                                                                
        If Not fNewRuleProduct(sProductKey) Then
            arrOutput(lEachRow, dictRptColIndex("SalesTaxPerUnit")) = dblGrossPrice * fGetTaxRate
            arrOutput(lEachRow, dictRptColIndex("PurchaeTaxPerUnit")) = 0
        Else
            Call fGetNewRuleProductTaxRate(sProductKey, dblNewRSalesTaxRate, dblNewRPurchaseTaxRate)
            
            arrOutput(lEachRow, dictRptColIndex("SalesTaxPerUnit")) = (dblGrossPrice - dblCostPrice) * dblNewRSalesTaxRate
            arrOutput(lEachRow, dictRptColIndex("PurchaeTaxPerUnit")) = (dblGrossPrice - dblCostPrice - arrOutput(lEachRow, dictRptColIndex("SalesTaxPerUnit"))) _
                                                    * dblNewRPurchaseTaxRate
        End If
        
        arrOutput(lEachRow, dictRptColIndex("GrossProfitPerUnit")) = dblGrossPrice - dblCostPrice _
                                                            - arrOutput(lEachRow, dictRptColIndex("SalesTaxPerUnit")) _
                                                            - arrOutput(lEachRow, dictRptColIndex("PurchaeTaxPerUnit"))
        
        dblGrossProfitAmt = arrOutput(lEachRow, dictRptColIndex("GrossProfitPerUnit")) * dblQuantity
        arrOutput(lEachRow, dictRptColIndex("GrossProfitAmt")) = dblGrossProfitAmt
        '-----------------------------------------------------------------------------------------------
        
        '==== salesman commission ==========================================
        sSalesManKey = sSalesCompName & DELIMITER & sHospital & DELIMITER _
                    & sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries & DELIMITER & dblSellPrice
        If Not fCalculateSalesManCommissionFromshtSalesManCommConfig(sSalesManKey _
                                , sSalesMan_1, sSalesMan_2, sSalesMan_3, dblComm_1, dblComm_2, dblComm_3 _
                                , sSalesManager, dblSalesMgrComm) Then
            If Not dictNoSalesManConf.Exists(sSalesManKey) Then
                dictNoSalesManConf.Add sSalesManKey, "'" & lEachRow + 1
            Else
                dictNoSalesManConf(sSalesManKey) = dictNoSalesManConf(sSalesManKey) & "," & (lEachRow + 1)
            End If
            
            mlExcepCnt = mlExcepCnt + 1:    arrExceptionRows(mlExcepCnt) = lEachRow + 1
            mlExcepCnt = mlExcepCnt + 1:    arrExceptionRows(mlExcepCnt) = dictRptColIndex("SalesMan_1")
        End If

        arrOutput(lEachRow, dictRptColIndex("SalesMan_1")) = sSalesMan_1
        arrOutput(lEachRow, dictRptColIndex("SalesMan_2")) = sSalesMan_2
        arrOutput(lEachRow, dictRptColIndex("SalesMan_3")) = sSalesMan_3
        arrOutput(lEachRow, dictRptColIndex("SalesManList")) = sSalesMan_1 _
                                                             & IIf(Len(sSalesMan_2) > 0, ", " & sSalesMan_2, "") _
                                                             & IIf(Len(sSalesMan_3) > 0, ", " & sSalesMan_3, "")
        arrOutput(lEachRow, dictRptColIndex("SalesCommission_1")) = dblComm_1
        arrOutput(lEachRow, dictRptColIndex("SalesCommission_2")) = dblComm_2
        arrOutput(lEachRow, dictRptColIndex("SalesCommission_3")) = dblComm_3
        '-----------------------------------------------------------------------------------------------
        
        arrOutput(lEachRow, dictRptColIndex("NetProfitPerUnit")) = arrOutput(lEachRow, dictRptColIndex("GrossProfitPerUnit")) _
                                            - arrOutput(lEachRow, dictRptColIndex("SalesCommission_1")) _
                                            - arrOutput(lEachRow, dictRptColIndex("SalesCommission_2")) _
                                            - arrOutput(lEachRow, dictRptColIndex("SalesCommission_3"))
        arrOutput(lEachRow, dictRptColIndex("NetProfitAmount")) = arrOutput(lEachRow, dictRptColIndex("NetProfitPerUnit")) _
                                                                * dblQuantity
        
        arrOutput(lEachRow, dictRptColIndex("SalesManagerCommissoin")) = arrOutput(lEachRow, dictRptColIndex("NetProfitAmount")) _
                                                                * dblSalesMgrComm

        arrOutput(lEachRow, dictRptColIndex("SalesManList")) = fComposeSalesManList(sSalesManager, sSalesMan_1, sSalesMan_2, sSalesMan_3)
        arrOutput(lEachRow, dictRptColIndex("SalesCommissionAmt_1")) = dblComm_1 * dblQuantity
        arrOutput(lEachRow, dictRptColIndex("SalesCommissionAmt_2")) = dblComm_2 * dblQuantity
        arrOutput(lEachRow, dictRptColIndex("SalesCommissionAmt_3")) = dblComm_3 * dblQuantity
next_sales:
    Next
    
    Dim lStartRow As Long
    Dim lRecCount As Long
    Dim arrTmp()
    
    lRecCount = fGetDictionayDelimiteredItemsCount(dictMissedFirstLComm)
    If lRecCount > 0 Then
        lStartRow = fGetValidMaxRow(shtFirstLevelCommission) + 1
        
        arrTmp = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictMissedFirstLComm)
        Call fAppendArray2Sheet(shtFirstLevelCommission, arrTmp)
        
        fMsgBox lRecCount & "�����������¼�Ĳ�֥�ֵ����ͷ�û�����ã�ϵͳ�Ѿ��Զ���������ӵ��ˡ�" & shtFirstLevelCommission.Name & "��" _
            & vbCr & "�����Բ鿴�ñ�������������"
    End If
    
    lRecCount = fGetDictionayDelimiteredItemsCount(dictMissedSecondLComm)
    If lRecCount > 0 Then
        lStartRow = fGetValidMaxRow(shtSecondLevelCommission) + 1
        
        arrTmp = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictMissedSecondLComm)
        Call fAppendArray2Sheet(shtSecondLevelCommission, arrTmp)
        
        fMsgBox lRecCount & "�����������¼����ҵ��˾�����ͷ�û�����ã�ϵͳ�Ѿ��Զ���������ӵ��ˡ�" & shtSecondLevelCommission.Name & "��" _
            & vbCr & "�����Բ鿴�ñ�������������"
    End If
    
    Call fSetBackToshtSelfSalesCalWithDeductedData
    Call fAddNoValidSelfSalesToSheetException(dictNoValidSelfSales)
    Call fAddNoSalesManConfToSheetException(dictNoSalesManConf)
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
        
        shtException.Columns(4).ColumnWidth = 100
        shtException.Cells(lStartRow - 1, 1).Value = "�Ҳ����ɿ۵ı���˾������¼(��Ϊ�˻��������Ҳ���ҽԺ���۵ֿ�)"
        Call fPrepareHeaderToSheet(shtException, Array("ҩƷ����", "ҩƷ����", "���", "�к�"), lStartRow)
        shtException.Rows(lStartRow - 1 & ":" & lStartRow).Font.Color = RGB(255, 0, 0)
        shtException.Rows(lStartRow - 1 & ":" & lStartRow).Font.Bold = True
        Call fAppendArray2Sheet(shtException, arrNewProductSeries)
        'sErr = fUbound(arrNewProductSeries)
        
        lRecCount = fGetDictionayDelimiteredItemsCount(dictNoValidSelfSales)
        
        shtException.Cells(lStartRow + 1, 4).Resize(dictNoValidSelfSales.Count, 1).Value = fConvertDictionaryItemsTo2DimenArrayForPaste(dictNoValidSelfSales, False)
        Erase arrNewProductSeries
        If lStartRow = 2 Then Call fFreezeSheet(shtException, , 2)
        
        shtException.Visible = xlSheetVisible
        shtException.Activate
        
        If fNzero(gsBusinessErrorMsg) Then gsBusinessErrorMsg = gsBusinessErrorMsg & vbCr & vbCr & vbCr & "===============================" & vbCr & vbCr

        gsBusinessErrorMsg = gsBusinessErrorMsg & lUniqRecCnt & "��ҩƷ" & lRecCount & "�����������ڱ���˾������¼���޳���ɿ۳�(��Ϊ�˻��������Ҳ���ҽԺ���۵ֿ�)��������Ҫ��" & vbCr _
            & "(1). �ڡ�����˾�����������һ���滻��¼" & vbCr _
            & "(2). �ڡ�ҩƷ�������޸������¼۸�" & vbCr & vbCr _
            & "������Щ����������е�һ�룬û�п��Կ۵ĳ�����¼�����԰����ǵĳɱ��۸��ע0��"
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
        
        shtException.Cells(lStartRow - 1, 1).Value = "�Ҳ���ҵ��Ա�ļ�¼  --> �п���ֻ���м۱�û�У�"
        Call fPrepareHeaderToSheet(shtException, Array("��ҵ��˾", "ҽԺ", "ҩƷ����", "ҩƷ����", "���", "�б��", "�к�"), lStartRow)
        shtException.Rows(lStartRow - 1 & ":" & lStartRow).Font.Color = RGB(255, 0, 0)
        shtException.Rows(lStartRow - 1 & ":" & lStartRow).Font.Bold = True
        
        arrNoSalesMan = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictNoSalesManConf, , False)
        Call fAppendArray2Sheet(shtException, arrNoSalesMan)
        
        shtException.Cells(lStartRow + 1, 7).Resize(lUniqRecCnt, 1).Value = fConvertDictionaryItemsTo2DimenArrayForPaste(dictNoSalesManConf)

        If lStartRow = 2 Then Call fFreezeSheet(shtException, , 2)
        
        If fNzero(gsBusinessErrorMsg) Then gsBusinessErrorMsg = gsBusinessErrorMsg & vbCr & vbCr & vbCr & "===============================" & vbCr & vbCr
        
        gsBusinessErrorMsg = gsBusinessErrorMsg & lUniqRecCnt & "�����������Ҳ���ҵ��Ա�ļ�¼������Ҫ��" & vbCr _
            & "(1). �ڡ�ҵ��ԱӶ�����á������ҵ��ԱӶ������"
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

'Function fCalculateGrossPrice(lEachRow As Long, ByRef dictMissedFirstLComm As Dictionary, ByRef dictMissedSecondLComm As Dictionary) As Double
'    Dim dblGrossPrice As Double
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
'    dblGrossPrice = dblSellPrice * (1 - dblFirstLevelComm) * (1 - dblSecondLevelComm)
'
'    fCalculateGrossPrice = dblGrossPrice
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
    If lNewRow = 0 Then
        lNewRow = lNewRow + 2
    Else
        lNewRow = lNewRow + 5
    End If
    
    fGetshtExceptionNewRow = lNewRow
End Function

Sub subMain_CalculateProfit_MonthEnd()
    Dim response As VbMsgBoxResult
    response = MsgBox(Prompt:="�ü����ۼ����⣬�޷���������ȷ��Ҫ���м��������Ӷ����" _
                        & vbCr & "��������㡾Yes��" & vbCr & "������㡾No��" _
                        , Buttons:=vbCritical + vbYesNo + vbDefaultButton2)
    If response <> vbYes Then Exit Sub
    
    Set shtSelfSalesCal = shtSelfSalesOrder
    Call subMain_CalculateProfit
End Sub

Sub subMain_CalculateProfit_PreCal()
    Set shtSelfSalesCal = shtSelfSalesPreDeduct
    Call fRemoveFilterForSheet(shtSelfSalesOrder)
    Call fRemoveFilterForSheet(shtSelfSalesPreDeduct)
    shtSelfSalesOrder.Cells.Copy shtSelfSalesCal.Cells
    Call subMain_CalculateProfit
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


Function fGetReplaceUnifyErrorRowCount() As Long
    fGetReplaceUnifyErrorRowCount = CLng(fGetSpecifiedConfigCellValue(shtSysConf, "[Facility For Testing]", "Value", "Setting Item ID=REPLACE_UNIFY_ERR_ROW_COUNT"))
End Function

