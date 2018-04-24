Attribute VB_Name = "MB5_ReplaceInventory"
Option Explicit
Option Base 1

Dim arrErrRows()
Dim arrWarningRows()
Dim alErrCnt As Long
Dim alWarningCnt As Long

Sub subMain_ReplaceInventory()
    'If Not fIsDev Then On Error GoTo error_handling
   'On Error GoTo error_handling
    Call fSetReplaceUnifyErrorRowCount_SCompInventory(999)
    
    shtInventoryRawDataRpt.Visible = xlSheetVisible
    shtException.Visible = xlSheetVeryHidden
    Call fUnProtectSheet(shtSalesCompInvUnified)
    Call fCleanSheetOutputResetSheetOutput(shtSalesCompInvUnified)
    Call fCleanSheetOutputResetSheetOutput(shtException)
    'shtException.Cells.NumberFormat = "@"
    shtException.Cells.WrapText = True
    
    fInitialization

    gsRptID = "REPLACE_UNIFY_INVENTORY"

    Call fReadSysConfig_InputTxtSheetFile

    gsRptFilePath = fReadSysConfig_Output(, gsRptType)
    Call fLoadFilesAndRead2Variables

    Call fPrepareOutputSheetHeaderAndTextColumns(shtSalesCompInvUnified)

    ReDim arrErrRows(1 To UBound(arrMaster, 1) * 2)
    alErrCnt = 0
    ReDim arrWarningRows(1 To UBound(arrMaster, 1) * 2)
    alWarningCnt = 0
    
    Call fProcessData
    If alErrCnt > 0 Then
        ReDim Preserve arrErrRows(1 To alErrCnt)
    Else
        arrErrRows = Array()
    End If
    If alWarningCnt > 0 Then
        ReDim Preserve arrWarningRows(1 To alWarningCnt)
    Else
        arrWarningRows = Array()
    End If
    
    If alErrCnt > 0 Or alWarningCnt > 0 Then
        shtException.Visible = xlSheetVisible
    End If
    
error_handling:
    'If shtException.Visible = xlSheetVisible Then
        Call fAppendArray2Sheet(shtSalesCompInvUnified, arrOutput)
    
    
        'Call fReSequenceSeqNo
    
    '    Call fSortDataInSheetSortSheetData(shtInventoryRawDataRpt, Array(dictRptColIndex("SalesCompanyName") _
                                                                    , dictRptColIndex("Hospital") _
                                                                    , dictRptColIndex("SalesDate") _
                                                                    , dictRptColIndex("ProductProducer") _
                                                                    , dictRptColIndex("ProductName") _
                                                                    , dictRptColIndex("ProductUnit")))
        Call fFormatOutputSheet(shtSalesCompInvUnified)
    
       ' Call fProtectSheetAndAllowEdit(shtInventoryRawDataRpt, shtInventoryRawDataRpt.Columns(4), UBound(arrOutput, 1) + 1, UBound(arrOutput, 2), False)
        Call fPostProcess(shtSalesCompInvUnified)
        
        If alErrCnt > 0 Then Call fSetFormatForExceptionCells(shtSalesCompInvUnified, arrErrRows, "REPORT_ERROR_COLOR")
        If alWarningCnt > 0 Then Call fSetFormatForExceptionCells(shtSalesCompInvUnified, arrWarningRows, "REPORT_WARNING_COLOR")
        Call fSetReplaceUnifyErrorRowCount_SCompInventory(alErrCnt / 2)
    
        shtSalesCompInvUnified.Visible = xlSheetVisible
        shtSalesCompInvUnified.Activate
        shtSalesCompInvUnified.Range("A1").Select
        
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

    Call fSetReneratedReport(, shtSalesCompInvUnified.Name)
    fMsgBox "�ɹ������ڹ�����[" & shtSalesCompInvUnified.Name & "] �У����飡", vbInformation

reset_excel_options:
    
    Err.Clear
    fEnableExcelOptionsAll
    End
End Sub

Private Function fLoadFilesAndRead2Variables()
    'gsCompanyID
    Call fLoadFileByFileTag("IMPORTED_DATA")
    Call fReadMasterSheetData("IMPORTED_DATA", , , True)
End Function

Private Function fProcessData()
    Call fRedimArrOutputBaseArrMaster
    
    Dim lEachRow As Long
    Dim sCompanyLongID As String
    Dim sCompanyName As String
    
    'Dim sHospital As String
    'Dim sReplacedHospital As String
    
    Dim sProducer As String
    Dim sReplacedProducer As String
    
    Dim sProductName As String
    Dim sReplacedProductName As String
    
    Dim sProductSeries As String
    Dim sReplacedProductSeries As String
    
    Dim sProductUnit As String
    Dim sReplacedProductUnit As String
    Dim sProductMasterUnit As String
    
    Dim dblRatio As Double
    
    'Dim dictNewHospital As Dictionary
    Dim dictNewProducer As Dictionary
    Dim dictNewProductName As Dictionary
    Dim dictNewProductSeries As Dictionary
    Dim dictNewProductUnit As Dictionary
    Dim dictNewProductUnitOrig As Dictionary
    
    'Set dictNewHospital = New Dictionary
    Set dictNewProducer = New Dictionary
    Set dictNewProductName = New Dictionary
    Set dictNewProductSeries = New Dictionary
    Set dictNewProductUnit = New Dictionary
    Set dictNewProductUnitOrig = New Dictionary

    Dim sTmpKey As String
    For lEachRow = LBound(arrMaster, 1) To UBound(arrMaster, 1)
        If dictMstColIndex.Exists("OrigInventoryID") Then
            arrOutput(lEachRow, dictRptColIndex("OrigInventoryID")) = arrMaster(lEachRow, dictMstColIndex("OrigInventoryID"))
        End If
        
        If dictMstColIndex.Exists("SeqNo") Then
            arrOutput(lEachRow, dictRptColIndex("SeqNo")) = arrMaster(lEachRow, dictMstColIndex("SeqNo"))
        End If
        
        arrOutput(lEachRow, dictRptColIndex("SalesCompanyName")) = arrMaster(lEachRow, dictMstColIndex("SalesCompanyName"))
      '  arrOutput(lEachRow, dictRptColIndex("InventoryDate")) = arrMaster(lEachRow, dictMstColIndex("InventoryDate"))
        arrOutput(lEachRow, dictRptColIndex("Quantity")) = arrMaster(lEachRow, dictMstColIndex("Quantity"))
        'arrOutput(lEachRow, dictRptColIndex("SellPrice")) = arrMaster(lEachRow, dictMstColIndex("SellPrice"))
        'arrOutput(lEachRow, dictRptColIndex("SellAmount")) = arrMaster(lEachRow, dictMstColIndex("SellAmount"))
        'arrOutput(lEachRow, dictRptColIndex("SellAmount")) = arrMaster(lEachRow, dictMstColIndex("SellPrice")) _
                                                            * arrMaster(lEachRow, dictMstColIndex("Quantity"))
        arrOutput(lEachRow, dictRptColIndex("LotNum")) = "'" & arrMaster(lEachRow, dictMstColIndex("LotNum"))
        'sHospital = Trim(arrMaster(lEachRow, dictMstColIndex("Hospital")))
        'arrOutput(lEachRow, dictRptColIndex("Hospital")) = sHospital
        sProducer = Trim(arrMaster(lEachRow, dictMstColIndex("ProductProducer")))
        arrOutput(lEachRow, dictRptColIndex("ProductProducer")) = sProducer
        sProductName = Trim(arrMaster(lEachRow, dictMstColIndex("ProductName")))
        arrOutput(lEachRow, dictRptColIndex("ProductName")) = sProductName
        sProductSeries = Trim(arrMaster(lEachRow, dictMstColIndex("ProductSeries")))
        arrOutput(lEachRow, dictRptColIndex("ProductSeries")) = sProductSeries
        sProductUnit = Trim(arrMaster(lEachRow, dictMstColIndex("ProductUnit")))
        arrOutput(lEachRow, dictRptColIndex("ProductUnit")) = sProductUnit
        
'        ' Hospital replace -----------------
'        If Not fReplaceAndValidateInHospitalMaster(sHospital, sReplacedHospital) Then
'            If Not dictNewHospital.Exists(sReplacedHospital) Then
'                dictNewHospital.Add sReplacedHospital, "'" & (lEachRow + 1)
'            Else
'                dictNewHospital(sReplacedHospital) = dictNewHospital(sReplacedHospital) & "," & (lEachRow + 1)
'            End If
'
'            alWarningCnt = alWarningCnt + 1
'            arrWarningRows(alWarningCnt) = lEachRow + 1
'            alWarningCnt = alWarningCnt + 1
'            arrWarningRows(alWarningCnt) = dictRptColIndex("Hospital")
'        End If
'        arrOutput(lEachRow, dictRptColIndex("MatchedHospital")) = sReplacedHospital
'        ' Hospital replace end -----------------
        
        ' Product producer replace -----------------
        If Not fReplaceAndValidateInProducerMaster(sProducer, sReplacedProducer) Then
            If Not dictNewProducer.Exists(sReplacedProducer) Then
                dictNewProducer.Add sReplacedProducer, "'" & (lEachRow + 1)
            Else
                dictNewProducer(sReplacedProducer) = dictNewProducer(sReplacedProducer) & "," & (lEachRow + 1)
            End If
            arrOutput(lEachRow, dictRptColIndex("MatchedProductProducer")) = ""
            
            alErrCnt = alErrCnt + 1
            arrErrRows(alErrCnt) = lEachRow + 1
            alErrCnt = alErrCnt + 1
            arrErrRows(alErrCnt) = dictRptColIndex("ProductProducer")
            GoTo next_sales
        Else
            arrOutput(lEachRow, dictRptColIndex("MatchedProductProducer")) = sReplacedProducer
        End If
        ' Product producer end -----------------
        
        ' Product Name replace -----------------
        If Not fReplaceAndValidateInProductNameMaster(sReplacedProducer, sProductName, sReplacedProductName) Then
            sTmpKey = sReplacedProducer & DELIMITER & sReplacedProductName
            If Not dictNewProductName.Exists(sTmpKey) Then
                dictNewProductName.Add sTmpKey, "'" & (lEachRow + 1)
            Else
                dictNewProductName(sTmpKey) = dictNewProductName(sTmpKey) & "," & lEachRow + 1
            End If
            arrOutput(lEachRow, dictRptColIndex("MatchedProductName")) = ""
            
            alErrCnt = alErrCnt + 1
            arrErrRows(alErrCnt) = lEachRow + 1
            alErrCnt = alErrCnt + 1
            arrErrRows(alErrCnt) = dictRptColIndex("ProductName")
            GoTo next_sales
        Else
            arrOutput(lEachRow, dictRptColIndex("MatchedProductName")) = sReplacedProductName
        End If
        ' Product Name end -----------------
        
        ' Product Series replace -----------------
        If Not fReplaceAndValidateInProductSeriesMaster(sReplacedProducer, sReplacedProductName, sProductSeries, sReplacedProductSeries) Then
            sTmpKey = sReplacedProducer & DELIMITER & sReplacedProductName & DELIMITER & sReplacedProductSeries
            If Not dictNewProductSeries.Exists(sTmpKey) Then
                dictNewProductSeries.Add sTmpKey, "'" & (lEachRow + 1)
            Else
                dictNewProductSeries(sTmpKey) = dictNewProductSeries(sTmpKey) & "," & (lEachRow + 1)
            End If
            arrOutput(lEachRow, dictRptColIndex("MatchedProductSeries")) = ""
            
            alErrCnt = alErrCnt + 1
            arrErrRows(alErrCnt) = lEachRow + 1
            alErrCnt = alErrCnt + 1
            arrErrRows(alErrCnt) = dictRptColIndex("ProductSeries")
            GoTo next_sales
        Else
            arrOutput(lEachRow, dictRptColIndex("MatchedProductSeries")) = sReplacedProductSeries
        End If
        ' Product Series end -----------------
        
        ' Product Unit ration -----------------
'        Call fGetConvertUnitAndUnitRatio

        sProductMasterUnit = fGetProductUnit(sReplacedProducer, sReplacedProductName, sReplacedProductSeries)

        dblRatio = 0
        If Len(Trim(sProductUnit)) <= 0 Then
            arrOutput(lEachRow, dictRptColIndex("MatchedProductUnit")) = sProductMasterUnit
            dblRatio = 1
        Else
'            Call fReplaceAndValidateInProductUnitMaster(sReplacedProducer, sReplacedProductName, sReplacedProductSeries, sProductUnit _
'                                                , sReplacedProductUnit, dblRatio)
            sReplacedProductUnit = fFindInConfigedReplaceProductUnit(sReplacedProducer, sReplacedProductName, sReplacedProductSeries _
                                                        , sProductUnit, dblRatio)
            If Len(Trim(sReplacedProductUnit)) <= 0 Then
                sReplacedProductUnit = sProductUnit
            Else
                If sReplacedProductUnit <> sProductMasterUnit Then
                    fErr "��" & shtProductUnitRatio.Name & "����Ƶ�λ��ҩƷ����һ�������顾" & shtProductUnitRatio.Name & "��" _
                        & vbCr & sReplacedProducer _
                        & vbCr & sReplacedProductName _
                        & vbCr & sReplacedProductSeries & vbCr & sProductMasterUnit
                End If

                If sProductUnit = sProductMasterUnit Then
                    If dblRatio <> 1 Then
                        fErr "ԭʼ�ļ���λ�ͻ�Ƶ�λһ�������Ǳ���ȴ����1�����顾" & shtProductUnitRatio.Name & "��" _
                            & vbCr & sReplacedProducer _
                            & vbCr & sReplacedProductName _
                            & vbCr & sReplacedProductSeries & vbCr & sProductMasterUnit
                    End If
                End If
            End If

            If sReplacedProductUnit <> sProductMasterUnit Then
                sTmpKey = sReplacedProducer & DELIMITER & sReplacedProductName & DELIMITER & sReplacedProductSeries & DELIMITER _
                                & sProductMasterUnit '& DELIMITER & sReplacedProductUnit
                If Not dictNewProductUnit.Exists(sTmpKey) Then
                    dictNewProductUnit.Add sTmpKey, "'" & (lEachRow + 1)
                    dictNewProductUnitOrig.Add sTmpKey, sReplacedProductUnit
                Else
                    dictNewProductUnit(sTmpKey) = dictNewProductUnit(sTmpKey) & "," & (lEachRow + 1)
                    
                    If InStr(dictNewProductUnitOrig(sTmpKey), sReplacedProductUnit) <= 0 Then _
                    dictNewProductUnitOrig(sTmpKey) = dictNewProductUnitOrig(sTmpKey) & "," & sReplacedProductUnit
                End If

                arrOutput(lEachRow, dictRptColIndex("MatchedProductUnit")) = ""
            
                alErrCnt = alErrCnt + 1
                arrErrRows(alErrCnt) = lEachRow + 1
                alErrCnt = alErrCnt + 1
                arrErrRows(alErrCnt) = dictRptColIndex("ProductUnit")
                GoTo next_sales
            End If

            arrOutput(lEachRow, dictRptColIndex("MatchedProductUnit")) = sReplacedProductUnit
        End If
        ' Product Unit ration -----------------
        
        If dblRatio <> 1 Then
            Debug.Print lEachRow
        End If
        
        arrOutput(lEachRow, dictRptColIndex("ConvertQuantity")) = arrMaster(lEachRow, dictMstColIndex("Quantity")) * dblRatio
       ' arrOutput(lEachRow, dictRptColIndex("ConvertSellPrice")) = arrMaster(lEachRow, dictMstColIndex("SellPrice")) / dblRatio
        'arrOutput(lEachRow, dictRptColIndex("RecalSellAmount")) = arrOutput(lEachRow, dictRptColIndex("ConvertQuantity")) _
                                                                * arrOutput(lEachRow, dictRptColIndex("ConvertSellPrice"))
next_sales:
    Next
    
    'Call fAddNewFoundMissedHospitalToSheet(dictNewHospital, False)
    'Call fAddNewFoundHospitalToSheetException(dictNewHospital, True)
    Call fAddNewFoundMissedProducerToSheetException(dictNewProducer)
    Call fAddNewFoundMissedProductNameToSheetException(dictNewProductName)
    Call fAddNewFoundMissedProductSeriesToSheetException(dictNewProductSeries)
    Call fAddNewFoundMissedProductUnitToSheetException(dictNewProductUnit, dictNewProductUnitOrig)
End Function

Function fAddNewFoundMissedHospitalToSheet(dictNewHospital As Dictionary, Optional bSetDictToNothing As Boolean = True)
    Dim arrNewHospital()
    arrNewHospital = fConvertDictionaryKeysTo2DimenArrayForPaste(dictNewHospital, bSetDictToNothing)
    Call fAppendArray2Sheet(shtHospital, arrNewHospital)
    
    If fUbound(arrNewHospital, 1) > 0 Then
        fMsgBox fUbound(arrNewHospital, 1) & "��ҽԺ�Ҳ�����" & vbCr & "���Ǳ��Զ����뵽�˱�" & shtHospital.Name & "������." _
                & vbCr & "�ñ������������Ϊ�����¼ӵġ�" & vbCr _
                & ""
    End If
    Erase arrNewHospital
End Function
Function fAddNewFoundHospitalToSheetException(ByRef dictNewHospital As Dictionary, Optional bSetDictToNothing As Boolean = True)
    Dim arrNewHospital()
    'Dim sErr As String
    Dim lStartRow As Long
    Dim lRecCount As Long
    
    lRecCount = fGetDictionayDelimiteredItemsCount(dictNewHospital)
        
    If lRecCount > 0 Then
'        shtException.Cells.NumberFormat = "@"
'        shtException.Cells.WrapText = True
        
        lStartRow = fGetshtExceptionNewRow
        arrNewHospital = fConvertDictionaryKeysTo2DimenArrayForPaste(dictNewHospital, False)
'        lStartRow = fGetValidMaxRow(shtException)
'        If lStartRow = 0 Then
'            lStartRow = lStartRow + 1
'        Else
'            lStartRow = lStartRow + 5
'        End If
        
        Call fPrepareHeaderToSheet(shtException, Array("��ϵͳ���Ҳ�����ҽԺ", "�к�"), lStartRow)
        
        shtException.Rows(lStartRow).Font.Color = RGB(255, 0, 0)
        shtException.Rows(lStartRow).Font.Bold = True
        
        Call fAppendArray2Sheet(shtException, arrNewHospital)
        shtException.Cells(lStartRow + 1, 2).Resize(dictNewHospital.Count, 1).Value = fConvertDictionaryItemsTo2DimenArrayForPaste(dictNewHospital, False)

        Call fFreezeSheet(shtException)
        
'        shtException.Visible = xlSheetVisible
'        shtException.Activate
        
        If fNzero(gsBusinessErrorMsg) Then gsBusinessErrorMsg = gsBusinessErrorMsg & vbCr & vbCr & vbCr & "===============================" & vbCr & vbCr
        
        gsBusinessErrorMsg = gsBusinessErrorMsg & lRecCount & "����ҽԺ���ڱ�ϵͳ���Ҳ�����" & vbCr _
            & "�������Ƕ����Զ����뵽�ˡ�ҽԺ��������" & vbCr _
            & "����Ҫ�ڡ�ҽԺ�滻�����������滻�����������У�����ҽԺ�ͻ�����ͳһ��" & vbCr & vbCr _
            & "����ҽԺ�����еĲ��õļ�¼ɾ����"
    End If
End Function

Function fAddNewFoundMissedProducerToSheetException(dictNewProducer As Dictionary)
    '======= Producer Validation ===============================================
    Dim arrNewProducer()
    'Dim sErr As String
    Dim lStartRow As Long
    Dim lRecCount As Long
    
    lRecCount = fGetDictionayDelimiteredItemsCount(dictNewProducer)
    If lRecCount > 0 Then
'        shtException.Cells.NumberFormat = "@"
'        shtException.Cells.WrapText = True
        
        lStartRow = fGetshtExceptionNewRow
        arrNewProducer = fConvertDictionaryKeysTo2DimenArrayForPaste(dictNewProducer, False)
'        lStartRow = fGetValidMaxRow(shtException)
'        If lStartRow = 0 Then
'            lStartRow = lStartRow + 1
'        Else
'            lStartRow = lStartRow + 5
'        End If
        
        Call fPrepareHeaderToSheet(shtException, Array("��ϵͳ���Ҳ�����ҩƷ��������", "�к�"), lStartRow)
        
        shtException.Rows(lStartRow).Font.Color = RGB(255, 0, 0)
        shtException.Rows(lStartRow).Font.Bold = True
        
        Call fAppendArray2Sheet(shtException, arrNewProducer)
        'sErr = fUbound(arrNewProducer)
        shtException.Cells(lStartRow + 1, 2).Resize(dictNewProducer.Count, 1).Value = fConvertDictionaryItemsTo2DimenArrayForPaste(dictNewProducer, False)
       ' Erase arrNewProducer
        Call fFreezeSheet(shtException)
        
        shtException.Visible = xlSheetVisible
        shtException.Activate
        
        If fNzero(gsBusinessErrorMsg) Then gsBusinessErrorMsg = gsBusinessErrorMsg & vbCr & vbCr & vbCr & "===============================" & vbCr & vbCr
        
        gsBusinessErrorMsg = gsBusinessErrorMsg & lRecCount & "��ҩƷ���������ҡ��ڱ�ϵͳ���Ҳ�����������Ҫ��" & vbCr _
            & "(1). �ڡ�ҩƷ�����滻�������һ���滻��¼" & vbCr _
            & "(2). �ڡ�ҩƷ��������������һ������" & vbCr & vbCr _
            & "���ε���ʧ�ܣ��������ݺ����ٴε����ť���С�ƥ���滻ͳһ��"
    End If
    '======= Producer end ===============================================
End Function

Private Function fAddNewFoundMissedProductNameToSheetException(dictNewProductName As Dictionary)
    '======= ProductName Validation ===============================================
    Dim arrNewProductName()
    'Dim sErr As String
    Dim lStartRow As Long
    Dim lRecCount As Long
    
    lRecCount = fGetDictionayDelimiteredItemsCount(dictNewProductName)
        
    If lRecCount > 0 Then
'        shtException.Cells.NumberFormat = "@"
'        shtException.Cells.WrapText = True
        lStartRow = fGetshtExceptionNewRow
        
        arrNewProductName = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictNewProductName, , False)
'        lStartRow = fGetValidMaxRow(shtException)
'        If lStartRow = 0 Then
'            lStartRow = lStartRow + 1
'        Else
'            lStartRow = lStartRow + 5
'        End If
        
        Call fPrepareHeaderToSheet(shtException, Array("ҩƷ����", "��ϵͳ���Ҳ�����ҩƷ����", "�к�"), lStartRow)
        shtException.Rows(lStartRow).Font.Color = RGB(255, 0, 0)
        shtException.Rows(lStartRow).Font.Bold = True
        Call fAppendArray2Sheet(shtException, arrNewProductName)
        'sErr = fUbound(arrNewProductName)
        shtException.Cells(lStartRow + 1, 3).Resize(dictNewProductName.Count, 1).Value = fConvertDictionaryItemsTo2DimenArrayForPaste(dictNewProductName, False)
        'Erase arrNewProductName
        Call fFreezeSheet(shtException)
        
'        shtException.Visible = xlSheetVisible
'        shtException.Activate
        
        If fNzero(gsBusinessErrorMsg) Then gsBusinessErrorMsg = gsBusinessErrorMsg & vbCr & vbCr & vbCr & "===============================" & vbCr & vbCr
        
        gsBusinessErrorMsg = gsBusinessErrorMsg & lRecCount & "��ҩƷ�����ơ��ڱ�ϵͳ���Ҳ�����������Ҫ��" & vbCr _
            & "(1). �ڡ�ҩƷ�����滻�������һ���滻��¼" & vbCr _
            & "(2). �ڡ�ҩƷ��������������һ������" & vbCr _
            & "** ��ע�⣺ҩƷ����û�����⣬��ƥ�䵽�ˡ�" & vbCr & vbCr _
            & "���ε���ʧ�ܣ��������ݺ����ٴε����ť���С�ƥ���滻ͳһ��"
    End If
    '======= ProductName end ===============================================
End Function

Function fAddNewFoundMissedProductSeriesToSheetException(dictNewProductSeries As Dictionary)
    '======= ProductSeries Validation ===============================================
    Dim arrNewProductSeries()
    'Dim sErr As String
    Dim lRecCount As Long
        Dim lStartRow As Long
    
    lRecCount = fGetDictionayDelimiteredItemsCount(dictNewProductSeries)
    If lRecCount > 0 Then
        lStartRow = fGetshtExceptionNewRow
        
        arrNewProductSeries = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictNewProductSeries, , False)
        
        Call fPrepareHeaderToSheet(shtException, Array("ҩƷ����", "ҩƷ����", "��ϵͳ���Ҳ�����ҩƷ�����", "�к�"), lStartRow)
        shtException.Rows(lStartRow).Font.Color = RGB(255, 0, 0)
        shtException.Rows(lStartRow).Font.Bold = True
        Call fAppendArray2Sheet(shtException, arrNewProductSeries)
        'sErr = fUbound(arrNewProductSeries)
        
        shtException.Cells(lStartRow + 1, 4).Resize(dictNewProductSeries.Count, 1).Value = fConvertDictionaryItemsTo2DimenArrayForPaste(dictNewProductSeries, False)
       ' Erase arrNewProductSeries
        Call fFreezeSheet(shtException)
        
'        shtException.Visible = xlSheetVisible
'        shtException.Activate
        
        If fNzero(gsBusinessErrorMsg) Then gsBusinessErrorMsg = gsBusinessErrorMsg & vbCr & vbCr & vbCr & "===============================" & vbCr & vbCr

        gsBusinessErrorMsg = gsBusinessErrorMsg & lRecCount & "��ҩƷ������ڱ�ϵͳ���Ҳ�����������Ҫ��" & vbCr _
            & "(1). �ڡ�ҩƷ����滻�������һ���滻��¼" & vbCr _
            & "(2). �ڡ�ҩƷ����������һ�����" & vbCr _
            & "** ��ע�⣺ҩƷ���Һ�ҩƷ����û�����⣬��ƥ�䵽�ˡ�" & vbCr & vbCr _
            & "���ε���ʧ�ܣ��������ݺ����ٴε����ť���С�ƥ���滻ͳһ��"
    End If
End Function

Function fAddNewFoundMissedProductUnitToSheetException(dictNewProductUnit As Dictionary, dictNewProductUnitOrig As Dictionary)
    Dim arrNewProductUnit()
    'Dim sErr As String
    Dim lRecCount As Long
        Dim lStartRow As Long
    
    lRecCount = fGetDictionayDelimiteredItemsCount(dictNewProductUnit)
    
    If lRecCount > 0 Then
'        shtException.Cells.NumberFormat = "@"
'        shtException.Cells.WrapText = True
        lStartRow = fGetshtExceptionNewRow
        
        arrNewProductUnit = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictNewProductUnit, , False)
'        lStartRow = fGetValidMaxRow(shtException)
'        If lStartRow = 0 Then
'            lStartRow = lStartRow + 1
'        Else
'            lStartRow = lStartRow + 5
'        End If
        
        Call fPrepareHeaderToSheet(shtException, Array("ҩƷ����", "ҩƷ����", "ҩƷ���", "ҩƷ��Ƶ�λ", "ԭʼ�ļ���λ", "�к�"), lStartRow)
        shtException.Rows(lStartRow).Font.Color = RGB(255, 0, 0)
        shtException.Rows(lStartRow).Font.Bold = True
        Call fAppendArray2Sheet(shtException, arrNewProductUnit)
        'sErr = fUbound(arrNewProductUnit)
            
        shtException.Cells(lStartRow + 1, 5).Resize(dictNewProductUnitOrig.Count, 1).Value = fConvertDictionaryItemsTo2DimenArrayForPaste(dictNewProductUnitOrig, False)
        shtException.Cells(lStartRow + 1, 6).Resize(dictNewProductUnit.Count, 1).Value = fConvertDictionaryItemsTo2DimenArrayForPaste(dictNewProductUnit, False)
       ' Erase arrNewProductUnit
        Call fFreezeSheet(shtException)
        
'        shtException.Visible = xlSheetVisible
'        shtException.Activate
        
        If fNzero(gsBusinessErrorMsg) Then gsBusinessErrorMsg = gsBusinessErrorMsg & vbCr & vbCr & vbCr & "===============================" & vbCr & vbCr

        gsBusinessErrorMsg = gsBusinessErrorMsg & lRecCount & "��ҩƷ����λ�����趨�Ļ�Ƶ�λ��һ�£�������Ҫ��" & vbCr _
            & "(1). �ڡ�ҩƷ��λ�����������һ����¼" & vbCr _
            & "(2). �ڡ�ҩƷ�������޸��䵥λ" & vbCr _
            & "** ��ע�⣺ҩƷ���ҡ����ơ����û�����⣬��ƥ�䵽�ˡ�" & vbCr & vbCr _
            & "���ε���ʧ�ܣ��������ݺ����ٴε����ť���С�ƥ���滻ͳһ��"
    End If
    '======= ProductUnit end ===============================================
End Function
Function fReplaceAndValidateInHospitalMaster(sHospital As String, ByRef sReplacedHospital As String) As Boolean
    sReplacedHospital = fFindInConfigedReplaceHospital(sHospital)
    If fZero(sReplacedHospital) Then sReplacedHospital = sHospital
    
    fReplaceAndValidateInHospitalMaster = fHospitalExistsInHospitalMaster(sReplacedHospital)
End Function

Function fReplaceAndValidateInProducerMaster(sProducer As String, ByRef sReplacedProducer As String) As Boolean
    sReplacedProducer = fFindInConfigedReplaceProducer(sProducer)
    If fZero(sReplacedProducer) Then sReplacedProducer = sProducer
    
    fReplaceAndValidateInProducerMaster = fProducerExistsInProducerMaster(sReplacedProducer)
End Function

Function fReplaceAndValidateInProductNameMaster(sReplacedProducer As String, sProductName As String, ByRef sReplacedProductName As String) As Boolean
    sReplacedProductName = fFindInConfigedReplaceProductName(sReplacedProducer, sProductName)
    If fZero(sReplacedProductName) Then sReplacedProductName = sProductName
    
    fReplaceAndValidateInProductNameMaster = fProductNameExistsInProductNameMaster(sReplacedProducer, sReplacedProductName)
End Function
'ProductSeriesMaster = ProductMaster
Function fReplaceAndValidateInProductSeriesMaster(sReplacedProducer As String, sReplacedProductName As String _
                                        , sProductSeries As String, ByRef sReplacedProductSeries As String) As Boolean
    sReplacedProductSeries = fFindInConfigedReplaceProductSeries(sReplacedProducer, sReplacedProductName, sProductSeries)
    If fZero(sReplacedProductSeries) Then sReplacedProductSeries = sProductSeries
    
    fReplaceAndValidateInProductSeriesMaster = fProductSeriesExistsInProductMaster(sReplacedProducer, sReplacedProductName, sReplacedProductSeries)
End Function

''ProductUnitMaster = ProductMaster
'Function fReplaceAndValidateInProductUnitMaster(sReplacedProducer As String, sReplacedProductName As String, sReplacedProductSeries As String _
'                                        , sProductUnit As String _
'                                        , ByRef sReplacedProductUnit As String _
'                                        , ByRef dblRatio As Double) As Boolean
'    sReplacedProductUnit = fFindInConfigedReplaceProductUnit(sReplacedProducer, sReplacedProductName, sReplacedProductSeries _
'                                                        , sProductUnit, dblRatio)
'  '  sProductMasterUnit = fGetProductUnit(sReplacedProducer, sReplacedProductName, sReplacedProductSeries)
'
''    If Len(Trim(sReplacedProductUnit)) <= 0 Then
''        sReplacedProductUnit = sProductUnit
''    Else
''        If sReplacedProductUnit <> sProductMasterUnit Then
''            fErr "��" & shtProductUnitRatio.Name & "����Ƶ�λ��ҩƷ����һ�������顾" & shtProductUnitRatio.Name & "��" _
''                & vbCr & sReplacedProducer _
''                & vbCr & sReplacedProductName _
''                & vbCr & sReplacedProductSeries & vbCr & sProductMasterUnit
''        End If
''
''        If sProductUnit = sProductMasterUnit Then
''            If dblRatio = 1 Then
''                fErr "ԭʼ�ļ���λ�ͻ�Ƶ�λһ�������Ǳ���ȴ����1�����顾" & shtProductUnitRatio.Name & "��" _
''                    & vbCr & sReplacedProducer _
''                    & vbCr & sReplacedProductName _
''                    & vbCr & sReplacedProductSeries & vbCr & sProductMasterUnit
''            End If
''        End If
'    End If
'
'    'fReplaceAndValidateInProductUnitMaster = (sReplacedProductUnit = sProductMasterUnit)
'End Function


'Function fGetConvertUnitAndUnitRatio() As Boolean
'    Dim bOut As Boolean
'
'        sProductMasterUnit = fGetProductUnit(sReplacedProducer, sReplacedProductName, sReplacedProductSeries)
'
'        dblRatio = 0
'        If Len(Trim(sProductUnit)) <= 0 Then
'            arrOutput(lEachRow, dictRptColIndex("MatchedProductUnit")) = sProductMasterUnit
'            dblRatio = 1
'        Else
''            Call fReplaceAndValidateInProductUnitMaster(sReplacedProducer, sReplacedProductName, sReplacedProductSeries, sProductUnit _
''                                                , sReplacedProductUnit, dblRatio)
'
'            sReplacedProductUnit = fFindInConfigedReplaceProductUnit(sReplacedProducer, sReplacedProductName, sReplacedProductSeries _
'                                                        , sProductUnit, dblRatio)
'            If Len(Trim(sReplacedProductUnit)) <= 0 Then
'                sReplacedProductUnit = sProductUnit
'            Else
'                If sReplacedProductUnit <> sProductMasterUnit Then
'                    fErr "��" & shtProductUnitRatio.Name & "����Ƶ�λ��ҩƷ����һ�������顾" & shtProductUnitRatio.Name & "��" _
'                        & vbCr & sReplacedProducer _
'                        & vbCr & sReplacedProductName _
'                        & vbCr & sReplacedProductSeries & vbCr & sProductMasterUnit
'                End If
'
'                If sProductUnit = sProductMasterUnit Then
'                    If dblRatio = 1 Then
'                        fErr "ԭʼ�ļ���λ�ͻ�Ƶ�λһ�������Ǳ���ȴ����1�����顾" & shtProductUnitRatio.Name & "��" _
'                            & vbCr & sReplacedProducer _
'                            & vbCr & sReplacedProductName _
'                            & vbCr & sReplacedProductSeries & vbCr & sProductMasterUnit
'                    End If
'                End If
'            End If
'
'            If sReplacedProductUnit <> sProductMasterUnit Then
'                sTmpKey = sReplacedProducer & DELIMITER & sReplacedProductName & DELIMITER & sReplacedProductSeries & DELIMITER _
'                                & sProductMasterUnit & DELIMITER & sReplacedProductUnit
'                If Not dictNewProductUnit.Exists(sTmpKey) Then
'                    dictNewProductUnit.Add sTmpKey, lEachRow + 1
'                End If
'
'                arrOutput(lEachRow, dictRptColIndex("MatchedProductUnit")) = ""
'                GoTo next_sales
'            End If
'
'            arrOutput(lEachRow, dictRptColIndex("MatchedProductUnit")) = sReplacedProductUnit
'        End If
'End Function

