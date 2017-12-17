Attribute VB_Name = "MB2_ReplaceSalesInfos"
Option Explicit
Option Base 1

Sub subMain_ReplaceSalesInfos()
    'If Not fIsDev Then On Error GoTo error_handling
    'On Error GoTo error_handling
    shtSalesRawDataRpt.Visible = xlSheetVisible
    shtException.Visible = xlSheetVeryHidden
    Call fUnProtectSheet(shtSalesInfos)
    Call fCleanSheetOutputResetSheetOutput(shtSalesInfos)
    Call fCleanSheetOutputResetSheetOutput(shtException)

    fInitialization

    gsRptID = "REPLACE_UNIFY_SALES_INFO"

    Call fReadSysConfig_InputTxtSheetFile

    gsRptFilePath = fReadSysConfig_Output(, gsRptType)
    Call fLoadFilesAndRead2Variables

    Call fPrepareOutputSheetHeaderAndTextColumns(shtSalesInfos)

    Call fProcessData
    
    If Not shtException.Visible = xlSheetVisible Then shtException.Visible = xlSheetVeryHidden
    
error_handling:
    'If shtException.Visible = xlSheetVisible Then
        Call fAppendArray2Sheet(shtSalesInfos, arrOutput)
    
    
        'Call fReSequenceSeqNo
    
    '    Call fSortDataInSheetSortSheetData(shtSalesRawDataRpt, Array(dictRptColIndex("SalesCompanyName") _
                                                                    , dictRptColIndex("Hospital") _
                                                                    , dictRptColIndex("SalesDate") _
                                                                    , dictRptColIndex("ProductProducer") _
                                                                    , dictRptColIndex("ProductName") _
                                                                    , dictRptColIndex("ProductUnit")))
        Call fFormatOutputSheet(shtSalesInfos)
    
       ' Call fProtectSheetAndAllowEdit(shtSalesRawDataRpt, shtSalesRawDataRpt.Columns(4), UBound(arrOutput, 1) + 1, UBound(arrOutput, 2), False)
        Call fPostProcess(shtSalesInfos)
    
        shtSalesInfos.Visible = xlSheetVisible
        shtSalesInfos.Activate
        shtSalesInfos.Range("A1").Select
        
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

    Call fSetReneratedReport(, shtSalesInfos.Name)
    fMsgBox "�ɹ������ڹ�����[" & shtSalesInfos.Name & "] �У����飡", vbInformation

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
    
    Dim sHospital As String
    Dim sReplacedHospital As String
    
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
    
    Dim dictNewHospital As Dictionary
    Dim dictNewProducer As Dictionary
    Dim dictNewProductName As Dictionary
    Dim dictNewProductSeries As Dictionary
    Dim dictNewProductUnit As Dictionary
    
    Set dictNewHospital = New Dictionary
    Set dictNewProducer = New Dictionary
    Set dictNewProductName = New Dictionary
    Set dictNewProductSeries = New Dictionary
    Set dictNewProductUnit = New Dictionary

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
        
        
        arrOutput(lEachRow, dictRptColIndex("Quantity")) = arrMaster(lEachRow, dictMstColIndex("Quantity"))
        arrOutput(lEachRow, dictRptColIndex("SellPrice")) = arrMaster(lEachRow, dictMstColIndex("SellPrice"))
        'arrOutput(lEachRow, dictRptColIndex("SellAmount")) = arrMaster(lEachRow, dictMstColIndex("SellAmount"))
        arrOutput(lEachRow, dictRptColIndex("SellAmount")) = arrMaster(lEachRow, dictMstColIndex("SellPrice")) _
                                                            * arrMaster(lEachRow, dictMstColIndex("Quantity"))
        
        ' Hospital replace -----------------
        sHospital = arrMaster(lEachRow, dictMstColIndex("Hospital"))
        arrOutput(lEachRow, dictRptColIndex("Hospital")) = sHospital
        
        If Not fReplaceAndValidateInHospitalMaster(sHospital, sReplacedHospital) Then
            If Not dictNewHospital.Exists(sReplacedHospital) Then dictNewHospital.Add sReplacedHospital, 0
        End If
        arrOutput(lEachRow, dictRptColIndex("MatchedHospital")) = sReplacedHospital
        ' Hospital replace end -----------------
        
        ' Product producer replace -----------------
        sProducer = arrMaster(lEachRow, dictMstColIndex("ProductProducer"))
        arrOutput(lEachRow, dictRptColIndex("ProductProducer")) = sProducer
        
        If Not fReplaceAndValidateInProducerMaster(sProducer, sReplacedProducer) Then
            If Not dictNewProducer.Exists(sReplacedProducer) Then dictNewProducer.Add sReplacedProducer, lEachRow + 1
            arrOutput(lEachRow, dictRptColIndex("MatchedProductProducer")) = ""
            GoTo next_sales
        Else
            arrOutput(lEachRow, dictRptColIndex("MatchedProductProducer")) = sReplacedProducer
        End If
        
        'arrOutput(lEachRow, dictRptColIndex("MatchedProductProducer")) = sReplacedProducer
        ' Product producer end -----------------
        
        ' Product Name replace -----------------
        sProductName = arrMaster(lEachRow, dictMstColIndex("ProductName"))
        arrOutput(lEachRow, dictRptColIndex("ProductName")) = sProductName
        
        If Not fReplaceAndValidateInProductNameMaster(sReplacedProducer, sProductName, sReplacedProductName) Then
            If Not dictNewProductName.Exists(sReplacedProducer & DELIMITER & sReplacedProductName) Then
                dictNewProductName.Add sReplacedProducer & DELIMITER & sReplacedProductName, lEachRow + 1
            End If
            arrOutput(lEachRow, dictRptColIndex("MatchedProductName")) = ""
            GoTo next_sales
        Else
            arrOutput(lEachRow, dictRptColIndex("MatchedProductName")) = sReplacedProductName
        End If
        ' Product Name end -----------------
        
        ' Product Series replace -----------------
        sProductSeries = arrMaster(lEachRow, dictMstColIndex("ProductSeries"))
        arrOutput(lEachRow, dictRptColIndex("ProductSeries")) = sProductSeries
        
        If Not fReplaceAndValidateInProductSeriesMaster(sReplacedProducer, sReplacedProductName, sProductSeries, sReplacedProductSeries) Then
            If Not dictNewProductSeries.Exists(sReplacedProducer & DELIMITER & sReplacedProductName & DELIMITER & sReplacedProductSeries) Then
                dictNewProductSeries.Add sReplacedProducer & DELIMITER & sReplacedProductName & DELIMITER & sReplacedProductSeries, lEachRow + 1
            End If
            arrOutput(lEachRow, dictRptColIndex("MatchedProductSeries")) = ""
            GoTo next_sales
        Else
            arrOutput(lEachRow, dictRptColIndex("MatchedProductSeries")) = sReplacedProductSeries
        End If
        ' Product Series end -----------------
        
        ' Product Unit ration -----------------
        sProductUnit = arrMaster(lEachRow, dictMstColIndex("ProductUnit"))
        arrOutput(lEachRow, dictRptColIndex("ProductUnit")) = sProductUnit
        
'        Call fGetConvertUnitAndUnitRatio

        sProductMasterUnit = fGetProductMasterUnit(sReplacedProducer, sReplacedProductName, sReplacedProductSeries)

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
                    If dblRatio = 1 Then
                        fErr "ԭʼ�ļ���λ�ͻ�Ƶ�λһ�������Ǳ���ȴ����1�����顾" & shtProductUnitRatio.Name & "��" _
                            & vbCr & sReplacedProducer _
                            & vbCr & sReplacedProductName _
                            & vbCr & sReplacedProductSeries & vbCr & sProductMasterUnit
                    End If
                End If
            End If

            If sReplacedProductUnit <> sProductMasterUnit Then
                sTmpKey = sReplacedProducer & DELIMITER & sReplacedProductName & DELIMITER & sReplacedProductSeries & DELIMITER _
                                & sProductMasterUnit & DELIMITER & sReplacedProductUnit
                If Not dictNewProductUnit.Exists(sTmpKey) Then
                    dictNewProductUnit.Add sTmpKey, lEachRow + 1
                End If

                arrOutput(lEachRow, dictRptColIndex("MatchedProductUnit")) = ""
                GoTo next_sales
            End If

            arrOutput(lEachRow, dictRptColIndex("MatchedProductUnit")) = sReplacedProductUnit
        End If
        ' Product Unit ration -----------------
        
        If dblRatio <> 1 Then
            Debug.Print lEachRow
        End If
        
        arrOutput(lEachRow, dictRptColIndex("ConvertQuantity")) = arrMaster(lEachRow, dictMstColIndex("Quantity")) * dblRatio
        arrOutput(lEachRow, dictRptColIndex("ConvertSellPrice")) = arrMaster(lEachRow, dictMstColIndex("SellPrice")) / dblRatio
        arrOutput(lEachRow, dictRptColIndex("RecalSellAmount")) = arrOutput(lEachRow, dictRptColIndex("ConvertQuantity")) _
                                                                * arrOutput(lEachRow, dictRptColIndex("ConvertSellPrice"))
        
next_sales:
    Next
    
    Call fAddNewFoundMissedHospitalToSheet(dictNewHospital)
    Call fAddNewFoundMissedProducerToSheetException(dictNewProducer)
    Call fAddNewFoundMissedProductNameToSheetException(dictNewProductName)
    Call fAddNewFoundMissedProductSeriesToSheetException(dictNewProductSeries)
    Call fAddNewFoundMissedProductUnitToSheetException(dictNewProductUnit)
End Function

Function fAddNewFoundMissedHospitalToSheet(dictNewHospital As Dictionary)
    '======= Hospitals paste to master sheet  ===============================================
    Dim arrNewHospital()
    arrNewHospital = fConvertDictionaryKeysTo2DimenArrayForPaste(dictNewHospital)
    Call fAppendArray2Sheet(shtHospital, arrNewHospital)
    
    If fUbound(arrNewHospital, 1) > 0 Then
        fMsgBox fUbound(arrNewHospital, 1) & "��ҽԺ�Ҳ�����" & vbCr & "���Ǳ��Զ����뵽�˱�" & shtHospital.Name & "������." _
                & vbCr & "�ñ������������Ϊ�����¼ӵġ�" & vbCr _
                & ""
    End If
    Erase arrNewHospital
    '======= Hospitals paste to master sheet   end ===============================================
End Function

Function fAddNewFoundMissedProducerToSheetException(dictNewProducer As Dictionary)
    '======= Producer Validation ===============================================
    Dim arrNewProducer()
    Dim sErr As String
    If dictNewProducer.Count > 0 Then
        arrNewProducer = fConvertDictionaryKeysTo2DimenArrayForPaste(dictNewProducer, False)
        Call fPrepareHeaderToSheet(shtException, Array("��ϵͳ���Ҳ�����ҩƷ��������", "�к�"))
        
        shtException.Rows(1).Font.Color = RGB(255, 0, 0)
        shtException.Rows(1).Font.Bold = True
        
        Call fAppendArray2Sheet(shtException, arrNewProducer)
        sErr = fUbound(arrNewProducer)
        shtException.Cells(2, 2).Resize(sErr, 1).Value = fConvertDictionaryItemsTo2DimenArrayForPaste(dictNewProducer)
        Erase arrNewProducer
        Call fFreezeSheet(shtException)
        
        shtException.Visible = xlSheetVisible
        shtException.Activate
        
        If fNzero(gsBusinessErrorMsg) Then gsBusinessErrorMsg = gsBusinessErrorMsg & vbCr & vbCr & vbCr & "===============================" & vbCr & vbCr
        
        gsBusinessErrorMsg = gsBusinessErrorMsg & sErr & "��ҩƷ���������ҡ��ڱ�ϵͳ���Ҳ�����������Ҫ��" & vbCr _
            & "(1). �ڡ�ҩƷ�����滻�������һ���滻��¼" & vbCr _
            & "(2). �ڡ�ҩƷ��������������һ������" & vbCr & vbCr _
            & "���ε���ʧ�ܣ��������ݺ����ٴε����ť���С�ƥ���滻ͳһ��"
    End If
    '======= Producer end ===============================================
End Function

Function fAddNewFoundMissedProductNameToSheetException(dictNewProductName As Dictionary)
    '======= ProductName Validation ===============================================
    Dim arrNewProductName()
    Dim sErr As String
    If dictNewProductName.Count > 0 Then
        arrNewProductName = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictNewProductName, , False)
        Dim lStartRow As Long
        lStartRow = fGetValidMaxRow(shtException)
        If lStartRow = 0 Then
            lStartRow = lStartRow + 1
        Else
            lStartRow = lStartRow + 5
        End If
        
        Call fPrepareHeaderToSheet(shtException, Array("ҩƷ����", "��ϵͳ���Ҳ�����ҩƷ����", "�к�"), lStartRow)
        shtException.Rows(lStartRow).Font.Color = RGB(255, 0, 0)
        shtException.Rows(lStartRow).Font.Bold = True
        Call fAppendArray2Sheet(shtException, arrNewProductName)
        sErr = fUbound(arrNewProductName)
        
        shtException.Cells(lStartRow + 1, 3).Resize(sErr, 1).Value = fConvertDictionaryItemsTo2DimenArrayForPaste(dictNewProductName)
        Erase arrNewProductName
        Call fFreezeSheet(shtException)
        
        shtException.Visible = xlSheetVisible
        shtException.Activate
        
        If fNzero(gsBusinessErrorMsg) Then gsBusinessErrorMsg = gsBusinessErrorMsg & vbCr & vbCr & vbCr & "===============================" & vbCr & vbCr
        
        gsBusinessErrorMsg = gsBusinessErrorMsg & sErr & "��ҩƷ�����ơ��ڱ�ϵͳ���Ҳ�����������Ҫ��" & vbCr _
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
    Dim sErr As String
    If dictNewProductSeries.Count > 0 Then
        arrNewProductSeries = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictNewProductSeries, , False)
        Dim lStartRow As Long
        lStartRow = fGetValidMaxRow(shtException)
        If lStartRow = 0 Then
            lStartRow = lStartRow + 1
        Else
            lStartRow = lStartRow + 5
        End If
        
        Call fPrepareHeaderToSheet(shtException, Array("ҩƷ����", "ҩƷ����", "��ϵͳ���Ҳ�����ҩƷ�����", "�к�"), lStartRow)
        shtException.Rows(lStartRow).Font.Color = RGB(255, 0, 0)
        shtException.Rows(lStartRow).Font.Bold = True
        Call fAppendArray2Sheet(shtException, arrNewProductSeries)
        sErr = fUbound(arrNewProductSeries)
        
        shtException.Cells(lStartRow + 1, 4).Resize(sErr, 1).Value = fConvertDictionaryItemsTo2DimenArrayForPaste(dictNewProductSeries)
        Erase arrNewProductSeries
        Call fFreezeSheet(shtException)
        
        shtException.Visible = xlSheetVisible
        shtException.Activate
        
        If fNzero(gsBusinessErrorMsg) Then gsBusinessErrorMsg = gsBusinessErrorMsg & vbCr & vbCr & vbCr & "===============================" & vbCr & vbCr

        gsBusinessErrorMsg = gsBusinessErrorMsg & sErr & "��ҩƷ������ڱ�ϵͳ���Ҳ�����������Ҫ��" & vbCr _
            & "(1). �ڡ�ҩƷ����滻�������һ���滻��¼" & vbCr _
            & "(2). �ڡ�ҩƷ����������һ�����" & vbCr _
            & "** ��ע�⣺ҩƷ���Һ�ҩƷ����û�����⣬��ƥ�䵽�ˡ�" & vbCr & vbCr _
            & "���ε���ʧ�ܣ��������ݺ����ٴε����ť���С�ƥ���滻ͳһ��"
    End If
    '======= ProductSeries end ===============================================
End Function

Function fAddNewFoundMissedProductUnitToSheetException(dictNewProductUnit As Dictionary)
    '======= ProductUnit Validation ===============================================
    Dim arrNewProductUnit()
    Dim sErr As String
    If dictNewProductUnit.Count > 0 Then
        arrNewProductUnit = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictNewProductUnit, , False)
        Dim lStartRow As Long
        lStartRow = fGetValidMaxRow(shtException)
        If lStartRow = 0 Then
            lStartRow = lStartRow + 1
        Else
            lStartRow = lStartRow + 5
        End If
        
        Call fPrepareHeaderToSheet(shtException, Array("ҩƷ����", "ҩƷ����", "ҩƷ���", "ҩƷ��Ƶ�λ", "ԭʼ�ļ���λ", "�к�"), lStartRow)
        shtException.Rows(lStartRow).Font.Color = RGB(255, 0, 0)
        shtException.Rows(lStartRow).Font.Bold = True
        Call fAppendArray2Sheet(shtException, arrNewProductUnit)
        sErr = fUbound(arrNewProductUnit)
        
        shtException.Cells(lStartRow + 1, 6).Resize(sErr, 1).Value = fConvertDictionaryItemsTo2DimenArrayForPaste(dictNewProductUnit)
        Erase arrNewProductUnit
        Call fFreezeSheet(shtException)
        
        shtException.Visible = xlSheetVisible
        shtException.Activate
        
        If fNzero(gsBusinessErrorMsg) Then gsBusinessErrorMsg = gsBusinessErrorMsg & vbCr & vbCr & vbCr & "===============================" & vbCr & vbCr

        gsBusinessErrorMsg = gsBusinessErrorMsg & sErr & "��ҩƷ����λ�����趨�Ļ�Ƶ�λ��һ�£�������Ҫ��" & vbCr _
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
'  '  sProductMasterUnit = fGetProductMasterUnit(sReplacedProducer, sReplacedProductName, sReplacedProductSeries)
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
'        sProductMasterUnit = fGetProductMasterUnit(sReplacedProducer, sReplacedProductName, sReplacedProductSeries)
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
