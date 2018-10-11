Attribute VB_Name = "MB6_ImportCZL2CompFile"
Option Explicit
Option Base 1
Dim arrQualifiedRows()

Sub subMain_ImportCZL2CompanySalesFile()
    If Not fIsDev() Then On Error GoTo error_handling
    
    Set shtCurrMenu = shtImportCZL2SalesCompSales
    fInitialization
    
    gsRptID = "IMPORT_CZL_SALES_TO_COMPANIES_FILE"
    Call fUnProtectSheet(shtCZLSales2CompRawData)
    
    Call fReadSysConfig_InputTxtSheetFile
    
    gsCompanyID = "CZL"
    'Set dictCompList = fReadConfigCompanyList
    Call fValidateUserInputAndSetToConfigSheet
    
'    Call fSetBackToConfigSheetAndUpdategDict_UserTicket
'    Call fSetBackToConfigSheetAndUpdategDict_InputFiles
    
    gsRptFilePath = fReadSysConfig_Output(, gsRptType)
    
   ' Dim i As Integer
'    Dim iCnt As Integer
'    iCnt = 0
'    For i = 0 To dictCompList.Count - 1
'        gsCompanyID = dictCompList.Keys(i)
'
'        If fGetCompany_UserTicked(gsCompanyID) = "Y" Then
'            iCnt = iCnt + 1
'        End If
'    Next
'
'    If iCnt <= 0 Then fErr "No Company is selected."
    
    If fIfClearImport Then
        Call fCleanSheetOutputResetSheetOutput(shtCZLSales2CompRawData)
        Call fPrepareOutputSheetHeaderAndTextColumns(shtCZLSales2CompRawData)
    End If
    
    fClearContentLeaveHeader shtCZLSales2Companies
    Call fSetReplaceUnifyErrorRowCount_CZLSales2Comp(100)

    Call fLoadFilesAndRead2Variables
    
'            If gsCompanyID = "PW" Then
'                arrMaster = fFileterTwoDimensionArray(arrMaster, dictMstColIndex("RecordType"), "���۳���")
'            ElseIf gsCompanyID = "SYY" Then
'                arrMaster = fFileterOutTwoDimensionArray(arrMaster, dictMstColIndex("Hospital"), "����ҽҩ���޹�˾����ҩƷ���۷ֹ�˾")
'            End If

    Call fGetQualfiedRows
    Call fProcessDataAll
    
    Erase arrMaster
    
    Call fDeleteSheet(gsCompanyID)
    
    Call fAppendArray2Sheet(shtCZLSales2CompRawData, arrOutput)
    
    
'    Call fSortDataInSheetSortSheetData(shtCZLSales2CompRawData, Array(dictRptColIndex("SalesCompanyName") _
                                                                , dictRptColIndex("Hospital") _
                                                                , dictRptColIndex("SalesDate") _
                                                                , dictRptColIndex("ProductProducer") _
                                                                , dictRptColIndex("ProductName") _
                                                                , dictRptColIndex("ProductUnit")))
    Call fFormatOutputSheet(shtCZLSales2CompRawData)
    
   ' Call fProtectSheetAndAllowEdit(shtCZLSales2CompRawData, shtCZLSales2CompRawData.Columns(4), UBound(arrOutput, 1) + 1, UBound(arrOutput, 2), False)
    Call fPostProcess(shtCZLSales2CompRawData)
    
    shtCZLSales2CompRawData.Rows(1).RowHeight = 25
    fShowAndActiveSheet shtCZLSales2CompRawData
    Application.Goto shtCZLSales2CompRawData.Range("A2"), True
    
    Call fModifyMoveActiveXButtonOnSheet(shtCZLSales2CompRawData.Cells(1, fGetValidMaxCol(shtCZLSales2CompRawData) + 1) _
                                        , "btnReplaceUnify", 1, 1, , 25, RGB(255, 20, 134), RGB(255, 255, 255))
error_handling:
    If fCheckIfGotBusinessError Then
        GoTo reset_excel_options
    End If
    
    If fCheckIfUnCapturedExceptionAbnormalError Then
        GoTo reset_excel_options
    End If
     
    fMsgBox "�ɹ������ڹ�����[" & shtCZLSales2CompRawData.Name & "] �У����飡", vbInformation
    
reset_excel_options:
    Err.Clear
    fClearRefVariables
    fEnableExcelOptionsAll
    'End
End Sub

Private Function fLoadFilesAndRead2Variables()
    'gsCompanyID
    Call fLoadFileByFileTag(gsCompanyID)
    Call fReadMasterSheetData(gsCompanyID)

End Function



Private Function fValidateUserInputAndSetToConfigSheet()
    Dim sEachFilePath  As String
     
    sEachFilePath = Trim(shtImportCZL2SalesCompSales.Range("rngCZL2CompSalesFile").Value)
    
    If Not fFileExists(sEachFilePath) Then
        shtImportCZL2SalesCompSales.Activate
        shtImportCZL2SalesCompSales.Range("rngCZL2CompSalesFile").Select
        fErr "������ļ������ڣ����飺" & vbCr & sEachFilePath
    End If
    
    Call fSetValueBackToSysConf_InputFile_FileName(gsCompanyID, sEachFilePath)
    Call fUpdateGDictInputFile_FileName(gsCompanyID, sEachFilePath)
End Function

'Function fSetSalesInfoFileToMainConfig(sCompanyId As String, sFile As String)
'    Call fSetSpecifiedConfigCellAddress(shtSysConf, "[Input Files]", "File Full Path", "Company ID=" & sCompanyId, sFile)
'End Function

Private Function fProcessDataAll()
    Dim lEachOutputRow As Long
    Dim lEachSourceRow As Long
'    Dim sCompanyLongID As String
'    Dim sCompanyName As String
'    Dim iCnt As Long
    
    If fUbound(arrQualifiedRows) <= 0 Then Exit Function
    
    'Call fRedimArrOutputBaseArrMaster
    ReDim arrOutput(1 To fUbound(arrQualifiedRows), 1 To fGetReportMaxColumn())
    
'    sCompanyLongID = fGetCompany_CompanyLongID(gsCompanyID)
'    sCompanyName = fGetCompany_CompanyName(gsCompanyID)
    
'    iCnt = 0
'    For lEachRow = LBound(arrMaster, 1) To UBound(arrMaster, 1)
    For lEachOutputRow = LBound(arrQualifiedRows) To UBound(arrQualifiedRows)
        lEachSourceRow = arrQualifiedRows(lEachOutputRow)
        
        arrOutput(lEachOutputRow, dictRptColIndex("OrigSalesInfoID")) = "'" & format(arrMaster(lEachSourceRow, dictMstColIndex("SalesDate")), "YYYYMMDD") _
                                                                & format(lEachSourceRow, "0000000")
        
        arrOutput(lEachOutputRow, dictRptColIndex("SalesDate")) = arrMaster(lEachSourceRow, dictMstColIndex("SalesDate"))
        arrOutput(lEachOutputRow, dictRptColIndex("ProductProducer")) = Trim(arrMaster(lEachSourceRow, dictMstColIndex("ProductProducer")))
        arrOutput(lEachOutputRow, dictRptColIndex("ProductName")) = Trim(arrMaster(lEachSourceRow, dictMstColIndex("ProductName")))
        arrOutput(lEachOutputRow, dictRptColIndex("ProductSeries")) = Trim(arrMaster(lEachSourceRow, dictMstColIndex("ProductSeries")))
        arrOutput(lEachOutputRow, dictRptColIndex("SalesCompanyName")) = Trim(arrMaster(lEachSourceRow, dictMstColIndex("SalesCompanyName")))
        arrOutput(lEachOutputRow, dictRptColIndex("Quantity")) = arrMaster(lEachSourceRow, dictMstColIndex("Quantity"))
        arrOutput(lEachOutputRow, dictRptColIndex("SellPrice")) = arrMaster(lEachSourceRow, dictMstColIndex("SellPrice"))
        arrOutput(lEachOutputRow, dictRptColIndex("LotNum")) = "'" & arrMaster(lEachSourceRow, dictMstColIndex("LotNum"))
        
        arrOutput(lEachOutputRow, dictRptColIndex("ProductUnit")) = arrMaster(lEachSourceRow, dictMstColIndex("ProductUnit"))
    Next
End Function

Private Function fGetQualfiedRows()
    Dim sProductProducer As String
    Dim sProductName As String
    Dim sProductSeries As String
    Dim iQualifiedCnt As Long
    Dim lEachRow As Long
    
    ReDim arrQualifiedRows(LBound(arrMaster, 1) To UBound(arrMaster, 1))
    
    iQualifiedCnt = 0
    
    For lEachRow = LBound(arrMaster, 1) To UBound(arrMaster, 1)
        sProductProducer = Trim(arrMaster(lEachRow, dictMstColIndex("ProductProducer")))
        sProductName = Trim(arrMaster(lEachRow, dictMstColIndex("ProductName")))
        sProductSeries = Trim(arrMaster(lEachRow, dictMstColIndex("ProductSeries")))
        
        If fProductExistsInExcludingProductListConfig(sProductProducer, sProductName, sProductSeries) Then GoTo next_row
        
        iQualifiedCnt = iQualifiedCnt + 1
        arrQualifiedRows(iQualifiedCnt) = lEachRow
next_row:
    Next
    
    If iQualifiedCnt > 0 Then
        ReDim Preserve arrQualifiedRows(1 To iQualifiedCnt)
    Else
        arrQualifiedRows = Array()
    End If
End Function

'Function fReSequenceSeqNo()
'    Dim arr()
'    Dim lMaxRow As Long
'    Dim lCol As Long
'    Dim eachRow As Long
'
'    lMaxRow = fGetValidMaxRow(shtCZLSales2CompRawData)
'    lCol = dictRptColIndex("SeqNo")
'
'    arr = fReadRangeDatatoArrayByStartEndPos(shtCZLSales2CompRawData, 2, lCol, lMaxRow, lCol)
'
'    lMaxRow = lMaxRow - 1
'    For eachRow = LBound(arr, 1) To UBound(arr, 1)
'        arr(eachRow, 1) = lMaxRow & "_" & format(arr(eachRow, 1), "0000")
'    Next
'
'    shtCZLSales2CompRawData.Cells(2, lCol).Resize(UBound(arr, 1), 1).Value = arr
'    Erase arr
'End Function

'Function fIfClearImport() As Boolean
'    Dim bClearImport As Boolean
'
'    If shtCurrMenu.Name = shtMenu.Name Then
'        bClearImport = shtMenu.OBClearImport.Value
'    ElseIf shtCurrMenu.Name = shtMenuCompInvt.Name Then
'        bClearImport = shtMenuCompInvt.OBClearImport.Value
'    Else
'        fErr "shtCurrMenu is neither shtMenu nor shtMenuCompInvt."
'    End If
'
'    'fIfClearImport = shtCurrMenu.OBClearImport.Value
'
'    Dim response As VbMsgBoxResult
'
'    If bClearImport Then
'        response = MsgBox(Prompt:="��ȷ��Ҫ������е�����������޷�������Ŷ" _
'                            & vbCr & "��������㡾Yes��" & vbCr & "������㡾No��" _
'                            , Buttons:=vbCritical + vbYesNo + vbDefaultButton2)
'        If response <> vbYes Then fErr
'    Else
'        response = MsgBox(Prompt:="������ѡ�����׷�ӵ��룬������Ҫ����������Ƿ������⣬������ܻ���Ϊ�ظ�����������ظ�����������" _
'                            & vbCr & "��������㡾Yes��" & vbCr & "������㡾No��" _
'                            , Buttons:=vbCritical + vbYesNo + vbDefaultButton2)
'        If response <> vbYes Then fErr
'    End If
'
'    fIfClearImport = bClearImport
'End Function

'Public Function fAddDictionaryToSheetException(dictNewProductSeries As Dictionary)
'    '======= ProductSeries Validation ===============================================
'    Dim arrNewProductSeries()
'    'Dim sErr As String
'    Dim lRecCount As Long
'        Dim lStartRow As Long
'
'    lRecCount = fGetDictionayDelimiteredItemsCount(dictNewProductSeries)
'    If lRecCount > 0 Then
'        lStartRow = fGetshtExceptionNewRow
'
'        arrNewProductSeries = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictNewProductSeries, , False)
'
'        Call fPrepareHeaderToSheet(shtException, Array("ҩƷ����", "ҩƷ����", "��ϵͳ���Ҳ�����ҩƷ�����", "�к�"), lStartRow)
'        shtException.Rows(lStartRow).Font.Color = RGB(255, 0, 0)
'        shtException.Rows(lStartRow).Font.Bold = True
'        Call fAppendArray2Sheet(shtException, arrNewProductSeries)
'        'sErr = fUbound(arrNewProductSeries)
'
'        shtException.Cells(lStartRow + 1, 4).Resize(dictNewProductSeries.Count, 1).Value = fConvertDictionaryItemsTo2DimenArrayForPaste(dictNewProductSeries, False)
'       ' Erase arrNewProductSeries
'        Call fFreezeSheet(shtException)
'
''        shtException.Visible = xlSheetVisible
''        shtException.Activate
'
'        If fNzero(gsBusinessErrorMsg) Then gsBusinessErrorMsg = gsBusinessErrorMsg & vbCr & vbCr & vbCr & "===============================" & vbCr & vbCr
'
'        gsBusinessErrorMsg = gsBusinessErrorMsg & lRecCount & "��ҩƷ������ڱ�ϵͳ���Ҳ�����������Ҫ��" & vbCr _
'            & "(1). �ڡ�ҩƷ����滻�������һ���滻��¼" & vbCr _
'            & "(2). �ڡ�ҩƷ����������һ�����" & vbCr _
'            & "** ��ע�⣺ҩƷ���Һ�ҩƷ����û�����⣬��ƥ�䵽�ˡ�" & vbCr & vbCr _
'            & "���ε���ʧ�ܣ��������ݺ����ٴε����ť���С�ƥ���滻ͳһ��"
'    End If
'End Function
