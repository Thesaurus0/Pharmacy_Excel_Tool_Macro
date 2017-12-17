Attribute VB_Name = "MB1_ImportSalesFiles"
Option Explicit
Option Base 1

Public gsCompanyID As String
Public dictCompList As Dictionary

Sub subMain_ImportSalesInfoFiles()
    'If Not fIsDev Then On Error GoTo error_handling
    'On Error GoTo error_handling
    
    fInitialization
    
    gsRptID = "IMPORT_SALES_INFO"
    Call fUnProtectSheet(shtSalesRawDataRpt)
    
    Call fReadSysConfig_InputTxtSheetFile
    
    Set dictCompList = fReadConfigCompanyList
    Call fValidateUserInputAndSetToConfigSheet
    
    Call fSetBackToConfigSheetAndUpdategDict_UserTicket
    Call fSetBackToConfigSheetAndUpdategDict_InputFiles
    
    gsRptFilePath = fReadSysConfig_Output(, gsRptType)
    
    Dim i As Integer
    Dim iCnt As Integer
    iCnt = 0
    For i = 0 To dictCompList.Count - 1
        gsCompanyID = dictCompList.Keys(i)
        
        If fGetCompany_UserTicked(gsCompanyID) = "Y" Then
            iCnt = iCnt + 1
        End If
    Next
    
    If iCnt <= 0 Then fErr "No Company is selected."
    
    Call fCleanSheetOutputResetSheetOutput(shtSalesRawDataRpt)
    Call fPrepareOutputSheetHeaderAndTextColumns(shtSalesRawDataRpt)
    
    For i = 0 To dictCompList.Count - 1
        gsCompanyID = dictCompList.Keys(i)
        
        If fGetCompany_UserTicked(gsCompanyID) = "Y" Then
            Call fLoadFilesAndRead2Variables
            
            If gsCompanyID = "PW" Then
                arrMaster = fFileterTwoDimensionArray(arrMaster, dictMstColIndex("RecordType"), "销售出库")
            End If

            Call fProcessDataAll
            
            Erase arrMaster
            
            Call fDeleteSheet(gsCompanyID)
            
            Call fAppendArray2Sheet(shtSalesRawDataRpt, arrOutput)
        End If
    Next
    
    Call fReSequenceSeqNo
    
'    Call fSortDataInSheetSortSheetData(shtSalesRawDataRpt, Array(dictRptColIndex("SalesCompanyName") _
                                                                , dictRptColIndex("Hospital") _
                                                                , dictRptColIndex("SalesDate") _
                                                                , dictRptColIndex("ProductProducer") _
                                                                , dictRptColIndex("ProductName") _
                                                                , dictRptColIndex("ProductUnit")))
    Call fFormatOutputSheet(shtSalesRawDataRpt)
    
   ' Call fProtectSheetAndAllowEdit(shtSalesRawDataRpt, shtSalesRawDataRpt.Columns(4), UBound(arrOutput, 1) + 1, UBound(arrOutput, 2), False)
    Call fPostProcess(shtSalesRawDataRpt)
    
    shtSalesRawDataRpt.Rows(1).RowHeight = 25
    shtSalesRawDataRpt.Visible = xlSheetVisible
    shtSalesRawDataRpt.Activate
    shtSalesRawDataRpt.Range("A1").Select
    
    Call fModifyMoveActiveXButtonOnSheet(shtSalesRawDataRpt.Cells(1, fGetValidMaxCol(shtSalesRawDataRpt) + 1) _
                                        , "btnReplaceUnify", 1, 1, , 25, RGB(255, 20, 134), RGB(255, 255, 255))
error_handling:
    If fCheckIfGotBusinessError Then
        Call fSetReneratedReport(, "-")
        GoTo reset_excel_options
    End If
    
    If fCheckIfUnCapturedExceptionAbnormalError Then
        Call fSetReneratedReport(, "-")
        GoTo reset_excel_options
    End If
    
    Call fSetReneratedReport(, shtSalesRawDataRpt.Name)
    fMsgBox "成功整合在工作表：[" & shtSalesRawDataRpt.Name & "] 中，请检查！", vbInformation
    
reset_excel_options:
    Err.Clear
    fEnableExcelOptionsAll
    End
End Sub

'Function fImportAllSalesInfoFiles()
'    Dim i As Integer
'
'    For i = LBound(arrSalesCompanys, 1) To UBound(arrSalesCompanys, 1)
'        Call fImportSalesInfoFileForComapnay(CStr(arrSalesCompanys(i, 0)) _
'                                            , CStr(arrSalesCompanys(i, 1)) _
'                                            , CStr(arrSalesCompanys(i, 2)))
'    Next
'End Function

Private Function fLoadFilesAndRead2Variables()
    'gsCompanyID
    Call fLoadFileByFileTag(gsCompanyID)
    Call fReadMasterSheetData(gsCompanyID)

End Function
 

'Function fImportSalesInfoFileForComapnay(asCompanyID As String, asCompanyName As String, sSalesInfoFile As String)
'    Dim sTmpSht As String
'    sTmpSht = fGenRandomUniqueString
'
'
'
'End Function


Function fValidateUserInputAndSetToConfigSheet()
    Dim i As Integer
    Dim sEachCompanyID As String
    Dim sFilePathRange As String
    Dim sEachFilePath  As String
    
    For i = 0 To dictCompList.Count - 1
        sEachCompanyID = dictCompList.Keys(i)
        'sFilePathRange = "rngSalesFilePath_" & sEachCompanyID
        
        sFilePathRange = fGetCompany_InputFileTextBoxName(sEachCompanyID)
        sEachFilePath = Trim(shtMenu.Range(sFilePathRange).Value)
        
        If Not fFileExists(sEachFilePath) Then
            shtMenu.Activate
            shtMenu.Range(sFilePathRange).Select
            fErr Split(dictCompList(sEachCompanyID), DELIMITER)(1) & ": 输入的文件不存在，请检查：" & vbCr & sEachFilePath
        End If
        
        'Call fSetSalesInfoFileToMainConfig(sEachCompanyID, sEachFilePath)
    Next
End Function

'Function fSetSalesInfoFileToMainConfig(sCompanyId As String, sFile As String)
'    Call fSetSpecifiedConfigCellAddress(shtSysConf, "[Input Files]", "File Full Path", "Company ID=" & sCompanyId, sFile)
'End Function

Private Function fProcessDataAll()
    
    Call fRedimArrOutputBaseArrMaster
    
    Dim lEachRow As Long
    Dim sCompanyLongID As String
    Dim sCompanyName As String
    Dim iCnt As Long
    
    sCompanyLongID = fGetCompany_CompanyLongID(gsCompanyID)
    sCompanyName = fGetCompany_CompanyName(gsCompanyID)
    
    iCnt = 0
    For lEachRow = LBound(arrMaster, 1) To UBound(arrMaster, 1)
        arrOutput(lEachRow, dictRptColIndex("SalesCompanyID")) = sCompanyLongID
        arrOutput(lEachRow, dictRptColIndex("SalesCompanyName")) = sCompanyName
        arrOutput(lEachRow, dictRptColIndex("OrigSalesInfoID")) = Left(sCompanyLongID & String(15, "_"), 12) _
                                                                & Format(arrMaster(lEachRow, dictMstColIndex("SalesDate")), "YYYYMMDD") _
                                                                & Format(lEachRow, "00000")
        arrOutput(lEachRow, dictRptColIndex("SeqNo")) = lEachRow
        
        arrOutput(lEachRow, dictRptColIndex("SalesDate")) = arrMaster(lEachRow, dictMstColIndex("SalesDate"))
        arrOutput(lEachRow, dictRptColIndex("ProductProducer")) = arrMaster(lEachRow, dictMstColIndex("ProductProducer"))
        arrOutput(lEachRow, dictRptColIndex("ProductName")) = arrMaster(lEachRow, dictMstColIndex("ProductName"))
        arrOutput(lEachRow, dictRptColIndex("ProductSeries")) = arrMaster(lEachRow, dictMstColIndex("ProductSeries"))
        arrOutput(lEachRow, dictRptColIndex("Hospital")) = arrMaster(lEachRow, dictMstColIndex("Hospital"))
        arrOutput(lEachRow, dictRptColIndex("Quantity")) = arrMaster(lEachRow, dictMstColIndex("Quantity"))
        arrOutput(lEachRow, dictRptColIndex("SellPrice")) = arrMaster(lEachRow, dictMstColIndex("SellPrice"))
        
        If dictMstColIndex.Exists("ProductUnit") Then
            arrOutput(lEachRow, dictRptColIndex("ProductUnit")) = arrMaster(lEachRow, dictMstColIndex("ProductUnit"))
        End If
        
        If dictMstColIndex.Exists("SellAmount") Then
            arrOutput(lEachRow, dictRptColIndex("SellAmount")) = arrMaster(lEachRow, dictMstColIndex("SellAmount"))
        End If
    Next
End Function

Function fReSequenceSeqNo()
    Dim arr()
    Dim lMaxRow As Long
    Dim lCol As Long
    Dim eachRow As Long
    
    lMaxRow = fGetValidMaxRow(shtSalesRawDataRpt)
    lCol = dictRptColIndex("SeqNo")
    
    arr = fReadRangeDatatoArrayByStartEndPos(shtSalesRawDataRpt, 2, lCol, lMaxRow, lCol)
    
    lMaxRow = lMaxRow - 1
    For eachRow = LBound(arr, 1) To UBound(arr, 1)
        arr(eachRow, 1) = lMaxRow & "_" & Format(arr(eachRow, 1), "0000")
    Next
    
    shtSalesRawDataRpt.Cells(2, lCol).Resize(UBound(arr, 1), 1).Value = arr
    Erase arr
End Function

Function fProcessDataQualified()
    
End Function
