Attribute VB_Name = "MB1_ImportSalesFiles"
Option Explicit
Option Base 1

Dim arrQualifiedRows()
Public gsCompanyID As String
Public dictCompList As Dictionary

Sub subMain_ImportSalesInfoFiles()
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
            
'            If gsCompanyID = "PW" Then
'                arrMaster = fFileterTwoDimensionArray(arrMaster, dictMstColIndex("RecordType"), "销售出库")
'            ElseIf gsCompanyID = "SYY" Then
'                arrMaster = fFileterOutTwoDimensionArray(arrMaster, dictMstColIndex("Hospital"), "广州医药有限公司大众药品销售分公司")
'            End If

            Call fGetQualfiedRows
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

Private Function fLoadFilesAndRead2Variables()
    'gsCompanyID
    Call fLoadFileByFileTag(gsCompanyID)
    Call fReadMasterSheetData(gsCompanyID)

End Function



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
    Dim lEachOutputRow As Long
    Dim lEachSourceRow As Long
    Dim sCompanyLongID As String
    Dim sCompanyName As String
'    Dim iCnt As Long
    
    If fUbound(arrQualifiedRows) <= 0 Then Exit Function
    
    'Call fRedimArrOutputBaseArrMaster
    ReDim arrOutput(1 To fUbound(arrQualifiedRows), 1 To fGetReportMaxColumn())
    
    sCompanyLongID = fGetCompany_CompanyLongID(gsCompanyID)
    sCompanyName = fGetCompany_CompanyName(gsCompanyID)
    
'    iCnt = 0
'    For lEachRow = LBound(arrMaster, 1) To UBound(arrMaster, 1)
    For lEachOutputRow = LBound(arrQualifiedRows) To UBound(arrQualifiedRows)
        lEachSourceRow = arrQualifiedRows(lEachOutputRow)
        
        arrOutput(lEachOutputRow, dictRptColIndex("SalesCompanyID")) = sCompanyLongID
        arrOutput(lEachOutputRow, dictRptColIndex("SalesCompanyName")) = sCompanyName
        arrOutput(lEachOutputRow, dictRptColIndex("OrigSalesInfoID")) = Left(sCompanyLongID & String(15, "_"), 12) _
                                                                & format(arrMaster(lEachSourceRow, dictMstColIndex("SalesDate")), "YYYYMMDD") _
                                                                & format(lEachSourceRow, "00000")
        arrOutput(lEachOutputRow, dictRptColIndex("SeqNo")) = lEachOutputRow
        
        arrOutput(lEachOutputRow, dictRptColIndex("SalesDate")) = arrMaster(lEachSourceRow, dictMstColIndex("SalesDate"))
        arrOutput(lEachOutputRow, dictRptColIndex("ProductProducer")) = Trim(arrMaster(lEachSourceRow, dictMstColIndex("ProductProducer")))
        arrOutput(lEachOutputRow, dictRptColIndex("ProductName")) = Trim(arrMaster(lEachSourceRow, dictMstColIndex("ProductName")))
        arrOutput(lEachOutputRow, dictRptColIndex("ProductSeries")) = Trim(arrMaster(lEachSourceRow, dictMstColIndex("ProductSeries")))
        arrOutput(lEachOutputRow, dictRptColIndex("Hospital")) = Trim(arrMaster(lEachSourceRow, dictMstColIndex("Hospital")))
        arrOutput(lEachOutputRow, dictRptColIndex("Quantity")) = arrMaster(lEachSourceRow, dictMstColIndex("Quantity"))
        arrOutput(lEachOutputRow, dictRptColIndex("SellPrice")) = arrMaster(lEachSourceRow, dictMstColIndex("SellPrice"))
        
        If dictMstColIndex.Exists("ProductUnit") Then
            arrOutput(lEachOutputRow, dictRptColIndex("ProductUnit")) = arrMaster(lEachSourceRow, dictMstColIndex("ProductUnit"))
        End If
        
        If dictMstColIndex.Exists("SellAmount") Then
            arrOutput(lEachOutputRow, dictRptColIndex("SellAmount")) = arrMaster(lEachSourceRow, dictMstColIndex("SellAmount"))
        End If
    Next
End Function

Function fGetQualfiedRows()
    Dim sProductProducer As String
    Dim sProductName As String
    Dim sProductSeries As String
    Dim iQualifiedCnt As Long
    Dim lEachRow As Long
    
    ReDim arrQualifiedRows(LBound(arrMaster, 1) To UBound(arrMaster, 1))
    
    iQualifiedCnt = 0
    
    For lEachRow = LBound(arrMaster, 1) To UBound(arrMaster, 1)
        If gsCompanyID = "PW" Then
            If Trim(arrMaster(lEachRow, dictMstColIndex("RecordType"))) = "销售出库" Then
            Else
                GoTo next_row
            End If
        ElseIf gsCompanyID = "SYY" Then
            If Trim(arrMaster(lEachRow, dictMstColIndex("Hospital"))) = "广州医药有限公司大众药品销售分公司" Then
                GoTo next_row
            End If
        Else
        End If
        
        sProductProducer = Trim(arrMaster(lEachRow, dictMstColIndex("ProductProducer")))
        sProductName = Trim(arrMaster(lEachRow, dictMstColIndex("ProductName")))
        sProductSeries = Trim(arrMaster(lEachRow, dictMstColIndex("ProductSeries")))
        
        If sProductProducer = "津金世" And sProductName = "金世力德(匹多莫德颗粒)" And sProductSeries = "2g:0.4g*6袋" Then
            GoTo next_row
        End If
        
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
        arr(eachRow, 1) = lMaxRow & "_" & format(arr(eachRow, 1), "0000")
    Next
    
    shtSalesRawDataRpt.Cells(2, lCol).Resize(UBound(arr, 1), 1).Value = arr
    Erase arr
End Function

Function fProcessDataQualified()
    
End Function
