Attribute VB_Name = "MB6_ImportCZL2CompFile"
Option Explicit
Option Base 1
Dim arrQualifiedRows()

Sub subMain_ImportCZL2CompanySalesFile()
    fResetdictCZLSalesOD
    If Not fIsDev() Then On Error GoTo error_handling
    'On Error GoTo error_handling
    
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
    shtCZLSales2CompRawData.Visible = xlSheetVisible
    shtCZLSales2CompRawData.Activate
    shtCZLSales2CompRawData.Range("A1").Select
    
    Call fModifyMoveActiveXButtonOnSheet(shtCZLSales2CompRawData.Cells(1, fGetValidMaxCol(shtCZLSales2CompRawData) + 1) _
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
    
    Call fSetReneratedReport(, shtCZLSales2CompRawData.Name)
    fMsgBox "成功整合在工作表：[" & shtCZLSales2CompRawData.Name & "] 中，请检查！", vbInformation
    
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



Private Function fValidateUserInputAndSetToConfigSheet()
    Dim sEachFilePath  As String
     
    sEachFilePath = Trim(shtImportCZL2SalesCompSales.Range("rngCZL2CompSalesFile").Value)
    
    If Not fFileExists(sEachFilePath) Then
        shtImportCZL2SalesCompSales.Activate
        shtImportCZL2SalesCompSales.Range("rngCZL2CompSalesFile").Select
        fErr "输入的文件不存在，请检查：" & vbCr & sEachFilePath
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
'        response = MsgBox(Prompt:="您确定要清空现有导入的数据吗？无法撤消的哦" _
'                            & vbCr & "继续，请点【Yes】" & vbCr & "否则，请点【No】" _
'                            , Buttons:=vbCritical + vbYesNo + vbDefaultButton2)
'        If response <> vbYes Then fErr
'    Else
'        response = MsgBox(Prompt:="您现在选择的是追加导入，请检查您要导入的数据是否有问题，否则可能会因为重复导入而出现重复的销售流向。" _
'                            & vbCr & "继续，请点【Yes】" & vbCr & "否则，请点【No】" _
'                            , Buttons:=vbCritical + vbYesNo + vbDefaultButton2)
'        If response <> vbYes Then fErr
'    End If
'
'    fIfClearImport = bClearImport
'End Function

Sub subMain_CompareCZLInventory()
    If fGetReplaceUnifyCZLSales2CompaniesErrorRowCount > 0 Then
        fErr "采芝林的销售数据中有药品在系统中找不到，无法计算库存，请先处理这些错误。"
        shtSalesInfos.Visible = xlSheetVisible
        shtException.Visible = xlSheetVisible:         shtException.Activate
        End
    End If
    
    If Not fIsDev() Then On Error GoTo error_handling
    
    fRemoveFilterForSheet shtCZLInventory
    fRemoveFilterForSheet shtProductMaster
    fClearContentLeaveHeader shtCZLInvDiff
    
    fInitialization
    
    gsRptID = "COMPARE_CZL_INVENTORY"
    Call fReadSysConfig_InputTxtSheetFile
    
    gsRptFilePath = fReadSysConfig_Output(, gsRptType)
    
    Call fLoadFileByFileTag("PRODUCT_MASTER")
    Call fReadMasterSheetData("PRODUCT_MASTER", shtProductMaster)
    
    Dim dictCZLInformedInv As Dictionary
    Set dictCZLInformedInv = fReadArray2DictionaryWithMultipleKeyColsSingleItemCol(arrMaster _
                            , Array(dictMstColIndex("ProductProducer"), dictMstColIndex("ProductName"), dictMstColIndex("ProductSeries")) _
                            , CLng(dictMstColIndex("CZLInformedInventory")), DELIMITER)
    Erase arrMaster
    
    Dim dictCZLInv As Dictionary
    Dim arrCZLInv()
    Call fCopyReadWholeSheetData2Array(shtCZLInventory, arrCZLInv)
    'fReadArray2DictionaryWithMultipleKeyColsSingleItemCol
    Set dictCZLInv = fReadArray2DictionaryWithMultipleKeyColsSingleItemColSum(arrCZLInv _
                            , Array(1, 2, 3) _
                            , 6, DELIMITER)
    Erase arrCZLInv
    
    Dim dictInventoryDiff As Dictionary
    Set dictInventoryDiff = fCompare2Inventory(dictCZLInformedInv, dictCZLInv)
    
    arrOutput = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictInventoryDiff, , False)
    Call fAppendArray2Sheet(shtCZLInvDiff, arrOutput)
    arrOutput = fConvertDictionaryDelimiteredItemsTo2DimenArrayForPaste(dictInventoryDiff)
    shtCZLInvDiff.Cells(2, 5).Resize(UBound(arrOutput, 1), UBound(arrOutput, 2)).Value = arrOutput
     
    Call fFormatOutputSheet(shtCZLInvDiff)
    
    shtCZLInvDiff.Rows(1).RowHeight = 25
    shtCZLInvDiff.Visible = xlSheetVisible
    shtCZLInvDiff.Activate
    shtCZLInvDiff.Range("A1").Select
error_handling:
    If fCheckIfGotBusinessError Then
        Call fSetReneratedReport(, "-")
        GoTo reset_excel_options
    End If
    
    If fCheckIfUnCapturedExceptionAbnormalError Then
        Call fSetReneratedReport(, "-")
        GoTo reset_excel_options
    End If
    
    Call fSetReneratedReport(, shtCZLSales2CompRawData.Name)
    fMsgBox "采芝林库存差异计算结果在表：[" & shtCZLInvDiff.Name & "] 中，请检查！", vbInformation
    
reset_excel_options:
    Err.Clear
    fEnableExcelOptionsAll
    End
    
End Sub

Function fCompare2Inventory(dictCZLInformedInv As Dictionary, dictCZLInv As Dictionary) As Dictionary
    Dim dictOut As Dictionary
     
    Dim i As Long
    Dim sProdLotKey As String
    Dim dblInformedInv As Double
    Dim dblCalculatedInv As Double
    
    Set dictOut = New Dictionary
    
    For i = 0 To dictCZLInformedInv.Count - 1
        sProdLotKey = dictCZLInformedInv.Keys(i)
        dblInformedInv = CDbl(dictCZLInformedInv.Items(i))
        
        If Not dictCZLInv.Exists(sProdLotKey) Then
            dictOut.Add sProdLotKey, dblInformedInv & DELIMITER & "0" & DELIMITER & dblInformedInv
        Else
            dblCalculatedInv = dictCZLInv(sProdLotKey)
            dictOut.Add sProdLotKey, dblInformedInv & DELIMITER & dblCalculatedInv & DELIMITER & (dblInformedInv - dblCalculatedInv)
        End If
    Next
    
    Set fCompare2Inventory = dictOut
    Set dictOut = Nothing
End Function


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
'        Call fPrepareHeaderToSheet(shtException, Array("药品厂家", "药品名称", "本系统中找不到的药品【规格】", "行号"), lStartRow)
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
'        gsBusinessErrorMsg = gsBusinessErrorMsg & lRecCount & "个药品【规格】在本系统中找不到，您可能要：" & vbCr _
'            & "(1). 在【药品规格替换表】中添加一条替换记录" & vbCr _
'            & "(2). 在【药品主表】中新增一个规格" & vbCr _
'            & "** 请注意：药品厂家和药品名称没有问题，都匹配到了。" & vbCr & vbCr _
'            & "本次导入失败，完善数据后，请再次点击按钮进行【匹配替换统一】"
'    End If
'End Function
