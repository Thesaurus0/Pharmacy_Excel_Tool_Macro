Attribute VB_Name = "MB2_ReplaceSalesInfos"
Option Explicit
Option Base 1

Sub subMain_ReplaceSalesInfos()
    'If Not fIsDev Then On Error GoTo error_handling
    'On Error GoTo error_handling
    shtSalesRawDataRpt.Visible = xlSheetVisible
    Call fUnProtectSheet(shtSalesInfos)
    Call fCleanSheetOutputResetSheetOutput(shtSalesInfos)
    
    fInitialization
    
    gsRptID = "REPLACE_UNIFY_SALES_INFO"
    
    Call fReadSysConfig_InputTxtSheetFile
    
    'gsRptFilePath = fReadSysConfig_Output(, gsRptType)
    Call fLoadFilesAndRead2Variables
    
    'Call fPrepareOutputSheetHeaderAndTextColumns(shtSalesInfos)
    
    'Call fProcessDataAll
            
    
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
error_handling:
    If fCheckIfGotBusinessError Then GoTo reset_excel_options
    If fCheckIfUnCapturedExceptionAbnormalError Then GoTo reset_excel_options
    
    fMsgBox "成功整合在工作表：[ ] 中，请检查！", vbInformation
    
reset_excel_options:
    err.Clear
    fEnableExcelOptionsAll
    End
End Sub


Private Function fLoadFilesAndRead2Variables()
    'gsCompanyID
    Call fLoadFileByFileTag("IMPORTED_DATA")
    Call fReadMasterSheetData("IMPORTED_DATA", , , True)

End Function

