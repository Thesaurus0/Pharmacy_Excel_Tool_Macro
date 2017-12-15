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

    gsRptFilePath = fReadSysConfig_Output(, gsRptType)
    Call fLoadFilesAndRead2Variables

    Call fPrepareOutputSheetHeaderAndTextColumns(shtSalesInfos)

    Call fProcessData
    
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
error_handling:
    If fCheckIfGotBusinessError Then GoTo reset_excel_options
    If fCheckIfUnCapturedExceptionAbnormalError Then GoTo reset_excel_options

    fMsgBox "成功整合在工作表：[" & shtSalesInfos.Name & "] 中，请检查！", vbInformation

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

Private Function fProcessData()
    Call fRedimArrOutputBaseArrMaster
    
    Dim lEachRow As Long
    Dim sCompanyLongID As String
    Dim sCompanyName As String
    
    Dim sHospital As String
    Dim sReplacedHospital As String
    
    Dim dictNewHospital As Dictionary
    
    Set dictNewHospital = New Dictionary

    For lEachRow = LBound(arrMaster, 1) To UBound(arrMaster, 1)
        If dictMstColIndex.Exists("OrigSalesInfoID") Then
            arrOutput(lEachRow, dictRptColIndex("OrigSalesInfoID")) = arrMaster(lEachRow, dictMstColIndex("OrigSalesInfoID"))
        End If
        
        If dictMstColIndex.Exists("SeqNo") Then
            arrOutput(lEachRow, dictRptColIndex("SeqNo")) = arrMaster(lEachRow, dictMstColIndex("SeqNo"))
        End If
        
        arrOutput(lEachRow, dictRptColIndex("SalesCompanyName")) = arrMaster(lEachRow, dictMstColIndex("SalesCompanyName"))
        arrOutput(lEachRow, dictRptColIndex("SalesDate")) = arrMaster(lEachRow, dictMstColIndex("SalesDate"))
        arrOutput(lEachRow, dictRptColIndex("ProductProducer")) = arrMaster(lEachRow, dictMstColIndex("ProductProducer"))
        arrOutput(lEachRow, dictRptColIndex("ProductName")) = arrMaster(lEachRow, dictMstColIndex("ProductName"))
        arrOutput(lEachRow, dictRptColIndex("ProductSeries")) = arrMaster(lEachRow, dictMstColIndex("ProductSeries"))
        
        arrOutput(lEachRow, dictRptColIndex("Quantity")) = arrMaster(lEachRow, dictMstColIndex("Quantity"))
        arrOutput(lEachRow, dictRptColIndex("SellPrice")) = arrMaster(lEachRow, dictMstColIndex("SellPrice"))
        arrOutput(lEachRow, dictRptColIndex("ProductUnit")) = arrMaster(lEachRow, dictMstColIndex("ProductUnit"))
        arrOutput(lEachRow, dictRptColIndex("SellAmount")) = arrMaster(lEachRow, dictMstColIndex("SellAmount"))
        arrOutput(lEachRow, dictRptColIndex("SellAmount")) = arrMaster(lEachRow, dictMstColIndex("SellPrice")) * arrMaster(lEachRow, dictMstColIndex("Quantity"))
        
        sHospital = arrMaster(lEachRow, dictMstColIndex("Hospital"))
        arrOutput(lEachRow, dictRptColIndex("Hospital")) = sHospital
        
        If Not fReplaceAndValidateInHospitalMaster(sHospital, sReplacedHospital) Then
            If Not dictNewHospital.Exists(sReplacedHospital) Then dictNewHospital.Add sReplacedHospital, 0
        End If
        
        arrOutput(lEachRow, dictRptColIndex("MatchedHospital")) = sReplacedHospital
    Next
    
    '======= Hospital ===============================================
    Dim arrNewHoispital
    Dim arrHospitalForPaste()
    
    arrNewHoispital = dictNewHospital.Keys
    Set dictNewHospital = Nothing
    
    arrHospitalForPaste = fTranspose1DimenArrayTo2DimenArrayVertically(arrNewHoispital)
    
    If fUbound(arrNewHoispital) > 0 Then
        fMsgBox fUbound(arrNewHoispital) & "个医院找不到，" _
            & vbCr & "他们被自动加入到了表【" & shtHospital.Name & "】中了."
    End If
    
    Erase arrNewHoispital
    
    Call fAppendArray2Sheet(shtHospital, arrHospitalForPaste)
    Erase arrHospitalForPaste
    '======= Hospital end ===============================================
End Function

Function fReplaceAndValidateInHospitalMaster(sHospital As String, ByRef sReplacedHospital As String) As Boolean
    sReplacedHospital = fFindInConfigedReplaceHospital(sHospital)
    If fZero(sReplacedHospital) Then sReplacedHospital = sHospital
    
    fReplaceAndValidateInHospitalMaster = fHospitalExistsInHospitalMaster(sReplacedHospital)
End Function
