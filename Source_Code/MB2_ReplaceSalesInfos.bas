Attribute VB_Name = "MB2_ReplaceSalesInfos"
Option Explicit
Option Base 1

Sub subMain_ReplaceSalesInfos()
    'If Not fIsDev Then On Error GoTo error_handling
   ' On Error GoTo error_handling
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

    fMsgBox "成功整合在工作表：[" & shtSalesInfos.Name & "] 中，请检查！", vbInformation

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
    
    Dim dictNewHospital As Dictionary
    Dim dictNewProducer As Dictionary
    Dim dictNewProductName As Dictionary
    
    Set dictNewHospital = New Dictionary
    Set dictNewProducer = New Dictionary
    Set dictNewProductName = New Dictionary

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
        arrOutput(lEachRow, dictRptColIndex("SellAmount")) = arrMaster(lEachRow, dictMstColIndex("SellAmount"))
        arrOutput(lEachRow, dictRptColIndex("SellAmount")) = arrMaster(lEachRow, dictMstColIndex("SellPrice")) * arrMaster(lEachRow, dictMstColIndex("Quantity"))
        
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
        End If
        
        arrOutput(lEachRow, dictRptColIndex("MatchedProductProducer")) = sReplacedProducer
        ' Product producer end -----------------
        
        ' Product Name replace -----------------
        sProductName = arrMaster(lEachRow, dictMstColIndex("ProductName"))
        arrOutput(lEachRow, dictRptColIndex("ProductName")) = sProductName
        
        If Not fReplaceAndValidateInProductNameMaster(sReplacedProducer, sProductName, sReplacedProductName) Then
            If Not dictNewProductName.Exists(sReplacedProducer & DELIMITER & sReplacedProductName) Then
                dictNewProductName.Add sReplacedProducer & DELIMITER & sReplacedProductName, lEachRow + 1
            End If
        End If
        arrOutput(lEachRow, dictRptColIndex("MatchedProductName")) = sReplacedProductName
        ' Product Name end -----------------
        
        arrOutput(lEachRow, dictRptColIndex("ProductSeries")) = arrMaster(lEachRow, dictMstColIndex("ProductSeries"))
        arrOutput(lEachRow, dictRptColIndex("ProductUnit")) = arrMaster(lEachRow, dictMstColIndex("ProductUnit"))
        
        arrOutput(lEachRow, dictRptColIndex("MatchedProductName")) = ""
        arrOutput(lEachRow, dictRptColIndex("MatchedProductSeries")) = ""
        arrOutput(lEachRow, dictRptColIndex("MatchedProductUnit")) = ""
    Next
    
    Call fAddNewFoundMissedHospitalToSheet(dictNewHospital)
    Call fAddNewFoundMissedProducerToSheetException(dictNewProducer)
    Call fAddNewFoundMissedProductNameToSheetException(dictNewProductName)
    
    
End Function

Function fAddNewFoundMissedHospitalToSheet(dictNewHospital As Dictionary)
    '======= Hospitals paste to master sheet  ===============================================
    Dim arrNewHospital()
    arrNewHospital = fConvertDictionaryKeysTo2DimenArrayForPaste(dictNewHospital)
    Call fAppendArray2Sheet(shtHospital, arrNewHospital)
    
    If fUbound(arrNewHospital, 1) > 0 Then
        fMsgBox fUbound(arrNewHospital, 1) & "个医院找不到，" & vbCr & "他们被自动加入到了表【" & shtHospital.Name & "】中了." _
                & vbCr & "该表的最后面的数据为本次新加的。" & vbCr _
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
        arrNewProducer = fConvertDictionaryKeysTo2DimenArrayForPaste(dictNewProducer)
        Call fPrepareHeaderToSheet(shtException, "本系统中找不到的药品生产厂家")
        Call fAppendArray2Sheet(shtException, arrNewProducer)
        sErr = fUbound(arrNewProducer)
        shtException.Cells(2, 2).Resize(sErr, 1).Value = fConvertDictionaryItemsTo2DimenArrayForPaste(dictNewProducer)
        Erase arrNewProducer
        Call fFreezeSheet(shtException)
        
        shtException.Visible = xlSheetVisible
        shtException.Activate
        
        gsBusinessErrorMsg = sErr & "个药品生产厂家在本系统中找不到，您可能要：" & vbCr _
            & "1. 在【药品厂家替换表】中添加一条替换记录" & vbCr _
            & "2. 在【药品厂家主表】中新增一个厂家" & vbCr & vbCr _
            & "本次导入失败，完善数据后，请再次点击按钮进行【匹配替换统一】"
    End If
    '======= Producer end ===============================================
End Function

Function fAddNewFoundMissedProductNameToSheetException(dictNewProductName As Dictionary)
    '======= ProductName Validation ===============================================
    Dim arrNewProductName()
    Dim sErr As String
    If dictNewProductName.Count > 0 Then
        arrNewProductName = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictNewProductName)
        Call fPrepareHeaderToSheet(shtException, "药品厂家", "本系统中找不到的药品名称")
        Call fAppendArray2Sheet(shtException, arrNewProductName)
        sErr = fUbound(arrNewProductName)
        shtException.Cells(2, 3).Resize(sErr, 1).Value = fConvertDictionaryItemsTo2DimenArrayForPaste(dictNewProductName)
        Erase arrNewProductName
        Call fFreezeSheet(shtException)
        
        shtException.Visible = xlSheetVisible
        shtException.Activate
        
        gsBusinessErrorMsg = sErr & "个药品名称在本系统中找不到，您可能要：" & vbCr _
            & "1. 在【药品名称替换表】中添加一条替换记录" & vbCr _
            & "2. 在【药品名称主表】中新增一个名称" & vbCr _
            & "** 请注意：药品厂家没有问题，都匹配到了。" & vbCr & vbCr _
            & "本次导入失败，完善数据后，请再次点击按钮进行【匹配替换统一】"
    End If
    '======= ProductName end ===============================================
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

