Attribute VB_Name = "MB1_ImportSalesFiles_Common"
Option Explicit
Option Base 1

Dim sCompanyName As String
Dim dictCompList As Dictionary
Dim dictCompIDName As Dictionary
Dim SalesFileCol As typeSalesFileCol
Dim arrQualifiedRows()

Private Enum CompanyComm
    SalesDate = 1
    ProductProducer = 2
    ProductName = 3
    ProductSeries = 4
    Hospital = 5
    ProductUnit = 6
    LotNum = 7
    Quantity = 8
    SellPrice = 9
    SellAmount = 10
    RecordType = 11
    id = 12
    Name = 13
End Enum

Private Type typeSalesFileCol
    SalesDate As Integer
    ProductProducer As Integer
    ProductName As Integer
    ProductSeries As Integer
    Hospital As Integer
    Quantity As Integer
    SellPrice As Integer
    LotNum As Integer
    SellAmount As Integer
    ProductUnit As Integer
    RecordType As Integer
End Type

Sub subMain_ImportSalesInfoFiles_Common()
    Dim sFilePath As String
    Dim shtMaster As Worksheet
    
    fResetdictCZLSalesHospital
    If Not fIsDev() Then On Error GoTo error_handling
    'On Error GoTo error_handling
    
    Set shtCurrMenu = shtMenu
    fInitialization
    
    gsRptID = "IMPORT_SALES_INFO"
    Call fUnProtectSheet(shtSalesRawDataRpt)
    
'    Call fReadSysConfig_InputTxtSheetFile
    'gsRptFilePath = fReadSysConfig_Output(, gsRptType)
    
'    Set dictCompList = fReadConfigCompanyList
    Call fValidateUserInput
    
    sCompanyName = Trim(shtMenu.cbbCompanyList.Value)
    sFilePath = Trim(shtMenu.Range("rngSalesFilePathComm").Value)
    gsCompanyID = fGetCompanyIDByName_Common(sCompanyName)
    
    fReadSysConfig_Output
    
    fSetSalesFileCol
    
    
'    Call fSetBackToConfigSheetAndUpdategDict_UserTicket
'    Call fSetBackToConfigSheetAndUpdategDict_InputFiles
    
'    gsRptFilePath = fReadSysConfig_Output(, gsRptType)
    
    Dim i As Integer
    Dim iCnt As Integer
    
    If fIfClearImport Then
        Call fCleanSheetOutputResetSheetOutput(shtSalesRawDataRpt)
        Call fPrepareOutputSheetHeaderAndTextColumns(shtSalesRawDataRpt)
    End If
    
    fClearContentLeaveHeader shtSalesInfos
    Call fSetReplaceUnifyErrorRowCount_SCompSalesInfo(100)
    
    'Call fLoadFileByFileTag(gsCompanyID)
    If fSheetExists(gsCompanyID) Then Call fDeleteSheet(gsCompanyID)
    Call fImportSingleSheetExcelFileToThisWorkbook(sFilePath, gsCompanyID)
    
    Set shtMaster = ThisWorkbook.Worksheets(gsCompanyID)
    Call fPreprocessImportedSalesInfoSheet(shtMaster)
    
    Call fRemoveFilterForSheet(shtMaster)
    
    shtMaster.Cells.Replace _
    What:="￥", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    
    'Call fReadMasterSheetData(gsCompanyID)
    
    Dim lMaxCol As Long
    lMaxCol = fGetConfigMaxCol
    Call fCopyReadWholeSheetData2Array(shtMaster, arrMaster, , , lMaxCol)
    
    Call fCheckIfSheetHasNodata_RaiseErrToStop(arrMaster, shtMaster)
    
    Call fGetQualfiedRows
    Call fProcessDataAll
    
    Erase arrMaster
    
    Call fDeleteSheet(shtMaster.Name)
    
    Call fAppendArray2Sheet(shtSalesRawDataRpt, arrOutput)
    
  '  Call fReSequenceSeqNo
    
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
      '  Call fSetReneratedReport(, "-")
        GoTo reset_excel_options
    End If
    
    If fCheckIfUnCapturedExceptionAbnormalError Then
       ' Call fSetReneratedReport(, "-")
        GoTo reset_excel_options
    End If
    
    Call fSetReneratedReport(, shtSalesRawDataRpt.Name)
    fMsgBox "成功整合在工作表：[" & shtSalesRawDataRpt.Name & "] 中，请检查！", vbInformation
    
    Application.Goto shtSalesRawDataRpt.Range("A" & fGetValidMaxRow(shtSalesRawDataRpt)), True
reset_excel_options:
    Err.Clear
    fEnableExcelOptionsAll
    End
End Sub

Private Function fValidateUserInput()
    Dim i As Integer
    Dim sCompany  As String
    Dim sFilePath  As String
    
    sCompany = Trim(shtMenu.cbbCompanyList.Value)
    If Len(sCompany) <= 0 Then
        shtMenu.cbbCompanyList.Activate
        shtMenu.cbbCompanyList.Select
        fErr "公司不能为空，请选择公司。"
    Else
        If Not fCompanyNameExists(sCompany) Then
            shtMenu.cbbCompanyList.Activate
            shtMenu.cbbCompanyList.SelStart = 0
            shtMenu.cbbCompanyList.SelLength = Len(shtMenu.cbbCompanyList.Value)
            fErr "公司不存在，请重新选择。"
        End If
    End If
    
    sFilePath = Trim(shtMenu.Range("rngSalesFilePathComm").Value)
     
    If Not fFileExists(sFilePath) Then
        shtMenu.Activate
        shtMenu.Range("rngSalesFilePathComm").Select
        fErr "输入的文件不存在，请检查：" & vbCr & sFilePath
    End If
End Function


Private Function fIfClearImport() As Boolean
    Dim bClearImport As Boolean
    
    bClearImport = shtMenu.OBClearImport_Comm.Value
    
    Dim response As VbMsgBoxResult
    
    If bClearImport Then
        response = MsgBox(Prompt:="您确定要清空现有导入的所有数据吗？无法撤消的哦" _
                            & vbCr & "继续，请点【Yes】" & vbCr & "否则，请点【No】" _
                            , Buttons:=vbCritical + vbYesNo + vbDefaultButton2)
        If response <> vbYes Then fErr
    ElseIf shtMenu.OBAppendImport_Comm.Value Then
'        response = MsgBox(Prompt:="您现在选择的是追加导入，请检查您要导入的数据是否有问题，否则可能会因为重复导入而出现重复的销售流向。" _
'                            & vbCr & "继续，请点【Yes】" & vbCr & "否则，请点【No】" _
'                            , Buttons:=vbCritical + vbYesNo + vbDefaultButton2)
'        If response <> vbYes Then fErr
    Else
    End If
    
    fIfClearImport = bClearImport
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
        If gsCompanyID = "PW" Then
            If Trim(arrMaster(lEachRow, SalesFileCol.RecordType)) = "销售出库" Then
            Else
                GoTo next_row
            End If
        ElseIf gsCompanyID = "SYY" Then
            If Trim(arrMaster(lEachRow, SalesFileCol.Hospital)) = "广州医药有限公司大众药品销售分公司" Then
                GoTo next_row
            End If
        ElseIf gsCompanyID = "JMXH" Then
            If Trim(arrMaster(lEachRow, 1)) Like "*类小计*" Or Trim(arrMaster(lEachRow, 1)) Like "*合计*" _
            Or Trim(arrMaster(lEachRow, 1)) Like "*总计*" Or Not IsDate(Trim(arrMaster(lEachRow, 1))) Then
                GoTo next_row
            End If
        Else
        End If
        
        sProductProducer = Trim(arrMaster(lEachRow, SalesFileCol.ProductProducer))
        sProductName = Trim(arrMaster(lEachRow, SalesFileCol.ProductName))
        sProductSeries = Trim(arrMaster(lEachRow, SalesFileCol.ProductSeries))
'
'        If sProductProducer = "津金世" And sProductName = "金世力德(匹多莫德颗粒)" And sProductSeries = "2g:0.4g*6袋" Then
'            GoTo next_row
'        End If
        
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



Private Function fProcessDataAll()
    Dim lEachOutputRow As Long
    Dim lEachSourceRow As Long
    
    Dim dtSalesDate As Date
    
    If dictCompList Is Nothing Then Set dictCompList = fReadConfigCompanyList_Comon
    
    If fUbound(arrQualifiedRows) <= 0 Then Exit Function
    
    'Call fRedimArrOutputBaseArrMaster
    ReDim arrOutput(1 To fUbound(arrQualifiedRows), 1 To fGetReportMaxColumn())
     
     
    For lEachOutputRow = LBound(arrQualifiedRows) To UBound(arrQualifiedRows)
        lEachSourceRow = arrQualifiedRows(lEachOutputRow)
        
        dtSalesDate = arrMaster(lEachSourceRow, SalesFileCol.SalesDate)
        
        arrOutput(lEachOutputRow, dictRptColIndex("SalesCompanyID")) = gsCompanyID
        arrOutput(lEachOutputRow, dictRptColIndex("SalesCompanyName")) = sCompanyName
        arrOutput(lEachOutputRow, dictRptColIndex("OrigSalesInfoID")) = Left(gsCompanyID & String(10, "_"), 12) _
                                                                & format(dtSalesDate, "YYYYMMDD") _
                                                                & format(lEachSourceRow, "00000")
        arrOutput(lEachOutputRow, dictRptColIndex("SeqNo")) = lEachOutputRow
        
        arrOutput(lEachOutputRow, dictRptColIndex("SalesDate")) = dtSalesDate
        arrOutput(lEachOutputRow, dictRptColIndex("ProductProducer")) = Trim(arrMaster(lEachSourceRow, SalesFileCol.ProductProducer))
        arrOutput(lEachOutputRow, dictRptColIndex("ProductName")) = Trim(arrMaster(lEachSourceRow, SalesFileCol.ProductName))
        arrOutput(lEachOutputRow, dictRptColIndex("ProductSeries")) = Trim(arrMaster(lEachSourceRow, SalesFileCol.ProductSeries))
        arrOutput(lEachOutputRow, dictRptColIndex("Hospital")) = Trim(arrMaster(lEachSourceRow, SalesFileCol.Hospital))
        arrOutput(lEachOutputRow, dictRptColIndex("Quantity")) = arrMaster(lEachSourceRow, SalesFileCol.Quantity)
        arrOutput(lEachOutputRow, dictRptColIndex("SellPrice")) = arrMaster(lEachSourceRow, SalesFileCol.SellPrice)
        arrOutput(lEachOutputRow, dictRptColIndex("LotNum")) = "'" & arrMaster(lEachSourceRow, SalesFileCol.LotNum)
        
        If SalesFileCol.ProductUnit > 0 Then
            arrOutput(lEachOutputRow, dictRptColIndex("ProductUnit")) = arrMaster(lEachSourceRow, SalesFileCol.ProductUnit)
        End If
        
        If SalesFileCol.SellAmount > 0 Then
            arrOutput(lEachOutputRow, dictRptColIndex("SellAmount")) = arrMaster(lEachSourceRow, SalesFileCol.SellAmount)
        End If
    Next
End Function

Private Function fSetSalesFileCol()
    SalesFileCol.SalesDate = fLetter2Num(Split(dictCompList(sCompanyName), DELIMITER)(CompanyComm.SalesDate - 1))
    SalesFileCol.ProductProducer = fLetter2Num(Split(dictCompList(sCompanyName), DELIMITER)(CompanyComm.ProductProducer - 1))
    SalesFileCol.ProductName = fLetter2Num(Split(dictCompList(sCompanyName), DELIMITER)(CompanyComm.ProductName - 1))
    SalesFileCol.ProductSeries = fLetter2Num(Split(dictCompList(sCompanyName), DELIMITER)(CompanyComm.ProductSeries - 1))
    SalesFileCol.Hospital = fLetter2Num(Split(dictCompList(sCompanyName), DELIMITER)(CompanyComm.Hospital - 1))
    SalesFileCol.Quantity = fLetter2Num(Split(dictCompList(sCompanyName), DELIMITER)(CompanyComm.Quantity - 1))
    SalesFileCol.SellPrice = fLetter2Num(Split(dictCompList(sCompanyName), DELIMITER)(CompanyComm.SellPrice - 1))
    SalesFileCol.LotNum = fLetter2Num(Split(dictCompList(sCompanyName), DELIMITER)(CompanyComm.LotNum - 1))
    
    If Len(Split(dictCompList(sCompanyName), DELIMITER)(CompanyComm.SellAmount - 1)) > 0 Then
        SalesFileCol.SellAmount = fLetter2Num(Split(dictCompList(sCompanyName), DELIMITER)(CompanyComm.SellAmount - 1))
    Else
        SalesFileCol.SellAmount = 0
    End If
    If Len(Split(dictCompList(sCompanyName), DELIMITER)(CompanyComm.ProductUnit - 1)) > 0 Then
        SalesFileCol.ProductUnit = fLetter2Num(Split(dictCompList(sCompanyName), DELIMITER)(CompanyComm.ProductUnit - 1))
    Else
        SalesFileCol.ProductUnit = 0
    End If
    If Len(Split(dictCompList(sCompanyName), DELIMITER)(CompanyComm.RecordType - 1)) > 0 Then
        SalesFileCol.RecordType = fLetter2Num(Split(dictCompList(sCompanyName), DELIMITER)(CompanyComm.RecordType - 1))
    Else
        SalesFileCol.RecordType = 0
    End If
End Function


Private Function fReadConfigCompanyList_Comon(Optional ByRef dictCompanyIDName As Dictionary) As Dictionary
    Dim asTag As String
    Dim arrColsName()
    Dim rngToFindIn As Range
    Dim arrConfigData()
    Dim lConfigStartRow As Long
    Dim lConfigStartCol As Long
    Dim lConfigEndRow As Long
    Dim lConfigHeaderAtRow As Long

    asTag = "[Sales Company List - Common Importing - Sales File]"
    ReDim arrColsName(CompanyComm.SalesDate To CompanyComm.Name)
    
    arrColsName(CompanyComm.id) = "Company ID"
    arrColsName(CompanyComm.Name) = "Company Name"
    arrColsName(CompanyComm.SalesDate) = "SalesDate"
    arrColsName(CompanyComm.ProductProducer) = "ProductProducer"
    arrColsName(CompanyComm.ProductName) = "ProductName"
    arrColsName(CompanyComm.ProductSeries) = "ProductSeries"
    arrColsName(CompanyComm.Hospital) = "Hospital"
    arrColsName(CompanyComm.ProductUnit) = "ProductUnit"
    arrColsName(CompanyComm.LotNum) = "LotNum"
    arrColsName(CompanyComm.Quantity) = "Quantity"
    arrColsName(CompanyComm.SellPrice) = "SellPrice"
    arrColsName(CompanyComm.SellAmount) = "SellAmount"
    arrColsName(CompanyComm.RecordType) = "RecordType"
    
    arrConfigData = fReadConfigBlockToArrayNet(asTag:=asTag, shtParam:=shtStaticData _
                                , arrColsName:=arrColsName _
                                , lConfigStartRow:=lConfigStartRow _
                                , lConfigStartCol:=lConfigStartCol _
                                , lConfigEndRow:=lConfigEndRow _
                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
                                , abNoDataConfigThenError:=True _
                                )
    Call fValidateDuplicateInArray(arrConfigData, CompanyComm.id, False, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "Company ID In DB")
    Call fValidateDuplicateInArray(arrConfigData, CompanyComm.Name, False, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "Company Name")

    Call fValidateBlankInArray(arrConfigData, CompanyComm.SalesDate, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "SalesDate")
    Call fValidateBlankInArray(arrConfigData, CompanyComm.ProductProducer, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "ProductProducer")
    Call fValidateBlankInArray(arrConfigData, CompanyComm.ProductName, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "ProductName")
    Call fValidateBlankInArray(arrConfigData, CompanyComm.ProductSeries, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "ProductSeries")
    Call fValidateBlankInArray(arrConfigData, CompanyComm.Hospital, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "Hospital ")
    Call fValidateBlankInArray(arrConfigData, CompanyComm.ProductUnit, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "ProductUnit ")
    Call fValidateBlankInArray(arrConfigData, CompanyComm.LotNum, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, " LotNum")
    Call fValidateBlankInArray(arrConfigData, CompanyComm.Quantity, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "Quantity ")
    Call fValidateBlankInArray(arrConfigData, CompanyComm.SellPrice, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "SellPrice ")

'    Set fReadConfigCompanyList = fReadArray2DictionaryWithMultipleColsCombined(arrConfigData, CompanyComm.Report_ID _
'            , Array(CompanyComm.ID, CompanyComm.Name, CompanyComm.Commission, CompanyComm.CheckBoxName, CompanyComm.InputFileTextBoxName, CompanyComm.Selected) _
'            , DELIMITER)

    Dim dictOut As Dictionary
    Set dictOut = New Dictionary
    
    Set dictCompanyIDName = New Dictionary
    
    Dim lEachRow As Long
    Dim sCompName As String
    Dim sValueStr As String
    
    For lEachRow = LBound(arrConfigData, 1) To UBound(arrConfigData, 1)
        If fArrayRowIsBlankHasNoData(arrConfigData, lEachRow) Then GoTo next_row
        
'        sRptNameStr = DELIMITER & arrConfigData(lEachRow, 1) & DELIMITER
'        If InStr(sRptNameStr, DELIMITER & asReportID & DELIMITER) <= 0 Then GoTo next_row
        
        'lActualRow = lConfigHeaderAtRow + lEachRow
        
        sCompName = Trim(arrConfigData(lEachRow, CompanyComm.Name))
        sValueStr = fComposeStrForDictCompanyList(arrConfigData, lEachRow)
        
        dictOut.Add sCompName, sValueStr
        
        dictCompanyIDName.Add arrConfigData(lEachRow, CompanyComm.id), arrConfigData(lEachRow, CompanyComm.Name)
next_row:
    Next
    
    Erase arrColsName
    Erase arrConfigData
    Set fReadConfigCompanyList_Comon = dictOut
    Set dictOut = Nothing
End Function

Private Function fComposeStrForDictCompanyList(arrConfigData, lEachRow As Long) As String
    Dim sOut As String
    Dim i As Integer
    Dim sCol As String
    
    For i = CompanyComm.SalesDate To CompanyComm.id
        sCol = Trim(arrConfigData(lEachRow, i))
        
        If i >= CompanyComm.SalesDate And i <= CompanyComm.RecordType Then
            If Len(sCol) > 0 Then
                If Range(sCol & "1") Is Nothing Then
                    fErr "[Sales Company List - Common Importing - Sales File]中列不正确，请检查。"
                End If
            End If
        End If
        
        sOut = sOut & DELIMITER & sCol
    Next
    
    fComposeStrForDictCompanyList = Right(sOut, Len(sOut) - 1)
End Function

Function fGetCompanyListCommon() As Dictionary
    If dictCompList Is Nothing Then Set dictCompList = fReadConfigCompanyList_Comon
    Set fGetCompanyListCommon = dictCompList
End Function

Function fCompanyNameExists(sCompName As String) As Boolean
    If dictCompList Is Nothing Then Set dictCompList = fReadConfigCompanyList_Comon
    fCompanyNameExists = dictCompList.Exists(sCompName)
End Function

Function fGetCompanyIDByName_Common(sCompName As String) As String
    If dictCompList Is Nothing Then Set dictCompList = fReadConfigCompanyList_Comon
    
    sCompName = Trim(sCompName)
    
    If Not dictCompList.Exists(sCompName) Then fErr "公司名称不存在于商业公司名称配置块rngStaticSalesCompanyNames_Comm中，请检查。" & vbCr & vbCr & sCompName

    fGetCompanyIDByName_Common = Split(dictCompList(sCompName), DELIMITER)(CompanyComm.id - 1)
End Function

Function fGetCompanyNameByID_Common(sCompanyID As String) As String
    If dictCompIDName Is Nothing Then Set dictCompList = fReadConfigCompanyList_Comon(dictCompIDName)
    fGetCompanyNameByID_Common = dictCompIDName(sCompanyID)
End Function


Private Function fGetConfigMaxCol() As Long
    Dim lOut As Long
    Dim i As Integer
    
    Dim arr
    arr = Split(dictCompList(sCompanyName), DELIMITER)
    
    Dim arrNum()
    ReDim arrNum(CompanyComm.SalesDate - 1 To CompanyComm.RecordType - 1)
    
    For i = LBound(arrNum) To UBound(arrNum)
        If Len(arr(i)) > 0 Then
            arrNum(i) = fLetter2Num(arr(i))
        Else
            arrNum(i) = 0
        End If
    Next
    
    Erase arr
    fGetConfigMaxCol = WorksheetFunction.Max(arrNum)
    Erase arrNum
End Function

Private Function fPreprocessImportedSalesInfoSheet(shtRaw As Worksheet)
    If gsCompanyID = "JMXH" Then
        'shtRaw.Rows(1).Delete shift:=xlUp
    End If
End Function

Sub test()
    Set dictCompList = fReadConfigCompanyList_Comon()
End Sub


