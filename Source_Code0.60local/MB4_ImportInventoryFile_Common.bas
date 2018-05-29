Attribute VB_Name = "MB4_ImportInventoryFile_Common"
Option Explicit
Option Base 1

Dim sCompanyName As String
Dim dictCompList As Dictionary
Dim dictCompIDName As Dictionary
Dim InvFileCol As typeInvFileCol
Dim arrQualifiedRows()

Private Enum InvColIDs
    ProductProducer = 1     'start
    ProductName = 2
    ProductSeries = 3
    ProductUnit = 4
    LotNum = 5
    InventoryDate = 6
    Quantity = 7            'end
    id = 8
    Name = 9
End Enum

Private Type typeInvFileCol
    ProductProducer As Integer
    ProductName As Integer
    ProductSeries As Integer
    ProductUnit As Integer
    LotNum As Integer
    Quantity As Integer
    InventoryDate As Integer
End Type

Sub subMain_ImportInventoryFiles_Common()
    Dim sFilePath As String
    Dim shtMaster As Worksheet
    
    fResetdictCZLSalesHospital
    If Not fIsDev() Then On Error GoTo error_handling
    'On Error GoTo error_handling
    
    Set shtCurrMenu = shtMenuCompInvt
    fInitialization
    
    gsRptID = "IMPORT_INVENTORY_FILE"
    Call fUnProtectSheet(shtInventoryRawDataRpt)
    
'    Call fReadSysConfig_InputTxtSheetFile
    'gsRptFilePath = fReadSysConfig_Output(, gsRptType)
    
'    Set dictCompList = fReadConfigCompanyList
    Call fValidateUserInput
    
    sCompanyName = Trim(shtMenuCompInvt.cbbCompanyList.Value)
    sFilePath = Trim(shtMenuCompInvt.Range("rngInventoryFilePathComm").Value)
    gsCompanyID = fGetCompanyIDByName_Common(sCompanyName)
    
    fReadSysConfig_Output
    
    fSetInvFileCol
    
    
'    Call fSetBackToConfigSheetAndUpdategDict_UserTicket
'    Call fSetBackToConfigSheetAndUpdategDict_InputFiles
    
'    gsRptFilePath = fReadSysConfig_Output(, gsRptType)
    
    Dim i As Integer
    Dim iCnt As Integer
    
    If fIfClearImport Then
        Call fCleanSheetOutputResetSheetOutput(shtInventoryRawDataRpt)
        Call fPrepareOutputSheetHeaderAndTextColumns(shtInventoryRawDataRpt)
    End If
    
    fClearContentLeaveHeader shtSalesCompInvUnified
    Call fSetReplaceUnifyErrorRowCount_SCompInventory(100)
    
    'Call fLoadFileByFileTag(gsCompanyID)
    If fSheetExists(gsCompanyID) Then Call fDeleteSheet(gsCompanyID)
    Call fImportSingleSheetExcelFileToThisWorkbook(sFilePath, gsCompanyID)
    
    Set shtMaster = ThisWorkbook.Worksheets(gsCompanyID)
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
    
    Call fAppendArray2Sheet(shtInventoryRawDataRpt, arrOutput)
    
  '  Call fReSequenceSeqNo
    
'    Call fSortDataInSheetSortSheetData(shtInventoryRawDataRpt, Array(dictRptColIndex("SalesCompanyName") _
                                                                , dictRptColIndex("Hospital") _
                                                                , dictRptColIndex("SalesDate") _
                                                                , dictRptColIndex("ProductProducer") _
                                                                , dictRptColIndex("ProductName") _
                                                                , dictRptColIndex("ProductUnit")))
    Call fFormatOutputSheet(shtInventoryRawDataRpt)
    
   ' Call fProtectSheetAndAllowEdit(shtInventoryRawDataRpt, shtInventoryRawDataRpt.Columns(4), UBound(arrOutput, 1) + 1, UBound(arrOutput, 2), False)
    Call fPostProcess(shtInventoryRawDataRpt)
    
    shtInventoryRawDataRpt.Rows(1).RowHeight = 25
    shtInventoryRawDataRpt.Visible = xlSheetVisible
    shtInventoryRawDataRpt.Activate
    shtInventoryRawDataRpt.Range("A1").Select
    
    Call fModifyMoveActiveXButtonOnSheet(shtInventoryRawDataRpt.Cells(1, fGetValidMaxCol(shtInventoryRawDataRpt) + 1) _
                                        , "btnReplaceUnify", 1, 1, , 25, RGB(255, 20, 134), RGB(255, 255, 255))
error_handling:
    If fCheckIfGotBusinessError Then
        GoTo reset_excel_options
    End If
    
    If fCheckIfUnCapturedExceptionAbnormalError Then
        GoTo reset_excel_options
    End If
     
    fMsgBox "成功整合在工作表：[" & shtInventoryRawDataRpt.Name & "] 中，请检查！", vbInformation
    
    Application.Goto shtInventoryRawDataRpt.Range("A" & fGetValidMaxRow(shtInventoryRawDataRpt)), True
reset_excel_options:
    Err.Clear
    fClearRefVariables
    fEnableExcelOptionsAll
    'End
End Sub

Private Function fValidateUserInput()
    Dim i As Integer
    Dim sCompany  As String
    Dim sFilePath  As String
    
    sCompany = Trim(shtMenuCompInvt.cbbCompanyList.Value)
    If Len(sCompany) <= 0 Then
        shtMenuCompInvt.cbbCompanyList.Activate
        shtMenuCompInvt.cbbCompanyList.Select
        fErr "公司不能为空，请选择公司。"
    Else
        If Not fCompanyNameExists(sCompany) Then
            shtMenuCompInvt.cbbCompanyList.Activate
            shtMenuCompInvt.cbbCompanyList.SelStart = 0
            shtMenuCompInvt.cbbCompanyList.SelLength = Len(shtMenuCompInvt.cbbCompanyList.Value)
            fErr "公司不存在，请重新选择。"
        End If
    End If
    
    sFilePath = Trim(shtMenuCompInvt.Range("rngInventoryFilePathComm").Value)
     
    If Not fFileExists(sFilePath) Then
        shtMenuCompInvt.Activate
        shtMenuCompInvt.Range("rngInventoryFilePathComm").Select
        fErr "输入的文件不存在，请检查：" & vbCr & sFilePath
    End If
End Function


Private Function fIfClearImport() As Boolean
    Dim bClearImport As Boolean
    
    bClearImport = shtMenuCompInvt.OBClearImport_Comm.Value
    
    Dim response As VbMsgBoxResult
    
    If bClearImport Then
        response = MsgBox(Prompt:="您确定要清空现有导入的所有数据吗？无法撤消的哦" _
                            & vbCr & "继续，请点【Yes】" & vbCr & "否则，请点【No】" _
                            , Buttons:=vbCritical + vbYesNo + vbDefaultButton2)
        If response <> vbYes Then fErr
    ElseIf shtMenuCompInvt.OBAppendImport_Comm.Value Then
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
'        If gsCompanyID = "PW" Then
'            If Trim(arrMaster(lEachRow, InvFileCol.RecordType)) = "销售出库" Then
'            Else
'                GoTo next_row
'            End If
'        ElseIf gsCompanyID = "SYY" Then
'            If Trim(arrMaster(lEachRow, InvFileCol.Hospital)) = "广州医药有限公司大众药品销售分公司" Then
'                GoTo next_row
'            End If
'        Else
'        End If
        
        sProductProducer = Trim(arrMaster(lEachRow, InvFileCol.ProductProducer))
        sProductName = Trim(arrMaster(lEachRow, InvFileCol.ProductName))
        sProductSeries = Trim(arrMaster(lEachRow, InvFileCol.ProductSeries))
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
    
    'Dim dtSalesDate As Date
    Dim sProducer As String
    Dim sProductNameOrig As String
    Dim sProductName As String
    Dim sProductSeries As String
    
    
    If dictCompList Is Nothing Then Set dictCompList = fReadConfigCompanyList_Comon
    
    If fUbound(arrQualifiedRows) <= 0 Then Exit Function
    
    'Call fRedimArrOutputBaseArrMaster
    ReDim arrOutput(1 To fUbound(arrQualifiedRows), 1 To fGetReportMaxColumn())
     
     
    For lEachOutputRow = LBound(arrQualifiedRows) To UBound(arrQualifiedRows)
        lEachSourceRow = arrQualifiedRows(lEachOutputRow)
        
        'dtSalesDate = arrMaster(lEachSourceRow, InvFileCol.SalesDate)
        
        arrOutput(lEachOutputRow, dictRptColIndex("SalesCompanyID")) = gsCompanyID
        arrOutput(lEachOutputRow, dictRptColIndex("SalesCompanyName")) = sCompanyName
         
        arrOutput(lEachOutputRow, dictRptColIndex("OrigInventoryID")) = Left(gsCompanyID & String(15, "_"), 12) _
                                                                & format(lEachSourceRow, "00000")
        arrOutput(lEachOutputRow, dictRptColIndex("SeqNo")) = lEachOutputRow
        
        sProducer = Trim(arrMaster(lEachSourceRow, InvFileCol.ProductProducer))
        sProductNameOrig = Trim(arrMaster(lEachSourceRow, InvFileCol.ProductName))
        sProductSeries = Trim(arrMaster(lEachSourceRow, InvFileCol.ProductSeries))
        
        If gsCompanyID = "ZSY" Then
            If InStr(sProducer, "/") > 0 Then sProducer = Trim(Split(sProducer, "/")(1))
        End If
         
        arrOutput(lEachOutputRow, dictRptColIndex("ProductProducer")) = sProducer
        
        sProductName = sProductNameOrig
        If gsCompanyID = "GKYX" Then
            If InStr(sProductNameOrig, "/") > 0 Then
                sProductName = Trim(Split(sProductNameOrig, "/")(0))
                sProductSeries = Trim(Split(sProductNameOrig, "/")(1))
            End If
        End If
        
        arrOutput(lEachOutputRow, dictRptColIndex("ProductName")) = sProductName
        arrOutput(lEachOutputRow, dictRptColIndex("ProductSeries")) = sProductSeries
          
        arrOutput(lEachOutputRow, dictRptColIndex("Quantity")) = arrMaster(lEachSourceRow, InvFileCol.Quantity)
        arrOutput(lEachOutputRow, dictRptColIndex("LotNum")) = "'" & arrMaster(lEachSourceRow, InvFileCol.LotNum)
        
        If InvFileCol.ProductUnit > 0 Then
            arrOutput(lEachOutputRow, dictRptColIndex("ProductUnit")) = arrMaster(lEachSourceRow, InvFileCol.ProductUnit)
        End If
        If InvFileCol.InventoryDate > 0 Then
            arrOutput(lEachOutputRow, dictRptColIndex("InventoryDate")) = arrMaster(lEachSourceRow, InvFileCol.InventoryDate)
        End If
    Next
End Function

Private Function fSetInvFileCol()
    InvFileCol.ProductProducer = fLetter2Num(Split(dictCompList(sCompanyName), DELIMITER)(InvColIDs.ProductProducer - 1))
    InvFileCol.ProductName = fLetter2Num(Split(dictCompList(sCompanyName), DELIMITER)(InvColIDs.ProductName - 1))
    InvFileCol.ProductSeries = fLetter2Num(Split(dictCompList(sCompanyName), DELIMITER)(InvColIDs.ProductSeries - 1))
    InvFileCol.Quantity = fLetter2Num(Split(dictCompList(sCompanyName), DELIMITER)(InvColIDs.Quantity - 1))
    InvFileCol.LotNum = fLetter2Num(Split(dictCompList(sCompanyName), DELIMITER)(InvColIDs.LotNum - 1))
    InvFileCol.ProductUnit = fLetter2Num(Split(dictCompList(sCompanyName), DELIMITER)(InvColIDs.ProductUnit - 1))
    
    If Len(Split(dictCompList(sCompanyName), DELIMITER)(InvColIDs.ProductUnit - 1)) > 0 Then
        InvFileCol.ProductUnit = fLetter2Num(Split(dictCompList(sCompanyName), DELIMITER)(InvColIDs.ProductUnit - 1))
    Else
        InvFileCol.ProductUnit = 0
    End If
    If Len(Split(dictCompList(sCompanyName), DELIMITER)(InvColIDs.InventoryDate - 1)) > 0 Then
        InvFileCol.InventoryDate = fLetter2Num(Split(dictCompList(sCompanyName), DELIMITER)(InvColIDs.InventoryDate - 1))
    Else
        InvFileCol.InventoryDate = 0
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

    asTag = "[Sales Company List - Common Importing - Inventory File]"
    ReDim arrColsName(InvColIDs.ProductProducer To InvColIDs.Name)
    
    arrColsName(InvColIDs.id) = "Company ID"
    arrColsName(InvColIDs.Name) = "Company Name"
    arrColsName(InvColIDs.ProductProducer) = "ProductProducer"
    arrColsName(InvColIDs.ProductName) = "ProductName"
    arrColsName(InvColIDs.ProductSeries) = "ProductSeries"
    arrColsName(InvColIDs.ProductUnit) = "ProductUnit"
    arrColsName(InvColIDs.LotNum) = "LotNum"
    arrColsName(InvColIDs.Quantity) = "Quantity"
    arrColsName(InvColIDs.InventoryDate) = "InventoryDate"
    
    arrConfigData = fReadConfigBlockToArrayNet(asTag:=asTag, shtParam:=shtStaticData _
                                , arrColsName:=arrColsName _
                                , lConfigStartRow:=lConfigStartRow _
                                , lConfigStartCol:=lConfigStartCol _
                                , lConfigEndRow:=lConfigEndRow _
                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
                                , abNoDataConfigThenError:=True _
                                )
    Call fValidateDuplicateInArray(arrConfigData, InvColIDs.id, False, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "Company ID In DB")
    Call fValidateDuplicateInArray(arrConfigData, InvColIDs.Name, False, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "Company Name")

    'Call fValidateBlankInArray(arrConfigData, InvColIDs.SalesDate, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "SalesDate")
    Call fValidateBlankInArray(arrConfigData, InvColIDs.ProductProducer, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "ProductProducer")
    Call fValidateBlankInArray(arrConfigData, InvColIDs.ProductName, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "ProductName")
    Call fValidateBlankInArray(arrConfigData, InvColIDs.ProductSeries, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "ProductSeries")
   ' Call fValidateBlankInArray(arrConfigData, InvColIDs.Hospital, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "Hospital ")
    Call fValidateBlankInArray(arrConfigData, InvColIDs.ProductUnit, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "ProductUnit ")
    Call fValidateBlankInArray(arrConfigData, InvColIDs.LotNum, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, " LotNum")
    Call fValidateBlankInArray(arrConfigData, InvColIDs.Quantity, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "Quantity ")
  '  Call fValidateBlankInArray(arrConfigData, InvColIDs.SellPrice, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "SellPrice ")

'    Set fReadConfigCompanyList = fReadArray2DictionaryWithMultipleColsCombined(arrConfigData, InvColIDs.Report_ID _
'            , Array(InvColIDs.ID, InvColIDs.Name, InvColIDs.Commission, InvColIDs.CheckBoxName, InvColIDs.InputFileTextBoxName, InvColIDs.Selected) _
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
        
        sCompName = Trim(arrConfigData(lEachRow, InvColIDs.Name))
        sValueStr = fComposeStrForDictCompanyList(arrConfigData, lEachRow)
        
        dictOut.Add sCompName, sValueStr
        
        dictCompanyIDName.Add arrConfigData(lEachRow, InvColIDs.id), arrConfigData(lEachRow, InvColIDs.Name)
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
    
    For i = InvColIDs.ProductProducer To InvColIDs.id
        sCol = Trim(arrConfigData(lEachRow, i))
        
        If i >= InvColIDs.ProductProducer And i <= InvColIDs.Quantity Then
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

'Function fGetCompanyListCommon() As Dictionary
'    If dictCompList Is Nothing Then Set dictCompList = fReadConfigCompanyList_Comon
'    Set fGetCompanyListCommon = dictCompList
'End Function

'Function fCompanyNameExists(sCompName As String) As Boolean
'    If dictCompList Is Nothing Then Set dictCompList = fReadConfigCompanyList_Comon
'    fCompanyNameExists = dictCompList.Exists(sCompName)
'End Function

'Function fGetCompanyIDByName_Common(sCompName As String) As String
'    If dictCompList Is Nothing Then Set dictCompList = fReadConfigCompanyList_Comon
'
'    sCompName = Trim(sCompName)
'
'    If Not dictCompList.Exists(sCompName) Then fErr "公司名称不存在于商业公司名称配置块rngStaticSalesCompanyNames_Comm中，请检查。" & vbCr & vbCr & sCompName
'
'    fGetCompanyIDByName_Common = Split(dictCompList(sCompName), DELIMITER)(InvColIDs.ID - 1)
'End Function

'Function fGetCompanyNameByID_Common(sCompanyID As String) As String
'    If dictCompIDName Is Nothing Then Set dictCompList = fReadConfigCompanyList_Comon(dictCompIDName)
'    fGetCompanyNameByID_Common = dictCompIDName(sCompanyID)
'End Function


Private Function fGetConfigMaxCol() As Long
    Dim lOut As Long
    Dim i As Integer
    
    Dim arr
    arr = Split(dictCompList(sCompanyName), DELIMITER)
    
    Dim arrNum()
    ReDim arrNum(InvColIDs.ProductProducer - 1 To InvColIDs.Quantity - 1)
    
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
Sub Test()
    Set dictCompList = fReadConfigCompanyList_Comon()
End Sub


