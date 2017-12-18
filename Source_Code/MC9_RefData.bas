Attribute VB_Name = "MC9_RefData"
Option Explicit
Option Base 1

Enum Company
    REPORT_ID = 1
    ID = 2
    Name = 3
    Commission = 4
    CheckBoxName = 5
    InputFileTextBoxName = 6
    Selected = 7
End Enum
 
Dim dictHospitalMaster As Dictionary
Dim dictHospitalReplace As Dictionary

Dim dictProducerMaster As Dictionary
Dim dictProducerReplace As Dictionary

Dim dictProductNameMaster As Dictionary
Dim dictProductNameReplace As Dictionary

Dim dictProductMaster As Dictionary
Dim dictProductSeriesReplace As Dictionary
'Dim dictProductUnit As Dictionary
Dim dictProductUnitRatio As Dictionary
'Dim dictProductUnitRatio As Dictionary

Dim dictFirstLevelComm As Dictionary
Dim dictSecondLevelComm As Dictionary

Dim dictCompanyNameID As Dictionary

'Dim bFirstLCommDefaultGot As Boolean
'Dim dblFirstLCommDefault As Double
Dim dictDefaultCommConfiged As Dictionary
'Dim bSecondLCommDefaultGot As Boolean
'Dim SecondLCommDefault As TypeSecondLCommDefault

Dim dictSelfSalesDeductFrom As Dictionary
Dim dictSelfSalesColIndex As Dictionary
Dim arrSelfSales()
    
Function fReadConfigCompanyList(Optional ByRef dictCompanyNameID As Dictionary) As Dictionary
    Dim asTag As String
    Dim arrColsName()
    Dim rngToFindIn As Range
    Dim arrConfigData()
    Dim lConfigStartRow As Long
    Dim lConfigStartCol As Long
    Dim lConfigEndRow As Long
    Dim lConfigHeaderAtRow As Long

    asTag = "[Sales Company List]"
    ReDim arrColsName(Company.REPORT_ID To Company.Selected)
    
    arrColsName(Company.REPORT_ID) = "Company ID"
    arrColsName(Company.ID) = "Company ID In DB"
    arrColsName(Company.Name) = "Company Name"
    arrColsName(Company.Commission) = "Default Commission"
    arrColsName(Company.CheckBoxName) = "CheckBox Name"
    arrColsName(Company.InputFileTextBoxName) = "Input File TextBox Name"
    arrColsName(Company.Selected) = "User Ticked"
     
    arrConfigData = fReadConfigBlockToArrayNet(asTag:=asTag, shtParam:=shtStaticData _
                                , arrColsName:=arrColsName _
                                , lConfigStartRow:=lConfigStartRow _
                                , lConfigStartCol:=lConfigStartCol _
                                , lConfigEndRow:=lConfigEndRow _
                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
                                , abNoDataConfigThenError:=True _
                                )
    Call fValidateDuplicateInArray(arrConfigData, Company.REPORT_ID, False, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "Company ID")
    Call fValidateDuplicateInArray(arrConfigData, Company.ID, False, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "Company ID In DB")
    Call fValidateDuplicateInArray(arrConfigData, Company.Name, False, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "Company Name")
    Call fValidateDuplicateInArray(arrConfigData, Company.CheckBoxName, False, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "Company Name")
    Call fValidateDuplicateInArray(arrConfigData, Company.InputFileTextBoxName, False, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "Company Name")
    
'    Call fValidateBlankInArray(arrConfigData, Company.Report_ID, shtstaticdata, lConfigHeaderAtRow, lConfigStartCol, "Company ID")
'    Call fValidateBlankInArray(arrConfigData, Company.ID, shtstaticdata, lConfigHeaderAtRow, lConfigStartCol, "Company ID In DB")
'    Call fValidateBlankInArray(arrConfigData, Company.Name, shtstaticdata, lConfigHeaderAtRow, lConfigStartCol, "Company Name")
    
'    Set fReadConfigCompanyList = fReadArray2DictionaryWithMultipleColsCombined(arrConfigData, Company.Report_ID _
'            , Array(Company.ID, Company.Name, Company.Commission, Company.CheckBoxName, Company.InputFileTextBoxName, Company.Selected) _
'            , DELIMITER)

    Dim dictOut As Dictionary
    Set dictOut = New Dictionary
    
    Set dictCompanyNameID = New Dictionary
    
    Dim lEachRow As Long
    Dim sFileTag As String
    Dim sValueStr As String
    
    For lEachRow = LBound(arrConfigData, 1) To UBound(arrConfigData, 1)
        If fArrayRowIsBlankHasNoData(arrConfigData, lEachRow) Then GoTo next_row
        
'        sRptNameStr = DELIMITER & arrConfigData(lEachRow, 1) & DELIMITER
'        If InStr(sRptNameStr, DELIMITER & asReportID & DELIMITER) <= 0 Then GoTo next_row
        
        'lActualRow = lConfigHeaderAtRow + lEachRow
        
        sFileTag = Trim(arrConfigData(lEachRow, Company.REPORT_ID))
        sValueStr = fComposeStrForDictCompanyList(arrConfigData, lEachRow)
        
        dictOut.Add sFileTag, sValueStr
        
        dictCompanyNameID.Add arrConfigData(lEachRow, Company.Name), arrConfigData(lEachRow, Company.REPORT_ID)
next_row:
    Next
    
    Erase arrColsName
    Erase arrConfigData
    Set fReadConfigCompanyList = dictOut
    Set dictOut = Nothing
End Function

Function fComposeStrForDictCompanyList(arrConfigData, lEachRow As Long) As String
    Dim sOut As String
    Dim i As Integer
    
    For i = Company.ID To Company.Selected
        sOut = sOut & DELIMITER & Trim(arrConfigData(lEachRow, i))
    Next
    
    fComposeStrForDictCompanyList = Right(sOut, Len(sOut) - 1)
End Function

Function fGetCompany_InputFileTextBoxName(asCompanyID As String) As String
    fGetCompany_InputFileTextBoxName = Split(dictCompList(asCompanyID), DELIMITER)(Company.InputFileTextBoxName - Company.REPORT_ID - 1)
End Function
Function fGetCompany_CheckBoxName(asCompanyID As String) As String
    fGetCompany_CheckBoxName = Split(dictCompList(asCompanyID), DELIMITER)(Company.CheckBoxName - Company.REPORT_ID - 1)
End Function
Function fGetCompany_UserTicked(asCompanyID As String) As String
    fGetCompany_UserTicked = Split(dictCompList(asCompanyID), DELIMITER)(Company.Selected - Company.REPORT_ID - 1)
End Function
Function fGetCompany_CompanyLongID(asCompanyID As String) As String
    fGetCompany_CompanyLongID = Split(dictCompList(asCompanyID), DELIMITER)(Company.ID - Company.REPORT_ID - 1)
End Function
Function fGetCompany_CompanyName(asCompanyID As String) As String
    fGetCompany_CompanyName = Split(dictCompList(asCompanyID), DELIMITER)(Company.Name - Company.REPORT_ID - 1)
End Function

'====================== Hospital Master =================================================================
Function fReadSheetHospitalMaster2Dictionary()
    Dim arrData()
    Dim dictColIndex As Dictionary
    
    Call fReadSheetDataByConfig("HOSPITAL_MASTER", dictColIndex, arrData, , , , , shtHospital)
    Set dictHospitalMaster = fReadArray2DictionaryOnlyKeys(arrData, dictColIndex("Hospital"))
    
    Set dictColIndex = Nothing
End Function
Function fHospitalExistsInHospitalMaster(sHospital As String) As Boolean
    If dictHospitalMaster Is Nothing Then Call fReadSheetHospitalMaster2Dictionary
    
    fHospitalExistsInHospitalMaster = dictHospitalMaster.Exists(sHospital)
End Function
'------------------------------------------------------------------------------

'====================== Hospital Replacement =================================================================
Function fReadSheetHospitalReplace2Dictionary()
    Dim arrData()
    Dim dictColIndex As Dictionary
    
    Call fReadSheetDataByConfig("HOSPITAL_REPLACE_SHEET", dictColIndex, arrData, , , , , shtHospitalReplace)
    Set dictHospitalReplace = fReadArray2DictionaryWithSingleCol(arrData, dictColIndex("FromHospital"), dictColIndex("ToHospital"))
    
    Set dictColIndex = Nothing
End Function
Function fFindInConfigedReplaceHospital(sHospital As String) As String
    If dictHospitalReplace Is Nothing Then Call fReadSheetHospitalReplace2Dictionary
    
    If dictHospitalReplace.Exists(sHospital) Then
        fFindInConfigedReplaceHospital = dictHospitalReplace(sHospital)
    Else
        fFindInConfigedReplaceHospital = ""
    End If
End Function
'------------------------------------------------------------------------------

'====================== Producer Master =================================================================
Function fReadSheetProducerMaster2Dictionary()
    Dim arrData()
    Dim dictColIndex As Dictionary
    
    Call fReadSheetDataByConfig("PRODUCER_MASTER", dictColIndex, arrData, , , , , shtProductProducerMaster)
    Set dictProducerMaster = fReadArray2DictionaryOnlyKeys(arrData, dictColIndex("ProductProducer"), True, False)
    
    Set dictColIndex = Nothing
End Function
Function fProducerExistsInProducerMaster(sProducer As String) As Boolean
    If dictProducerMaster Is Nothing Then Call fReadSheetProducerMaster2Dictionary
    
    fProducerExistsInProducerMaster = dictProducerMaster.Exists(sProducer)
End Function
'------------------------------------------------------------------------------

'====================== Producer Replacement =================================================================
Function fReadSheetProducerReplace2Dictionary()
    Dim arrData()
    Dim dictColIndex As Dictionary
    
    Call fReadSheetDataByConfig("PRODUCER_REPLACE_SHEET", dictColIndex, arrData, , , , , shtProductProducerReplace)
    Set dictProducerReplace = fReadArray2DictionaryWithSingleCol(arrData, dictColIndex("FromProducer"), dictColIndex("ToProducer"))
    
    Set dictColIndex = Nothing
End Function
Function fFindInConfigedReplaceProducer(sProducer As String) As String
    If dictProducerReplace Is Nothing Then Call fReadSheetProducerReplace2Dictionary
    
    If dictProducerReplace.Exists(sProducer) Then
        fFindInConfigedReplaceProducer = dictProducerReplace(sProducer)
    Else
        fFindInConfigedReplaceProducer = ""
    End If
End Function
'------------------------------------------------------------------------------


'====================== ProductName Master =================================================================
Function fReadSheetProductNameMaster2Dictionary()
    Dim arrData()
    Dim dictColIndex As Dictionary
    
    Call fReadSheetDataByConfig("PRODUCER_NAME_MASTER", dictColIndex, arrData, , , , , shtProductNameMaster)
    Call fValidateDuplicateInArray(arrData, Array(dictColIndex("ProductProducer"), dictColIndex("ProductName")), False, shtProductNameMaster, 1, 1, "厂家 + 名称")
    
    Set dictProductNameMaster = fReadArray2DictionaryMultipleKeysWithMultipleColsCombined(arrData _
                                    , Array(dictColIndex("ProductProducer"), dictColIndex("ProductName")) _
                                    , Array(dictColIndex("ProductName")), DELIMITER, DELIMITER)
    Set dictColIndex = Nothing
End Function
Function fProductNameExistsInProductNameMaster(sProductProducer As String, sProductName As String) As Boolean
    If dictProductNameMaster Is Nothing Then Call fReadSheetProductNameMaster2Dictionary
    
    fProductNameExistsInProductNameMaster = dictProductNameMaster.Exists(sProductProducer & DELIMITER & sProductName)
End Function
'------------------------------------------------------------------------------

'====================== ProductName Replacement =================================================================
Function fReadSheetProductNameReplace2Dictionary()
    Dim arrData()
    Dim dictColIndex As Dictionary
    
    Call fReadSheetDataByConfig("PRODUCT_NAME_REPLACE_SHEET", dictColIndex, arrData, , , , , shtProductNameReplace)
    Set dictProductNameReplace = fReadArray2DictionaryMultipleKeysWithMultipleColsCombined(arrData _
                                    , Array(dictColIndex("ProductProducer"), dictColIndex("FromProductName")) _
                                    , Array(dictColIndex("ToProductName")), DELIMITER, DELIMITER)
    Set dictColIndex = Nothing
End Function
Function fFindInConfigedReplaceProductName(sProductProducer As String, sProductName As String) As String
    Dim sKey As String
    
    If dictProductNameReplace Is Nothing Then Call fReadSheetProductNameReplace2Dictionary
    
    sKey = sProductProducer & DELIMITER & sProductName
    
    If dictProductNameReplace.Exists(sKey) Then
        fFindInConfigedReplaceProductName = dictProductNameReplace(sKey)
    Else
        fFindInConfigedReplaceProductName = ""
    End If
End Function
'------------------------------------------------------------------------------


'====================== ProductSeries Master =================================================================
Function fReadSheetProductSeriesMaster2Dictionary()
    Dim arrData()
    Dim dictColIndex As Dictionary
    
    Call fReadSheetDataByConfig("PRODUCT_MASTER", dictColIndex, arrData, , , , , shtProductMaster)
    Call fValidateDuplicateInArray(arrData, Array(dictColIndex("ProductProducer"), dictColIndex("ProductName"), dictColIndex("ProductSeries")), False, shtProductMaster, 1, 1, "厂家 + 名称 + 规格")
    
    Set dictProductMaster = fReadArray2DictionaryMultipleKeysWithMultipleColsCombined(arrData _
                                    , Array(dictColIndex("ProductProducer"), dictColIndex("ProductName"), dictColIndex("ProductSeries")) _
                                    , Array(dictColIndex("ProductUnit")), DELIMITER, DELIMITER)
    Set dictColIndex = Nothing
End Function
Function fProductSeriesExistsInProductMaster(sProductProducer As String, sProductName As String, sProductSeries As String) As Boolean
    If dictProductMaster Is Nothing Then Call fReadSheetProductSeriesMaster2Dictionary
    
    fProductSeriesExistsInProductMaster = dictProductMaster.Exists(sProductProducer & DELIMITER & sProductName & DELIMITER & sProductSeries)
End Function
'------------------------------------------------------------------------------

'====================== ProductSeries Replacement =================================================================
Function fReadSheetProductSeriesReplace2Dictionary()
    Dim arrData()
    Dim dictColIndex As Dictionary
    
    Call fReadSheetDataByConfig("PRODUCT_SERIES_REPLACE_SHEET", dictColIndex, arrData, , , , , shtProductSeriesReplace)
    Set dictProductSeriesReplace = fReadArray2DictionaryMultipleKeysWithMultipleColsCombined(arrData _
                                    , Array(dictColIndex("ProductProducer"), dictColIndex("ProductName"), dictColIndex("FromProductSeries")) _
                                    , Array(dictColIndex("ToProductSeries")), DELIMITER, DELIMITER)
    Set dictColIndex = Nothing
End Function
Function fFindInConfigedReplaceProductSeries(sProductProducer As String, sProductName As String, sOrigProductSeries As String) As String
    Dim sKey As String
    
    If dictProductSeriesReplace Is Nothing Then Call fReadSheetProductSeriesReplace2Dictionary
    
    sKey = sProductProducer & DELIMITER & sProductName & DELIMITER & sOrigProductSeries
    
    If dictProductSeriesReplace.Exists(sKey) Then
        fFindInConfigedReplaceProductSeries = dictProductSeriesReplace(sKey)
    Else
        fFindInConfigedReplaceProductSeries = ""
    End If
End Function
'------------------------------------------------------------------------------

'====================== ProductUnit Ratio =================================================================
Function fReadSheetProductUnitRatio2Dictionary()
    Dim arrData()
    Dim dictColIndex As Dictionary
    
    Call fReadSheetDataByConfig("PRODUCT_UNIT_RATIO_SHEET", dictColIndex, arrData, , , , , shtProductUnitRatio)
    Set dictProductUnitRatio = fReadArray2DictionaryMultipleKeysWithMultipleColsCombined(arrData _
                                    , Array(dictColIndex("ProductProducer"), dictColIndex("ProductName"), dictColIndex("ProductSeries"), dictColIndex("FromUnit")) _
                                    , Array(dictColIndex("ProductUnit"), dictColIndex("Ratio")), DELIMITER, DELIMITER)
    Set dictColIndex = Nothing
End Function
Function fFindInConfigedReplaceProductUnit(sProductProducer As String, sProductName As String, sProductSeries As String _
                            , sOrigProductUnit As String _
                            , ByRef dblRatio As Double) As String
    Dim sKey As String
    
    If dictProductUnitRatio Is Nothing Then Call fReadSheetProductUnitRatio2Dictionary
    
    sKey = sProductProducer & DELIMITER & sProductName & DELIMITER & sProductSeries & DELIMITER & sOrigProductUnit
    
    If dictProductUnitRatio.Exists(sKey) Then
        dblRatio = Split(dictProductUnitRatio(sKey), DELIMITER)(1)
        fFindInConfigedReplaceProductUnit = Split(dictProductUnitRatio(sKey), DELIMITER)(0)
    Else
        dblRatio = 1
        fFindInConfigedReplaceProductUnit = ""
    End If
End Function

Function fGetProductMasterUnit(sProductProducer As String, sProductName As String, sProductSeries As String) As String
    Dim sKey As String
    
    If dictProductMaster Is Nothing Then Call fReadSheetProductSeriesMaster2Dictionary
    
    sKey = sProductProducer & DELIMITER & sProductName & DELIMITER & sProductSeries
    
    If Not dictProductMaster.Exists(sKey) Then _
        fErr "药品厂家+名称+规格 还不存在于药品主表中, 会计单位找不到的情况下，计算无法进行：" & vbCr & sProductProducer & vbCr & sProductName & vbCr & sProductSeries
    
    fGetProductMasterUnit = dictProductMaster(sKey)
End Function
'------------------------------------------------------------------------------

'====================== 1st Level Commission =================================================================
Function fReadSheetFirstLevelComm2Dictionary()
    Dim arrData()
    Dim dictColIndex As Dictionary
    
    Call fReadSheetDataByConfig("FIRST_LEVEL_COMMISSION", dictColIndex, arrData, , , , , shtFirstLevelCommission)
    Set dictFirstLevelComm = fReadArray2DictionaryMultipleKeysWithMultipleColsCombined(arrData _
            , Array(dictColIndex("SalesCompany") _
                  , dictColIndex("ProductProducer") _
                  , dictColIndex("ProductName") _
                  , dictColIndex("ProductSeries")) _
            , Array(dictColIndex("Commission")), DELIMITER, DELIMITER)
    Set dictColIndex = Nothing
End Function
Function fGetFirstLevelComm(sSalesCompName As String, sProducer As String, sProductName As String _
                            , sProductSeries As String, ByRef dblFirstComm As Double) As Boolean
    If dictFirstLevelComm Is Nothing Then Call fReadSheetFirstLevelComm2Dictionary
    
    Dim sKey As String
    sKey = sSalesCompName & DELIMITER & sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries
    
    If dictFirstLevelComm.Exists(sKey) Then
        dblFirstComm = dictFirstLevelComm(sKey)
        fGetFirstLevelComm = True
    Else
        dblFirstComm = 0
        fGetFirstLevelComm = False
    End If
End Function
'------------------------------------------------------------------------------


'====================== 2nd Level Commission =================================================================
Function fReadSheetSecondLevelComm2Dictionary()
    Dim arrData()
    Dim dictColIndex As Dictionary
    
    Call fReadSheetDataByConfig("SECOND_LEVEL_COMMISSION", dictColIndex, arrData, , , , , shtSecondLevelCommission)
    Set dictSecondLevelComm = fReadArray2DictionaryMultipleKeysWithMultipleColsCombined(arrData _
            , Array(dictColIndex("SalesCompany") _
                  , dictColIndex("Hospital") _
                  , dictColIndex("ProductProducer") _
                  , dictColIndex("ProductName") _
                  , dictColIndex("ProductSeries")) _
            , Array(dictColIndex("Commission")), DELIMITER, DELIMITER)
    Set dictColIndex = Nothing
End Function
Function fGetSecondLevelComm(sSalesCompName As String, sHospital, sProducer As String, sProductName As String _
                            , sProductSeries As String, ByRef dblSecondComm As Double) As Boolean
    If dictSecondLevelComm Is Nothing Then Call fReadSheetSecondLevelComm2Dictionary
    
    Dim sKey As String
    sKey = sSalesCompName & DELIMITER & sHospital & DELIMITER & sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries
    
    If dictSecondLevelComm.Exists(sKey) Then
        dblSecondComm = dictSecondLevelComm(sKey)
        fGetSecondLevelComm = True
    Else
        dblSecondComm = 0
        fGetSecondLevelComm = False
    End If
End Function
'------------------------------------------------------------------------------

Function fGetConfigFirstLevelDefaultComm() As Double
    'If Not bFirstLCommDefaultGot Then dblFirstLCommDefault = fGetSpecifiedConfigCellValue(shtSysConf, "[System Misc Settings]", "Value", "Setting Item ID=FIRST_LEVEL_COMMISSION_DEFAULT")
    
    If dictDefaultCommConfiged Is Nothing Then Call fReadConfigSecondLCommDefault2Dictionary
    
    fGetConfigFirstLevelDefaultComm = dictDefaultCommConfiged("FIRST_LEVEL_COMMISSION_DEFAULT")
End Function

Function fGetConfigSecondLevelDefaultComm(sSalesCompName As String) As Double
    If dictDefaultCommConfiged Is Nothing Then Call fReadConfigSecondLCommDefault2Dictionary

    Dim sCompID As String
    
    sCompID = fGetCompanyIdByCompanyName(sSalesCompName)
    
    fGetConfigSecondLevelDefaultComm = dictDefaultCommConfiged("SECOND_LEVEL_COMMISSION_DEFAULT_" & sCompID)
End Function

Function fReadConfigSecondLCommDefault2Dictionary()
    Dim arrConfigData()
    arrConfigData = fReadConfigBlockToArrayNet("[System Misc Settings]", shtSysConf, Array("Setting Item ID", "Value"))
    
    Set dictDefaultCommConfiged = fReadArray2DictionaryWithSingleCol(arrConfigData, 1, 2)
    Erase arrConfigData
End Function

Function fGetCompanyIdByCompanyName(sSalesCompName As String) As String
    sSalesCompName = Trim(sSalesCompName)
    If dictCompanyNameID Is Nothing Then Call fReadConfigCompanyList(dictCompanyNameID)
    
    If Not dictCompanyNameID.Exists(sSalesCompName) Then fErr "公司名称有错误，请检查。"
    
    fGetCompanyIdByCompanyName = Trim(dictCompanyNameID(sSalesCompName))
End Function

'====================== Self Sales =================================================================
Function fReadSelfSalesOrder2Dictionary()
    Dim sTmpKey As String
    Dim sProducer As String, sProductName As String, sProductSeries As String
    Dim dblSellQuantity As Double
    Dim dblHospitalQuantity As Double
                            
    Call fSortDataInSheetSortSheetDataByFileSpec("SELF_SALES_ORDER", Array("ProductProducer" _
                                    , "ProductName" _
                                    , "ProductSeries" _
                                    , "SalesDate"))
    
    Call fReadSheetDataByConfig("SELF_SALES_ORDER", dictSelfSalesColIndex, arrSelfSales, , , , , shtSelfSalesOrder)
    
    For lEachRow = LBound(arrSelfSales, 1) To UBound(arrSelfSales, 1)
        dblSellQuantity = arrSelfSales(lEachRow, dictSelfSalesColIndex("SellQuantity"))
        dblHospitalQuantity = arrSelfSales(lEachRow, dictSelfSalesColIndex("HospitalSellQuantity"))
        
        
        If dblSellQuantity < dblHospitalQuantity Then fErr "数据出错，医院销售数量不应该大于出货数量" _
                        & vbCr & "工作表：" & shtSelfSalesOrder.Name _
                        & vbCr & "行号：" & lEachRow + 1
        If dblSellQuantity = dblHospitalQuantity Then GoTo next_row
        
        sProducer = arrSelfSales(lEachRow, dictSelfSalesColIndex("ProductProducer"))
        sProductName = arrSelfSales(lEachRow, dictSelfSalesColIndex("ProductName"))
        sProductSeries = arrSelfSales(lEachRow, dictSelfSalesColIndex("ProductSeries"))
        
        sTmpKey = sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries
        
        If Not dictSelfSalesDeductFrom.Exists(sTmpKey) Then
            dictSelfSalesDeductFrom.Add sTmpKey, lEachRow
        End If
next_row:
    Next
    
    Set dictSelfSalesColIndex = Nothing
    
    'dictSelfSalesDeductFrom
End Function
Function fCalculateCostPriceFromSelfSalesOrder(sProducer As String, sProductName As String, sProductSeries As String _
                           , ByRef dblSalesQuantity As Double, ByRef dblSecondComm As Double) As Boolean
    If dictSelfSalesDeductFrom Is Nothing Then Call fReadSelfSalesOrder2Dictionary
    
    Dim bOut As Boolean
    Dim lDeductStartRow As Long
    Dim dblSellQuantity As Double
    Dim dblHospitalQuantity As Double
    
    bOut = False
    
    Dim sTmpKey As String
    sTmpKey = sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries
    
    If Not dictSelfSalesDeductFrom.Exists(sTmpKey) Then GoTo exit_fun
    
    lDeductStartRow = dictSelfSalesDeductFrom(sTmpKey)
    
    For lEachRow = lDeductStartRow To UBound(arrSelfSales)
        dblSellQuantity = arrSelfSales(lEachRow, dictSelfSalesColIndex("SellQuantity"))
        dblHospitalQuantity = arrSelfSales(lEachRow, dictSelfSalesColIndex("HospitalSellQuantity"))
        
        If dblSellQuantity <= dblHospitalQuantity Then fErr "这一行的日期晚，不应该出现抵扣" _
                        & vbCr & "工作表：" & shtSelfSalesOrder.Name _
                        & vbCr & "行号：" & lEachRow + 1
        
        arrSelfSales(dictSelfSalesColIndex("HospitalSellQuantity")) = aaaa
    Next
    
exit_fun:
    fCalculateCostPriceFromSelfSalesOrder = bOut
End Function
'------------------------------------------------------------------------------
