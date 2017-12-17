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

Function fReadConfigCompanyList() As Dictionary
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


