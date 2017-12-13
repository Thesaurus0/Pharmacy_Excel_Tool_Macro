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



