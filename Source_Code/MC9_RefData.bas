Attribute VB_Name = "MC9_RefData"
Option Explicit
Option Base 1

Enum Company
    Report_ID = 1
    ID = 2
    Name = 3
    Commission = 4
    Selected = 5
End Enum


Function fReadConfigCompanyList() As Dictionary
    Dim asTag As String
    Dim arrColsName()
    Dim arrKeyColsForValidation()
    Dim rngToFindIn As Range
    Dim arrConfigData()
    Dim lConfigStartRow As Long
    Dim lConfigStartCol As Long
    Dim lConfigEndRow As Long
    Dim lConfigHeaderAtRow As Long

    asTag = "[Sales Company List]"
    ReDim arrColsName(Company.Report_ID To Company.Selected)
    
    arrColsName(Company.Report_ID) = "Company ID"
    arrColsName(Company.ID) = "Company ID In DB"
    arrColsName(Company.Name) = "Company Name"
    arrColsName(Company.Commission) = "Default Commission"
    arrColsName(Company.Selected) = "User Ticked"
     
    arrConfigData = fReadConfigBlockToArrayNet(asTag:=asTag, rngToFindIn:=shtStaticData.Cells _
                                , arrColsName:=arrColsName _
                                , lConfigStartRow:=lConfigStartRow _
                                , lConfigStartCol:=lConfigStartCol _
                                , lConfigEndRow:=lConfigEndRow _
                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
                                , abNoDataConfigThenError:=True _
                                )
    Call fValidateDuplicateInArray(arrConfigData, Company.Report_ID, False, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "Company ID")
    Call fValidateDuplicateInArray(arrConfigData, Company.ID, False, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "Company ID In DB")
    Call fValidateDuplicateInArray(arrConfigData, Company.Name, False, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "Company Name")
    
'    Call fValidateBlankInArray(arrConfigData, Company.Report_ID, shtstaticdata, lConfigHeaderAtRow, lConfigStartCol, "Company ID")
'    Call fValidateBlankInArray(arrConfigData, Company.ID, shtstaticdata, lConfigHeaderAtRow, lConfigStartCol, "Company ID In DB")
'    Call fValidateBlankInArray(arrConfigData, Company.Name, shtstaticdata, lConfigHeaderAtRow, lConfigStartCol, "Company Name")
    
    Set fReadConfigCompanyList = fReadArray2DictionaryWithMultipleColsCombined(arrConfigData, Company.Report_ID _
            , Array(Company.ID, Company.Name, Company.Commission, Company.Selected), DELIMITER)
    Erase arrColsName
    Erase arrConfigData
End Function

