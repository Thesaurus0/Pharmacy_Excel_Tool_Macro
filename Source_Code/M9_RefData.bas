Attribute VB_Name = "M9_RefData"
Option Explicit
Option Base 1

Function fReadConfigCompanyList() As Variant
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
    arrColsName = Array("Company ID", "Company Name", "Company ID In DB", "User Ticked")
     
    arrConfigData = fReadConfigBlockToArrayNet(asTag:=asTag, rngToFindIn:=shtMainConf.Cells _
                                , arrColsName:=arrColsName _
                                , lConfigStartRow:=lConfigStartRow _
                                , lConfigStartCol:=lConfigStartCol _
                                , lConfigEndRow:=lConfigEndRow _
                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
                                , abNoDataConfigThenError:=True _
                                )
    Call fValidateDuplicateInArray(arrConfigData, 1, False, shtMainConf, lConfigHeaderAtRow, lConfigStartCol, "Company ID")
    Call fValidateDuplicateInArray(arrConfigData, 3, False, shtMainConf, lConfigHeaderAtRow, lConfigStartCol, "Company ID In DB")
    
    fReadConfigCompanyList = arrConfigData
    Erase arrColsName
    Erase arrConfigData
End Function

