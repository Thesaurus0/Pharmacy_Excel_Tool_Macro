Attribute VB_Name = "M1_ImportFiles"
Option Explicit
Option Base 1
   
Dim arrSalesCompanys()

Function fImportAllSalesInfoFiles()
    Dim i As Integer
    
    For i = LBound(arrSalesCompanys, 1) To UBound(arrSalesCompanys, 1)
        Call fImportSalesInfoFileForComapnay(CStr(arrSalesCompanys(i, 0)) _
                                            , CStr(arrSalesCompanys(i, 1)) _
                                            , CStr(arrSalesCompanys(i, 2)))
    Next
End Function
 

Function fImportSalesInfoFileForComapnay(asCompanyID As String, asCompanyName As String, sSalesInfoFile As String)
    Dim sTmpSht As String
    sTmpSht = fGenRandomUniqueString
    
    
    
End Function

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
    arrColsName = Array("Company ID", "Company Name")
    arrKeyColsForValidation = Array(1)
     
    arrConfigData = fReadConfigBlockToArrayValidated(asTag:=asTag, rngToFindIn:=shtMainConf.Cells _
                                , arrColsName:=arrColsName _
                                , lConfigStartRow:=lConfigStartRow _
                                , lConfigStartCol:=lConfigStartCol _
                                , lConfigEndRow:=lConfigEndRow _
                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
                                , abNoDataConfigThenError:=True _
                                , arrKeyCols:=arrKeyColsForValidation _
                                , bNetValues:=True _
                                )
    fReadConfigCompanyList = arrConfigData
    Erase arrColsName
    Erase arrConfigData
End Function

Function fValidationAndSetToConfigSheet()
    arrSalesCompanys = fReadConfigCompanyList()
    
    Dim i As Integer
    Dim sEachCompanyID As String
    Dim sFilePathRange As String
    Dim sEachFilePath  As String
    
    For i = LBound(arrSalesCompanys, 1) To UBound(arrSalesCompanys, 1)
        sEachCompanyID = arrSalesCompanys(i, 0)
        sFilePathRange = "rngSalesFilePath_" & sEachCompanyID
        sEachFilePath = Trim(shtMenu.Range(sFilePathRange).Value)
        
        If Not fFileExists(sEachFilePath) Then
            fMsgRaiseErr arrSalesCompanys(i, 1) & "的文件不存在，请检查：" & vbCr & sEachFilePath
        End If
        
        Call fSetSalesInfoFileToMainConfig(sEachCompanyID, sEachFilePath)
    Next
End Function

Function fSetSalesInfoFileToMainConfig(sCompanyId As String, sFile As String)
    Call fSetSpecifiedConfigCellAddress(shtMainConf, "[Input Files]", "File Full Path", "Company ID=" & sCompanyId, sFile)
End Function
