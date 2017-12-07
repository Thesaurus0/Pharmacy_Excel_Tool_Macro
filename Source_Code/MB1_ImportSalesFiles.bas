Attribute VB_Name = "MB1_ImportSalesFiles"
Option Explicit
Option Base 1
   
'Dim arrSalesCompanys()

Public gsCompanyID As String
Dim dictCompList As Dictionary

Sub subMain_ImportSalesInfoFiles()
    gsEnv = "DEV"
    
    Set dictCompList = fReadConfigCompanyList
    Call fValidationAndSetToConfigSheet
    
   ' On Error GoTo error_handling
    gsCompanyID = "GY"
    
    Call fReadConfigInputFiles
    
    Call fSetUserTicketToConfigSheetAndgDict
    
    Call fOverWriteGDictWithUserInputOnMenuScreen
    
error_handling:
End Sub

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


Function fValidationAndSetToConfigSheet()
    Dim i As Integer
    Dim sEachCompanyID As String
    Dim sFilePathRange As String
    Dim sEachFilePath  As String
    
    For i = 0 To dictCompList.Count - 1
        sEachCompanyID = dictCompList.Keys(i)
        sFilePathRange = "rngSalesFilePath_" & sEachCompanyID
        sEachFilePath = Trim(shtMenu.Range(sFilePathRange).Value)
        
        If Not fFileExists(sEachFilePath) Then
            shtMenu.Activate
            shtMenu.Range(sFilePathRange).Select
            fErr Split(dictCompList(sEachCompanyID), DELIMITER)(1) & ": 输入的文件不存在，请检查：" & vbCr & sEachFilePath
        End If
        
        'Call fSetSalesInfoFileToMainConfig(sEachCompanyID, sEachFilePath)
    Next
End Function

'Function fSetSalesInfoFileToMainConfig(sCompanyId As String, sFile As String)
'    Call fSetSpecifiedConfigCellAddress(shtSysConf, "[Input Files]", "File Full Path", "Company ID=" & sCompanyId, sFile)
'End Function

Function fSetUserTicketToConfigSheetAndgDict()
    If shtMenu.ckb_GY Then
        Call fSetSpecifiedConfigCellValue(shtStaticData, "[Sales Company List]", "User Ticked", "Company ID=GY", "Y")
        Call fUpdateItemValueForDelimitedElement(dictCompList, "GY", Company.Selected - Company.Report_ID, "Y")
    Else
        Call fSetSpecifiedConfigCellValue(shtStaticData, "[Sales Company List]", "User Ticked", "Company ID=GY", "N")
        Call fUpdateItemValueForDelimitedElement(dictCompList, "GY", Company.Selected - Company.Report_ID, "Y")
    End If
End Function
