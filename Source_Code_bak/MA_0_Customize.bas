Attribute VB_Name = "MA_0_Customize"
Option Explicit
Option Base 1

Function fSetBackToConfigSheetAndUpdategDict_UserTicket()
    
    Dim ckb As Object
    
    Dim eachObj As Object
    
    'for each eachobj in shtmenu.
    Dim i As Long
    Dim sCompanyID As String
    Dim sTickValue As String
    
    For i = 0 To dictCompList.Count - 1
        sCompanyID = dictCompList.Keys(i)
         
        If Not fActiveXControlExistsInSheet(shtMenu, fGetCompany_CheckBoxName(sCompanyID), ckb) Then GoTo next_company
        
        sTickValue = IIf(ckb.Value, "Y", "N")
        
        Call fSetSpecifiedConfigCellValue(shtStaticData, "[Sales Company List]", "User Ticked", "Company ID=" & sCompanyID, sTickValue)
        Call fUpdateDictionaryItemValueForDelimitedElement(dictCompList, sCompanyID, Company.Selected - Company.REPORT_ID, sTickValue)
next_company:
    Next
End Function

Function fSetBackToConfigSheetAndUpdategDict_InputFiles()
    Dim i As Integer
    Dim sEachCompanyID As String
    Dim sFilePathRange As String
    Dim sEachFilePath  As String
    
    For i = 0 To dictCompList.Count - 1
        sEachCompanyID = dictCompList.Keys(i)
        'sFilePathRange = "rngSalesFilePath_" & sEachCompanyID
        
        If fGetCompany_UserTicked(sEachCompanyID) = "Y" Then
            sFilePathRange = fGetCompany_InputFileTextBoxName(sEachCompanyID)
            sEachFilePath = Trim(shtMenu.Range(sFilePathRange).Value)
        Else
            sEachFilePath = "User not selected."
        End If
         
        Call fSetValueBackToSysConf_InputFile_FileName(sEachCompanyID, sEachFilePath)
        Call fUpdateGDictInputFile_FileName(sEachCompanyID, sEachFilePath)
        
        'Call fSetSalesInfoFileToMainConfig(sEachCompanyID, sEachFilePath)
    Next
    
    
'    sFile = Trim(shtMenu.Range("rngSalesFilePath_GY").Value)
'
'    Call fSetValueBackToSysConf_InputFile_FileName("GY", sFile)
'    Call fUpdateGDictInputFile_FileName("GY", sFile)
    
    
End Function

Function fSetIntialValueForShtMenuInitialize()
    
End Function

Function fInitialization()
    err.Clear
    gbNoData = False
    gbBusinessError = False
    gbUserCanceled = False
    gbCheckCompatibility = True
    
    If fZero(gsEnv) Then gsEnv = fGetEnvFromSysConf
    
    Call fDisableExcelOptionsAll
    
    If fIsDev Then Application.ScreenUpdating = True
    Application.ScreenUpdating = True   ' for testing
    
    Call fRevmoeFilterForAllSheets(ThisWorkbook)
End Function

