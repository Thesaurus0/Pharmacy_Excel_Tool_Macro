Attribute VB_Name = "MG01_ME_Hist_Profit"
Option Explicit
Option Base 1

Sub subMain_OpenHistProfitFile()
    Dim sHistFileFullPath As String
    Dim wbMonthly As Workbook
    Dim shtMonth As Worksheet
    
    If Not fIsDev() Then On Error GoTo error_handling

    fInitialization
    
    sHistFileFullPath = fGetLatestCreatedMEProfitFileAndUpdateConfig
    Set wbMonthly = fOpenWorkbook(sHistFileFullPath, , False)
    
    If Not fSheetExistsByCodeName("shtProfit", shtMonth, wbMonthly) Then _
    fErr "您选择的文件中没有找到代码名称为[" & "shtProfit" & "]的工作表, 请选择当初创建的文件" & vbCr & sHistFileFullPath
    
error_handling:
    Set shtMonth = Nothing
    
    If fCheckIfGotBusinessError Then fCloseWorkBookWithoutSave wbMonthly: GoTo reset_excel_options
    If fCheckIfUnCapturedExceptionAbnormalError Then GoTo reset_excel_options
    
reset_excel_options:
    Err.Clear
    fClearRefVariables
    fEnableExcelOptionsAll
    'End
End Sub
Sub subMain_CreateMonthEndFile_Profit()
    Dim sNewFile As String
    Dim response As VbMsgBoxResult
    Dim wb As Workbook
    
    If Not fIsDev() Then On Error GoTo error_handling
    
    fInitialization
    
    Do While True
        sNewFile = fGetMEProfitFileConfigAndUpdateConfigForCreation
        
        If fFileExists(sNewFile) Then
            response = MsgBox("您输入的文件已经存在，你想覆盖它吗？", vbYesNoCancel + vbCritical + vbDefaultButton2)
            If response = vbYes Then
                fIfExcelFileOpenedToCloseIt sNewFile
                fDeleteFile sNewFile
                Exit Do
            ElseIf response = vbNo Then
                
            Else
                fErr
            End If
        Else
            Exit Do
        End If
    Loop
    
    Set wb = fCopySingleSheet2NewWorkbookFile(shtProfit, sNewFile)
    
    Call fDeleteRowsFromSheetLeaveHeader(fGetSheetByCodeName("shtProfit", wb))
    
    Call fSaveAndCloseWorkBook(wb)
    
    Call fSetSpecifiedConfigCellValue(shtSysConf, "[System Misc Settings]", "Value", "Setting Item ID=MONTHEND_PROFIT_FILE_NAME_CREATED" _
            , sNewFile)
    
    fMsgBox "一个新的空文件已经创建： " & vbCr & sNewFile, vbInformation
    
error_handling:
    If Not wb Is Nothing Then fCloseWorkBookWithoutSave wb
    If fCheckIfGotBusinessError Then GoTo reset_excel_options
    If fCheckIfUnCapturedExceptionAbnormalError Then GoTo reset_excel_options
reset_excel_options:
    Err.Clear
    fClearRefVariables
    fEnableExcelOptionsAll
    'End
End Sub

Sub subMain_SaveProfitTableToHistory()
    Dim sHistFileFullPath As String
    Dim wbMonthly As Workbook
    Dim shtMonth As Worksheet
    
    If Not fIsDev() Then On Error GoTo error_handling

    fInitialization
    
    sHistFileFullPath = fGetLatestCreatedMEProfitFileAndUpdateConfig
    
    If Not fPromptToConfirmToContinue("您确定要把本软件中的利润表添加到历史利润文件中去吗？" & vbCr & vbCr & sHistFileFullPath) Then fErr
    
    Set wbMonthly = fOpenWorkbook(sHistFileFullPath, , False)
    
    If Not fSheetExistsByCodeName("shtProfit", shtMonth, wbMonthly) Then _
    fErr "您选择的文件中没有找到代码名称为[" & "shtProfit" & "]的工作表, 请选择当初创建的文件" & vbCr & sHistFileFullPath
    
    Dim arrData()
    Call fCopyReadWholeSheetData2Array(shtProfit, arrData)
    Call fAppendArray2Sheet(shtMonth, arrData)
    Erase arrData
    
    Call fBasicCosmeticFormatSheet(shtMonth)
    
    Call fSetConditionFormatForOddEvenLine(shtMonth)
    
    Call fSetBorderLineForSheet(shtMonth)

    Set shtMonth = Nothing
    fSaveAndCloseWorkBook wbMonthly
    
error_handling:
    Set shtMonth = Nothing
    If Not wbMonthly Is Nothing Then fCloseWorkBookWithoutSave wbMonthly
    If fCheckIfGotBusinessError Then GoTo reset_excel_options
    If fCheckIfUnCapturedExceptionAbnormalError Then GoTo reset_excel_options
    
    fMsgBox "本月利润已经保存到历史利润文件中： " & vbCr & sHistFileFullPath, vbInformation
reset_excel_options:
    Err.Clear
    fClearRefVariables
    fEnableExcelOptionsAll
    'End
End Sub


Private Function fGetLatestCreatedMEProfitFileAndUpdateConfig() As String
    Dim sHistFileFullPath As String
    Dim response As VbMsgBoxResult
    
    sHistFileFullPath = fGetSysMiscConfig("MONTHEND_PROFIT_FILE_NAME_CREATED")
    
    If Len(Trim(sHistFileFullPath)) <= 0 Then fErr "您还没有创建过历史文件，请点击按钮【第一次创建历史利润表】创建该文件。"

    If Not fFileExists(sHistFileFullPath) Then
        response = MsgBox("您上次创建的历史文件找不到，请确认您是否移动了位置。" & vbCr _
                & sHistFileFullPath & vbCr & vbCr _
               & "如果您要手动选择它，请点【Yes】" & vbCr _
               & "如果您想再创建一个，请点【No】,然后点击按钮【第一次创建历史利润表】创建该文件。" _
               , vbYesNo + vbCritical + vbDefaultButton1)
        If response = vbYes Then
            sHistFileFullPath = fSelectFileDialog(, "Excel File=*.xlsx;*.xls", "历史利润表")
            If Len(sHistFileFullPath) <= 0 Then fErr
            
            Call fSetSpecifiedConfigCellValue(shtSysConf, "[System Misc Settings]", "Value", "Setting Item ID=MONTHEND_PROFIT_FILE_NAME_CREATED" _
                    , sHistFileFullPath)
        Else
            fErr
        End If
    End If
    
    fGetLatestCreatedMEProfitFileAndUpdateConfig = sHistFileFullPath
End Function

Function fReplaceVariablesInConfiguration(ByRef sValue As String)
    If InStr(sValue, "$CURRENT_FOLDER$") > 0 Then sValue = Replace(sValue, "$CURRENT_FOLDER$", ThisWorkbook.Path)
    
    sValue = fReplaceDatePattern(sValue, Now())
End Function


Private Function fGetMEProfitFileConfigAndUpdateConfigForCreation() As String
    Dim sDefaultFolder As String
    Dim sFileName  As String
    Dim sFileFull As String
    
    sDefaultFolder = fGetSysMiscConfig("MONTHEND_PROFIT_FILE_DEFAULT_FOLDER")
    sFileName = fGetSysMiscConfig("MONTHEND_PROFIT_FILE_NAME_Pattern")
    
    fReplaceVariablesInConfiguration sDefaultFolder
    fReplaceVariablesInConfiguration sFileName
    
    sFileFull = fCheckPath(sDefaultFolder) & sFileName
    
    sFileFull = fSelectSaveAsFileDialog(sFileFull, "Excel File(*.xlsx),*.xlsx", "创建历史利润文件")
    
    If Len(sFileFull) <= 0 Then fErr
    
    Dim sParentFolder As String
    
    fGetFSO
    sParentFolder = gFSO.GetParentFolderName(sFileFull)
    sParentFolder = Replace(sParentFolder, ThisWorkbook.Path, "$CURRENT_FOLDER$")
    
    Call fSetSpecifiedConfigCellValue(shtSysConf, "[System Misc Settings]", "Value", "Setting Item ID=MONTHEND_PROFIT_FILE_DEFAULT_FOLDER" _
            , fCheckPath(sParentFolder))
    
    fGetMEProfitFileConfigAndUpdateConfigForCreation = sFileFull
End Function

