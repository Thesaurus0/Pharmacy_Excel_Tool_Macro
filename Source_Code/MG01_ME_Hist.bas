Attribute VB_Name = "MG01_ME_Hist"
Option Explicit
Option Base 1

Sub subMain_SaveProfitTableToHistory()
    Dim sHistFolder As String
    Dim sHistFile As String
    Dim sHistFileFullPath As String
    
    If Not fIsDev() Then On Error GoTo error_handling
    On Error GoTo error_handling
    fInitialization
    
    fGetMonthEndProfitFileConfigAndUpdateConfig
    
    sHistFolder = fCheckPath(fGetSysMiscConfig("MONTHEND_PROFIT_FILE_SAVE_FOLDER"))
    sHistFile = fGetSysMiscConfig("MONTHEND_PROFIT_FILE_NAME")
    
    fReplaceVariablesInConfiguration sHistFolder
    fReplaceVariablesInConfiguration sHistFile
    
    sHistFileFullPath = fCheckPath(sHistFolder) & sHistFile
        
    fGetFSO
    If Not gFSO.FileExists(sHistFileFullPath) Then
        fMsgBox "设定的文件找不到。关闭该信息后，请选择您保存的历史文件。"
        
        sHistFileFullPath = fSelectSaveAsFileDialog(sHistFileFullPath, "Excel File(*.xls*),*.xlsx;*.xls", "请选择历史利润文件")
        If Len(sHistFileFullPath) <= 0 Then fErr
    End If
    
    Dim wbME As Workbook
    Dim shtMEProfit As Worksheet
    
    Set wbME = fOpenWorkbook(sHistFileFullPath)
    
    If Not fSheetExistsByCodeName("shtMEProfit", shtMEProfit, wbME) Then _
    fErr "您选择的文件中没有找到代码名称为[" & "shtMEProfit" & "]的工作表, 请选择当初创建的文件"
    
    
    
    
    fSaveAndCloseWorkBook wbME
    
error_handling:
    If Not wbME Is Nothing Then fCloseWorkBookWithoutSave wbME
    If fCheckIfGotBusinessError Then GoTo reset_excel_options
    If fCheckIfUnCapturedExceptionAbnormalError Then GoTo reset_excel_options
    
    fMsgBox "本月利润已经保存到历史利润文件中： " & vbCr & sHistFileFullPath, vbInformation
reset_excel_options:
    Err.Clear
    fEnableExcelOptionsAll
    End
End Sub

Private Function fGetMonthEndProfitFileConfigAndUpdateConfig() As String
    Dim sHistFolder As String
    Dim sHistFile As String
    Dim sHistFileFullPath As String
    
    sHistFolder = fCheckPath(fGetSysMiscConfig("MONTHEND_PROFIT_FILE_SAVE_FOLDER"))
    sHistFile = fGetSysMiscConfig("MONTHEND_PROFIT_FILE_NAME")
    
    fReplaceVariablesInConfiguration sHistFolder
    fReplaceVariablesInConfiguration sHistFile
    
    sHistFileFullPath = fCheckPath(sHistFolder) & sHistFile
        
    fGetFSO
    If Not gFSO.FileExists(sHistFileFullPath) Then
        fMsgBox "设定的文件找不到。关闭该信息后，请选择您保存的历史文件。"
        
        sHistFileFullPath = fSelectSaveAsFileDialog(sHistFileFullPath, "Excel File(*.xls*),*.xlsx;*.xls", "请选择历史利润文件")
        If Len(sHistFileFullPath) <= 0 Then fErr
    End If
    
    fGetMonthEndProfitFileConfigAndUpdateConfig = sHistFileFullPath
End Function

Function fReplaceVariablesInConfiguration(ByRef sValue As String)
    If InStr(sValue, "$CURRENT_FOLDER$") > 0 Then sValue = Replace(sValue, "$CURRENT_FOLDER$", ThisWorkbook.Path)
    
    sValue = fReplaceDatePattern(sValue, Now())
End Function



