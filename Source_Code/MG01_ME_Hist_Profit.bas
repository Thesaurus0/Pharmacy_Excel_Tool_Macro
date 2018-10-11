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
    fErr "��ѡ����ļ���û���ҵ���������Ϊ[" & "shtProfit" & "]�Ĺ�����, ��ѡ�񵱳��������ļ�" & vbCr & sHistFileFullPath
    
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
            response = MsgBox("��������ļ��Ѿ����ڣ����븲������", vbYesNoCancel + vbCritical + vbDefaultButton2)
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
    
    fMsgBox "һ���µĿ��ļ��Ѿ������� " & vbCr & sNewFile, vbInformation
    
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
    
    If Not fPromptToConfirmToContinue("��ȷ��Ҫ�ѱ�����е��������ӵ���ʷ�����ļ���ȥ��" & vbCr & vbCr & sHistFileFullPath) Then fErr
    
    Set wbMonthly = fOpenWorkbook(sHistFileFullPath, , False)
    
    If Not fSheetExistsByCodeName("shtProfit", shtMonth, wbMonthly) Then _
    fErr "��ѡ����ļ���û���ҵ���������Ϊ[" & "shtProfit" & "]�Ĺ�����, ��ѡ�񵱳��������ļ�" & vbCr & sHistFileFullPath
    
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
    
    fMsgBox "���������Ѿ����浽��ʷ�����ļ��У� " & vbCr & sHistFileFullPath, vbInformation
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
    
    If Len(Trim(sHistFileFullPath)) <= 0 Then fErr "����û�д�������ʷ�ļ���������ť����һ�δ�����ʷ������������ļ���"

    If Not fFileExists(sHistFileFullPath) Then
        response = MsgBox("���ϴδ�������ʷ�ļ��Ҳ�������ȷ�����Ƿ��ƶ���λ�á�" & vbCr _
                & sHistFileFullPath & vbCr & vbCr _
               & "�����Ҫ�ֶ�ѡ��������㡾Yes��" & vbCr _
               & "��������ٴ���һ������㡾No��,Ȼ������ť����һ�δ�����ʷ������������ļ���" _
               , vbYesNo + vbCritical + vbDefaultButton1)
        If response = vbYes Then
            sHistFileFullPath = fSelectFileDialog(, "Excel File=*.xlsx;*.xls", "��ʷ�����")
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
    
    sFileFull = fSelectSaveAsFileDialog(sFileFull, "Excel File(*.xlsx),*.xlsx", "������ʷ�����ļ�")
    
    If Len(sFileFull) <= 0 Then fErr
    
    Dim sParentFolder As String
    
    fGetFSO
    sParentFolder = gFSO.GetParentFolderName(sFileFull)
    sParentFolder = Replace(sParentFolder, ThisWorkbook.Path, "$CURRENT_FOLDER$")
    
    Call fSetSpecifiedConfigCellValue(shtSysConf, "[System Misc Settings]", "Value", "Setting Item ID=MONTHEND_PROFIT_FILE_DEFAULT_FOLDER" _
            , fCheckPath(sParentFolder))
    
    fGetMEProfitFileConfigAndUpdateConfigForCreation = sFileFull
End Function

