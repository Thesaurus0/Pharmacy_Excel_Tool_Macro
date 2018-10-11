Attribute VB_Name = "MG02_ME_Hist_CZLSales2SComp"
Option Explicit
Option Base 1
Sub subMain_CreateMonthEndFile_CZLSales2SComp()
    Dim sNewFile As String
    Dim response As VbMsgBoxResult
    Dim wb As Workbook
    
    If Not fIsDev() Then On Error GoTo error_handling
    fInitialization
    
    Do While True
        sNewFile = fGetMECZLSales2SCompFileConfigAndUpdateConfigForCreation
        
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
    
    Set wb = fCopySingleSheet2NewWorkbookFile(shtCZLSales2SCompAll, sNewFile)
    
    Dim shtHist As Worksheet
    
    Set shtHist = fGetSheetByCodeName("shtCZLSales2SCompAll", wb)
    
    Call fDeleteRowsFromSheetLeaveHeader(shtHist)
    
'    Dim rngTmp As Range
'    Set rngTmp = Union(shtHist.Columns(CZLSales2Comp.OrigSalesCompanyName), shtHist.Columns(CZLSales2Comp.OrigProductProducer) _
'         , shtHist.Columns(CZLSales2Comp.OrigProductName), shtHist.Columns(CZLSales2Comp.OrigProductSeries) _
'         , shtHist.Columns(CZLSales2Comp.OrigProductUnit), shtHist.Columns(CZLSales2Comp.OrigQuantity) _
'         , shtHist.Columns(CZLSales2Comp.OrigPrice), shtHist.Columns(CZLSales2Comp.OrigAmount), shtHist.Columns(CZLSales2Comp.OrigSalesInfoID))
'    rngTmp.Delete shift:=xlToLeft
'
'    shtHist.Rows(1).Replace What:="匹配后", Replacement:="", LookAt:=xlPart, _
'        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'    shtHist.Rows(1).Replace What:="单位换算后", Replacement:="", LookAt:=xlPart, _
'        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'    shtHist.Rows(1).Replace What:="重新计算", Replacement:="", LookAt:=xlPart, _
'        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    
    Call fSaveAndCloseWorkBook(wb)
    
    Call fSetSpecifiedConfigCellValue(shtSysConf, "[System Misc Settings]", "Value", "Setting Item ID=MONTHEND_CZL2SCOMP_FILE_NAME_CREATED" _
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

Sub subMain_OpenHistFile_CZLSales2SComp()
    Dim sHistFileFullPath As String
    Dim wbMonthly As Workbook
    
    If Not fIsDev() Then On Error GoTo error_handling

    fInitialization
    
    Call fGetLatestCreatedMEFileCZLSales2SCompAndUpdateConfig(sHistFileFullPath, wbMonthly)
error_handling:
    
    If fCheckIfGotBusinessError Then fCloseWorkBookWithoutSave wbMonthly: GoTo reset_excel_options
    If fCheckIfUnCapturedExceptionAbnormalError Then GoTo reset_excel_options
    
reset_excel_options:
    Err.Clear
    fClearRefVariables
    fEnableExcelOptionsAll
End Sub


Function fGetLatestCreatedMEFileCZLSales2SCompAndUpdateConfig(ByRef sHistFileFullPath As String _
                            , ByRef wbMonthly As Workbook, Optional ByRef shtMonth As Worksheet)
    Dim response As VbMsgBoxResult
    
    sHistFileFullPath = fGetSysMiscConfig("MONTHEND_CZL2SCOMP_FILE_NAME_CREATED")
    
    If Len(Trim(sHistFileFullPath)) <= 0 Then fErr "您还没有创建过历史文件，请点击按钮【第一次创建采芝林销售流向历史表】创建该文件。"

    If Not fFileExists(sHistFileFullPath) Then
        response = MsgBox("您上次创建的历史文件找不到，请确认您是否移动了位置。" & vbCr _
                & sHistFileFullPath & vbCr & vbCr _
               & "如果您要手动选择它，请点【Yes】" & vbCr _
               & "如果您想再创建一个，请点【No】,然后点击按钮【第一次创建采芝林销售流向历史表】创建该文件。" _
               , vbYesNo + vbCritical + vbDefaultButton1)
        If response = vbYes Then
            sHistFileFullPath = fSelectFileDialog(, "Excel File=*.xlsx;*.xls", "采芝林销售流向历史表")
            If Len(sHistFileFullPath) <= 0 Then fErr
            
        Else
            fErr
        End If
    End If
    
    Do While True
        Set wbMonthly = fOpenWorkbook(sHistFileFullPath, , False)
        
        If Not fSheetExistsByCodeName("shtCZLSales2SCompAll", shtMonth, wbMonthly) Then
            fCloseWorkBookWithoutSave wbMonthly
            fMsgBox "您选择的文件中没有找到代码名称为[" & "shtCZLSales2SCompAll" & "]的工作表, 请选择当初用软件创建的文件" & vbCr & sHistFileFullPath
            
            sHistFileFullPath = fSelectFileDialog(, "Excel File=*.xlsx;*.xls", "采芝林销售流向历史表")
            If Len(sHistFileFullPath) <= 0 Then fErr
        Else
            Exit Do
        End If
    Loop
    
    Call fSetSpecifiedConfigCellValue(shtSysConf, "[System Misc Settings]", "Value", "Setting Item ID=MONTHEND_CZL2SCOMP_FILE_NAME_CREATED" _
            , sHistFileFullPath)
End Function

Private Function fGetMECZLSales2SCompFileConfigAndUpdateConfigForCreation() As String
    Dim sDefaultFolder As String
    Dim sFileName  As String
    Dim sFileFull As String
    
    sDefaultFolder = fGetSysMiscConfig("MONTHEND_CZL2SCOMP_FILE_DEFAULT_FOLDER")
    sFileName = fGetSysMiscConfig("MONTHEND_CZL2SCOMP_FILE_NAME_Pattern")
    
    fReplaceVariablesInConfiguration sDefaultFolder
    fReplaceVariablesInConfiguration sFileName
    
    sFileFull = fCheckPath(sDefaultFolder) & sFileName
    
    sFileFull = fSelectSaveAsFileDialog(sFileFull, "Excel File(*.xlsx),*.xlsx", "创建采芝林销售流向历史文件")
    
    If Len(sFileFull) <= 0 Then fErr
    
    Dim sParentFolder As String
    
    fGetFSO
    sParentFolder = gFSO.GetParentFolderName(sFileFull)
    sParentFolder = Replace(sParentFolder, ThisWorkbook.Path, "$CURRENT_FOLDER$")
    
    Call fSetSpecifiedConfigCellValue(shtSysConf, "[System Misc Settings]", "Value", "Setting Item ID=MONTHEND_CZL2SCOMP_FILE_DEFAULT_FOLDER" _
            , fCheckPath(sParentFolder))
    
    fGetMECZLSales2SCompFileConfigAndUpdateConfigForCreation = sFileFull
End Function


