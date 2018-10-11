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
'    shtHist.Rows(1).Replace What:="ƥ���", Replacement:="", LookAt:=xlPart, _
'        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'    shtHist.Rows(1).Replace What:="��λ�����", Replacement:="", LookAt:=xlPart, _
'        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'    shtHist.Rows(1).Replace What:="���¼���", Replacement:="", LookAt:=xlPart, _
'        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    
    Call fSaveAndCloseWorkBook(wb)
    
    Call fSetSpecifiedConfigCellValue(shtSysConf, "[System Misc Settings]", "Value", "Setting Item ID=MONTHEND_CZL2SCOMP_FILE_NAME_CREATED" _
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
    
    If Len(Trim(sHistFileFullPath)) <= 0 Then fErr "����û�д�������ʷ�ļ���������ť����һ�δ�����֥������������ʷ���������ļ���"

    If Not fFileExists(sHistFileFullPath) Then
        response = MsgBox("���ϴδ�������ʷ�ļ��Ҳ�������ȷ�����Ƿ��ƶ���λ�á�" & vbCr _
                & sHistFileFullPath & vbCr & vbCr _
               & "�����Ҫ�ֶ�ѡ��������㡾Yes��" & vbCr _
               & "��������ٴ���һ������㡾No��,Ȼ������ť����һ�δ�����֥������������ʷ���������ļ���" _
               , vbYesNo + vbCritical + vbDefaultButton1)
        If response = vbYes Then
            sHistFileFullPath = fSelectFileDialog(, "Excel File=*.xlsx;*.xls", "��֥������������ʷ��")
            If Len(sHistFileFullPath) <= 0 Then fErr
            
        Else
            fErr
        End If
    End If
    
    Do While True
        Set wbMonthly = fOpenWorkbook(sHistFileFullPath, , False)
        
        If Not fSheetExistsByCodeName("shtCZLSales2SCompAll", shtMonth, wbMonthly) Then
            fCloseWorkBookWithoutSave wbMonthly
            fMsgBox "��ѡ����ļ���û���ҵ���������Ϊ[" & "shtCZLSales2SCompAll" & "]�Ĺ�����, ��ѡ�񵱳�������������ļ�" & vbCr & sHistFileFullPath
            
            sHistFileFullPath = fSelectFileDialog(, "Excel File=*.xlsx;*.xls", "��֥������������ʷ��")
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
    
    sFileFull = fSelectSaveAsFileDialog(sFileFull, "Excel File(*.xlsx),*.xlsx", "������֥������������ʷ�ļ�")
    
    If Len(sFileFull) <= 0 Then fErr
    
    Dim sParentFolder As String
    
    fGetFSO
    sParentFolder = gFSO.GetParentFolderName(sFileFull)
    sParentFolder = Replace(sParentFolder, ThisWorkbook.Path, "$CURRENT_FOLDER$")
    
    Call fSetSpecifiedConfigCellValue(shtSysConf, "[System Misc Settings]", "Value", "Setting Item ID=MONTHEND_CZL2SCOMP_FILE_DEFAULT_FOLDER" _
            , fCheckPath(sParentFolder))
    
    fGetMECZLSales2SCompFileConfigAndUpdateConfigForCreation = sFileFull
End Function


