VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' version : 201712 - 001 lost the validate blank -- 23:45 -- 0��11
Option Explicit
Option Base 1

Dim arrAllCmdBarList()

'Function fIfSomeFundamentalSheetsWereDeleted() As Boolean
'    Dim sht As Worksheet
'
'    On Error Resume Next
'
'    Set sht = Sheet4
'    fIfSomeFundamentalSheetsWereDeleted = (sht Is Nothing)
'End Function

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call fRemoveAllCommandbarsByConfig

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.AskToUpdateLinks = True
    ThisWorkbook.CheckCompatibility = False
    
    Application.OnKey "^{BACKSPACE}"
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    'If fIfSomeFundamentalSheetsWereDeleted() Then Cancel = False
End Sub

Private Sub Workbook_Open()
    Application.OnKey "^{BACKSPACE}", "Sub_ToHomeSheet"
    fGetProgressBar
    gProBar.ShowBar
    gsEnv = fGetEnvFromSysConf
    
    Application.EnableEvents = False

    Call fUpdateRangeAddressForName_rngStaticSalesCompanyNames_Comm
    
    gProBar.ChangeProcessBarValue 0.1, "�����������Ͱ�ť"

    Call fRefreshGetAllCommandbarsList
    Call sub_WorkBookInitialization

    gProBar.ChangeProcessBarValue 0.7, "Ϊ�������ó�ʼ����"
    Call fSetIntialValueForShtMenuInitialize

    Call sub_RemoveCommandBar("Team")
    ThisWorkbook.Saved = True
    ThisWorkbook.CheckCompatibility = False

    gProBar.ChangeProcessBarValue 0.9, "��ͨ�ù��ܳ�ʼ�������б��"
    shtMenu.sub_Initialize_CompanyListCombobox_SalesInfo
    shtMenuCompInvt.sub_Initialize_CompanyListCombobox_Inventory

    gProBar.ChangeProcessBarValue 1, "�Ѿ�������"
'    Application.CommandBars("cell").FindControl(ID:=19).OnAction = ""
'    Application.OnKey "^c", ""
    gProBar.SleepBar 500
    'gProBar.DestroyBar

    shtSelfSalesA.Range("A1").Value = shtSelfSalesA.Range("A1").Value2 + 1
    Application.EnableEvents = True
    ThisWorkbook.Saved = True
    gProBar.DestroyBar
    'End
End Sub

Sub sub_WorkBookInitialization()
    
    Call fReadConfigRibbonCommandBarMenuAndCreateCommandBarButton
    
    If fIsDev() Then
        shtSysConf.Visible = xlSheetVisible
        shtStaticData.Visible = xlSheetVisible
        shtFileSpec.Visible = xlSheetVisible
    Else
        shtSysConf.Visible = xlSheetHidden
        shtStaticData.Visible = xlSheetHidden
        shtFileSpec.Visible = xlSheetHidden
    End If
    
    shtMenu.AutoFilterMode = False
    shtMenuCompInvt.AutoFilterMode = False
    fHideSheet shtDataStage

    fGetProgressBar
    gProBar.ChangeProcessBarValue 0.2, "ȥ�����й�����Ĺ�������"
    Call fRemoveFilterForAllSheets
    
    gProBar.ChangeProcessBarValue 0.25, "ɾ�����й�����������Ŀհ���"
    Call fDeleteBlankRowsFromAllSheets
    
    gProBar.ChangeProcessBarValue 0.28, "��������ҵ������"
    Call subMain_InvisibleHideAllBusinessSheets
    gProBar.ChangeProcessBarValue 0.4, "Ϊ���й��������������б��"
    Call fSetValidationListForAllSheets
    Call fSetValidationForNumberAndDateColumnsForAllSheets
    gProBar.ChangeProcessBarValue 0.5, "Ϊ���й��������ã���������ʽ"
    Call fSetConditionFormatForFundamentalSheets
    
    gProBar.ChangeProcessBarValue 0.6, "Ϊ���й����������Զ�ɸѡ"
    Call fAutoFileterAllSheets
        
'    Application.CommandBars("cell").FindControl(ID:=19).OnAction = "fGetCopyAddress"
'    Application.OnKey "^c", "fGetCopyAddress"
    shtDataStage.UsedRange.ClearComments
    shtDataStage.UsedRange.ClearContents
    shtDataStage.UsedRange.ClearFormats
    'shtDataStage.UsedRange.ClearHyperlinks
'    shtDataStage.UsedRange.ClearNotes
'    shtDataStage.UsedRange.ClearOutline
End Sub

Function fRefreshGetAllCommandbarsList()
    Dim arrCBarsInfo()
    Dim dict As Dictionary

    Debug.Print "fRefreshGetAllCommandbarsList " & vbTab & Now()

    arrCBarsInfo = fReadConfigCommandBarsInfo

    Erase arrAllCmdBarList

    Set dict = fReadArray2DictionaryOnlyKeys(arrCBarsInfo, 1, False, False)
    Call fCopyDictionaryKeys2Array(dict, arrAllCmdBarList)

    Set dict = Nothing
End Function

Function fGetThisWorkBookVariable(sVariable As String) As Variant
    Select Case sVariable
        Case "CMDBAR"
            If fArrayIsEmptyOrNoData(arrAllCmdBarList) Then
                Call fRefreshGetAllCommandbarsList
            End If

            fGetThisWorkBookVariable = arrAllCmdBarList
        Case Else
            fErr "wrong param"
    End Select
End Function


Private Sub Workbook_Activate()
    Call fEnableOrDisableAllCommandBarsByConfig(True)
End Sub
Private Sub Workbook_Deactivate()
    Call fEnableOrDisableAllCommandBarsByConfig(False)
    
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
   ' fGetRibbonReference.Invalidate
End Sub

Private Sub Workbook_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    Dim btn As CommandBarButton
    Dim cmdbar As CommandBar
    
    Set cmdbar = Application.CommandBars("cell")
    cmdbar.Reset
    
    Set btn = cmdbar.Controls.Add(msoControlButton)
    btn.Caption = "Active Sheet Code Pane"
    btn.OnAction = "subMain_ActiveSheetInfo"
    Set btn = Nothing
    
    
    Set btn = cmdbar.Controls.Add(msoControlButton)
    btn.Caption = "��ȡ��ǰ��ҵ����Ϣ"
    btn.OnAction = "subMain_GetCurrentRowBusinessInfo"
    Set btn = Nothing
    
    Set cmdbar = Nothing
End Sub
  

Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)
    'If Sh.Parent Is ThisWorkbook Then
        Call fAppendDataToLastCellOfColumn(shtDataStage, 2, Sh.Name)
    'End If
    'fGetRibbonReference.Invalidate
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    'fReGetValue_tgGetSearchBy
'    If tgGetSearchBy_Val Then
'        'fGetRibbonReference.InvalidateControl "tgGetSearchBy"
'        fPresstgGetSearchBy (tgGetSearchBy_Val)
'    End If
End Sub

Function fUpdateRangeAddressForName_rngStaticSalesCompanyNames_Comm()
    Dim asTag As String
    Dim arrColsName()
    Dim arrColsIndex()
    Dim rngToFindIn As Range
    Dim arrConfigData()
    Dim lConfigStartRow As Long
    Dim lConfigStartCol As Long
    Dim lConfigEndRow As Long
    Dim lConfigHeaderAtRow As Long

    asTag = "[Sales Company List - Common Importing - Sales File]"
    ReDim arrColsName(1 To 2)
    arrColsName(1) = "Company Name"
    arrColsName(2) = "SalesDate"
    
    Call fReadConfigBlockStartEnd(asTag, shtStaticData, lConfigStartRow, lConfigStartCol, lConfigEndRow)
    
    If lConfigEndRow < lConfigStartRow + 1 Then
        fErr "No data is configured under tag " & asTag & " in sheet " & shtStaticData.Name & vbCr & "You must leave at least one blank line after the tag."
    End If
     
    Set rngToFindIn = fGetRangeByStartEndPos(shtStaticData, lConfigStartRow, lConfigStartCol, lConfigEndRow, Columns.count)
    Call fFindAllColumnsIndexByColNames(rngToFindIn, arrColsName, arrColsIndex, lConfigHeaderAtRow)
    
    Dim lCol As Long
    lCol = arrColsIndex(1)
    Call fSetName("rngStaticSalesCompanyNames_Comm" _
                , "=" & fGetRangeByStartEndPos(shtStaticData, lConfigHeaderAtRow + 1, lCol, lConfigEndRow, lCol).Address(external:=True))
End Function


