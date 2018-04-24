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
'Option Explicit
Option Base 1

Dim arrAllCmdBarList()

Function fIfSomeFundamentalSheetsWereDeleted() As Boolean
    Dim sht As Worksheet
    
    On Error Resume Next
    
    Set sht = Sheet4
    fIfSomeFundamentalSheetsWereDeleted = (sht Is Nothing)
    
End Function

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
    If fIfSomeFundamentalSheetsWereDeleted() Then Cancel = False

End Sub

Private Sub Workbook_Open()
    Application.OnKey "^{BACKSPACE}", "Sub_ToHomeSheet"
    fGetProgressBar
    gProBar.ShowBar
    gsEnv = fGetEnvFromSysConf

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



'Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
'    Application.EnableEvents = False
'
'    On Error GoTo exit_function
'
'    If Sh Is shtSysConf Then
'        Call fShtSysConf_SheetChange_DevProdChange(Target)
'        'Call shtSysConf_SheetChange_CommandBarConfig(Target)
'    End If
'
'exit_function:
'    Application.EnableEvents = True
'End Sub
'
'Private Sub Workbook_SheetActivate(ByVal Sh As Object)
'End Sub

Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)
    'If Sh.Parent Is ThisWorkbook Then
        Call fAppendDataToLastCellOfColumn(shtDataStage, 2, Sh.Name)
    'End If
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    'If Sh.CodeName = "shtSalesCompInvDiff" Then
    fReGetValue_tgSearchBy
        If tgSearchBy_Val Then
            'fReGetRibbonReference.InvalidateControl "tgSearchBy"
            fPresstgSearchBy (tgSearchBy_Val)
        End If
    'End If
End Sub
