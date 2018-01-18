VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' version : 201712 - 001 lost the validate blank -- 23:45 -- 0£º11
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
    
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If fIfSomeFundamentalSheetsWereDeleted() Then Cancel = False

End Sub

Private Sub Workbook_Open()
    gsEnv = fGetEnvFromSysConf

    Call fRefreshGetAllCommandbarsList
    Call sub_WorkBookInitialization
    Call fSetIntialValueForShtMenuInitialize

    Call sub_RemoveCommandBar("Team")
    ThisWorkbook.Saved = True
    ThisWorkbook.CheckCompatibility = False
    
'    Application.CommandBars("cell").FindControl(ID:=19).OnAction = ""
'    Application.OnKey "^c", ""
    End
End Sub

Sub sub_WorkBookInitialization()
    Call fReadConfigRibbonCommandBarMenuAndCreateCommandBarButton
    
    If fIsDev() Then
        shtSysConf.Visible = xlSheetVisible
        shtStaticData.Visible = xlSheetVisible
        shtFileSpec.Visible = xlSheetVisible
    Else
        shtSysConf.Visible = xlSheetVeryHidden
        shtStaticData.Visible = xlSheetVeryHidden
        shtFileSpec.Visible = xlSheetVeryHidden
    End If
    
    fHideSheet shtDataStage

    Call fRemoveFilterForAllSheets
    Call fDeleteBlankRowsFromAllSheets
    
    Call subMain_InvisibleHideAllBusinessSheets
    Call fSetValidationListForAllSheets
    Call fSetConditionFormatForFundamentalSheets
    
        
'    Application.CommandBars("cell").FindControl(ID:=19).OnAction = "fGetCopyAddress"
'    Application.OnKey "^c", "fGetCopyAddress"
    shtDataStage.UsedRange.ClearComments
    shtDataStage.UsedRange.ClearContents
    shtDataStage.UsedRange.ClearFormats
    shtDataStage.UsedRange.ClearHyperlinks
    shtDataStage.UsedRange.ClearNotes
    shtDataStage.UsedRange.ClearOutline
    
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



Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Application.EnableEvents = False

    On Error GoTo exit_function

    If Sh Is shtSysConf Then
        Call fShtSysConf_SheetChange_DevProdChange(Target)
        'Call shtSysConf_SheetChange_CommandBarConfig(Target)
    End If

exit_function:
    Application.EnableEvents = True
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
End Sub
Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)
    'If Sh.Parent Is ThisWorkbook Then
        Call fAppendDataToLastCellOfColumn(shtDataStage, 2, Sh.Name)
    'End If
End Sub
