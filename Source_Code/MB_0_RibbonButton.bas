Attribute VB_Name = "MB_0_RibbonButton"
Option Explicit
Option Base 1

Private Const DELETED_FROM_NEW_VERSION = "�����汾�ж��°汾��û��(��ɾ����)"
Private Const SAME_IN_BOTH = "�����汾����(��ͬ��)"
Private Const NEWLY_ADDED_IN_NEW_VERSION = "�°汾���ж������汾��û��(�����ӵ�)"
Private Const BOTH_HAVE_BUT_DIFF_VALUE = "�����汾���е�����ֵ��ͬ(���޸ĵ�)"

Sub subMain_NewRuleProducts()
    fActiveVisibleSwitchSheet shtNewRuleProducts, , False
End Sub

Sub subMain_ImportSalesCompanyInventory()
    fActiveVisibleSwitchSheet shtMenuCompInvt, "A63", False
End Sub
Sub subMain_Ribbon_ImportSalesInfoFiles()
    fActiveVisibleSwitchSheet shtMenu, "A74", False
End Sub

Sub subMain_Hospital()
    fActiveVisibleSwitchSheet shtHospital, , False
    'Call fHideAllSheetExcept(shtHospital, shtHospitalReplace)
End Sub

Sub subMain_HideHospital()
    On Error Resume Next
    shtHospital.Visible = xlSheetVeryHidden
    Err.Clear
End Sub
Sub subMain_HospitalReplacement()
    fActiveVisibleSwitchSheet shtHospitalReplace, , False
    'Call fHideAllSheetExcept(shtHospital, shtHospitalReplace)
End Sub

Sub subMain_Exception()
    fActiveVisibleSwitchSheet shtException, , False
    'Call fHideAllSheetExcept(shtHospital, shtHospitalReplace)
End Sub
Sub subMain_RawSalesInfos()
    fActiveVisibleSwitchSheet shtSalesRawDataRpt, , False
End Sub

Sub subMain_SalesInfos()
    fActiveVisibleSwitchSheet shtSalesInfos, , False
End Sub

Sub subMain_ProductMaster()
    fActiveVisibleSwitchSheet shtProductMaster, , False
    'Call fHideAllSheetExcept(shtProductMaster, shtProductProducerReplace, shtProductNameReplace, shtProductSeriesReplace, shtProductUnitRatio)
End Sub

Sub subMain_HideProductMaster()
    On Error Resume Next
    shtProductMaster.Visible = xlSheetVeryHidden
    Err.Clear
End Sub
Sub subMain_HideProducerMaster()
    On Error Resume Next
    shtProductProducerMaster.Visible = xlSheetVeryHidden
    Err.Clear
End Sub

Sub subMain_ProducerMaster()
    fActiveVisibleSwitchSheet shtProductProducerMaster, , False
    'Call fHideAllSheetExcept(shtProductMaster, shtProductProducerReplace, shtProductNameReplace, shtProductSeriesReplace, shtProductUnitRatio)
End Sub
Sub subMain_ProductNameMaster()
    fActiveVisibleSwitchSheet shtProductNameMaster, , False
    'Call fHideAllSheetExcept(shtProductMaster, shtProductProducerReplace, shtProductNameReplace, shtProductSeriesReplace, shtProductUnitRatio)
End Sub
Sub subMain_HideProductNameMaster()
    On Error Resume Next
    shtProductNameMaster.Visible = xlSheetVeryHidden
    Err.Clear
End Sub
Sub subMain_ProductProducerReplace()
    fActiveVisibleSwitchSheet shtProductProducerReplace, , False
    'Call fHideAllSheetExcept(shtProductMaster, shtProductProducerReplace, shtProductNameReplace, shtProductSeriesReplace, shtProductUnitRatio)
End Sub
Sub subMain_ProductNameReplace()
    fActiveVisibleSwitchSheet shtProductNameReplace, , False
    'Call fHideAllSheetExcept(shtProductMaster, shtProductProducerReplace, shtProductNameReplace, shtProductSeriesReplace, shtProductUnitRatio)
End Sub
Sub subMain_ProductSeriesReplace()
    fActiveVisibleSwitchSheet shtProductSeriesReplace, , False
    'Call fHideAllSheetExcept(shtProductMaster, shtProductProducerReplace, shtProductNameReplace, shtProductSeriesReplace, shtProductUnitRatio)
End Sub
Sub subMain_ProductUnitRatio()
    fActiveVisibleSwitchSheet shtProductUnitRatio, , False
    'Call fHideAllSheetExcept(shtProductMaster, shtProductProducerReplace, shtProductNameReplace, shtProductSeriesReplace, shtProductUnitRatio)
End Sub

Sub subMain_SalesMan()
    fActiveVisibleSwitchSheet shtSalesManMaster, , False
End Sub
Sub subMain_SalesManCommissionConfig()
    fActiveVisibleSwitchSheet shtSalesManCommConfig, , False
End Sub

Sub subMain_Profit()
    fActiveVisibleSwitchSheet shtProfit, , False
End Sub

Sub subMain_SelfSalesPreDeduct()
    fActiveVisibleSwitchSheet shtSelfSalesPreDeduct, , False
End Sub


Sub subMain_SelfPurchaseOrder()
    fActiveVisibleSwitchSheet shtSelfPurchaseOrder, , False
End Sub

Sub subMain_SelfSalesOrder()
    fActiveVisibleSwitchSheet shtSelfSalesOrder, , False
End Sub


Sub subMain_FirstLevelCommission()
    fActiveVisibleSwitchSheet shtFirstLevelCommission, , False
End Sub

Sub subMain_SecondLevelCommission()
    fActiveVisibleSwitchSheet shtSecondLevelCommission, , False
End Sub

Sub subMain_InvisibleHideAllBusinessSheets()
    fVeryHideSheet shtCompanyNameReplace
    fVeryHideSheet shtHospital
    fVeryHideSheet shtHospitalReplace
    fVeryHideSheet shtSalesRawDataRpt
    fVeryHideSheet shtSalesInfos
    fVeryHideSheet shtProductMaster
    fVeryHideSheet shtProductNameReplace
    fVeryHideSheet shtProductProducerReplace
    fVeryHideSheet shtProductSeriesReplace
    fVeryHideSheet shtProductUnitRatio
    fVeryHideSheet shtProductProducerMaster
    fVeryHideSheet shtProductNameMaster
    fVeryHideSheet shtException
    fVeryHideSheet shtProfit
    
    If Not gProBar Is Nothing Then gProBar.ChangeProcessBarValue 0.8, "��������ҵ������: shtProfit"
    
    fVeryHideSheet shtSelfSalesOrder
    
    If Not gProBar Is Nothing Then gProBar.ChangeProcessBarValue 0.8, "��������ҵ������: 1"
    
    fVeryHideSheet shtSelfSalesPreDeduct
    gProBar.ChangeProcessBarValue 0.8, "��������ҵ������: 2"
    fVeryHideSheet shtSelfPurchaseOrder
    gProBar.ChangeProcessBarValue 0.8, "��������ҵ������: 3"
    fVeryHideSheet shtSalesManMaster
    gProBar.ChangeProcessBarValue 0.8, "��������ҵ������: 4"
    fVeryHideSheet shtFirstLevelCommission
    gProBar.ChangeProcessBarValue 0.8, "��������ҵ������: 5"
    fVeryHideSheet shtSecondLevelCommission
    gProBar.ChangeProcessBarValue 0.8, "��������ҵ������: 6"
    fVeryHideSheet shtSalesManCommConfig
    gProBar.ChangeProcessBarValue 0.8, "��������ҵ������: 7"
    fVeryHideSheet shtSelfInventory
    gProBar.ChangeProcessBarValue 0.8, "��������ҵ������: 8"
    fVeryHideSheet shtMenuCompInvt
    gProBar.ChangeProcessBarValue 0.8, "��������ҵ������: 9"
    fVeryHideSheet shtMenu
    gProBar.ChangeProcessBarValue 0.8, "��������ҵ������: 10"
    fVeryHideSheet shtInventoryRawDataRpt
    gProBar.ChangeProcessBarValue 0.8, "��������ҵ������: 11"
    fVeryHideSheet shtImportCZL2SalesCompSales
    gProBar.ChangeProcessBarValue 0.8, "��������ҵ������: 12"
    fVeryHideSheet shtCZLSales2CompRawData
    gProBar.ChangeProcessBarValue 0.8, "��������ҵ������: 13"
    fVeryHideSheet shtCZLSales2Companies
    gProBar.ChangeProcessBarValue 0.8, "��������ҵ������: 14"
    fVeryHideSheet shtNewRuleProducts
    gProBar.ChangeProcessBarValue 0.8, "��������ҵ������: 15"
    fVeryHideSheet shtPV
    gProBar.ChangeProcessBarValue 0.8, "��������ҵ������: 16"
    fVeryHideSheet shtPromotionProduct
    gProBar.ChangeProcessBarValue 0.8, "��������ҵ������: 17"
    fVeryHideSheet shtCZLInvDiff
    gProBar.ChangeProcessBarValue 0.8, "��������ҵ������: 18"
    fVeryHideSheet shtCZLInventory
    gProBar.ChangeProcessBarValue 0.8, "��������ҵ������: 19"
    fVeryHideSheet shtCZLPurchaseOrder
    gProBar.ChangeProcessBarValue 0.8, "��������ҵ������: 20"
    'fVeryHideSheet shtCZLInformedInvInput
    fVeryHideSheet shtCZLRolloverInv
    gProBar.ChangeProcessBarValue 0.8, "��������ҵ������: 21"
    fVeryHideSheet shtSalesCompInvCalcd
    gProBar.ChangeProcessBarValue 0.8, "��������ҵ������: 22"
    fVeryHideSheet shtSalesCompInvUnified
    gProBar.ChangeProcessBarValue 0.8, "��������ҵ������: 23"
    fVeryHideSheet shtSalesCompRolloverInv
    gProBar.ChangeProcessBarValue 0.8, "��������ҵ������: 24"
    fVeryHideSheet shtSalesCompInvDiff
    gProBar.ChangeProcessBarValue 0.8, "��������ҵ������: 25"
    fVeryHideSheet shtProductTaxRate
    gProBar.ChangeProcessBarValue 0.8, "��������ҵ������: 26"
    fVeryHideSheet shtRefund
    gProBar.ChangeProcessBarValue 0.28, "��������ҵ������: 27"
    'fVeryHideSheet shtMenuRefund
    fVeryHideSheet shtCZLSales2SCompAll
    gProBar.ChangeProcessBarValue 0.8, "��������ҵ������: 28"
    fVeryHideSheet shtSellPriceInAdv
    gProBar.ChangeProcessBarValue 0.8, "��������ҵ������: 30"
    
    fShowSheet shtMainMenu
    gProBar.ChangeProcessBarValue 0.8, "��������ҵ������: 31"
    shtMainMenu.Activate
    
    gProBar.ChangeProcessBarValue 0.28, "��������ҵ������: done"
    'If Not mRibbonObj Is Nothing Then fGetRibbonReference.Invalidate
End Sub

Sub subMain_ShowAllBusinessSheets()
    fShowSheet shtCompanyNameReplace
    fShowSheet shtHospital
    fShowSheet shtHospitalReplace
    fShowSheet shtSalesRawDataRpt
    fShowSheet shtSalesInfos
    fShowSheet shtProductMaster
    fShowSheet shtProductNameReplace
    fShowSheet shtProductProducerReplace
    fShowSheet shtProductSeriesReplace
    fShowSheet shtProductUnitRatio
    fShowSheet shtProductProducerMaster
    fShowSheet shtProductNameMaster
    fShowSheet shtException
    fShowSheet shtProfit
    fShowSheet shtSelfSalesOrder
    fShowSheet shtSelfSalesPreDeduct
    fShowSheet shtSelfPurchaseOrder
    fShowSheet shtSalesManMaster
    fShowSheet shtFirstLevelCommission
    fShowSheet shtSecondLevelCommission
    fShowSheet shtSellPriceInAdv
    fShowSheet shtSalesManCommConfig
    fShowSheet shtSelfInventory
    fShowSheet shtMenuCompInvt
    fShowSheet shtMenu
    fShowSheet shtInventoryRawDataRpt
    'fShowSheet shtSalesCompInventory
    fShowSheet shtImportCZL2SalesCompSales
    fShowSheet shtCZLSales2CompRawData
    fShowSheet shtCZLSales2Companies
    fShowSheet shtPromotionProduct
    fShowSheet shtCZLInvDiff
    'fShowSheet shtCZLInformedInvInput
    fShowSheet shtCZLRolloverInv
    'fShowSheet shtSalesCompInv
    fShowSheet shtSalesCompRolloverInv
    fShowSheet shtProductTaxRate
    
    fShowSheet shtMainMenu
    shtMainMenu.Activate
End Sub

Function fActiveVisibleSwitchSheet(shtToSwitch As Worksheet, Optional sRngAddrToSelect As String = "A1" _
                            , Optional bHidePreviousActiveSheet As Boolean = False)
    Dim shtCurr As Worksheet
    Set shtCurr = ActiveSheet

    On Error Resume Next
    
    If shtToSwitch.Visible = xlSheetVisible Then
        If Not ActiveSheet Is shtToSwitch Then
            shtToSwitch.Visible = xlSheetVisible
            shtToSwitch.Activate
            Range(sRngAddrToSelect).Select
        Else
            shtToSwitch.Visible = xlSheetVeryHidden
        End If
    Else
        shtToSwitch.Visible = xlSheetVisible
        shtToSwitch.Activate
        Range(sRngAddrToSelect).Select
    End If

    If bHidePreviousActiveSheet Then
        If Not shtCurr Is shtToSwitch Then shtCurr.Visible = xlSheetVeryHidden
    End If

    Err.Clear
End Function

Function fShowActivateSheet(shtToSwitch As Worksheet, Optional sRngAddrToSelect As String = "A1" _
                            , Optional bHidePreviousActiveSheet As Boolean = False)
    Dim shtCurr As Worksheet
    Set shtCurr = ActiveSheet

    'On Error Resume Next
    
    If shtToSwitch.Visible <> xlSheetVisible Then shtToSwitch.Visible = xlSheetVisible
    
    shtToSwitch.Activate
    'Range(sRngAddrToSelect).Select

    If bHidePreviousActiveSheet Then
        If Not shtCurr Is shtToSwitch Then shtCurr.Visible = xlSheetVeryHidden
    End If

    Err.Clear
End Function
Function fShowAndActiveSheet(sht As Worksheet)
    sht.Visible = xlSheetVisible
    sht.Activate
End Function
'Function fActiveVisibleSwitchSheet(shtToSwitch As Worksheet, Optional sRngAddrToSelect As String = "A1")
'    Dim shtCurr As Worksheet
'    Set shtCurr = ActiveSheet
'
'    On Error Resume Next
'
'    If shtToSwitch.Visible = xlSheetVisible Then
'        If ActiveSheet Is shtToSwitch Then
'            shtToSwitch.Visible = xlSheetVisible
'            shtToSwitch.Activate
'            Range(sRngAddrToSelect).Select
'        Else
'            shtToSwitch.Visible = xlSheetVeryHidden
'        End If
'    Else
'        shtToSwitch.Visible = xlSheetVisible
'        shtToSwitch.Activate
'        Range(sRngAddrToSelect).Select
'    End If
'
'    If bHidePreviousActiveSheet Then
'        If Not shtCurr Is shtToSwitch Then shtCurr.Visible = xlSheetVeryHidden
'    End If
'
'    err.Clear
'End Function
Function fHideAllSheetExcept(ParamArray arr())
    Dim sht 'As Worksheet
    Dim shtConvt 'As Worksheet
    Dim wbSht 'As Worksheet
    
    On Error Resume Next
    
    For Each wbSht In ThisWorkbook.Worksheets
        For Each sht In arr
            Set shtConvt = sht
            If wbSht Is shtConvt Then
                'sht.Visible = xlSheetVisible
                GoTo next_wbsheet
            End If
        Next
        
        wbSht.Visible = xlSheetVeryHidden
next_wbsheet:
    Next
    
    Set shtConvt = Nothing
    Err.Clear
End Function

Sub subMain_ValidateAllSheetsData()
    On Error GoTo exit_sub
    
    fGetProgressBar
    gProBar.ShowBar
    gProBar.ChangeProcessBarValue 0.1
    If Not shtCompanyNameReplace.fValidateSheet(False) Then GoTo exit_sub
    If Not shtHospital.fValidateSheet(False) Then GoTo exit_sub
    If Not shtProductMaster.fValidateSheet(False) Then GoTo exit_sub
    gProBar.ChangeProcessBarValue 0.2
    If Not shtProductNameMaster.fValidateSheet(False) Then GoTo exit_sub
    If Not shtProductProducerMaster.fValidateSheet(False) Then GoTo exit_sub
    gProBar.ChangeProcessBarValue 0.3
    If Not shtSalesManMaster.fValidateSheet(False) Then GoTo exit_sub
    If Not shtSalesManCommConfig.fValidateSheet(False) Then GoTo exit_sub
    
    gProBar.ChangeProcessBarValue 0.4
    If Not shtNewRuleProducts.fValidateSheet(False) Then GoTo exit_sub
    
    If Not shtHospitalReplace.fValidateSheet(False) Then GoTo exit_sub
    gProBar.ChangeProcessBarValue 0.5
    If Not shtProductProducerReplace.fValidateSheet(False) Then GoTo exit_sub
    If Not shtProductNameReplace.fValidateSheet(False) Then GoTo exit_sub
    gProBar.ChangeProcessBarValue 0.7
    If Not shtProductSeriesReplace.fValidateSheet(False) Then GoTo exit_sub
    If Not shtProductUnitRatio.fValidateSheet(False) Then GoTo exit_sub
    gProBar.ChangeProcessBarValue 0.8
    If Not shtFirstLevelCommission.fValidateSheet(False) Then GoTo exit_sub
    If Not shtSecondLevelCommission.fValidateSheet(False) Then GoTo exit_sub
    gProBar.ChangeProcessBarValue 1
    If Not shtSelfPurchaseOrder.fValidateSheet(False) Then GoTo exit_sub
    If Not shtSelfSalesOrder.fValidateSheet(False) Then GoTo exit_sub
    If Not shtPromotionProduct.fValidateSheet(False) Then GoTo exit_sub
    If Not shtProductTaxRate.fValidateSheet(False) Then GoTo exit_sub
    If Not shtSellPriceInAdv.fValidateSheet(False) Then GoTo exit_sub
    
    gProBar.DestroyBar
    fMsgBox "û�з��ִ���", vbInformation
exit_sub:
    'If Err.Number <> 0 Then fMsgBox Err.Number
    gProBar.DestroyBar
End Sub

Sub subMain_BackToLastPosition()
    Dim sLastSheetName As String
    Dim lLastMaxRow As Long
    Dim lPrevMaxRow As Long
    Dim bFound As Boolean
    
    Const LAST_COL = 2
    Const PREV_COL = 3
    
    bFound = False
    On Error GoTo exit_sub
    
    Dim shtLast As Worksheet
    Dim lEachRow As Long
    
    lLastMaxRow = shtDataStage.Cells(Rows.count, LAST_COL).End(xlUp).Row
    
    For lEachRow = lLastMaxRow To 1 Step -1
        sLastSheetName = Trim(shtDataStage.Cells(lEachRow, LAST_COL).Value)
        shtDataStage.Cells(lEachRow, LAST_COL).ClearContents
        
        If fZero(sLastSheetName) Then GoTo previous_row
        
        If fSheetExists(sLastSheetName) Then
            Set shtLast = ThisWorkbook.Worksheets(sLastSheetName)
            
            If UCase(shtLast.Name) = UCase(ActiveSheet.Name) Then
                Call fAppendDataToLastCellOfColumn(shtDataStage, PREV_COL, sLastSheetName)
            Else
                If fSheetIsVisible(shtLast) Then
                    'Application.EnableEvents = False
                    shtLast.Activate
                    'Application.EnableEvents = True
                    bFound = True
                    Exit For
                End If
            End If
        End If
        
previous_row:
    Next
    
    If bFound Then
        Call fAppendDataToLastCellOfColumn(shtDataStage, PREV_COL, sLastSheetName)
    End If
    
exit_sub:
    Set shtLast = Nothing
    'Application.EnableEvents = True
End Sub

Sub subMain_BackToPreviousPosition()
    Dim sPrevSheetName As String
    Dim lPrevMaxRow As Long
    Dim lLastMaxRow As Long
    Dim bFound As Boolean
    
    Const LAST_COL = 2
    Const PREV_COL = 3
    
    bFound = False
    On Error GoTo exit_sub
    
    Dim shtPrev As Worksheet
    Dim lEachRow As Long
    
    lPrevMaxRow = shtDataStage.Cells(Rows.count, PREV_COL).End(xlUp).Row
    
    For lEachRow = lPrevMaxRow To 1 Step -1
        sPrevSheetName = Trim(shtDataStage.Cells(lEachRow, PREV_COL).Value)
        shtDataStage.Cells(lEachRow, PREV_COL).ClearContents
        
        If fZero(sPrevSheetName) Then GoTo previous_row
        
        If fSheetExists(sPrevSheetName) Then
            Set shtPrev = ThisWorkbook.Worksheets(sPrevSheetName)
            
            If UCase(shtPrev.Name) = UCase(ActiveSheet.Name) Then
                Call fAppendDataToLastCellOfColumn(shtDataStage, LAST_COL, sPrevSheetName)
            Else
                If fSheetIsVisible(shtPrev) Then
                    'Application.EnableEvents = False
                    shtPrev.Activate
                    'Application.EnableEvents = True
                    bFound = True
                    Exit For
                End If
            End If
        End If
        
previous_row:
    Next
    
    If bFound Then
        Call fAppendDataToLastCellOfColumn(shtDataStage, LAST_COL, sPrevSheetName)
    End If
    
exit_sub:
    Set shtPrev = Nothing
    'Application.EnableEvents = True
End Sub

Function fAppendDataToLastCellOfColumn(ByRef sht As Worksheet, alCol As Long, aValue)
    Dim lMaxRow As Long
    lMaxRow = sht.Cells(Rows.count, alCol).End(xlUp).Row
    
    If lMaxRow <= 1 Then
        If fZero(sht.Cells(lMaxRow, alCol).Value) Then
            sht.Cells(lMaxRow, alCol).Value = aValue
        Else
            sht.Cells(lMaxRow + 1, alCol).Value = aValue
        End If
    Else
        sht.Cells(lMaxRow + 1, alCol).Value = aValue
    End If
End Function

Sub Sub_DataMigration()
    On Error GoTo error_handling
    
    If Not fPromptToConfirmToContinue("��ȷ��Ҫ�������ļ�Ǩ�����ݽ�����?") Then fErr
    
    fInitialization

    Dim arrSource()
    Dim sOldFile As String
    Dim arrSheetsToMigr
    
    'to-do
    arrSheetsToMigr = Array(shtHospital _
                            , shtProductProducerMaster _
                            , shtProductNameMaster _
                            , shtProductMaster _
                            , shtSalesManMaster _
                            , shtHospitalReplace _
                            , shtProductProducerReplace _
                            , shtProductNameReplace _
                            , shtProductSeriesReplace _
                            , shtProductUnitRatio _
                            , shtSalesManCommConfig _
                            , shtSelfPurchaseOrder _
                            , shtSelfSalesOrder _
                            , shtFirstLevelCommission _
                            , shtSecondLevelCommission _
                            , shtNewRuleProducts _
                            , shtCompanyNameReplace _
                            , shtCZLRolloverInv _
                            , shtSalesCompRolloverInv _
                            , shtProductTaxRate _
                            , shtPromotionProduct _
                            , shtSellPriceInAdv _
                              )

    sOldFile = fSelectFileDialog(, "Macro File=*.xlsm", "Old Version With Latest User Data")
    If fZero(sOldFile) Then Exit Sub
    
    Call fIfExcelFileOpenedToCloseIt(sOldFile)
    
    Dim wbSource As Workbook
    Dim shtSource As Worksheet
    Dim eachSheet
    Dim shtTargetEach As Worksheet
    
    Application.AutomationSecurity = msoAutomationSecurityForceDisable

    Set wbSource = Workbooks.Open(Filename:=sOldFile, ReadOnly:=True)
    
    For Each eachSheet In arrSheetsToMigr
        Set shtTargetEach = eachSheet
        
        Set shtSource = fFindSheetBySheetCodeName(wbSource, shtTargetEach)
        Call fRemoveFilterForSheet(shtSource)
        
        Call fConvertFomulaToValueForSheetIfAny(shtSource)
        Call fCopyReadWholeSheetData2Array(shtSource, arrSource)
        'arrSource = wbSource.shtProductMaster.UsedRange.Value2
        Call fDeleteRemoveDataFormatFromSheetLeaveHeader(shtTargetEach)
        
        Call fWriteArray2Sheet(shtTargetEach, arrSource)
        
        If ArrLen(arrSource, 1) + 1 <> fGetValidMaxRow(shtTargetEach) Then
            fErr "ArrLen(arrSource, 1) + 1 <> fGetValidMaxRow(shtTargetEach): " & ArrLen(arrSource, 1) + 1 & " <> " & fGetValidMaxRow(shtTargetEach)
        End If
        
        Erase arrSource
    Next
    
    Call fCloseWorkBookWithoutSave(wbSource)
    MsgBox "done"
error_handling:
    If Err.Number <> gErrNum Then
        If Err.Number <> 0 Then MsgBox Err.Description
    End If
    
    Erase arrSource
    If Not wbSource Is Nothing Then Call fCloseWorkBookWithoutSave(wbSource)
    
    Application.AutomationSecurity = msoAutomationSecurityByUI
    
    If fCheckIfGotBusinessError Then Err.Clear
    If fCheckIfUnCapturedExceptionAbnormalError Then End
End Sub


Function fCompareDictionaryKeys(dictBase As Dictionary, dictThis As Dictionary) As Dictionary
    Dim dictOut As Dictionary
    Dim i As Long
    Dim sKey As String
    
    Set dictOut = New Dictionary
    
    'missed from right one
    For i = 0 To dictBase.count - 1
        sKey = dictBase.Keys(i)
        
        If Not dictThis.Exists(sKey) Then
            dictOut.Add DELETED_FROM_NEW_VERSION & DELIMITER & sKey, dictBase.Items(i) + 1
        Else
            'dictOut.Add SAME_IN_BOTH & DELIMITER & sKey, dictBase.Items(i) + 1 & DELIMITER & dictThis(sKey) + 1
            dictThis.Remove sKey
        End If
    Next
    
'    Dim iBlankColNum As Integer
'    If dictBase.Count > 0 Then iBlankColNum = UBound(Split(dictBase.Keys(0), DELIMITER)) - LBound(Split(dictBase.Keys(0), DELIMITER)) + 1
'    If dictThis <= 0 And dictThis.Count > 0 Then iBlankColNum = UBound(Split(dictThis.Keys(0), DELIMITER)) - LBound(Split(dictThis.Keys(0), DELIMITER)) + 1
    
    'missed from LEFT one
    For i = 0 To dictThis.count - 1
        sKey = dictThis.Keys(i)
        
        'If Not dictBase.Exists(sKey) Then
            'dictOut.Add "�°汾�ж������汾��û��" & String(DELIMITER, iBlankColNum) & sKey, dictThis.Items(i) + 1
            dictOut.Add NEWLY_ADDED_IN_NEW_VERSION & DELIMITER & sKey, dictThis.Items(i) + 1
        'End If
    Next
    
    Set fCompareDictionaryKeys = dictOut
    Set dictOut = Nothing
End Function

Function fCompareDictionaryKeysAndSingleItem(dictBase As Dictionary, dictThis As Dictionary) As Dictionary
    Dim dictOut As Dictionary
    Dim i As Long
    Dim sKey As String
    Dim sValue As String
    
    Set dictOut = New Dictionary
    
    'missed from right one
    For i = 0 To dictBase.count - 1
        sKey = dictBase.Keys(i)
        
        If Not dictThis.Exists(sKey) Then
            dictOut.Add DELETED_FROM_NEW_VERSION & DELIMITER & sKey, dictBase.Items(i) & DELIMITER & "�°汾��û������"
        Else
            If dictBase.Items(i) <> dictThis(sKey) Then
                dictOut.Add BOTH_HAVE_BUT_DIFF_VALUE & DELIMITER & sKey, dictBase.Items(i) & DELIMITER & dictThis(sKey)
            Else
                'dictOut.Add SAME_IN_BOTH & DELIMITER & sKey, dictBase.Items(i) & DELIMITER & dictThis(sKey)
            End If
            
            dictThis.Remove sKey
        End If
    Next
    
    'missed from LEFT one
    For i = 0 To dictThis.count - 1
        sKey = dictThis.Keys(i)
        
        'If Not dictBase.Exists(sKey) Then
            dictOut.Add NEWLY_ADDED_IN_NEW_VERSION & DELIMITER & sKey, dictThis.Items(i) & DELIMITER & "�����汾��û������"
        'End If
    Next
    
    Set fCompareDictionaryKeysAndSingleItem = dictOut
    Set dictOut = Nothing
End Function

Function fCompareDictionaryKeysAndMultipleItems(ByRef dictBase As Dictionary, ByRef dictThis As Dictionary) As Dictionary
    Dim dictOut As Dictionary
    Dim i As Long
    Dim sKey As String
    Dim sValue As String
    
    Set dictOut = New Dictionary
    
    'missed from right one
    For i = 0 To dictBase.count - 1
        sKey = dictBase.Keys(i)
        
        If Not dictThis.Exists(sKey) Then
            dictOut.Add DELETED_FROM_NEW_VERSION & DELIMITER & sKey, dictBase.Items(i) & vbLf & "�°汾��û������"
        Else
            If dictBase.Items(i) <> dictThis(sKey) Then
                dictOut.Add BOTH_HAVE_BUT_DIFF_VALUE & DELIMITER & sKey, dictBase.Items(i) & vbLf & dictThis(sKey)
            Else
                'dictOut.Add SAME_IN_BOTH & DELIMITER & sKey, dictBase.Items(i) & vbLf & dictThis(sKey)
            End If
            dictThis.Remove sKey
        End If
    Next
    
    'missed from LEFT one
    For i = 0 To dictThis.count - 1
        sKey = dictThis.Keys(i)
        
        'If Not dictBase.Exists(sKey) Then
            dictOut.Add NEWLY_ADDED_IN_NEW_VERSION & DELIMITER & sKey, dictThis.Items(i) & vbLf & "�����汾��û������"
        'End If
    Next
    
    Set fCompareDictionaryKeysAndMultipleItems = dictOut
    Set dictOut = Nothing
End Function
Function fFindSheetBySheetCodeName(wb As Workbook, shtToMatch As Worksheet) As Worksheet
    Dim shtMatched As Worksheet
    
    Dim shtEach As Worksheet
    
    For Each shtEach In wb.Worksheets
        If shtEach.CodeName = shtToMatch.CodeName Then
            Set shtMatched = shtEach
            Exit For
        End If
    Next
    
    If shtMatched Is Nothing Then fErr shtToMatch.CodeName & " cannot be found in the opened macro file."
    Set fFindSheetBySheetCodeName = shtMatched
    Set shtMatched = Nothing
End Function

Function fAutoFileterAllSheets()
    fResetAutoFilter shtCompanyNameReplace
    fResetAutoFilter shtHospital
    fResetAutoFilter shtHospitalReplace
    fResetAutoFilter shtSalesRawDataRpt
    fResetAutoFilter shtSalesInfos
    fResetAutoFilter shtProductMaster
    fResetAutoFilter shtProductNameReplace
    fResetAutoFilter shtProductProducerReplace
    fResetAutoFilter shtProductSeriesReplace
    fResetAutoFilter shtProductUnitRatio
    fResetAutoFilter shtProductProducerMaster
    fResetAutoFilter shtProductNameMaster
    fResetAutoFilter shtProfit
    fResetAutoFilter shtSelfSalesOrder
    fResetAutoFilter shtSelfSalesPreDeduct
    fResetAutoFilter shtSelfPurchaseOrder
    fResetAutoFilter shtSalesManMaster
    fResetAutoFilter shtFirstLevelCommission
    fResetAutoFilter shtSecondLevelCommission
    fResetAutoFilter shtSellPriceInAdv
    fResetAutoFilter shtSalesManCommConfig
    fResetAutoFilter shtSelfInventory
    fResetAutoFilter shtInventoryRawDataRpt
    fResetAutoFilter shtImportCZL2SalesCompSales
    fResetAutoFilter shtCZLSales2CompRawData
    fResetAutoFilter shtCZLSales2Companies
    fResetAutoFilter shtCZLInvDiff
    fResetAutoFilter shtPromotionProduct
    fResetAutoFilter shtSalesCompInvUnified
    fResetAutoFilter shtSalesCompInvCalcd
    fResetAutoFilter shtSalesCompInvDiff
    fResetAutoFilter shtProductTaxRate
    fResetAutoFilter shtRefund
End Function

Function fResetAutoFilter(sht As Worksheet)
    sht.Rows(1).AutoFilter
    sht.Rows(1).AutoFilter
End Function


Sub subMain_RefreshAllPvTables()
    ThisWorkbook.RefreshAll
    fShowAndActiveSheet shtPV
End Sub

Sub subMain_InvisibleHideCurrentSheet()
    If shtMainMenu.CodeName = ActiveSheet.CodeName Then Exit Sub
    
    'If ThisWorkbook.Worksheets.Count > 1 Then
        fVeryHideSheet ActiveSheet
    'End If
End Sub

Function fGetReplaceUnifyErrorRowCount_SCompSalesInfo() As Long
    fGetReplaceUnifyErrorRowCount_SCompSalesInfo = CLng(fGetSpecifiedConfigCellValue(shtSysConf, "[Facility For Testing]", "Value", "Setting Item ID=REPLACE_UNIFY_ERR_ROW_COUNT_SALES_INFO"))
End Function
Function fSetReplaceUnifyErrorRowCount_SCompSalesInfo(ByVal rowCnt As Long) As Long
    Call fSetSpecifiedConfigCellValue(shtSysConf, "[Facility For Testing]", "Value", "Setting Item ID=REPLACE_UNIFY_ERR_ROW_COUNT_SALES_INFO", CStr(rowCnt))
End Function

Function fGetReplaceUnifyErrorRowCount_SalesInventory() As Long
    fGetReplaceUnifyErrorRowCount_SalesInventory = CLng(fGetSpecifiedConfigCellValue(shtSysConf, "[Facility For Testing]", "Value", "Setting Item ID=REPLACE_UNIFY_ERR_ROW_COUNT_COMPNAY_INVENTORY"))
End Function
Function fSetReplaceUnifyErrorRowCount_SCompInventory(ByVal rowCnt As Long) As Long
    Call fSetSpecifiedConfigCellValue(shtSysConf, "[Facility For Testing]", "Value", "Setting Item ID=REPLACE_UNIFY_ERR_ROW_COUNT_COMPNAY_INVENTORY", CStr(rowCnt))
End Function

Function fGetReplaceUnifyErrorRowCount_CZLSales2Comp() As Long
    fGetReplaceUnifyErrorRowCount_CZLSales2Comp = CLng(fGetSpecifiedConfigCellValue(shtSysConf, "[Facility For Testing]", "Value", "Setting Item ID=REPLACE_UNIFY_ERR_ROW_COUNT_CZL_SALES_2_COMPANIES"))
End Function
Function fSetReplaceUnifyErrorRowCount_CZLSales2Comp(ByVal rowCnt As Long) As Long
    Call fSetSpecifiedConfigCellValue(shtSysConf, "[Facility For Testing]", "Value", "Setting Item ID=REPLACE_UNIFY_ERR_ROW_COUNT_CZL_SALES_2_COMPANIES", CStr(rowCnt))
End Function


Function fCheckIfErrCountNotZero_SCompSalesInfo()
    Dim iErr As Long
    iErr = fGetReplaceUnifyErrorRowCount_SCompSalesInfo
    If iErr <> 0 Then
        subMain_InvisibleHideAllBusinessSheets
            fShowSheet shtSalesRawDataRpt
            fShowSheet shtSalesInfos
        
        If iErr = 100 Then
            fShowSheet shtSalesRawDataRpt
            fShowSheet shtSalesInfos
            fErr "ԭʼ��������û�����滻�������滻ͳһԭʼ��������"
        Else    'If iErr = 999 Then
            fShowAndActiveSheet shtException
            fErr "����������������ҩƷ��ϵͳ���Ҳ������޷����������Ӷ�����ȴ�����Щ����"
'        Else
'            fShowSheet shtSalesRawDataRpt
'            fShowSheet shtSalesInfos
'            fErr "REPLACE_UNIFY_ERR_ROW_COUNT_SALES_INFO = " & iErr & ", but it was not covered in fCheckIfErrCountNotZero_SCompSalesInfo"
        End If
    End If
End Function

Function fCheckIfErrCountNotZero_CZLSales2Comp()
    Dim iErr As Long
    iErr = fGetReplaceUnifyErrorRowCount_CZLSales2Comp
    
    If iErr <> 0 Then
        subMain_InvisibleHideAllBusinessSheets
            fShowSheet shtCZLSales2CompRawData
            fShowSheet shtCZLSales2Companies
        
        If iErr = 100 Then
            fShowSheet shtCZLSales2CompRawData
            fShowSheet shtCZLSales2Companies
            fErr "��֥�ֵ�ԭʼ������������(����ҵ��˾)��û�����滻�������滻ͳһԭʼ��������"
        Else    'If iErr = 999 Then
            fShowAndActiveSheet shtException
            fErr "��֥�ֵ���������(����ҵ��˾)����ҩƷ��ϵͳ���Ҳ������޷������棬���ȴ�����Щ����"
'        Else
'            fShowSheet shtCZLSales2CompRawData
'            fShowSheet shtCZLSales2Companies
'            fErr "REPLACE_UNIFY_ERR_ROW_COUNT_CZL_SALES_2_COMPANIES = " & iErr & ", but it was not covered in fCheckIfErrCountNotZero_CZLSales2Comp"
        End If
    End If
End Function

Function fCheckIfErrCountNotZero_SalesInventory()
    Dim iErr As Long
    iErr = fGetReplaceUnifyErrorRowCount_SalesInventory
    
    If iErr <> 0 Then
        subMain_InvisibleHideAllBusinessSheets
            fShowSheet shtInventoryRawDataRpt
            fShowSheet shtSalesCompInvUnified
        
        If iErr = 100 Then
            fShowSheet shtInventoryRawDataRpt
            fShowSheet shtSalesCompInvUnified
            fErr "��ҵ��˾(��֥�ֵ�)��ԭʼ������ݻ�û�����滻�������滻ͳһ��"
        Else    'If iErr = 999 Then
            fShowAndActiveSheet shtException
            fErr "��ҵ��˾(��֥�ֵ�)�Ŀ����������ҩƷ��ϵͳ���Ҳ������޷����п��˶ԣ����ȴ�����Щ����"
'        Else
'            fShowSheet shtInventoryRawDataRpt
'            fShowSheet shtSalesCompInvUnified
'            fErr "REPLACE_UNIFY_ERR_ROW_COUNT_COMPNAY_INVENTORY = " & iErr & ", but it was not covered in fCheckIfErrCountNotZero_SalesInventory"
        End If
    End If
End Function


Sub subMain_CompareChangeWithPrevVersion()
    Dim arrBase()
    Dim arrThis()
    Dim arrDiff()
    Dim sBaseVersion As String
    Dim arrBaseVersion
    Dim wbBase As Workbook
    Dim shtBase As Worksheet
    Dim eachSheet
    Dim shtTargetEach As Worksheet
    Dim dictBase As Dictionary
    Dim dictThis As Dictionary
    Dim dictDiff As Dictionary
    Dim wbOutput As Workbook
    Dim shtOutput As Worksheet
    
    If Not fIsDev() Then On Error GoTo error_handling
    
    fInitialization
    
    'to-do
    arrBaseVersion = Array(shtHospital _
                        , shtProductProducerMaster _
                        , shtProductNameMaster _
                        , shtProductMaster _
                        , shtSalesManMaster _
                        , shtHospitalReplace _
                        , shtProductProducerReplace _
                        , shtProductNameReplace _
                        , shtProductSeriesReplace _
                        , shtProductUnitRatio _
                        , shtSalesManCommConfig _
                        , shtSelfPurchaseOrder _
                        , shtSelfSalesOrder _
                        , shtFirstLevelCommission _
                        , shtSecondLevelCommission _
                        , shtNewRuleProducts _
                        , shtCompanyNameReplace _
                        , shtCZLRolloverInv _
                        , shtSalesCompRolloverInv _
                        , shtProductTaxRate _
                        , shtPromotionProduct _
                        )

    sBaseVersion = fSelectFileDialog(, "���=*.xlsm", "��ѡ����бȽϵĻ����汾")
    If fZero(sBaseVersion) Then Exit Sub
    
    If sBaseVersion = ThisWorkbook.FullName Then fErr "��ѡ���˱��������������ѡ��"
    
    Call fIfExcelFileOpenedToCloseIt(sBaseVersion)
    
    Application.AutomationSecurity = msoAutomationSecurityForceDisable

    Set wbBase = Workbooks.Open(Filename:=sBaseVersion, ReadOnly:=True)
    
    For Each eachSheet In arrBaseVersion
        Set shtTargetEach = eachSheet
        
        fStartTimer
   
        Set shtBase = fFindSheetBySheetCodeName(wbBase, shtTargetEach)
        Call fRemoveFilterForSheet(shtBase)
        
        Call fConvertFomulaToValueForSheetIfAny(shtBase)
        
        If wbOutput Is Nothing Then
            Application.SheetsInNewWorkbook = 1
            Set wbOutput = Workbooks.Add(xlWBATWorksheet)
            wbOutput.Worksheets(wbOutput.Worksheets.count).Name = "Temp"
        Else
            'wbOutput.Worksheets.Add after:=wbOutput.Worksheets(wbOutput.Worksheets.Count)
        End If
            
        wbOutput.Worksheets.Add after:=wbOutput.Worksheets(wbOutput.Worksheets.count)
        Set shtOutput = wbOutput.Worksheets(wbOutput.Worksheets.count)
        shtOutput.Name = shtBase.Name & "_�ȽϽ��"
            
        Call fCopyReadWholeSheetData2Array(shtBase, arrBase)
        Call fCopyReadWholeSheetData2Array(shtTargetEach, arrThis)
            
        If shtTargetEach.CodeName = "shtHospital" Then
            Call fPrepareHeaderToSheet(shtOutput, Array("���ݱ�־", "ҽԺ����", "�����汾���к�", "", "", "�°汾���к�"), 1)
            
            '========================================
            Set dictBase = fReadArray2DictionaryWithRowNum(arrBase, 1, True, False)
            Set dictThis = fReadArray2DictionaryWithRowNum(arrThis, 1, True, False)
            Set dictDiff = fCompareDictionaryKeys(dictBase, dictThis)
            arrDiff = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictDiff, , False)
            Call fAppendArray2Sheet(shtOutput, arrDiff, , False)
            '========================================

            If dictDiff.count > 0 Then _
            shtOutput.Cells(2, UBound(arrDiff, 2) + 1).Resize(dictDiff.count, 1).Value = fConvertDictionaryItemsTo2DimenArrayForPaste(dictDiff, False)
        ElseIf shtTargetEach.CodeName = "shtSalesManMaster" Then
            Call fPrepareHeaderToSheet(shtOutput, Array("���ݱ�־", "ҵ��Ա����", "�����汾 ����", " �°汾�о���", "", "", ""), 1)
            
            '========================================
            Set dictBase = fReadArray2DictionaryWithSingleCol(arrBase, 1, 2, True, False)
            Set dictThis = fReadArray2DictionaryWithSingleCol(arrThis, 1, 2, True, False)
            Set dictDiff = fCompareDictionaryKeysAndSingleItem(dictBase, dictThis)
            arrDiff = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictDiff, , False)
            Call fAppendArray2Sheet(shtOutput, arrDiff, , False)
            '========================================
            
            If dictDiff.count > 0 Then _
            shtOutput.Cells(2, UBound(arrDiff, 2) + 1).Resize(dictDiff.count, 2).Value = fConvertDictionaryDelimiteredItemsTo2DimenArrayForPaste(dictDiff, , False)
        ElseIf shtTargetEach.CodeName = "shtProductProducerMaster" Then
            Call fPrepareHeaderToSheet(shtOutput, Array("���ݱ�־", "ҩƷ����", "�����汾���к�", "", "", "�°汾���к�"), 1)
            
            '========================================
            Set dictBase = fReadArray2DictionaryWithRowNum(arrBase, 1, True, False)
            Set dictThis = fReadArray2DictionaryWithRowNum(arrThis, 1, True, False)
            Set dictDiff = fCompareDictionaryKeys(dictBase, dictThis)
            arrDiff = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictDiff, , False)
            Call fAppendArray2Sheet(shtOutput, arrDiff, , False)
            '========================================

            If dictDiff.count > 0 Then _
            shtOutput.Cells(2, UBound(arrDiff, 2) + 1).Resize(dictDiff.count, 1).Value = fConvertDictionaryItemsTo2DimenArrayForPaste(dictDiff, False)
        ElseIf shtTargetEach.CodeName = "shtProductNameMaster" Then
            Call fPrepareHeaderToSheet(shtOutput, Array("���ݱ�־", "ҩƷ����", "ҩƷ����", "�����汾���к�", "", "", "�°汾���к�"), 1)
            
            '========================================
            Set dictBase = fReadArray2DictionaryMultipleKeysWithRowNum(arrBase, Array(1, 2), DELIMITER, True, False)
            Set dictThis = fReadArray2DictionaryMultipleKeysWithRowNum(arrThis, Array(1, 2), DELIMITER, True, False)
            Set dictDiff = fCompareDictionaryKeys(dictBase, dictThis)
            arrDiff = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictDiff, , False)
            Call fAppendArray2Sheet(shtOutput, arrDiff, , False)
            '========================================

            If dictDiff.count > 0 Then _
            shtOutput.Cells(2, UBound(arrDiff, 2) + 1).Resize(dictDiff.count, 1).Value = fConvertDictionaryItemsTo2DimenArrayForPaste(dictDiff, False)
        ElseIf shtTargetEach.CodeName = "shtHospitalReplace" Then
            Call fPrepareHeaderToSheet(shtOutput, Array("���ݱ�־", "ҽԺ����", "�����汾 �滻Ϊ", "�°汾�� �滻Ϊ", "", "", ""), 1)
            
            '========================================
            Set dictBase = fReadArray2DictionaryWithSingleCol(arrBase, 1, 2, True, False)
            Set dictThis = fReadArray2DictionaryWithSingleCol(arrThis, 1, 2, True, False)
            Set dictDiff = fCompareDictionaryKeysAndSingleItem(dictBase, dictThis)
            arrDiff = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictDiff, , False)
            Call fAppendArray2Sheet(shtOutput, arrDiff, , False)
            '========================================

            If dictDiff.count > 0 Then _
            shtOutput.Cells(2, UBound(arrDiff, 2) + 1).Resize(dictDiff.count, 2).Value = fConvertDictionaryDelimiteredItemsTo2DimenArrayForPaste(dictDiff, , False)
        ElseIf shtTargetEach.CodeName = "shtCompanyNameReplace" Then
            Call fPrepareHeaderToSheet(shtOutput, Array("���ݱ�־", "ԭʼ�ļ���ҵ��˾����", "�����汾 �滻Ϊ", "�°汾�� �滻Ϊ", "", "", ""), 1)
            
            '========================================
            Set dictBase = fReadArray2DictionaryWithSingleCol(arrBase, 1, 2, True, False)
            Set dictThis = fReadArray2DictionaryWithSingleCol(arrThis, 1, 2, True, False)
            Set dictDiff = fCompareDictionaryKeysAndSingleItem(dictBase, dictThis)
            arrDiff = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictDiff, , False)
            Call fAppendArray2Sheet(shtOutput, arrDiff, , False)
            '========================================

            If dictDiff.count > 0 Then _
            shtOutput.Cells(2, UBound(arrDiff, 2) + 1).Resize(dictDiff.count, 2).Value = fConvertDictionaryDelimiteredItemsTo2DimenArrayForPaste(dictDiff, , False)
        ElseIf shtTargetEach.CodeName = "shtProductMaster" Then
            Call fPrepareHeaderToSheet(shtOutput, Array("���ݱ�־", "ҩƷ����", "ҩƷ����", "ҩƷ���", "�����汾���к�", "", "", "�°汾���к�"), 1)
            
            '========================================
            Set dictBase = fReadArray2DictionaryMultipleKeysWithRowNum(arrBase, Array(1, 2, 3), DELIMITER, True, False)
            Set dictThis = fReadArray2DictionaryMultipleKeysWithRowNum(arrThis, Array(1, 2, 3), DELIMITER, True, False)
            Set dictDiff = fCompareDictionaryKeys(dictBase, dictThis)
            arrDiff = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictDiff, , False)
            Call fAppendArray2Sheet(shtOutput, arrDiff, , False)
            '========================================

            If dictDiff.count > 0 Then _
            shtOutput.Cells(2, UBound(arrDiff, 2) + 1).Resize(dictDiff.count, 1).Value = fConvertDictionaryItemsTo2DimenArrayForPaste(dictDiff, False)
        ElseIf shtTargetEach.CodeName = "shtProductProducerReplace" Then
            Call fPrepareHeaderToSheet(shtOutput, Array("���ݱ�־", "ԭʼҩƷ����", "�����汾 �滻Ϊ", "�°汾�� �滻Ϊ", "", "", ""), 1)
            
            '========================================
            Set dictBase = fReadArray2DictionaryWithSingleCol(arrBase, 1, 2, True, False)
            Set dictThis = fReadArray2DictionaryWithSingleCol(arrThis, 1, 2, True, False)
            Set dictDiff = fCompareDictionaryKeysAndSingleItem(dictBase, dictThis)
            arrDiff = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictDiff, , False)
            Call fAppendArray2Sheet(shtOutput, arrDiff, , False)
            '========================================

            If dictDiff.count > 0 Then _
            shtOutput.Cells(2, UBound(arrDiff, 2) + 1).Resize(dictDiff.count, 2).Value = fConvertDictionaryDelimiteredItemsTo2DimenArrayForPaste(dictDiff, , False)
        ElseIf shtTargetEach.CodeName = "shtProductNameReplace" Then
            Call fPrepareHeaderToSheet(shtOutput, Array("���ݱ�־", "ҩƷ����", "ԭʼҩƷ����", "�����汾 �滻Ϊ", "�°汾�� �滻Ϊ", "", "", ""), 1)
            
            '========================================
            Set dictBase = fReadArray2DictionaryWithMultipleKeyColsSingleItemCol(arrBase, Array(1, 2), 3, DELIMITER, True, False)
            Set dictThis = fReadArray2DictionaryWithMultipleKeyColsSingleItemCol(arrThis, Array(1, 2), 3, DELIMITER, True, False)
            Set dictDiff = fCompareDictionaryKeysAndSingleItem(dictBase, dictThis)
            arrDiff = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictDiff, , False)
            Call fAppendArray2Sheet(shtOutput, arrDiff, , False)
            '========================================

            If dictDiff.count > 0 Then _
            shtOutput.Cells(2, UBound(arrDiff, 2) + 1).Resize(dictDiff.count, 2).Value = fConvertDictionaryDelimiteredItemsTo2DimenArrayForPaste(dictDiff, , False)
        ElseIf shtTargetEach.CodeName = "shtProductSeriesReplace" Then
            Call fPrepareHeaderToSheet(shtOutput, Array("���ݱ�־", "ҩƷ����", "ҩƷ����", "ԭʼҩƷ���", "�����汾 �滻Ϊ", "�°汾�� �滻Ϊ", "", "", ""), 1)
            
            '========================================
            Set dictBase = fReadArray2DictionaryWithMultipleKeyColsSingleItemCol(arrBase, Array(1, 2, 3), 4, DELIMITER, True, False)
            Set dictThis = fReadArray2DictionaryWithMultipleKeyColsSingleItemCol(arrThis, Array(1, 2, 3), 4, DELIMITER, True, False)
            Set dictDiff = fCompareDictionaryKeysAndSingleItem(dictBase, dictThis)
            arrDiff = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictDiff, , False)
            Call fAppendArray2Sheet(shtOutput, arrDiff, , False)
            '========================================

            If dictDiff.count > 0 Then _
            shtOutput.Cells(2, UBound(arrDiff, 2) + 1).Resize(dictDiff.count, 2).Value = fConvertDictionaryDelimiteredItemsTo2DimenArrayForPaste(dictDiff, , False)
        ElseIf shtTargetEach.CodeName = "shtProductUnitRatio" Then
            Call fPrepareHeaderToSheet(shtOutput, Array("���ݱ�־", "ҩƷ����", "ҩƷ����", "ҩƷ���", "ͳһ��λ", "ԭʼ��λ", "�����汾 ����", "�°汾�� ����", "", "", ""), 1)
            
            '========================================
            Set dictBase = fReadArray2DictionaryWithMultipleKeyColsSingleItemCol(arrBase, Array(1, 2, 3, 4, 6), 5, DELIMITER, True, False)
            Set dictThis = fReadArray2DictionaryWithMultipleKeyColsSingleItemCol(arrThis, Array(1, 2, 3, 4, 6), 5, DELIMITER, True, False)
            Set dictDiff = fCompareDictionaryKeysAndSingleItem(dictBase, dictThis)
            arrDiff = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictDiff, , False)
            Call fAppendArray2Sheet(shtOutput, arrDiff, , False)
            '========================================

            If dictDiff.count > 0 Then _
            shtOutput.Cells(2, UBound(arrDiff, 2) + 1).Resize(dictDiff.count, 2).Value = fConvertDictionaryDelimiteredItemsTo2DimenArrayForPaste(dictDiff, , False)
        ElseIf shtTargetEach.CodeName = "shtProductTaxRate" Then
            Call fPrepareHeaderToSheet(shtOutput, Array("���ݱ�־", "ҩƷ����", "ҩƷ����", "ҩƷ���", "�����汾 ˰��", "�°汾�� ˰��", "", "", ""), 1)
            
            '========================================
            Set dictBase = fReadArray2DictionaryWithMultipleKeyColsSingleItemCol(arrBase, Array(1, 2, 3), 4, DELIMITER, True, False)
            Set dictThis = fReadArray2DictionaryWithMultipleKeyColsSingleItemCol(arrThis, Array(1, 2, 3), 4, DELIMITER, True, False)
            Set dictDiff = fCompareDictionaryKeysAndSingleItem(dictBase, dictThis)
            arrDiff = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictDiff, , False)
            Call fAppendArray2Sheet(shtOutput, arrDiff, , False)
            '========================================
            
            If dictDiff.count > 0 Then _
            shtOutput.Cells(2, UBound(arrDiff, 2) + 1).Resize(dictDiff.count, 2).Value = fConvertDictionaryDelimiteredItemsTo2DimenArrayForPaste(dictDiff, , False)
        ElseIf shtTargetEach.CodeName = "shtSalesManCommConfig" Then
            Call fPrepareHeaderToSheet(shtOutput, Array("���ݱ�־", "��ҵ��˾", "ҽԺ", "ҩƷ��������", "ҩƷ����", "���", "�б��", "ҵ��Ա1|Ӷ��1|ҵ��Ա2|Ӷ��2|ҵ��Ա3|Ӷ��3|����������|��������ɱ���", "", "", ""), 1)
            
            '========================================
            Set dictBase = fReadArray2DictionaryMultipleKeysWithMultipleColsCombined(arrBase, Array(1, 2, 3, 4, 5, 6), Array(7, 8, 9, 10, 11, 12, 13, 14), DELIMITER, DELIMITER, True, False)
            Set dictThis = fReadArray2DictionaryMultipleKeysWithMultipleColsCombined(arrThis, Array(1, 2, 3, 4, 5, 6), Array(7, 8, 9, 10, 11, 12, 13, 14), DELIMITER, DELIMITER, True, False)
            Set dictDiff = fCompareDictionaryKeysAndMultipleItems(dictBase, dictThis)
            arrDiff = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictDiff, , False)
            Call fAppendArray2Sheet(shtOutput, arrDiff, , False)
            '========================================

            If dictDiff.count > 0 Then _
            shtOutput.Cells(2, UBound(arrDiff, 2) + 1).Resize(dictDiff.count, 1).Value = fTranspose1DimenArrayTo2DimenArrayVertically(dictDiff.Items) 'fConvertDictionaryDelimiteredItemsTo2DimenArrayForPaste(dictDiff, vbLf, False)
        ElseIf shtTargetEach.CodeName = "shtSelfPurchaseOrder" Then
            Call fPrepareHeaderToSheet(shtOutput, Array("���ݱ�־", "ҩƷ��������", "ҩƷ����", "���", "��λ", "��������", "����", "��������|��������", "", "", ""), 1)
            
            '========================================
            Set dictBase = fReadArray2DictionaryMultipleKeysWithMultipleColsCombined(arrBase, Array(1, 2, 3, 4, 5, 8), Array(6, 7), DELIMITER, DELIMITER, True, False)
            Set dictThis = fReadArray2DictionaryMultipleKeysWithMultipleColsCombined(arrThis, Array(1, 2, 3, 4, 5, 8), Array(6, 7), DELIMITER, DELIMITER, True, False)
            Set dictDiff = fCompareDictionaryKeysAndMultipleItems(dictBase, dictThis)
            arrDiff = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictDiff, , False)
            Call fAppendArray2Sheet(shtOutput, arrDiff, , False)
            '========================================

            If dictDiff.count > 0 Then _
            shtOutput.Cells(2, UBound(arrDiff, 2) + 1).Resize(dictDiff.count, 1).Value = fTranspose1DimenArrayTo2DimenArrayVertically(dictDiff.Items) 'fConvertDictionaryDelimiteredItemsTo2DimenArrayForPaste(dictDiff, vbLf, False)
        ElseIf shtTargetEach.CodeName = "shtSelfSalesOrder" Then
            Call fPrepareHeaderToSheet(shtOutput, Array("���ݱ�־", "ҩƷ��������", "ҩƷ����", "���", "��λ", "��������", "����", "��������|��������|ҽԺ���۵�������", "", "", ""), 1)
            
            '========================================
            Set dictBase = fReadArray2DictionaryMultipleKeysWithMultipleColsCombined(arrBase, Array(1, 2, 3, 4, 5, 8), Array(6, 7, 9), DELIMITER, DELIMITER, True, False)
            Set dictThis = fReadArray2DictionaryMultipleKeysWithMultipleColsCombined(arrThis, Array(1, 2, 3, 4, 5, 8), Array(6, 7, 9), DELIMITER, DELIMITER, True, False)
            Set dictDiff = fCompareDictionaryKeysAndMultipleItems(dictBase, dictThis)
            arrDiff = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictDiff, , False)
            Call fAppendArray2Sheet(shtOutput, arrDiff, , False)
            '========================================

            If dictDiff.count > 0 Then _
            shtOutput.Cells(2, UBound(arrDiff, 2) + 1).Resize(dictDiff.count, 1).Value = fTranspose1DimenArrayTo2DimenArrayVertically(dictDiff.Items) 'fConvertDictionaryDelimiteredItemsTo2DimenArrayForPaste(dictDiff, vbLf, False)
        ElseIf shtTargetEach.CodeName = "shtFirstLevelCommission" Then
            Call fPrepareHeaderToSheet(shtOutput, Array("���ݱ�־", "��ҵ��˾", "ҩƷ����", "ҩƷ����", "ҩƷ���", "�����汾 ���ͷ�", "�°汾�� ���ͷ�", "", "", ""), 1)
            
            '========================================
            Set dictBase = fReadArray2DictionaryWithMultipleKeyColsSingleItemCol(arrBase, Array(1, 2, 3, 4), 5, DELIMITER, True, False)
            Set dictThis = fReadArray2DictionaryWithMultipleKeyColsSingleItemCol(arrThis, Array(1, 2, 3, 4), 5, DELIMITER, True, False)
            Set dictDiff = fCompareDictionaryKeysAndSingleItem(dictBase, dictThis)
            arrDiff = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictDiff, , False)
            Call fAppendArray2Sheet(shtOutput, arrDiff, , False)
            '========================================

            If dictDiff.count > 0 Then _
            shtOutput.Cells(2, UBound(arrDiff, 2) + 1).Resize(dictDiff.count, 2).Value = fConvertDictionaryDelimiteredItemsTo2DimenArrayForPaste(dictDiff, , False)
        ElseIf shtTargetEach.CodeName = "shtSecondLevelCommission" Then
            Call fPrepareHeaderToSheet(shtOutput, Array("���ݱ�־", "��ҵ��˾", "ҽԺ", "ҩƷ����", "ҩƷ����", "ҩƷ���", "�����汾 ���ͷ�", "�°汾�� ���ͷ�", "", "", ""), 1)
            
            '========================================
            Set dictBase = fReadArray2DictionaryMultipleKeysWithMultipleColsCombined(arrBase, Array(1, 2, 3, 4, 5), Array(6, 7), DELIMITER, True, False)
            Set dictThis = fReadArray2DictionaryMultipleKeysWithMultipleColsCombined(arrThis, Array(1, 2, 3, 4, 5), Array(6, 7), DELIMITER, True, False)
            Set dictDiff = fCompareDictionaryKeysAndMultipleItems(dictBase, dictThis)
            arrDiff = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictDiff, , False)
            Call fAppendArray2Sheet(shtOutput, arrDiff, , False)
            '========================================

            If dictDiff.count > 0 Then _
            shtOutput.Cells(2, UBound(arrDiff, 2) + 1).Resize(dictDiff.count, 2).Value = fConvertDictionaryDelimiteredItemsTo2DimenArrayForPaste(dictDiff, , False)
        ElseIf shtTargetEach.CodeName = "shtNewRuleProducts" Then
            Call fPrepareHeaderToSheet(shtOutput, Array("���ݱ�־", "ҩƷ��������", "ҩƷ����", "���", "����˰���� | ����˰����", "", "", ""), 1)
            
            '========================================
            Set dictBase = fReadArray2DictionaryMultipleKeysWithMultipleColsCombined(arrBase, Array(1, 2, 3), Array(4, 5), DELIMITER, DELIMITER, True, False)
            Set dictThis = fReadArray2DictionaryMultipleKeysWithMultipleColsCombined(arrThis, Array(1, 2, 3), Array(4, 5), DELIMITER, DELIMITER, True, False)
            Set dictDiff = fCompareDictionaryKeysAndMultipleItems(dictBase, dictThis)
            arrDiff = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictDiff, , False)
            Call fAppendArray2Sheet(shtOutput, arrDiff, , False)
            '========================================

            If dictDiff.count > 0 Then _
            shtOutput.Cells(2, UBound(arrDiff, 2) + 1).Resize(dictDiff.count, 1).Value = fTranspose1DimenArrayTo2DimenArrayVertically(dictDiff.Items) 'fConvertDictionaryDelimiteredItemsTo2DimenArrayForPaste(dictDiff, vbLf, False)
        ElseIf shtTargetEach.CodeName = "shtPromotionProduct" Then
            Call fPrepareHeaderToSheet(shtOutput, Array("���ݱ�־", "ҽԺ", "ҩƷ����", "ҩƷ����", "ҩƷ���", "�б��", "�����汾 ����", "�°汾�� ����", "", "", ""), 1)
            
            '========================================
            Set dictBase = fReadArray2DictionaryMultipleKeysWithMultipleColsCombined(arrBase, Array(1, 2, 3, 4, 5, 9), Array(6, 7, 8, 10), DELIMITER, True, False)
            Set dictThis = fReadArray2DictionaryMultipleKeysWithMultipleColsCombined(arrThis, Array(1, 2, 3, 4, 5, 9), Array(6, 7, 8, 10), DELIMITER, True, False)
            Set dictDiff = fCompareDictionaryKeysAndMultipleItems(dictBase, dictThis)
            arrDiff = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictDiff, , False)
            Call fAppendArray2Sheet(shtOutput, arrDiff, , False)
            '========================================

            If dictDiff.count > 0 Then _
            shtOutput.Cells(2, UBound(arrDiff, 2) + 1).Resize(dictDiff.count, 2).Value = fConvertDictionaryDelimiteredItemsTo2DimenArrayForPaste(dictDiff, , False)
        ElseIf shtTargetEach.CodeName = "shtCZLRolloverInv" Then
            Call fPrepareHeaderToSheet(shtOutput, Array("���ݱ�־", "ҩƷ����", "ҩƷ����", "ҩƷ���", "��λ", "����", "�����汾 �ڳ��������", "�°汾�� �ڳ��������", "", "", ""), 1)
            
            '========================================
            Set dictBase = fReadArray2DictionaryWithMultipleKeyColsSingleItemCol(arrBase, Array(1, 2, 3, 4, 5), 6, DELIMITER, True, False)
            Set dictThis = fReadArray2DictionaryWithMultipleKeyColsSingleItemCol(arrThis, Array(1, 2, 3, 4, 5), 6, DELIMITER, True, False)
            Set dictDiff = fCompareDictionaryKeysAndSingleItem(dictBase, dictThis)
            arrDiff = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictDiff, , False)
            Call fAppendArray2Sheet(shtOutput, arrDiff, , False)
            '========================================

            If dictDiff.count > 0 Then _
            shtOutput.Cells(2, UBound(arrDiff, 2) + 1).Resize(dictDiff.count, 2).Value = fConvertDictionaryDelimiteredItemsTo2DimenArrayForPaste(dictDiff, , False)
        ElseIf shtTargetEach.CodeName = "shtSalesCompRolloverInv" Then
            Call fPrepareHeaderToSheet(shtOutput, Array("���ݱ�־", "��ҵ��˾", "ҩƷ����", "ҩƷ����", "ҩƷ���", "��λ", "����", "�����汾 �ڳ��������", "�°汾�� �ڳ��������", "", "", ""), 1)
            
            '========================================
            Set dictBase = fReadArray2DictionaryWithMultipleKeyColsSingleItemCol(arrBase, Array(1, 2, 3, 4, 5, 6), 7, DELIMITER, True, False)
            Set dictThis = fReadArray2DictionaryWithMultipleKeyColsSingleItemCol(arrThis, Array(1, 2, 3, 4, 5, 6), 7, DELIMITER, True, False)
            Set dictDiff = fCompareDictionaryKeysAndSingleItem(dictBase, dictThis)
            arrDiff = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictDiff, , False)
            Call fAppendArray2Sheet(shtOutput, arrDiff, , False)
            '========================================

            If dictDiff.count > 0 Then _
            shtOutput.Cells(2, UBound(arrDiff, 2) + 1).Resize(dictDiff.count, 2).Value = fConvertDictionaryDelimiteredItemsTo2DimenArrayForPaste(dictDiff, , False)
        End If
        
        Erase arrBase
        Erase arrThis
        Set dictBase = Nothing: Set dictThis = Nothing
        Erase arrDiff
        
        If dictDiff.count <= 0 Then
            fDeleteSheet shtOutput.Name, wbOutput
        Else
            shtOutput.Rows(1).Font.Color = RGB(255, 0, 0)
            shtOutput.Rows(1).Font.Bold = True
            
            fHowLong shtTargetEach.Name
            Call fFreezeSheet(shtOutput)
            fAutoFilterAutoFitSheet shtOutput
            fSortDataInSheetSortSheetData shtOutput, Array(1, 2, 3, 4)
           ' If shtOutput.AutoFilterMode Then shtOutput.UsedRange.AutoFilter Field:=1, Criteria1:="<>" & SAME_IN_BOTH, Operator:=xlAnd
    '        Erase arrSource
        End If
    Next
    
    Call fCloseWorkBookWithoutSave(wbBase)
    wbOutput.Activate
error_handling:
    If Err.Number <> 0 Then MsgBox Err.Description
     
    If Not wbBase Is Nothing Then Call fCloseWorkBookWithoutSave(wbBase)
    
    Application.AutomationSecurity = msoAutomationSecurityByUI
    
    If fCheckIfGotBusinessError Then Err.Clear
    If fCheckIfUnCapturedExceptionAbnormalError Then End
    
    If wbOutput.Worksheets.count > 1 Then
        fDeleteSheet "Temp", wbOutput
        MsgBox "�ԱȽ�������鿴�����"
    Else
       ' wbOutput.Worksheets("Temp").Name = "��ǰһ�汾���û�з��ֲ�ͬ"
        Call fCloseWorkBookWithoutSave(wbOutput)
        MsgBox "�ԱȽ�������ǰһ�汾���û�з��ֲ�ͬ��"
    End If
End Sub

