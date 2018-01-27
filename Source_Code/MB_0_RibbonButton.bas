Attribute VB_Name = "MB_0_RibbonButton"
Option Explicit
Option Base 1



Sub subMain_ImportSalesCompanyInventory()
    fActiveVisibleSwitchSheet shtMenuCompInvt, "A63", False
End Sub
Sub subMain_Ribbon_ImportSalesInfoFiles()
    fActiveVisibleSwitchSheet shtMenu, "A63", False
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
    shtMenu.Visible = xlSheetVisible
    
    shtHospital.Visible = xlSheetVeryHidden
    shtHospitalReplace.Visible = xlSheetVeryHidden
    shtSalesRawDataRpt.Visible = xlSheetVeryHidden
    shtSalesInfos.Visible = xlSheetVeryHidden
    
    shtProductMaster.Visible = xlSheetVeryHidden
    shtProductNameReplace.Visible = xlSheetVeryHidden
    shtProductProducerReplace.Visible = xlSheetVeryHidden
    shtProductSeriesReplace.Visible = xlSheetVeryHidden
    shtProductUnitRatio.Visible = xlSheetVeryHidden
    shtProductProducerMaster.Visible = xlSheetVeryHidden
    shtProductNameMaster.Visible = xlSheetVeryHidden
    
    shtException.Visible = xlSheetVeryHidden
    shtProfit.Visible = xlSheetVeryHidden
    shtSelfSalesOrder.Visible = xlSheetVeryHidden
    shtSelfSalesPreDeduct.Visible = xlSheetVeryHidden
    shtSelfPurchaseOrder.Visible = xlSheetVeryHidden
    shtSalesManMaster.Visible = xlSheetVeryHidden
    shtFirstLevelCommission.Visible = xlSheetVeryHidden
    shtSecondLevelCommission.Visible = xlSheetVeryHidden
    shtSalesManCommConfig.Visible = xlSheetVeryHidden
    
    shtSelfInventory.Visible = xlSheetVeryHidden
    fVeryHideSheet shtMenuCompInvt
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
    'On Error GoTo Exit_Sub
    
    If Not shtHospital.fValidateSheet Then Exit Sub
    If Not shtProductMaster.fValidateSheet Then Exit Sub
    If Not shtProductNameMaster.fValidateSheet Then Exit Sub
    If Not shtProductProducerMaster.fValidateSheet Then Exit Sub
    If Not shtSalesManMaster.fValidateSheet Then Exit Sub
    If Not shtSalesManCommConfig.fValidateSheet Then Exit Sub
    
    If Not shtHospitalReplace.fValidateSheet Then Exit Sub
    If Not shtProductProducerReplace.fValidateSheet Then Exit Sub
    If Not shtProductNameReplace.fValidateSheet Then Exit Sub
    If Not shtProductSeriesReplace.fValidateSheet Then Exit Sub
    If Not shtProductUnitRatio.fValidateSheet Then Exit Sub
    
    If Not shtFirstLevelCommission.fValidateSheet Then Exit Sub
    If Not shtSecondLevelCommission.fValidateSheet Then Exit Sub
    
    If Not shtSelfSalesOrder.fValidateSheet Then Exit Sub
    If Not shtSelfPurchaseOrder.fValidateSheet Then Exit Sub
    
'Exit_Sub:
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
    
    lLastMaxRow = shtDataStage.Cells(Rows.Count, LAST_COL).End(xlUp).Row
    
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
    
    lPrevMaxRow = shtDataStage.Cells(Rows.Count, PREV_COL).End(xlUp).Row
    
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
    lMaxRow = sht.Cells(Rows.Count, alCol).End(xlUp).Row
    
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
    
    fInitialization

    Dim arrSource()
    Dim sOldFile As String
    Dim arrSheetsToMigr
    
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
    
        Call fConvertFomulaToValueForSheetIfAny(shtSource)
        Call fCopyReadWholeSheetData2Array(shtSource, arrSource)
        'arrSource = wbSource.shtProductMaster.UsedRange.Value2

        Call fWriteArray2Sheet(shtTargetEach, arrSource)
        Erase arrSource
    Next
    
    Call fCloseWorkBookWithoutSave(wbSource)
error_handling:
    Erase arrSource
    If Not wbSource Is Nothing Then Call fCloseWorkBookWithoutSave(wbSource)
    
    Application.AutomationSecurity = msoAutomationSecurityByUI
    
    If fCheckIfGotBusinessError Then Err.Clear
    If fCheckIfUnCapturedExceptionAbnormalError Then End
    
    MsgBox "done"
End Sub

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

Sub subMain_SelfInventory()
    Dim lMaxRow As Long
    lMaxRow = fGetValidMaxRow(shtSelfInventory)
    
    If lMaxRow > 2 Then
        With fGetRangeByStartEndPos(shtSelfInventory, 2, 1, lMaxRow, fGetValidMaxCol(shtSelfInventory))
            .ClearContents
            '.ClearFormats
            .ClearComments
            .ClearNotes
            .ClearOutline
        End With
    End If
    
    fCalculateSelfInventory
    fActiveVisibleSwitchSheet shtSelfInventory, , False
    
    fMsgBox "本公司库存计算完成！", vbInformation
End Sub
