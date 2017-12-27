Attribute VB_Name = "MB_0_RibbonButton"
Option Explicit
Option Base 1
  
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
End Sub

Function fActiveVisibleSwitchSheet(shtToSwitch As Worksheet, Optional sRngAddrToSelect As String = "A1" _
                            , Optional bHidePreviousActiveSheet As Boolean = True)
    Dim shtCurr As Worksheet
    Set shtCurr = ActiveSheet

    'On Error Resume Next
    
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
