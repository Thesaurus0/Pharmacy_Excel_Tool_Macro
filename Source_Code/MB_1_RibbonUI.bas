Attribute VB_Name = "MB_1_RibbonUI"
Option Explicit

#If VBA7 And Win64 Then  'Win64
    Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
#Else
    Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
#End If

Public mRibbonObj As IRibbonUI
Public ebSalesCompany_val As String
Public ebProductProducer_val As String
Public ebProductName_val As String
Public ebProductSeries_val As String
Public ebLotNum_val As String
Public ebHospital_val As String

Private arrGallery()
Private sSelectedGalleryId As String

'=============================================================
Sub subRefreshRibbon()
    Erase arrGallery
    sSelectedGalleryId = ""
    fGetRibbonReference.Invalidate
End Sub
Sub subActivateRibbonTab()
    fGetRibbonReference.ActivateTab "ERP_2010"
End Sub
Sub ERP_UI_Onload(ribbon As IRibbonUI)
  Set mRibbonObj = ribbon
  
  fSetName "nmRibbonPointer", ObjPtr(ribbon)
  'Names("nmRibbonPointer").RefersTo = ObjPtr(ribbon)
  
  mRibbonObj.ActivateTab "ERP_2010"
  ThisWorkbook.Saved = True
  Call fPrepareGallery
End Sub
Function fGetRibbonReference() As IRibbonUI
    If Not mRibbonObj Is Nothing Then Set fGetRibbonReference = mRibbonObj: Exit Function
    
    Dim objRibbon As Object
    Dim lRibPointer As LongPtr
    
    lRibPointer = [nmRibbonPointer]
    
    CopyMemory objRibbon, lRibPointer, LenB(lRibPointer)
    
    Set fGetRibbonReference = objRibbon
    Set mRibbonObj = objRibbon
    Set objRibbon = Nothing
End Function
'---------------------------------------------------------------------
Private Function fPrepareGallery()
    Dim i As Integer
    
    ReDim arrGallery(1 To 15, 1 To 3)
     
    i = i + 1
    arrGallery(i, 1) = "gal_Profit1":         arrGallery(i, 2) = "xx利润表":           i = i + 1
    arrGallery(i, 1) = "gal_Profit2":         arrGallery(i, 2) = "xx利润表":           i = i + 1
    arrGallery(i, 1) = "gal_Profit3":         arrGallery(i, 2) = "xx利润表":           i = i + 1
    arrGallery(i, 1) = "gal_Profit4":         arrGallery(i, 2) = "xx利润表":           i = i + 1
    arrGallery(i, 1) = "gal_Profit5":         arrGallery(i, 2) = "xx利润表":           i = i + 1
    arrGallery(i, 1) = "gal_Profit":          arrGallery(i, 2) = "利润表":              i = i + 1
    arrGallery(i, 1) = "gal_Profit6":         arrGallery(i, 2) = "xx利润表":           i = i + 1
    arrGallery(i, 1) = "gal_Profit7":         arrGallery(i, 2) = "xx利润表":           i = i + 1
    arrGallery(i, 1) = "gal_Profit8":         arrGallery(i, 2) = "xx利润表":           i = i + 1
    arrGallery(i, 1) = "gal_Profit9":         arrGallery(i, 2) = "xx利润表":           i = i + 1
End Function
Sub Gallery_GetLabel(control As IRibbonControl, ByRef label)
    label = "检索表"
End Sub
'Callback for GalSearchInSheet getImage
Sub Gallery_getImage(control As IRibbonControl, ByRef returnedVal)
    returnedVal = "ControlSubFormReport"
End Sub
Sub Gallery_getSize(control As IRibbonControl, ByRef size)
    size = False
End Sub
Sub Gallery_GetItemCount(control As IRibbonControl, ByRef count)
    count = 15
End Sub
Sub Gallery_getItemHeight(control As IRibbonControl, ByRef height)
    height = 20
End Sub
Sub Gallery_GetItemID(control As IRibbonControl, index As Integer, ByRef id)
    If fArrayIsEmpty(arrGallery) Then Call fPrepareGallery
    id = arrGallery(index + 1, 1)
End Sub
Sub Gallery_GetItemImage(control As IRibbonControl, index As Integer, ByRef image)
    image = "ControlSubFormReport"
End Sub
Sub Gallery_GetItemLabel(control As IRibbonControl, index As Integer, ByRef label)
    If fArrayIsEmpty(arrGallery) Then Call fPrepareGallery
    label = arrGallery(index + 1, 2)
End Sub
Sub Gallery_getItemWidth(control As IRibbonControl, ByRef width)
    width = 40
End Sub
'Sub Gallery_GetSelectedItemID(control As IRibbonControl, ByRef index)
''    If Len(sSelectedGalleryId) > 0 Then
''        index = sSelectedGalleryId
''    Else
''        index = "gal_Profit9"
''    End If
'End Sub
 
Sub Gallery_OnAction(control As IRibbonControl, selectedId As String, selectedIndex As Integer)
    'MsgBox selectedId & vbCr & selectedIndex
    Select Case UCase(selectedId)
        Case UCase("gal_Profit")
            Call Sub_SearchOnCurrentSheetBySearchBy(shtProfit)
        Case Else
            fMsgBox selectedId & " not covered in Gallery_OnAction"
    End Select
    
    sSelectedGalleryId = selectedId
End Sub
'======================================================================
Sub Button_onAction(control As IRibbonControl)
    Call fGetControlAttributes(control, "ACTION")
End Sub
Sub Button_getImage(control As IRibbonControl, ByRef imageMso)
    Call fGetControlAttributes(control, "IMAGE", imageMso)
End Sub
Sub Button_getLabel(control As IRibbonControl, ByRef label)
    Call fGetControlAttributes(control, "LABEL", label)
End Sub
Sub Button_getSize(control As IRibbonControl, ByRef size)
    Call fGetControlAttributes(control, "SIZE", size)
End Sub

'================== toggle button common function===========================================
Sub ToggleButtonToSwitchSheet_onAction(control As IRibbonControl, pressed As Boolean)
    Dim sht As Worksheet
    Set sht = fGetSheetByUIRibbonTag(control.Tag)
    
    If Not sht Is Nothing Then
        fToggleSheetVisibleFromUIRibbonControl pressed, sht, control
    End If
    Set sht = Nothing
End Sub

Sub ToggleButtonToSwitchSheet_getPressed(control As IRibbonControl, ByRef returnedVal)
    Dim sht As Worksheet
    Set sht = fGetSheetByUIRibbonTag(control.Tag)
    
    If sht Is Nothing Then
        returnedVal = False
    Else
        returnedVal = (sht.Visible = xlSheetVisible And ActiveSheet Is sht)
    End If
End Sub
Function fGetSheetByUIRibbonTag(ByVal asButtonTag As String) As Worksheet
    Dim sht As Worksheet
    
    If fSheetExistsByCodeName(asButtonTag, sht) Then
        Set fGetSheetByUIRibbonTag = sht
    Else
        MsgBox "The button's Tag is not corresponding to any worksheet in this workbook, please check the customUI.xml you prepared," _
            & " The design thought is that the button's tag is the name of a sheet, so that the common function ToggleButtonToSwitchSheet_onAction/getPressed can get a worksheet."
    End If
    Set sht = Nothing
End Function
Function fToggleSheetVisibleFromUIRibbonControl(ByVal pressed As Boolean, sht As Worksheet, control As IRibbonControl)
    If pressed Then
        If ActiveSheet.CodeName <> sht.CodeName Then
            fActiveVisibleSwitchSheet sht
        End If
    Else
        If ActiveSheet.CodeName <> sht.CodeName Then
            fActiveVisibleSwitchSheet sht
        Else
            fVeryHideSheet sht
        End If
    End If
    
    'fGetRibbonReference.InvalidateControl (control.id)
    fGetRibbonReference.Invalidate
End Function

'---------------------------------------------------------------------

'==========================dev prod switch===================================
Sub btnSwitchDevProd_onAction(control As IRibbonControl, pressed As Boolean)
    sub_SwitchDevProdMode
End Sub

Sub btnSwitchDevProd_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = fIsDev()
End Sub
Sub btnSwitchDevProd_getVisible(control As IRibbonControl, ByRef returnedVal)
    'returnedVal = fIsDev()
    returnedVal = True
End Sub
Sub grpDevFacilities_getVisible(control As IRibbonControl, ByRef returnedVal)
    returnedVal = fIsDev()
End Sub
'---------------------------------------------------------------------

'================ dev facilities ==============================================
Sub btnListAllFunctions_onAction(control As IRibbonControl)
    sub_ListAllFunctionsOfThisWorkbook
End Sub
Sub btnExportSourceCode_onAction(control As IRibbonControl)
    sub_ExportModulesSourceCodeToFolder
End Sub
Sub btnGenNumberList_onAction(control As IRibbonControl)
    sub_GenNumberList
End Sub
Sub btnGenAlphabetList_onAction(control As IRibbonControl)
    sub_GenAlpabetList
End Sub
Sub btnListAllActiveXOnCurrSheet_onAction(control As IRibbonControl)
    Sub_ListActiveXControlOnActiveSheet
End Sub
Sub btnResetOnError_onAction(control As IRibbonControl)
    sub_ResetOnError_Initialize
End Sub
'------------------------------------------------------------------------------

Sub EditBox_onChange(control As IRibbonControl, text As String)
    Select Case control.id
        Case "ebSalesCompany"
            ebSalesCompany_val = text
            shtDataStage.Range("K1").Value = ebSalesCompany_val
        Case "ebProductProducer"
            ebProductProducer_val = text
            shtDataStage.Range("K2").Value = ebProductProducer_val
        Case "ebProductName"
            ebProductName_val = text
            shtDataStage.Range("K3").Value = ebProductName_val
        Case "ebProductSeries"
            ebProductSeries_val = text
            shtDataStage.Range("K4").Value = ebProductSeries_val
        Case "ebLotnum"
            ebLotNum_val = text
            shtDataStage.Range("K5").Value = ebLotNum_val
        Case "ebHospital"
            ebHospital_val = text
            shtDataStage.Range("K6").Value = ebHospital_val
        Case Else
    End Select
     
End Sub

Function fRefreshEditBoxFromShtDataStage()
    If Len(ebProductProducer_val) <= 0 And Len(ebProductName_val) <= 0 And Len(ebProductSeries_val) <= 0 _
    And Len(ebSalesCompany_val) <= 0 And Len(ebLotNum_val) <= 0 And Len(ebHospital_val) <= 0 Then
        Debug.Print "fRefreshEditBoxFromShtDataStage : all becomes blank."
        
        ebSalesCompany_val = Trim(shtDataStage.Range("K1").Value)
        ebProductProducer_val = Trim(shtDataStage.Range("K2").Value)
        ebProductName_val = Trim(shtDataStage.Range("K3").Value)
        ebProductSeries_val = Trim(shtDataStage.Range("K4").Value)
        ebLotNum_val = Trim(shtDataStage.Range("K5").Value)
        ebHospital_val = Trim(shtDataStage.Range("K6").Value)
    End If
End Function

Private Sub EditBox_getText(control As IRibbonControl, ByRef returnedVal)
    Dim iColIndex As Integer
    
    Select Case control.id
        Case "ebHospital"
            Call fGetColIndexs(ActiveSheet, iColIndex)
            If iColIndex > 0 Then
                ebHospital_val = Trim(ActiveSheet.Cells(ActiveCell.Row, iColIndex).Value)
            End If
            returnedVal = ebHospital_val
            shtDataStage.Range("K6").Value = ebHospital_val
          '  Call fSetName("nmHospital", ebHospital_val)
           ' Debug.Print ebHospital_val
        Case "ebSalesCompany"
            Call fGetColIndexs(ActiveSheet, , iColIndex)
            If iColIndex > 0 Then
                ebSalesCompany_val = Trim(ActiveSheet.Cells(ActiveCell.Row, iColIndex).Value)
            End If
             
            returnedVal = ebSalesCompany_val
            shtDataStage.Range("K1").Value = ebSalesCompany_val
         '   Call fSetName("nmSalesCompany", ebSalesCompany_val)
          '  Debug.Print ebSalesCompany_val
        Case "ebProductProducer"
            Call fGetColIndexs(ActiveSheet, , , iColIndex)
            If iColIndex > 0 Then
                ebProductProducer_val = Trim(ActiveSheet.Cells(ActiveCell.Row, iColIndex).Value)
            End If
            returnedVal = ebProductProducer_val
            shtDataStage.Range("K2").Value = ebProductProducer_val
          '  Call fSetName("nmProductProducer", ebProductProducer_val)
         '   Debug.Print ebProductProducer_val
        Case "ebProductName"
            Call fGetColIndexs(ActiveSheet, , , , iColIndex)
            If iColIndex > 0 Then
                ebProductName_val = Trim(ActiveSheet.Cells(ActiveCell.Row, iColIndex).Value)
            End If
            returnedVal = ebProductName_val
            shtDataStage.Range("K3").Value = ebProductName_val
          '  Call fSetName("nmProductName", ebProductName_val)
          '  Debug.Print ebProductName_val
        Case "ebProductSeries"
            Call fGetColIndexs(ActiveSheet, , , , , iColIndex)
            If iColIndex > 0 Then
                ebProductSeries_val = Trim(ActiveSheet.Cells(ActiveCell.Row, iColIndex).Value)
            End If
            returnedVal = ebProductSeries_val
            shtDataStage.Range("K4").Value = ebProductSeries_val
          '  Call fSetName("nmProductSeries", ebProductSeries_val)
          '  Debug.Print ebProductSeries_val
        Case "ebLotnum"
            Call fGetColIndexs(ActiveSheet, , , , , , iColIndex)
            If iColIndex > 0 Then
                ebLotNum_val = Trim(ActiveSheet.Cells(ActiveCell.Row, iColIndex).Value)
            End If
            returnedVal = ebLotNum_val
            shtDataStage.Range("K5").Value = ebLotNum_val
           ' Call fSetName("nmLotNum", ebLotNum_val)
           ' Debug.Print ebLotNum_val
        Case Else
            fMsgBox "not covered in EditBox_getText: " & control.id
    End Select
End Sub

'Function fPresstgGetSearchBy(bPressed As Boolean)
Sub sub_PresstgGetSearchBy()
    If ActiveCell.Row <= 1 Then Exit Sub
    
    If ActiveSheet.CodeName = "shtMainMenu" Or ActiveSheet.CodeName = "shtDataStage" Or ActiveSheet.CodeName = "shtStaticData" _
    Or ActiveSheet.CodeName = "shtFileSpec" Or ActiveSheet.CodeName = "shtSysConf" Or ActiveSheet.CodeName = "" _
    Then
        fMsgBox "当前页没有业务数据!"
        Exit Sub
    End If
    
    If Not fGetColIndexs(ActiveSheet) Then
        fMsgBox "当前页还未设置该功能, 请联系开发人员添加此功能."
         Exit Sub
    End If
    
    If ActiveCell.Row > fGetValidMaxRow(ActiveSheet) Then fMsgBox "请先选中一行":  Exit Sub
     
    fGetRibbonReference.InvalidateControl "ebSalesCompany"
    fGetRibbonReference.InvalidateControl "ebProductProducer"
    fGetRibbonReference.InvalidateControl "ebProductName"
    fGetRibbonReference.InvalidateControl "ebProductSeries"
    fGetRibbonReference.InvalidateControl "ebLotnum"
    fGetRibbonReference.InvalidateControl "ebHospital"

    shtDataStage.Range("K1").Value = ebSalesCompany_val
    shtDataStage.Range("K2").Value = ebProductProducer_val
    shtDataStage.Range("K3").Value = ebProductName_val
    shtDataStage.Range("K4").Value = ebProductSeries_val
    shtDataStage.Range("K5").Value = ebLotNum_val
    shtDataStage.Range("K6").Value = ebHospital_val
            
'            Call fSetName("nmHospital", ebHospital_val)
'            Call fSetName("nmSalesCompany", ebSalesCompany_val)
'            Call fSetName("nmProductProducer", ebProductProducer_val)
'            Call fSetName("nmProductName", ebProductName_val)
'            Call fSetName("nmProductSeries", ebProductSeries_val)
'            Call fSetName("nmLotNum", ebLotNum_val)
End Sub
Sub Sub_SearchOnCurrentSheetBySearchBy(Optional sht As Worksheet)
    Dim iColIndex_Hospital As Integer
    Dim iColIndex_SalesCompany As Integer
    Dim iColIndex_ProductProducer As Integer
    Dim iColIndex_ProductName As Integer
    Dim iColIndex_ProductSeries As Integer
    Dim iColIndex_LotNum As Integer
    Dim lMaxRow As Long
    Dim lMaxCol As Long
    
    If sht Is Nothing Then Set sht = ActiveSheet
    
    If Not fGetColIndexs(sht, iColIndex_Hospital _
            , iColIndex_SalesCompany, iColIndex_ProductProducer _
            , iColIndex_ProductName, iColIndex_ProductSeries _
            , iColIndex_LotNum) Then Exit Sub
     
    fKeepCopyContent
    
    fRefreshEditBoxFromShtDataStage
    
    lMaxCol = sht.Cells(1, 1).End(xlToRight).Column
    lMaxRow = fGetValidMaxRow(sht)

    If sht.AutoFilterMode Then  'auto filter
        sht.AutoFilter.ShowAllData
    Else
        fGetRangeByStartEndPos(sht, 1, 1, 1, lMaxCol).AutoFilter
    End If
    
    Dim rg As Range
     
    Set rg = fGetRangeByStartEndPos(sht, 1, 1, lMaxRow, lMaxCol)
     
    If iColIndex_Hospital > 0 And Len(ebHospital_val) > 0 Then
        rg.AutoFilter Field:=iColIndex_Hospital, Criteria1:="=*" & ebHospital_val & "*", Operator:=xlAnd
    End If
    If iColIndex_SalesCompany > 0 And Len(ebSalesCompany_val) > 0 Then
        rg.AutoFilter Field:=iColIndex_ProductProducer, Criteria1:="=*" & ebSalesCompany_val & "*", Operator:=xlAnd
    End If
    If iColIndex_ProductProducer > 0 And Len(ebProductProducer_val) > 0 Then
        rg.AutoFilter Field:=iColIndex_ProductProducer, Criteria1:="=*" & ebProductProducer_val & "*", Operator:=xlAnd
    End If
    If iColIndex_ProductName > 0 And Len(ebProductName_val) > 0 Then
        rg.AutoFilter Field:=iColIndex_ProductName, Criteria1:="=*" & ebProductName_val & "*", Operator:=xlAnd
    End If
    If iColIndex_ProductSeries > 0 And Len(ebProductSeries_val) > 0 Then
        rg.AutoFilter Field:=iColIndex_ProductSeries, Criteria1:="=*" & ebProductSeries_val & "*", Operator:=xlAnd
    End If
    If iColIndex_LotNum > 0 And Len(ebLotNum_val) > 0 Then
        rg.AutoFilter Field:=iColIndex_LotNum, Criteria1:="=*" & ebLotNum_val & "*", Operator:=xlAnd
    End If
    
    fShowActivateSheet sht
    
    Set rg = Nothing
    'Set sht = Nothing  'this line will cause issue
      
    fCopyFromKept
End Sub

Private Sub UIbtnHome(control As IRibbonControl)
    Call Sub_ToHomeSheet
End Sub

Private Sub tgGetSearchBy_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = tgGetSearchBy_Val
End Sub
'
'Function fGetCurrSheetCurrRowSalesCompany(sOut As String) As Boolean
'
''    Dim iColIndex_Hospital As Integer
'    Dim iColIndex_SalesCompany As Integer
'    Dim iColIndex_ProductProducer As Integer
'    Dim iColIndex_ProductName As Integer
'    Dim iColIndex_ProductSeries As Integer
'    Dim iColIndex_LotNum As Integer
'
'    Call fGetColIndexs(ActiveSheet, , iColIndex_SalesCompany)
'    sOut = Trim(ActiveSheet.Cells(ActiveCell.Row, iColIndex_Hospital).Value)
'
'    'Call fGetValueFromCurrSheetCurrRow("SALES_COMPANY", sOut)
'End Function
'Function fGetCurrSheetCurrRowProductProducer(sOut As String) As Boolean
''    Dim iColIndex_Hospital As Integer
'    Dim iColIndex_SalesCompany As Integer
'    Dim iColIndex_ProductProducer As Integer
'    Dim iColIndex_ProductName As Integer
'    Dim iColIndex_ProductSeries As Integer
'    Dim iColIndex_LotNum As Integer
'
'    Call fGetColIndexs(ActiveSheet, , iColIndex_SalesCompany)
'    sOut = Trim(ActiveSheet.Cells(ActiveCell.Row, iColIndex_Hospital).Value)
'
'    Call fGetValueFromCurrSheetCurrRow("PRODUCT_PRODUCER", sOut)
'End Function
'Function fGetCurrSheetCurrRowProductName(sOut As String) As Boolean
''    Dim iColIndex_Hospital As Integer
'    Dim iColIndex_SalesCompany As Integer
'    Dim iColIndex_ProductProducer As Integer
'    Dim iColIndex_ProductName As Integer
'    Dim iColIndex_ProductSeries As Integer
'    Dim iColIndex_LotNum As Integer
'
'    Call fGetColIndexs(ActiveSheet, , iColIndex_SalesCompany)
'    sOut = Trim(ActiveSheet.Cells(ActiveCell.Row, iColIndex_Hospital).Value)
'
'    Call fGetValueFromCurrSheetCurrRow("PRODUCT_NAME", sOut)
'End Function
'Function fGetCurrSheetCurrRowProductSeries(sOut As String) As Boolean
''    Dim iColIndex_Hospital As Integer
'    Dim iColIndex_SalesCompany As Integer
'    Dim iColIndex_ProductProducer As Integer
'    Dim iColIndex_ProductName As Integer
'    Dim iColIndex_ProductSeries As Integer
'    Dim iColIndex_LotNum As Integer
'
'    Call fGetColIndexs(ActiveSheet, , iColIndex_SalesCompany)
'    sOut = Trim(ActiveSheet.Cells(ActiveCell.Row, iColIndex_Hospital).Value)
'
'    Call fGetValueFromCurrSheetCurrRow("PRODUCT_SERIES", sOut)
'End Function
'Function fGetCurrSheetCurrRowLotNum(sOut As String) As Boolean
''    Dim iColIndex_Hospital As Integer
'    Dim iColIndex_SalesCompany As Integer
'    Dim iColIndex_ProductProducer As Integer
'    Dim iColIndex_ProductName As Integer
'    Dim iColIndex_ProductSeries As Integer
'    Dim iColIndex_LotNum As Integer
'
'    Call fGetColIndexs(ActiveSheet, , iColIndex_SalesCompany)
'    sOut = Trim(ActiveSheet.Cells(ActiveCell.Row, iColIndex_Hospital).Value)
'    'Call fGetValueFromCurrSheetCurrRow("LOT_NUM", sOut)
'End Function
'Function fGetCurrSheetCurrRowHospital(sOut As String) As Boolean
'    Dim iColIndex_Hospital As Integer
'
'    Call fGetColIndexs(ActiveSheet, iColIndex_Hospital)
'    sOut = Trim(ActiveSheet.Cells(ActiveCell.Row, iColIndex_Hospital).Value)
'    'Call fGetValueFromCurrSheetCurrRow("HOSPITAL", sOut)
'End Function
 
'Function fReGetValue_tgGetSearchBy() As Boolean
'    fGetRibbonReference.InvalidateControl "tgGetSearchBy"
'    'fGetRibbonReference.Invalidate
'    fReGetValue_tgGetSearchBy = tgGetSearchBy_Val
'End Function
 
Function fGetControlAttributes(control As IRibbonControl, sType As String, Optional ByRef val)
    If Not (sType = "LABEL" Or sType = "IMAGE" Or sType = "SIZE" Or sType = "ACTION") Then
        fErr "wrong param to fGetControlAttributes: " & vbCr & "sType=" & sType & vbCr & "control=" & control.id
    End If
    
    Select Case control.id
        Case "btnHome"
            Select Case sType
                Case "LABEL":   val = "主菜单"
                Case "IMAGE":   val = "OpenStartPage"
                Case "SIZE":    val = "true"
                Case "ACTION":  Call Sub_ToHomeSheet
            End Select
        Case "btnCloseAllSheets"
            Select Case sType
                Case "LABEL":   val = "关闭所有"
                Case "IMAGE":   val = "DeclineInvitation"
                Case "SIZE":    val = "true"
                Case "ACTION":  Call subMain_InvisibleHideAllBusinessSheets
            End Select
        Case "btnFilterBySelected"
            Select Case sType
                Case "LABEL":   val = "以所选过滤"
                Case "IMAGE":   val = "FilterBySelection"
                Case "SIZE":    val = "true"
                Case "ACTION":  Call Sub_FilterBySelectedCells
            End Select
        Case "btnSortBySelected"
            Select Case sType
                Case "LABEL":   val = "以所选排序"
                Case "IMAGE":   val = "SortUp"
                Case "SIZE":    val = "true"
                Case "ACTION":  Call sub_SortBySelectedCells
            End Select
            
        Case "btnRemoveFilter"
            Select Case sType
                Case "LABEL":   val = "清除过滤"
                Case "IMAGE":   val = "FilterClearAllFilters"
                Case "SIZE":    val = "true"
                Case "ACTION":  Call Sub_RemoveFilterForAcitveSheet
            End Select
        Case "tgGetSearchBy"
            Select Case sType
                Case "LABEL":   val = "获取当前行"
                Case "IMAGE":   val = "RecurrenceEdit"
                Case "SIZE":    val = "false"
                Case "ACTION":  Call sub_PresstgGetSearchBy
            End Select
        Case "tgSearchOnCurrSheet"
            Select Case sType
                Case "LABEL":   val = "在当前表过滤"
                Case "IMAGE":   val = "FilterBySelection"
                Case "SIZE":    val = "false"
                Case "ACTION":  Call Sub_SearchOnCurrentSheetBySearchBy
            End Select
'        Case "btnSCompInvImported"
'            Select Case sType
'                Case "LABEL":   val = "(商业公司)库存表(导入的)"
'                Case "IMAGE":   val = "PageSetupSheetDialog"
'                Case "SIZE":    val = "false"
'                Case "ACTION":  Call btnSCompInvImported_Click
'            End Select
'        Case "btnPromotionProduct"
'            Select Case sType
'                Case "LABEL":   val = "推广品"
'                Case "IMAGE":   val = "PageSetupSheetDialog"
'                Case "SIZE":    val = "false"
'                Case "ACTION":  Call btnPromotionProduct_Click
'            End Select
        Case "btnFirstLevelComm"
            Select Case sType
                Case "LABEL":   val = "采芝林配送费"
                Case "IMAGE":   val = "PageSetupSheetDialog"
                Case "SIZE":    val = "false"
                Case "ACTION":  Call Sub_SearchOnCurrentSheetBySearchBy(shtFirstLevelCommission)
            End Select
        Case "btnSecondLevelComm"
            Select Case sType
                Case "LABEL":   val = "商业公司配送费"
                Case "IMAGE":   val = "PageSetupSheetDialog"
                Case "SIZE":    val = "false"
                Case "ACTION":  Call Sub_SearchOnCurrentSheetBySearchBy(shtSecondLevelCommission)
            End Select
        Case "btnSalePriceInAdv"
            Select Case sType
                Case "LABEL":   val = "药品预收供货价"
                Case "IMAGE":   val = "PageSetupSheetDialog"
                Case "SIZE":    val = "false"
                Case "ACTION":  Call Sub_SearchOnCurrentSheetBySearchBy(shtSellPriceInAdv)
            End Select
        Case "btnSalesInfo"
            Select Case sType
                Case "LABEL":   val = "替换统一的销售流向"
                Case "IMAGE":   val = "PageSetupSheetDialog"
                Case "SIZE":    val = "false"
                Case "ACTION":  Call Sub_SearchOnCurrentSheetBySearchBy(shtSalesInfos)
            End Select
        Case "btnProfit"
            Select Case sType
                Case "LABEL":   val = "利润"
                Case "IMAGE":   val = "PageSetupSheetDialog"
                Case "SIZE":    val = "false"
                Case "ACTION":  Call Sub_SearchOnCurrentSheetBySearchBy(shtProfit)
            End Select

        Case "tbtnProfitSheet"
            Select Case sType
                Case "LABEL":   val = "利润表"
                Case "IMAGE":   val = "ControlSubFormReport"
                'Case "ACTION":  Call ToggleButtonToSwitchSheet_onAction(control)
            End Select
        Case "tbtnRefundSheet"
            Select Case sType
                Case "LABEL":   val = "补差表"
                Case "IMAGE":   val = "ControlSubFormReport"
                'Case "ACTION":  Call ToggleButtonToSwitchSheet_onAction(control)
            End Select
        Case "tbtnSelfInv"
            Select Case sType
                Case "LABEL":   val = "(本公司)库存表"
                Case "IMAGE":   val = "ControlSubFormReport"
                'Case "ACTION":  Call ToggleButtonToSwitchSheet_onAction(control)
            End Select
        Case "tbtnCZLInvCalcd"
            Select Case sType
                Case "LABEL":   val = "(采芝林)库存表(本公司计算)"
                Case "IMAGE":   val = "ControlSubFormReport"
                'Case "ACTION":  Call ToggleButtonToSwitchSheet_onAction(control)
            End Select

    End Select
    
End Function

Sub testaaaaa()
    Dim control As IRibbonControl
    Dim sht As Worksheet
    Set sht = fGetSheetByUIRibbonTag("shtProductMaster")
    
    If Not sht Is Nothing Then
        fToggleSheetVisibleFromUIRibbonControl True, sht, control
    End If
    Set sht = Nothing
End Sub
