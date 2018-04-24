Attribute VB_Name = "MB_1_RibbonUI"
Option Explicit

#If VBA7 And Win64 Then  'Win64
    Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As Long)
#Else
    Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As Long)
#End If

Private mRibbonObj As IRibbonUI
Public tgSearchBy_Val As Boolean
Private ebSalesCompany_val As String
Private ebProductProducer_val As String
Private ebProductName_val As String
Private ebProductSeries_val As String
Private ebLotNum_val As String

Sub ERP_UI_Onload(ribbon As IRibbonUI)
  Set mRibbonObj = ribbon
  
  fAddNameWithConstValueIfNonExists "nmRibbonPointer", 0
  Names("nmRibbonPointer").RefersTo = ObjPtr(ribbon)
  
  mRibbonObj.ActivateTab "ERP_2010"
  tgSearchBy_Val = True
End Sub
Function fReGetRibbonReference() As IRibbonUI
    If Not mRibbonObj Is Nothing Then Set fReGetRibbonReference = mRibbonObj: Exit Function
    
    Dim objRibbon As Object
    Dim lRibPointer As LongPtr
    
    lRibPointer = [nmRibbonPointer]
    
    CopyMemory objRibbon, lRibPointer, LenB(lRibPointer)
    
    Set fReGetRibbonReference = objRibbon
    Set mRibbonObj = objRibbon
    Set objRibbon = Nothing
End Function
 
Sub subUIebSalesCompany_onChange(control As IRibbonControl, text As String)
    ebSalesCompany_val = text
End Sub
Sub subUIebProductProducer_onChange(control As IRibbonControl, text As String)
    ebProductProducer_val = text
End Sub
Sub subUIebProductName_onChange(control As IRibbonControl, text As String)
    ebProductName_val = text
End Sub
Sub subUIebProductSeries_onChange(control As IRibbonControl, text As String)
    ebProductSeries_val = text
End Sub
Sub subUIebLotnum_onChange(control As IRibbonControl, text As String)
    ebLotNum_val = text
End Sub

Private Sub subUIebSalesCompany_getText(control As IRibbonControl, ByRef returnedVal)
    If fGetCurrSheetCurrRowSalesCompany(ebSalesCompany_val) Then
        returnedVal = ebSalesCompany_val
    End If
End Sub
Private Sub subUIebProductProducer_getText(control As IRibbonControl, ByRef returnedVal)
    If fGetCurrSheetCurrRowProductProducer(ebProductProducer_val) Then
        returnedVal = ebProductProducer_val
    End If
End Sub

Private Sub subUIebProductName_getText(control As IRibbonControl, ByRef returnedVal)
    If fGetCurrSheetCurrRowProductName(ebProductName_val) Then
        returnedVal = ebProductName_val
    End If
End Sub
Private Sub subUIebProductSeries_getText(control As IRibbonControl, ByRef returnedVal)
    If fGetCurrSheetCurrRowProductSeries(ebProductSeries_val) Then
        returnedVal = ebProductSeries_val
    End If
End Sub
Private Sub subUIebLotnum_getText(control As IRibbonControl, ByRef returnedVal)
    If fGetCurrSheetCurrRowLotNum(ebLotNum_val) Then
        returnedVal = ebLotNum_val
    End If
End Sub

Private Sub UIbtnHome(control As IRibbonControl)
    Call Sub_ToHomeSheet
End Sub
 
Private Sub btnSCompInvImported_Click(control As IRibbonControl)
    If Len(ebProductProducer_val) > 0 Then
        Call fSetFilterForSheet(shtSalesCompInvUnified, Array(SCompUnifiedInv.SalesCompany, SCompUnifiedInv.ProductProducer, SCompUnifiedInv.ProductName, SCompUnifiedInv.ProductSeries, SCompUnifiedInv.LotNum) _
                , Array(ebSalesCompany_val, ebProductProducer_val, ebProductName_val, ebProductSeries_val, ebLotNum_val))
    Else
        fRemoveFilterForSheet shtSalesCompInvUnified
    End If
    
    fActiveVisibleSwitchSheet shtSalesCompInvUnified
End Sub
 
Private Sub tbSCompInvCalcd_Click(control As IRibbonControl)
    If Len(ebProductProducer_val) > 0 Then
        Call fSetFilterForSheet(shtSalesCompInvCalcd, Array(SCompInvCalcd.SalesCompany, SCompInvCalcd.ProductProducer, SCompInvCalcd.ProductName, SCompInvCalcd.ProductSeries, SCompInvCalcd.LotNum) _
                , Array(ebSalesCompany_val, ebProductProducer_val, ebProductName_val, ebProductSeries_val, ebLotNum_val))
    Else
        fRemoveFilterForSheet shtSalesCompInvCalcd
    End If
    
    fActiveVisibleSwitchSheet shtSalesCompInvCalcd
End Sub

Private Sub btnCZLSalesToSComp_Click(control As IRibbonControl)
    If Len(ebProductProducer_val) > 0 Then
        Call fSetFilterForSheet(shtCZLSales2Companies, Array(CZLSales2Comp.SalesCompany, CZLSales2Comp.ProductProducer, CZLSales2Comp.ProductName, CZLSales2Comp.ProductSeries, CZLSales2Comp.LotNum) _
                , Array(ebSalesCompany_val, ebProductProducer_val, ebProductName_val, ebProductSeries_val, ebLotNum_val))
    Else
        fRemoveFilterForSheet shtCZLSales2Companies
    End If
    
    fActiveVisibleSwitchSheet shtCZLSales2Companies
End Sub
Private Sub btnProductMaster_Click(control As IRibbonControl)
    If Len(ebProductProducer_val) > 0 Then
        Call fSetFilterForSheet(shtProductMaster, Array(ProductMst.ProductProducer, ProductMst.ProductName, ProductMst.ProductSeries) _
                , Array(ebProductProducer_val, ebProductName_val, ebProductSeries_val))
    Else
        fRemoveFilterForSheet shtProductMaster
    End If
    
    fActiveVisibleSwitchSheet shtProductMaster
End Sub
Private Sub tgSearchBy_Click(control As IRibbonControl, pressed As Boolean)
    tgSearchBy_Val = pressed
    Call fPresstgSearchBy(pressed)
End Sub

Private Sub tgSearchBy_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = tgSearchBy_Val
End Sub

Private Sub btnRemoveFilter_Click(control As IRibbonControl)
    Sub_RemoveFilterForAcitveSheet
End Sub

Function fPresstgSearchBy(bPressed As Boolean)
    If bPressed Then
        fReGetRibbonReference.InvalidateControl "ebSalesCompany"
        fReGetRibbonReference.InvalidateControl "ebProductProducer"
        fReGetRibbonReference.InvalidateControl "ebProductName"
        fReGetRibbonReference.InvalidateControl "ebProductSeries"
        fReGetRibbonReference.InvalidateControl "ebLotnum"
    Else
    End If
End Function

Function fGetCurrSheetCurrRowSalesCompany(sOut As String) As Boolean
    Dim iColIndex As Integer
    
    If ActiveCell.Row <= 1 Then Exit Function
    If Not fGetSalesCompanyColIndex(ActiveSheet, iColIndex) Then Exit Function
        
    sOut = Trim(ActiveSheet.Cells(ActiveCell.Row, iColIndex).Value)
    fGetCurrSheetCurrRowSalesCompany = True
End Function
Function fGetCurrSheetCurrRowProductProducer(sOut As String) As Boolean
    Dim iColIndex As Integer
    
    If ActiveCell.Row <= 1 Then Exit Function
    If Not fGetProductProcuderColIndex(ActiveSheet, iColIndex) Then Exit Function
        
    sOut = Trim(ActiveSheet.Cells(ActiveCell.Row, iColIndex).Value)
    
    fGetCurrSheetCurrRowProductProducer = True
End Function
Function fGetCurrSheetCurrRowProductName(sOut As String) As Boolean
    Dim iColIndex As Integer
    
    If ActiveCell.Row <= 1 Then Exit Function
    If Not fGetProductProductNameColIndex(ActiveSheet, iColIndex) Then Exit Function
        
    sOut = Trim(ActiveSheet.Cells(ActiveCell.Row, iColIndex).Value)
    
    fGetCurrSheetCurrRowProductName = True
End Function
Function fGetCurrSheetCurrRowProductSeries(sOut As String) As Boolean
    Dim iColIndex As Integer
    
    If ActiveCell.Row <= 1 Then Exit Function
    If Not fGetProductSeriesColIndex(ActiveSheet, iColIndex) Then Exit Function
        
    sOut = Trim(ActiveSheet.Cells(ActiveCell.Row, iColIndex).Value)
    
    fGetCurrSheetCurrRowProductSeries = True
End Function
Function fGetCurrSheetCurrRowLotNum(sOut As String) As Boolean
    Dim iColIndex As Integer
    
    If ActiveCell.Row <= 1 Then Exit Function
    If Not fGetLogNumColIndex(ActiveSheet, iColIndex) Then Exit Function
        
    sOut = Trim(ActiveSheet.Cells(ActiveCell.Row, iColIndex).Value)
    
    fGetCurrSheetCurrRowLotNum = True
End Function

Function fGetSalesCompanyColIndex(shtParam As Worksheet, ByRef iColIndex As Integer) As Boolean
    Select Case shtParam.CodeName
        Case "shtSalesCompInvUnified"
            iColIndex = SCompUnifiedInv.SalesCompany
        Case "shtSalesCompInvDiff"
            iColIndex = SCompInvDiff.SalesCompany
        Case "shtSalesCompInvCalcd"
            iColIndex = SCompInvCalcd.SalesCompany
        Case Else
            iColIndex = 0: Exit Function
    End Select
    
    fGetSalesCompanyColIndex = True
End Function
Function fGetProductProcuderColIndex(shtParam As Worksheet, ByRef iColIndex As Integer) As Boolean
    Select Case shtParam.CodeName
        Case "shtSalesCompInvUnified"
            iColIndex = SCompUnifiedInv.ProductProducer
        Case "shtSalesCompInvDiff"
            iColIndex = SCompInvDiff.ProductProducer
        Case "shtSalesCompInvCalcd"
            iColIndex = SCompInvCalcd.ProductProducer
        Case Else
            iColIndex = 0: Exit Function
    End Select
    
    fGetProductProcuderColIndex = True
End Function

Function fGetProductProductNameColIndex(shtParam As Worksheet, ByRef iColIndex As Integer) As Boolean
    Select Case shtParam.CodeName
        Case "shtSalesCompInvUnified"
            iColIndex = SCompUnifiedInv.ProductName
        Case "shtSalesCompInvDiff"
            iColIndex = SCompInvDiff.ProductName
        Case "shtSalesCompInvCalcd"
            iColIndex = SCompInvCalcd.ProductName
        Case Else
            iColIndex = 0: Exit Function
    End Select
    
    fGetProductProductNameColIndex = True
End Function

Function fGetProductSeriesColIndex(shtParam As Worksheet, ByRef iColIndex As Integer) As Boolean
    Select Case shtParam.CodeName
        Case "shtSalesCompInvUnified"
            iColIndex = SCompUnifiedInv.ProductSeries
        Case "shtSalesCompInvDiff"
            iColIndex = SCompInvDiff.ProductSeries
        Case "shtSalesCompInvCalcd"
            iColIndex = SCompInvCalcd.ProductSeries
        Case Else
            iColIndex = 0: Exit Function
    End Select
    
    fGetProductSeriesColIndex = True
End Function
Function fGetLogNumColIndex(shtParam As Worksheet, ByRef iColIndex As Integer) As Boolean
    Select Case shtParam.CodeName
        Case "shtSalesCompInvUnified"
            iColIndex = SCompUnifiedInv.LotNum
        Case "shtSalesCompInvDiff"
            iColIndex = SCompInvDiff.LotNum
        Case "shtSalesCompInvCalcd"
            iColIndex = SCompInvCalcd.LotNum
        Case Else
            iColIndex = 0: Exit Function
    End Select
    
    fGetLogNumColIndex = True
End Function

Function fReGetValue_tgSearchBy() As Boolean
    fReGetRibbonReference.InvalidateControl "tgSearchBy"
    'fReGetRibbonReference.Invalidate
    fReGetValue_tgSearchBy = tgSearchBy_Val
End Function

'Function fGetProdProducerCol(shtParam As Worksheet) As Integer
'    Select Case shtParam.CodeName
'        Case "shtSalesCompInvDiff"
'            fGetProdProducerCol = SCompInvDiff.ProductProducer
'        Case Else
'            fGetProdProducerCol = 9
'    End Select
'End Function

Sub ddd()
    Debug.Print tgSearchBy_Val
End Sub


'Private Sub dwSearchTables_getItemCount(control As IRibbonControl, ByRef returnedVal)
'    returnedVal = 8
'End Sub
'
'Private Sub dwSearchTables_Click(control As IRibbonControl, id As String, index As Integer)
'    Dim sTable
'   ' Call dwSearchTables_getItemLabel(control, index, sTable)
'    MsgBox id & vbCr & sTable
'End Sub
'
'Private Sub dwSearchTables_getItemID(control As IRibbonControl, index As Integer, ByRef id)
'    Select Case index
'        Case 0
'            id = "Table_" & index
'        Case 1
'            id = "tbSCompInvInformed"
'        Case 2
'            id = "tbSCompInvCalcd"
'    End Select
'End Sub
'
'Private Sub dwSearchTables_getItemLabel(control As IRibbonControl, index As Integer, ByRef returnedVal)
'    Select Case index
'        Case 0
'            returnedVal = "药品"
'        Case 1
'            returnedVal = "(商业公司)库存表(导入的)"
'        Case 2
'            returnedVal = "(商业公司)库存表(计算的)"
'        Case Else
'            returnedVal = "药品90"
'    End Select
'End Sub


