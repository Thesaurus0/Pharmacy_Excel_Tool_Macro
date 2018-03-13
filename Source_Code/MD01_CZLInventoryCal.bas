Attribute VB_Name = "MD01_CZLInventoryCal"
Option Explicit
Option Base 1

'Sub subMain_CalculateCZLInventory()
'    fClearDataFromSheetLeaveHeader shtCZLInventory
'
'    fCalculateCZLInventory
'    fActiveVisibleSwitchSheet shtCZLInventory, , False
'
'    fMsgBox "采芝林库存计算完成！", vbInformation
'End Sub
'
'
'Private Function fCalculateCZLInventory()
'    Call fRemoveFilterForSheet(shtSelfSalesOrder)
'    Call fRemoveFilterForSheet(shtCZLSales2Companies)
'
'    If Not shtSelfSalesOrder.fValidateSheet(False) Then Exit Function
'    If Not shtCZLSales2Companies.fValidateSheet(False) Then Exit Function
'
'    'If dictSelfPurchaseOD Is Nothing Then Call fReadSheetSelfPurchaseOrder2Dictionary
'    If dictSelfSalesOD Is Nothing Then Call fReadSheetSelfSalesOrder2Dictionary
'
'    Dim i As Long
'    Dim lEachRow As Long
'    Dim sKey As String
'    Dim sProducer As String
'    Dim sProductName As String
'    Dim sProductSeries As String
'    Dim sLotNum As String
'    Dim dblPurchaseQty As Double
'    Dim dblSellQty As Double
'    Dim arrOut()
'
'    ReDim arrOut(1 To dictSelfPurchaseOD.Count, 7)
'
'    For i = 0 To dictSelfPurchaseOD.Count - 1
'        sKey = dictSelfPurchaseOD.Keys(i)
'
'        dblPurchaseQty = CDbl(Split(dictSelfPurchaseOD(sKey), DELIMITER)(0))
'
'        If dictSelfSalesOD.Exists(sKey) Then
'            dblSellQty = CDbl(Split(dictSelfSalesOD(sKey), DELIMITER)(0))
'        Else
'            dblSellQty = 0
'        End If
'
'        arrOut(i + 1, 1) = Split(sKey, DELIMITER)(0)
'        arrOut(i + 1, 2) = Split(sKey, DELIMITER)(1)
'        arrOut(i + 1, 3) = Split(sKey, DELIMITER)(2)
'        arrOut(i + 1, 5) = Split(sKey, DELIMITER)(3)
'
'        arrOut(i + 1, 6) = dblPurchaseQty - dblSellQty
'        arrOut(i + 1, 7) = CDbl(Split(dictSelfPurchaseOD(sKey), DELIMITER)(1))
'    Next
'
'    'fCalculateSelfInventory = arrOut
'    Call fAppendArray2Sheet(shtSelfInventory, arrOut)
'    Erase arrOut
'End Function
Function fCalculateCZLInventory()
    Call fRemoveFilterForSheet(shtSelfSalesOrder)       'purchase
    Call fRemoveFilterForSheet(shtCZLSales2Companies)       'sales
    
    If Not shtSelfSalesOrder.fValidateSheet(False) Then fErr    'purchase
    'If Not shtCZLSales2Companies.fValidateSheet(False) Then fErr     'sales
    
    If dictSelfSalesOD Is Nothing Then Call fReadSheetSelfSalesOrder2Dictionary
    If dictCZLSalesOD Is Nothing Then Call fReadSheetCZLSalesOrder2Dictionary
    'If dictCZLSelfSalesOD Is Nothing Then
    Call fReadSheetCZLSalesOrder2Hospital2Dictionary
    
    Dim i As Long
    Dim lEachRow As Long
    Dim sKey As String
    Dim sProducer As String
    Dim sProductName As String
    Dim sProductSeries As String
    Dim sLotNum As String
    Dim dblPurchaseQty As Double
    Dim dblSellQty As Double
    Dim arrOut()
    
    Dim dictMissedLot As Dictionary
    Set dictMissedLot = New Dictionary
    
    For i = 0 To dictCZLSalesOD.Count - 1
        sKey = dictCZLSalesOD.Keys(i)
        
        If Not dictSelfSalesOD.Exists(sKey) Then
            dictMissedLot.Add sKey, 0
        End If
    Next
    For i = 0 To dictCZLSelfSalesOD.Count - 1
        sKey = dictCZLSelfSalesOD.Keys(i)
        
        If Not dictSelfSalesOD.Exists(sKey) Then
            If Not dictMissedLot.Exists(sKey) Then dictMissedLot.Add sKey, 0
        End If
    Next
    
    If dictMissedLot.Count > 0 Then
        Call fAddMissedSelfSalesLotNumToSheetException(dictMissedLot)
        'fErr gsBusinessErrorMsg
'        fMsgBox gsBusinessErrorMsg
        gsBusinessErrorMsg = gsBusinessErrorMsg
    End If
    
    Set dictMissedLot = Nothing
    
    ReDim arrOut(1 To dictSelfSalesOD.Count, 7) 'purchase
    
    For i = 0 To dictSelfSalesOD.Count - 1  'purchase
        sKey = dictSelfSalesOD.Keys(i)
        
        dblPurchaseQty = CDbl(Split(dictSelfSalesOD(sKey), DELIMITER)(0))
        
        If dictCZLSalesOD.Exists(sKey) Then
            dblSellQty = CDbl(Split(dictCZLSalesOD(sKey), DELIMITER)(0))
        Else
            dblSellQty = 0
        End If
        
        If dictCZLSelfSalesOD.Exists(sKey) Then
            dblSellQty = dblSellQty + CDbl(Split(dictCZLSelfSalesOD(sKey), DELIMITER)(0))
        End If
            
        arrOut(i + 1, 1) = Split(sKey, DELIMITER)(0)
        arrOut(i + 1, 2) = Split(sKey, DELIMITER)(1)
        arrOut(i + 1, 3) = Split(sKey, DELIMITER)(2)
        arrOut(i + 1, 4) = fGetProductUnit(arrOut(i + 1, 1), arrOut(i + 1, 2), arrOut(i + 1, 3))
        arrOut(i + 1, 5) = "'" & Split(sKey, DELIMITER)(3)  'lot num
        
        arrOut(i + 1, 6) = dblPurchaseQty - dblSellQty
        If IsNumeric(Split(dictSelfSalesOD(sKey), DELIMITER)(2)) Then
            arrOut(i + 1, 7) = CDbl(Split(dictSelfSalesOD(sKey), DELIMITER)(2))     'purcahse price
        Else
        End If
    Next
    
    'fCalculateCZLInventory = arrOut
    Call fAppendArray2Sheet(shtCZLInventory, arrOut)
    Erase arrOut
    Set dictCZLSalesOD = Nothing
End Function



