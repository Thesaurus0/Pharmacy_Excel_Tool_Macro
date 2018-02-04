Attribute VB_Name = "MD01_InventoryCalculate"
Option Explicit
Option Base 1

Sub subMain_CalculateCZLInventory()
    fClearDataFromSheetLeaveHeader shtCZLInventory
    
    fCalculateCZLInventory
    fActiveVisibleSwitchSheet shtCZLInventory, , False
    
    fMsgBox "采芝林库存计算完成！", vbInformation
End Sub


Private Function fCalculateCZLInventory()
    Call fRemoveFilterForSheet(shtSelfSalesOrder)
    Call fRemoveFilterForSheet(shtCZLSales2Companies)
    
    If Not shtSelfSalesOrder.fValidateSheet(False) Then Exit Function
    If Not shtCZLSales2Companies.fValidateSheet(False) Then Exit Function
    
    'If dictSelfPurchaseOD Is Nothing Then Call fReadSheetSelfPurchaseOrder2Dictionary
    If dictSelfSalesOD Is Nothing Then Call fReadSheetSelfSalesOrder2Dictionary
    
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
    
    ReDim arrOut(1 To dictSelfPurchaseOD.Count, 7)
    
    For i = 0 To dictSelfPurchaseOD.Count - 1
        sKey = dictSelfPurchaseOD.Keys(i)
        
        dblPurchaseQty = CDbl(Split(dictSelfPurchaseOD(sKey), DELIMITER)(0))
        
        If dictSelfSalesOD.Exists(sKey) Then
            dblSellQty = CDbl(Split(dictSelfSalesOD(sKey), DELIMITER)(0))
        Else
            dblSellQty = 0
        End If
            
        arrOut(i + 1, 1) = Split(sKey, DELIMITER)(0)
        arrOut(i + 1, 2) = Split(sKey, DELIMITER)(1)
        arrOut(i + 1, 3) = Split(sKey, DELIMITER)(2)
        arrOut(i + 1, 5) = Split(sKey, DELIMITER)(3)
        
        arrOut(i + 1, 6) = dblPurchaseQty - dblSellQty
        arrOut(i + 1, 7) = CDbl(Split(dictSelfPurchaseOD(sKey), DELIMITER)(1))
    Next
    
    'fCalculateSelfInventory = arrOut
    Call fAppendArray2Sheet(shtSelfInventory, arrOut)
    Erase arrOut
End Function
