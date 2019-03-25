Attribute VB_Name = "MD01_SelfInventoryCal"
Option Explicit
Option Base 1

Private dictSelfPurchaseOD As Dictionary
Private dictSelfSalesOD As Dictionary

Function fResetdictSelfPurchaseOD()
    Set dictSelfPurchaseOD = Nothing
End Function
Function fResetdictSelfSalesOD()
    Set dictSelfSalesOD = Nothing
End Function
Function fGetdictSelfSalesOD(ByRef dictOut As Dictionary)
    If dictSelfSalesOD Is Nothing Then Call fReadSheetSelfSalesOrder2Dictionary
    
    Set dictOut = dictSelfSalesOD
End Function

Sub subMain_CalculateSelfInventory()
    If Not fIsDev() Then On Error GoTo error_handling
    
    gsRptID = "CALCULATE_PROFIT"
    
    fClearContentLeaveHeader shtSelfInventory
    
    fCalculateSelfInventory
    
error_handling:
    If fCheckIfGotBusinessError Then
        GoTo reset_excel_options
    End If
    
    If fCheckIfUnCapturedExceptionAbnormalError Then
        GoTo reset_excel_options
    End If
     
    fMsgBox "本公司库存计算完成！", vbInformation
    fActiveVisibleSwitchSheet shtSelfInventory, shtSelfInventory.Range("A2"), False
    
reset_excel_options:
    Err.Clear
    fClearRefVariables
    fEnableExcelOptionsAll
    'End
End Sub

Private Function fCalculateSelfInventory()
    Call fRemoveFilterForSheet(shtSelfPurchaseOrder)
    Call fRemoveFilterForSheet(shtSelfSalesOrder)
    
    If Not shtSelfPurchaseOrder.fValidateSheet(False) Then fErr
    If Not shtSelfSalesOrder.fValidateSheet(False) Then fErr
    
    If dictSelfPurchaseOD Is Nothing Then Call fReadSheetSelfPurchaseOrder2Dictionary
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
    
    For i = 0 To dictSelfSalesOD.count - 1
        sKey = dictSelfSalesOD.Keys(i)
        
        If Not dictSelfPurchaseOD.Exists(sKey) Then
            dictSelfPurchaseOD.Add sKey, "0|" & dictSelfSalesOD.Items(i)
        End If
    Next
    
    ReDim arrOut(1 To dictSelfPurchaseOD.count, 7)
    
    For i = 0 To dictSelfPurchaseOD.count - 1
        sKey = dictSelfPurchaseOD.Keys(i)
        
        dblPurchaseQty = CDbl(Split(dictSelfPurchaseOD(sKey), DELIMITER)(0))
        
        If dictSelfSalesOD.Exists(sKey) Then
            dblSellQty = CDbl(Split(dictSelfSalesOD(sKey), DELIMITER)(0))
        Else
            dblSellQty = 0
        End If
            
        arrOut(i + 1, SelfInv.ProductProducer) = Split(sKey, DELIMITER)(0)    'producer
        arrOut(i + 1, SelfInv.ProductName) = Split(sKey, DELIMITER)(1)    'product name
        arrOut(i + 1, SelfInv.ProductSeries) = Split(sKey, DELIMITER)(2)    'product seriese
        arrOut(i + 1, SelfInv.ProductUnit) = fGetProductUnit(arrOut(i + 1, 1), arrOut(i + 1, 2), arrOut(i + 1, 3))    'unit
        'arrOut(i + 1, 5) = Split(sKey, DELIMITER)(3)    'lotnum
        
        arrOut(i + 1, SelfInv.Qty) = dblPurchaseQty - dblSellQty
        arrOut(i + 1, SelfInv.Price) = CDbl(Split(dictSelfPurchaseOD(sKey), DELIMITER)(1))
    Next
    
    'fCalculateSelfInventory = arrOut
    Call fAppendArray2Sheet(shtSelfInventory, arrOut)
    Erase arrOut
End Function

'====================== Self Purchase Order =================================================================
Private Function fReadSheetSelfPurchaseOrder2Dictionary()
    Dim arrData()
    Dim dictColIndex As Dictionary
    
    fSortDataInSheetSortSheetData shtSelfPurchaseOrder, Array(SelfPurchase.ProductProducer, SelfPurchase.ProductName, SelfPurchase.ProductSeries)
    Call fReadSheetDataByConfig("SELF_PURCHASE_ORDER", dictColIndex, arrData, , , , , shtSelfPurchaseOrder)
    'Call fValidateDuplicateInArray(arrData, Array(dictColIndex("ProductProducer"), dictColIndex("ProductName"), dictColIndex("ProductSeries")), False, shtSelfPurchaseOrder, 1, 1, "厂家 + 名称 + 规格")
    
    Dim lEachRow As Long
    Dim sProducer As String
    Dim sProductName As String
    Dim sProductSeries As String
    'Dim sLotNum As String
    Dim sKey As String
    Dim dictRowNoTmp As Dictionary
    
    Set dictSelfPurchaseOD = New Dictionary
    Set dictRowNoTmp = New Dictionary
    
    For lEachRow = LBound(arrData, 1) To UBound(arrData, 1)
        sProducer = Trim(arrData(lEachRow, dictColIndex("ProductProducer")))
        sProductName = Trim(arrData(lEachRow, dictColIndex("ProductName")))
        sProductSeries = Trim(arrData(lEachRow, dictColIndex("ProductSeries")))
'        sLotNum = Trim(arrData(lEachRow, dictColIndex("LotNum")))
        
        'sKey = sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries & DELIMITER & sLotNum
        sKey = sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries '& DELIMITER & sLotNum
        
        If Not dictSelfPurchaseOD.Exists(sKey) Then
            dictSelfPurchaseOD.Add sKey, CDbl(arrData(lEachRow, dictColIndex("PurchaseQuantity")))
        Else
            dictSelfPurchaseOD(sKey) = dictSelfPurchaseOD(sKey) + CDbl(arrData(lEachRow, dictColIndex("PurchaseQuantity")))
        End If
        If Len(Trim(arrData(lEachRow, dictColIndex("PurchasePrice")))) <= 0 Then
            If Not dictRowNoTmp.Exists(sKey) Then dictRowNoTmp(sKey) = 0
        Else
            dictRowNoTmp(sKey) = arrData(lEachRow, dictColIndex("PurchasePrice"))
        End If
    Next
    
    Dim i As Long
    For i = 0 To dictSelfPurchaseOD.count - 1
        sKey = dictSelfPurchaseOD.Keys(i)
        
        dictSelfPurchaseOD(sKey) = dictSelfPurchaseOD(sKey) & DELIMITER & dictRowNoTmp(sKey)
    Next
    
    Set dictColIndex = Nothing
    Set dictRowNoTmp = Nothing
End Function
Public Function fLotNumExistsInSelfPurchaseOrder(sProductProducer As String, sProductName As String, sProductSeries As String, sLotNum As String) As Boolean
    If dictSelfPurchaseOD Is Nothing Then Call fReadSheetSelfPurchaseOrder2Dictionary
    
    fLotNumExistsInSelfPurchaseOrder = dictSelfPurchaseOD.Exists(sProductProducer & DELIMITER & sProductName & DELIMITER & sProductSeries & DELIMITER & sLotNum)
End Function
'------------------------------------------------------------------------------


'====================== Self Sales Order =================================================================
Private Function fReadSheetSelfSalesOrder2Dictionary()
    Dim arrData()
    Dim dictColIndex As Dictionary
    
    Call fRemoveFilterForSheet(shtSelfSalesOrder)
    Call fReadSheetDataByConfig("SELF_SALES_ORDER", dictColIndex, arrData, , , , , shtSelfSalesOrder)
    'Call fValidateDuplicateInArray(arrData, Array(dictColIndex("ProductProducer"), dictColIndex("ProductName"), dictColIndex("ProductSeries")), False, shtSelfSalesOrder, 1, 1, "厂家 + 名称 + 规格")
    
    Dim lEachRow As Long
    Dim sProducer As String
    Dim sProductName As String
    Dim sProductSeries As String
   ' Dim sLotNum As String
    Dim sKey As String
    Dim dictRowNoTmp As Dictionary
    
    Set dictSelfSalesOD = New Dictionary
    Set dictRowNoTmp = New Dictionary
    
    For lEachRow = LBound(arrData, 1) To UBound(arrData, 1)
        sProducer = Trim(arrData(lEachRow, dictColIndex("ProductProducer")))
        sProductName = Trim(arrData(lEachRow, dictColIndex("ProductName")))
        sProductSeries = Trim(arrData(lEachRow, dictColIndex("ProductSeries")))
        'sLotNum = Trim(arrData(lEachRow, dictColIndex("LotNum")))
        
        'sKey = sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries & DELIMITER & sLotNum
        sKey = sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries ' & DELIMITER & sLotNum
        
        If Not dictSelfSalesOD.Exists(sKey) Then
            dictSelfSalesOD.Add sKey, CDbl(arrData(lEachRow, dictColIndex("SellQuantity")))
        Else
            dictSelfSalesOD(sKey) = dictSelfSalesOD(sKey) + CDbl(arrData(lEachRow, dictColIndex("SellQuantity")))
        End If
        dictRowNoTmp(sKey) = (lEachRow + 1) & DELIMITER & arrData(lEachRow, dictColIndex("SellPrice"))
    Next
    
    Dim i As Long
    For i = 0 To dictSelfSalesOD.count - 1
        sKey = dictSelfSalesOD.Keys(i)
        
        dictSelfSalesOD(sKey) = dictSelfSalesOD(sKey) & DELIMITER & dictRowNoTmp(sKey)
    Next
    
    Set dictColIndex = Nothing
    Set dictRowNoTmp = Nothing
End Function
'Function fLotNumExistsInSelfSalesOrder(sProductProducer As String, sProductName As String, sProductSeries As String, sLotNum As String) As Boolean
'    If dictSelfSalesOD Is Nothing Then Call fReadSheetSelfSalesOrder2Dictionary
'
'    fLotNumExistsInSelfSalesOrder = dictSelfSalesOD.Exists(sProductProducer & DELIMITER & sProductName & DELIMITER & sProductSeries & DELIMITER & sLotNum)
'End Function
'------------------------------------------------------------------------------


'====================== Self Sales Order - By Year Month =================================================================
Function fReadSheetSelfSalesOrderByYearMonth(adYearMonth As String) As Dictionary
'adYearMonth : 201803
    Dim dictOut As Dictionary
    
    Dim arrData()
    Dim dictColIndex As Dictionary
    
    Call fRemoveFilterForSheet(shtSelfSalesOrder)
    Call fReadSheetDataByConfig("SELF_SALES_ORDER", dictColIndex, arrData, , , , , shtSelfSalesOrder)
    'Call fValidateDuplicateInArray(arrData, Array(dictColIndex("ProductProducer"), dictColIndex("ProductName"), dictColIndex("ProductSeries")), False, shtSelfSalesOrder, 1, 1, "厂家 + 名称 + 规格")
    
    Dim lEachRow As Long
    Dim sProducer As String
    Dim sProductName As String
    Dim sProductSeries As String
    'Dim sLotNum As String
    Dim dtSalesDate As Date
    Dim sKey As String
    Dim dictRowNoTmp As Dictionary
    
    Set dictOut = New Dictionary
    Set dictRowNoTmp = New Dictionary
    
    For lEachRow = LBound(arrData, 1) To UBound(arrData, 1)
        sProducer = Trim(arrData(lEachRow, dictColIndex("ProductProducer")))
        sProductName = Trim(arrData(lEachRow, dictColIndex("ProductName")))
        sProductSeries = Trim(arrData(lEachRow, dictColIndex("ProductSeries")))
       ' sLotNum = Trim(arrData(lEachRow, dictColIndex("LotNum")))
        dtSalesDate = arrData(lEachRow, dictColIndex("SalesDate"))
        
        If Format(dtSalesDate, "YYYYMM") = adYearMonth Then
            'sKey = sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries & DELIMITER & sLotNum
            sKey = sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries
            
            If Not dictOut.Exists(sKey) Then
                dictOut.Add sKey, CDbl(arrData(lEachRow, dictColIndex("SellQuantity")))
            Else
                dictOut(sKey) = dictOut(sKey) + CDbl(arrData(lEachRow, dictColIndex("SellQuantity")))
            End If
            dictRowNoTmp(sKey) = (lEachRow + 1) & DELIMITER & arrData(lEachRow, dictColIndex("SellPrice"))
        End If
    Next
    
    Dim i As Long
    For i = 0 To dictOut.count - 1
        sKey = dictOut.Keys(i)
        
        dictOut(sKey) = dictOut(sKey) & DELIMITER & dictRowNoTmp(sKey)
    Next
    
    Set dictColIndex = Nothing
    Set dictRowNoTmp = Nothing
    
    Set fReadSheetSelfSalesOrderByYearMonth = dictOut
    Set dictOut = Nothing
End Function
'------------------------------------------------------------------------------

Function fCheckIfSelfSellAmountIsGreaterThanPurchaseByLotNumber(arrData, iColProducer As Integer, iColProductName As Integer, iColProductSeries As Integer _
                                        , iColLotNum As Integer, Optional ByRef alErrRowNo As Long, Optional ByRef alErrColNo As Long)
    Dim i As Long
    Dim sKey As String
    
    Call fRemoveFilterForSheet(shtSelfPurchaseOrder)
    
    If dictSelfPurchaseOD Is Nothing Then Call fReadSheetSelfPurchaseOrder2Dictionary
    If dictSelfSalesOD Is Nothing Then Call fReadSheetSelfSalesOrder2Dictionary
    
    For i = 0 To dictSelfSalesOD.count - 1
        sKey = dictSelfSalesOD.Keys(i)
        
        If Not dictSelfPurchaseOD.Exists(sKey) Then
            'fErr "【药品+批号】不存在于本公司进货表中" & vbCr & "行号：" & Split(dictSelfSalesOD(sKey), DELIMITER)(1) & vbCr & vbCr & sKey
            fErr "【药品】不存在于本公司进货表中" & vbCr & "行号：" & Split(dictSelfSalesOD(sKey), DELIMITER)(1) & vbCr & vbCr & sKey
        Else
            If CDbl(Split(dictSelfSalesOD(sKey), DELIMITER)(0)) > CDbl(Split(dictSelfPurchaseOD(sKey), DELIMITER)(0)) Then
                alErrRowNo = Split(dictSelfSalesOD(sKey), DELIMITER)(1)
                alErrColNo = iColLotNum
                'fErr "【药品+批号】的总出货数量大于本公司进货表中的进货总数量" & vbCr & vbCr & sKey & vbCr & "行号：" & alErrRowNo
                fErr "【药品】的总出货数量大于本公司进货表中的进货总数量" & vbCr & vbCr & sKey & vbCr & "行号：" & alErrRowNo
                Exit For
            End If
        End If
    Next
End Function

