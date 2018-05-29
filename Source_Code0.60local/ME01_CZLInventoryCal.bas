Attribute VB_Name = "ME01_CZLInventoryCal"
Option Explicit
Option Base 1


Private dictCZLRolloverInv As Dictionary
Private dictCZLSales2Companies As Dictionary
Private dictCZLSales2Hospital As Dictionary
Private sYearMonth As String

Function fResetdictCZLSalesHospital()
    Set dictCZLSales2Hospital = Nothing
End Function
Function fResetdictCZLSales2Companies()
    Set dictCZLSales2Companies = Nothing
End Function

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

Sub subMain_CalCZLInventory()
'    If fGetReplaceUnifyCZLSales2CompaniesErrorRowCount > 0 Then
'        fMsgBox "采芝林的销售流向(商业公司)中有药品在系统中找不到，无法计算库存，请先处理这些错误。" & vbCr & "并进行替换统一后再进行库存计算。"
'        shtCZLSales2Companies.Visible = xlSheetVisible
'        shtException.Visible = xlSheetVisible:         shtException.Activate
'        End
'    End If
'    If fGetReplaceUnifyErrorRowCount_SCompSalesInfo > 0 Then
'        fMsgBox "销售流向数据中有药品在系统中找不到，无法计算利润和佣金，请先处理这些错误。" & vbCr & "并进行替换统一后再进行库存计算。"
'        shtSalesInfos.Visible = xlSheetVisible
'        shtException.Visible = xlSheetVisible:         shtException.Activate
'        End
'    End If
    
    Call fInitialization
    
    If Not fIsDev() Then On Error GoTo err_handle
    
    fCheckIfErrCountNotZero_SCompSalesInfo
    fCheckIfErrCountNotZero_CZLSales2Comp
    
    Call fValidaterngYearMonth(sYearMonth)
    
    If Not fPromptToConfirmToContinue("你当前输入的年月是：【" & sYearMonth & "】," & vbCr & "你确定这个年月正确吗？" _
                & vbCr & vbCr & "因为" & vbCr & vbCr _
                & "1.采芝林的采购入库会用我们公司的这个月的出库，如果所填的月份不正确，库存计算就会不正确" & vbCr _
                & "2.采芝林的销售出库到商业公司和医院的销售流向文件请选择这个月的文件，如果与所填的月份不一致，库存计算就会不正确" _
    ) Then fErr
    
    gsRptID = "CALCULATE_PROFIT"
    Call fReadSysConfig_InputTxtSheetFile
    
    fVeryHideSheet shtException
    'Call fCleanSheetOutputResetSheetOutput(shtException)
    Call fDeleteRowsFromSheetLeaveHeader(shtException)
    Call fRemoveFilterForSheet(shtCZLInventory)
    fClearContentLeaveHeader shtCZLInventory
    
    Call fCalculateCZLInventory
'    fBasicCosmeticFormatSheet shtCZLInventory
'    fSetBorderLineForSheet shtCZLInventory
    fActiveVisibleSwitchSheet shtCZLInventory, , False
    Application.Goto shtCZLInventory.Range("A2"), True
    
    fSortDataInSheetSortSheetData shtCZLInventory, Array(CZLInv.ProductProducer, CZLInv.ProductName, CZLInv.ProductSeries)
    
    If fZero(gsBusinessErrorMsg) Then fMsgBox "【采芝林】库存计算完成！", vbInformation
err_handle:
    If fCheckIfUnCapturedExceptionAbnormalError Then GoTo reset_excel_options
    
    If shtException.Visible = xlSheetVisible Then
        Dim lExcepMaxCol As Long
        lExcepMaxCol = fGetValidMaxCol(shtException)
        Call fSetFormatBoldOrangeBorderForHeader(shtException, lExcepMaxCol)
        Call fSetBorderLineForSheet(shtException, lExcepMaxCol)
        Call fBasicCosmeticFormatSheet(shtException, lExcepMaxCol)
        Call fSetFormatForOddEvenLineByFixColor(shtException, lExcepMaxCol)
        shtException.Activate
    End If
    
    If fCheckIfGotBusinessError Then
        GoTo reset_excel_options
    End If
    
    If fCheckIfUnCapturedExceptionAbnormalError Then
        GoTo reset_excel_options
    End If
reset_excel_options:
    Err.Clear
    fClearRefVariables
    fEnableExcelOptionsAll
    'End
End Sub

Private Function fCalculateCZLInventory()
    Dim i As Long
    Dim lEachRow As Long
    Dim sKey As String
    Dim sProducer As String
    Dim sProductName As String
    Dim sProductSeries As String
    Dim sLotNum As String
    Dim dblRolloverQty As Double
    Dim dblPurchaseQty As Double
    Dim dblSellQty As Double
    Dim arrOut()
    
    Dim dictSelfPurchaseOD As Dictionary
    Dim dictSelfSalesOD As Dictionary

    Call fRemoveFilterForSheet(shtCZLRolloverInv)       'rollover inventory
    Call fRemoveFilterForSheet(shtSelfSalesOrder)       'purchase
    Call fRemoveFilterForSheet(shtCZLSales2Companies)       'sales
    
'    If Not shtSelfSalesOrder.fValidateSheet(False) Then fErr    'purchase
    'If Not shtCZLSales2Companies.fValidateSheet(False) Then fErr     'sales
    
    Set dictSelfSalesOD = fReadSheetSelfSalesOrderByYearMonth(sYearMonth)              'SelfSales = purchase
    Call fReadSheetCZLSalesToCompanies2Dictionary   ' sales to companies   dictCZLSales2Companies
    Call fReadSheetCZLSalesToHospital2Dictionary    ' sales to hospital    dictCZLSales2Hospital
    Call fReadCZLRolloverInventory2Dictionary      'dictCZLRolloverInv
    
    '================= verify lot number ========================================
'    Dim dictMissedLot As Dictionary
'    Set dictMissedLot = New Dictionary
    
    For i = 0 To dictCZLSales2Hospital.Count - 1
        sKey = dictCZLSales2Hospital.Keys(i)
        
        If Not dictSelfSalesOD.Exists(sKey) Then
           ' dictMissedLot.Add sKey, 0
            dictSelfSalesOD.Add sKey, "0|0|0"     'add for output
        End If
    Next
    For i = 0 To dictCZLSales2Companies.Count - 1
        sKey = dictCZLSales2Companies.Keys(i)
        
        If Not dictSelfSalesOD.Exists(sKey) Then
           ' If Not dictMissedLot.Exists(sKey) Then dictMissedLot.Add sKey, 0
            dictSelfSalesOD.Add sKey, "0|0|0"     'add for output
        End If
    Next
    
    'rollover
    For i = 0 To dictCZLRolloverInv.Count - 1
        sKey = dictCZLRolloverInv.Keys(i)
        
        If Not dictSelfSalesOD.Exists(sKey) Then
            dictSelfSalesOD.Add sKey, "0|0|0"     'add for output
        End If
    Next
    
'    If dictMissedLot.Count > 0 Then
'        Call fAddMissedSelfSalesLotNumToSheetException(dictMissedLot)
'        'fErr gsBusinessErrorMsg
''        fMsgBox gsBusinessErrorMsg
'    End If
'
'    Set dictMissedLot = Nothing
    '---------------------------------------------------------------
    
    '================= calculate inventory of this month  ========================================
    ReDim arrOut(1 To dictSelfSalesOD.Count, 7) 'SelfSales = purchase
    
    For i = 0 To dictSelfSalesOD.Count - 1  'SelfSales = purchase
        sKey = dictSelfSalesOD.Keys(i)
        
        dblPurchaseQty = CDbl(Split(dictSelfSalesOD(sKey), DELIMITER)(0))
        dblSellQty = 0
        dblRolloverQty = 0
        
        If dictCZLSales2Hospital.Exists(sKey) Then
            dblSellQty = CDbl(Split(dictCZLSales2Hospital(sKey), DELIMITER)(0))
        End If
        
        If dictCZLSales2Companies.Exists(sKey) Then
            dblSellQty = dblSellQty + CDbl(Split(dictCZLSales2Companies(sKey), DELIMITER)(0))
        End If
        
        If dictCZLRolloverInv.Exists(sKey) Then
            dblRolloverQty = dictCZLRolloverInv(sKey)
        End If
            
        arrOut(i + 1, CZLInv.ProductProducer) = Split(sKey, DELIMITER)(0)
        arrOut(i + 1, CZLInv.ProductName) = Split(sKey, DELIMITER)(1)
        arrOut(i + 1, CZLInv.ProductSeries) = Split(sKey, DELIMITER)(2)
        arrOut(i + 1, CZLInv.ProductUnit) = fGetProductUnit(arrOut(i + 1, 1), arrOut(i + 1, 2), arrOut(i + 1, 3))
        'arrOut(i + 1, 5) = "'" & Split(sKey, DELIMITER)(3)  'lot num
        
        arrOut(i + 1, CZLInv.InventoryQty) = dblPurchaseQty - dblSellQty + dblRolloverQty
        
'        If IsNumeric(Split(dictSelfSalesOD(sKey), DELIMITER)(2)) Then
'            arrOut(i + 1, CZLInv.PurchasePrice) = CDbl(Split(dictSelfSalesOD(sKey), DELIMITER)(2))     'purcahse price
'        Else
'        End If
    Next
    '---------------------------------------------------------------
    
    Set dictSelfSalesOD = Nothing
    Set dictCZLRolloverInv = Nothing
    
    'fCalculateCZLInventory = arrOut
    Call fAppendArray2Sheet(shtCZLInventory, arrOut)
    Erase arrOut
    Set dictCZLSales2Hospital = Nothing
End Function




'====================== CZL Sales TO companies =================================================================
Function fReadSheetCZLSalesToCompanies2Dictionary()
    Dim arrData()
    Dim dictColIndex As Dictionary
    
    Call fRemoveFilterForSheet(shtSelfSalesOrder)
    Call fReadSheetDataByConfig("CZL_SALES_TO_COMPANIES", dictColIndex, arrData, , , , , shtCZLSales2Companies)
    'Call fValidateDuplicateInArray(arrData, Array(dictColIndex("ProductProducer"), dictColIndex("ProductName"), dictColIndex("ProductSeries")), False, shtSelfSalesOrder, 1, 1, "厂家 + 名称 + 规格")
    
    Dim lEachRow As Long
    Dim sProducer As String
    Dim sProductName As String
    Dim sProductSeries As String
    'Dim sLotNum As String
    Dim sKey As String
    Dim dictRowNoTmp As Dictionary
    
    Set dictCZLSales2Companies = New Dictionary
    Set dictRowNoTmp = New Dictionary
    
    For lEachRow = LBound(arrData, 1) To UBound(arrData, 1)
        sProducer = Trim(arrData(lEachRow, dictColIndex("MatchedProductProducer")))
        sProductName = Trim(arrData(lEachRow, dictColIndex("MatchedProductName")))
        sProductSeries = Trim(arrData(lEachRow, dictColIndex("MatchedProductSeries")))
        'sLotNum = Trim(arrData(lEachRow, dictColIndex("LotNum")))
        
        'sKey = sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries & DELIMITER & sLotNum
        sKey = sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries
        
        If Not dictCZLSales2Companies.Exists(sKey) Then
            dictCZLSales2Companies.Add sKey, CDbl(arrData(lEachRow, dictColIndex("Quantity")))
        Else
            dictCZLSales2Companies(sKey) = dictCZLSales2Companies(sKey) + CDbl(arrData(lEachRow, dictColIndex("Quantity")))
        End If
        dictRowNoTmp(sKey) = (lEachRow + 1) & DELIMITER & arrData(lEachRow, dictColIndex("SellPrice"))
    Next
    
    Dim i As Long
    For i = 0 To dictCZLSales2Companies.Count - 1
        sKey = dictCZLSales2Companies.Keys(i)
        
        dictCZLSales2Companies(sKey) = dictCZLSales2Companies(sKey) & DELIMITER & dictRowNoTmp(sKey)
    Next
    
    Set dictColIndex = Nothing
    Set dictRowNoTmp = Nothing
End Function

Function fReadSheetCZLSalesToHospital2Dictionary()
    Dim arrData()
    Dim dictColIndex As Dictionary
    
    Call fRemoveFilterForSheet(shtSelfSalesOrder)
    Call fReadSheetDataByConfig("CZL_SALES_TO_HOSPITAL", dictColIndex, arrData, , , , , shtSalesInfos)
    'Call fValidateDuplicateInArray(arrData, Array(dictColIndex("ProductProducer"), dictColIndex("ProductName"), dictColIndex("ProductSeries")), False, shtSelfSalesOrder, 1, 1, "厂家 + 名称 + 规格")
    
    Dim sCZLCompName As String
    Dim lEachRow As Long
    Dim sProducer As String
    Dim sProductName As String
    Dim sProductSeries As String
    'Dim sLotNum As String
    Dim sKey As String
    Dim dictRowNoTmp As Dictionary
    
    'sCZLCompName = fGetCompany_CompanyName("CZL")
    sCZLCompName = fGetCompanyNameByID_Common("CZL")
    'If sCZLCompName <> fGetCompanyNameByID_Common("CZL") Then fErr "CZL的名字设置不一致： [Sales Company List - Common Importing - Sales File] 和 [Sales Company List]"
    
    Set dictCZLSales2Hospital = New Dictionary
    Set dictRowNoTmp = New Dictionary
    
    For lEachRow = LBound(arrData, 1) To UBound(arrData, 1)
        If Trim(arrData(lEachRow, dictColIndex("SalesCompanyName"))) = sCZLCompName Then
            sProducer = Trim(arrData(lEachRow, dictColIndex("MatchedProductProducer")))
            sProductName = Trim(arrData(lEachRow, dictColIndex("MatchedProductName")))
            sProductSeries = Trim(arrData(lEachRow, dictColIndex("MatchedProductSeries")))
            'sLotNum = Trim(arrData(lEachRow, dictColIndex("LotNum")))
            
            'sKey = sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries & DELIMITER & sLotNum
            sKey = sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries
            
            If Not dictCZLSales2Hospital.Exists(sKey) Then
                dictCZLSales2Hospital.Add sKey, CDbl(arrData(lEachRow, dictColIndex("Quantity")))
            Else
                dictCZLSales2Hospital(sKey) = dictCZLSales2Hospital(sKey) + CDbl(arrData(lEachRow, dictColIndex("Quantity")))
            End If
            
            dictRowNoTmp(sKey) = (lEachRow + 1) & DELIMITER & arrData(lEachRow, dictColIndex("SellPrice"))
        End If
    Next
    
    Dim i As Long
    For i = 0 To dictCZLSales2Hospital.Count - 1
        sKey = dictCZLSales2Hospital.Keys(i)
        
        dictCZLSales2Hospital(sKey) = dictCZLSales2Hospital(sKey) & DELIMITER & dictRowNoTmp(sKey)
    Next
    
    Set dictColIndex = Nothing
    Set dictRowNoTmp = Nothing
End Function
'------------------------------------------------------------------------------

Private Function fAddMissedSelfSalesLotNumToSheetException(dictMissedLotNum As Dictionary)
    Dim arrMissedLotNum()
    'Dim sErr As String
    Dim lRecCount As Long
    Dim lStartRow As Long
    
    lRecCount = fGetDictionayDelimiteredItemsCount(dictMissedLotNum)
    If lRecCount > 0 Then
        lStartRow = fGetshtExceptionNewRow
        
        arrMissedLotNum = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictMissedLotNum, , False)
        
        Call fPrepareHeaderToSheet(shtException, Array("药品厂家", "药品名称", "药品规格", "不存在的批号"), lStartRow)
        shtException.Rows(lStartRow).Font.Color = RGB(255, 0, 0)
        shtException.Rows(lStartRow).Font.Bold = True
        fGetRangeByStartEndPos(shtException, lStartRow + 1, 4, lStartRow + UBound(arrMissedLotNum, 1), 4).NumberFormat = "@"
        Call fAppendArray2Sheet(shtException, arrMissedLotNum)
        'sErr = fUbound(arrMissedLotNum)
        
'        shtException.Cells(lStartRow + 1, 8).Resize(dictMissedLotNum.Count, 1).Value = fConvertDictionaryItemsTo2DimenArrayForPaste(dictMissedLotNum, False)
       ' Erase arrMissedLotNum
        Call fFreezeSheet(shtException)
                
        fShowAndActiveSheet shtException
        If fNzero(gsBusinessErrorMsg) Then gsBusinessErrorMsg = gsBusinessErrorMsg & vbCr & vbCr & vbCr & "===============================" & vbCr & vbCr

        gsBusinessErrorMsg = gsBusinessErrorMsg & lRecCount & "个药品的出货的【批号】在本公司的出货表中找不到，" & vbCr _
            & "请检查【本公司的出货】表"
    End If
End Function

Function fValidaterngYearMonth(ByRef sYearMonth As String)
    'Dim sYearMonth As String
    sYearMonth = Trim(shtMainMenu.Range("rngYearMonth").Value)
    
    If Not fIsDate(sYearMonth & "01", "YYYYMMDD") Then
        fErr "输入了错误的年和月，请输入YYYYMMDD"
    End If
End Function

Function fReadCZLRolloverInventory2Dictionary()
    Dim arrData()
    
    Call fCopyReadWholeSheetData2Array(shtCZLRolloverInv, arrData)
    
'    Set dictCZLRolloverInv = fReadArray2DictionaryWithMultipleKeyColsSingleItemColSum(arrData _
'                            , Array(CZLRollover.ProductProducer, CZLRollover.ProductName, CZLRollover.ProductSeries, CZLRollover.LotNum) _
'                            , CLng(CZLRollover.RolloverQty), DELIMITER)
    Set dictCZLRolloverInv = fReadArray2DictionaryWithMultipleKeyColsSingleItemColSum(arrData _
                            , Array(CZLRollover.ProductProducer, CZLRollover.ProductName, CZLRollover.ProductSeries) _
                            , CLng(CZLRollover.RolloverQty), DELIMITER)
    Erase arrData
End Function


'====================== CZL Sales TO companies Except CZL =================================================================
Function fReadSheetCZLSalesToCompaniesByCompaniesExceptCZL() As Dictionary
    Dim dictOut As Dictionary
    
    Dim arrData()
    Dim dictColIndex As Dictionary
    
    Call fRemoveFilterForSheet(shtCZLSales2Companies)
    Call fReadSheetDataByConfig("CZL_SALES_TO_COMPANIES", dictColIndex, arrData, , , , , shtCZLSales2Companies)
    'Call fValidateDuplicateInArray(arrData, Array(dictColIndex("ProductProducer"), dictColIndex("ProductName"), dictColIndex("ProductSeries")), False, shtSelfSalesOrder, 1, 1, "厂家 + 名称 + 规格")
    
    Dim sCZLCompName As String
    Dim lEachRow As Long
    Dim sSalesCompany As String
    Dim sProducer As String
    Dim sProductName As String
    Dim sProductSeries As String
    'Dim sLotNum As String
    Dim sKey As String
    Dim dictRowNoTmp As Dictionary
    
    Set dictOut = New Dictionary
    Set dictRowNoTmp = New Dictionary
    
    sCZLCompName = fGetCompanyNameByID_Common("CZL")
    
    For lEachRow = LBound(arrData, 1) To UBound(arrData, 1)
        sSalesCompany = Trim(arrData(lEachRow, dictColIndex("MatchedCompanyName")))
        
        If sSalesCompany = sCZLCompName Then GoTo next_row
        
        sProducer = Trim(arrData(lEachRow, dictColIndex("MatchedProductProducer")))
        sProductName = Trim(arrData(lEachRow, dictColIndex("MatchedProductName")))
        sProductSeries = Trim(arrData(lEachRow, dictColIndex("MatchedProductSeries")))
'        sLotNum = Trim(arrData(lEachRow, dictColIndex("LotNum")))
        
        'sKey = sSalesCompany & DELIMITER & sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries & DELIMITER & sLotNum
        sKey = sSalesCompany & DELIMITER & sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries
        
        If Not dictOut.Exists(sKey) Then
            dictOut.Add sKey, CDbl(arrData(lEachRow, dictColIndex("ConvertQuantity")))
        Else
            dictOut(sKey) = dictOut(sKey) + CDbl(arrData(lEachRow, dictColIndex("ConvertQuantity")))
        End If
        dictRowNoTmp(sKey) = (lEachRow + 1) & DELIMITER & arrData(lEachRow, dictColIndex("SellPrice"))
next_row:
    Next
    
    Dim i As Long
    For i = 0 To dictOut.Count - 1
        sKey = dictOut.Keys(i)
        
        dictOut(sKey) = dictOut(sKey) & DELIMITER & dictRowNoTmp(sKey)
    Next
    
    Set dictColIndex = Nothing
    Set dictRowNoTmp = Nothing
    
    Set fReadSheetCZLSalesToCompaniesByCompaniesExceptCZL = dictOut
    Set dictOut = Nothing
End Function

Function fPrepareCZLSales2HospitalByFiltering()
    fActiveVisibleSwitchSheet shtSalesInfos
    
    Dim lMaxCol As Long
    Dim lMaxRow As Long
    Dim sCZLName As String

    If shtSalesInfos.AutoFilterMode Then  'auto filter
        shtSalesInfos.AutoFilter.ShowAllData
    Else
        fGetRangeByStartEndPos(shtSalesInfos, 1, 1, 1, lMaxCol).AutoFilter
    End If
    
    lMaxCol = shtSalesInfos.Cells(1, 1).End(xlToRight).Column
    lMaxRow = fGetValidMaxRow(shtSalesInfos)
    
    If lMaxRow < 2 Then GoTo err_handling
    
    sCZLName = fGetCompanyNameByID_Common("CZL")
    
    fGetRangeByStartEndPos(shtSalesInfos, 1, 1, lMaxRow, lMaxCol).AutoFilter _
                Field:=Sales2Hospital.SalesCompany, Criteria1:="=" & sCZLName, Operator:=xlAnd
    
    If Not fSheetHasDataAfterFilter(shtSalesInfos, , lMaxRow, lMaxCol) Then GoTo err_handling
    
    Exit Function
err_handling:
    fMsgBox "采芝林没有销售流向，请导入采芝林销售流向文件。"
End Function

Function fPrepareCZLPurchaseFromSelfSales()
    Dim arrSelf()
    
    fRemoveFilterForSheet shtSelfSalesOrder
    
    fClearContentLeaveHeader shtCZLPurchaseOrder
    Call fSortDataInSheetSortSheetData(shtSelfSalesOrder, Array(SelfSales.ProductProducer, SelfSales.ProductName, SelfSales.ProductSeries))
    Call fCopyReadWholeSheetData2Array(shtSelfSalesOrder, arrSelf, , , fLetter2Num("H"))
    Call fWriteArray2Sheet(shtCZLPurchaseOrder, arrSelf)
    Erase arrSelf
    fActiveVisibleSwitchSheet shtCZLPurchaseOrder
End Function


