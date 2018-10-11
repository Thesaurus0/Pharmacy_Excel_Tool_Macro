Attribute VB_Name = "ME02_CZLInventoryCompare"
Option Explicit
Option Base 1

Sub subMain_CompareCZLInventory()
    If Not fIsDev() Then On Error GoTo error_handling
    
    fCheckIfErrCountNotZero_SalesInventory
    fCheckIfErrCountNotZero_CZLSales2Comp
    fCheckIfErrCountNotZero_SCompSalesInfo
    
    fRemoveFilterForSheet shtCZLInventory
    fRemoveFilterForSheet shtSalesCompInvUnified
    fRemoveFilterForSheet shtCZLInvDiff
    fDeleteRowsFromSheetLeaveHeader shtCZLInvDiff
    
    fInitialization
    
    gsRptID = "COMPARE_CZL_INVENTORY"
    Call fReadSysConfig_InputTxtSheetFile
    
    gsRptFilePath = fReadSysConfig_Output(, gsRptType)

    Call fCopyReadWholeSheetData2Array(shtSalesCompInvUnified, arrMaster)
    
    Dim dictCZLInformedInv As Dictionary
    Set dictCZLInformedInv = fReadSCompUnifiedInvSumQuantityByCZL(True)
    Erase arrMaster
    
    Dim dictCZLCalInv As Dictionary
    Dim arrCZLInv()
    Call fCopyReadWholeSheetData2Array(shtCZLInventory, arrCZLInv)
'    Set dictCZLCalInv = fReadArray2DictionaryWithMultipleKeyColsSingleItemColSum(arrCZLInv _
'                            , Array(CZLInv.ProductProducer, CZLInv.ProductName, CZLInv.ProductSeries, CZLInv.LotNum) _
'                            , CZLInv.InventoryQty, DELIMITER)
    Set dictCZLCalInv = fReadArray2DictionaryWithMultipleKeyColsSingleItemColSum(arrCZLInv _
                            , Array(CZLInv.ProductProducer, CZLInv.ProductName, CZLInv.ProductSeries) _
                            , CZLInv.InventoryQty, DELIMITER)
    Erase arrCZLInv
    
    Dim dictInventoryDiff As Dictionary
    Set dictInventoryDiff = fCompare2Inventory(dictCZLInformedInv, dictCZLCalInv)
    Set dictCZLCalInv = Nothing
    Set dictCZLInformedInv = Nothing
    
    arrOutput = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictInventoryDiff, , False)
    Call fAppendArray2Sheet(shtCZLInvDiff, arrOutput)
    arrOutput = fConvertDictionaryDelimiteredItemsTo2DimenArrayForPaste(dictInventoryDiff)
    shtCZLInvDiff.Cells(2, CZLInvDiff.InformedQty).Resize(UBound(arrOutput, 1), UBound(arrOutput, 2)).Value = arrOutput
    
    Call fSortDataInSheetSortSheetData(shtCZLInvDiff, Array(CZLInvDiff.ProductProducer, CZLInvDiff.ProductName, CZLInvDiff.ProductSeries, CZLInvDiff.LotNum))
    
    Call fFormatOutputSheet(shtCZLInvDiff)
    
    shtCZLInvDiff.Rows(1).RowHeight = 25
    shtCZLInvDiff.Visible = xlSheetVisible
    shtCZLInvDiff.Activate
    Application.Goto shtCZLInvDiff.Range("A" & fGetValidMaxRow(shtCZLInvDiff)), True
error_handling:
    If fCheckIfUnCapturedExceptionAbnormalError Then GoTo reset_excel_options
    
    If fCheckIfGotBusinessError Then
        GoTo reset_excel_options
    End If
    
    If fCheckIfUnCapturedExceptionAbnormalError Then
        GoTo reset_excel_options
    End If
     
    fGotoCell shtCZLInvDiff.Range("A2")
    fMsgBox "采芝林库存差异计算结果在表：[" & shtCZLInvDiff.Name & "] 中，请检查！", vbInformation
    
reset_excel_options:
    Err.Clear
    fClearRefVariables
    fEnableExcelOptionsAll
    'End
End Sub

Private Function fCompare2Inventory(dictCZLInformedInv As Dictionary, dictCZLCalInv As Dictionary) As Dictionary
    Dim dictOut As Dictionary
     
    Dim i As Long
    Dim sProdLotKey As String
    Dim dblInformedInv As Double
    Dim dblCalculatedInv As Double
    
    Set dictOut = New Dictionary
    
    For i = 0 To dictCZLInformedInv.Count - 1
        sProdLotKey = dictCZLInformedInv.Keys(i)
        dblInformedInv = CDbl(Split(dictCZLInformedInv.Items(i), DELIMITER)(0))
        
        If Not dictCZLCalInv.Exists(sProdLotKey) Then
            dictOut.Add sProdLotKey, dblInformedInv & DELIMITER & "0" & DELIMITER & dblInformedInv
        Else
            dblCalculatedInv = dictCZLCalInv(sProdLotKey)
            dictOut.Add sProdLotKey, dblInformedInv & DELIMITER & dblCalculatedInv & DELIMITER & (dblInformedInv - dblCalculatedInv)
        End If
    Next
    
    
    For i = 0 To dictCZLCalInv.Count - 1
        sProdLotKey = dictCZLCalInv.Keys(i)
        dblCalculatedInv = CDbl(dictCZLCalInv.Items(i))
        
        If Not dictCZLInformedInv.Exists(sProdLotKey) Then
            dictOut.Add sProdLotKey, "0" & DELIMITER & dblCalculatedInv & DELIMITER & dblCalculatedInv * -1
        End If
    Next
    
    Set fCompare2Inventory = dictOut
    Set dictOut = Nothing
End Function


'====================== CZL Sales TO companies Except CZL =================================================================
'Function fReadSheetUnifiedSalesInfoSumQuantityByCZL(Optional CZLOnly As Boolean = True) As Dictionary
'    Dim dictOut As Dictionary
'
'    Dim arrData()
'    Dim dictColIndex As Dictionary
'
'    Call fRemoveFilterForSheet(shtSalesInfos)
'    Call fReadSheetDataByConfig("UNIFIED_SALES_INFO", dictColIndex, arrData, , , , , shtSalesInfos)
'    'Call fValidateDuplicateInArray(arrData, Array(dictColIndex("ProductProducer"), dictColIndex("ProductName"), dictColIndex("ProductSeries")), False, shtSelfSalesOrder, 1, 1, "厂家 + 名称 + 规格")
'
'    Dim sCZLCompName As String
'    Dim lEachRow As Long
'    Dim sSalesCompany As String
'    Dim sProducer As String
'    Dim sProductName As String
'    Dim sProductSeries As String
'   ' Dim sLotNum As String
'    Dim sKey As String
'    Dim dictRowNoTmp As Dictionary
'
'    Set dictOut = New Dictionary
'    Set dictRowNoTmp = New Dictionary
'
'    sCZLCompName = fGetCompanyNameByID_Common("CZL")
'
'    For lEachRow = LBound(arrData, 1) To UBound(arrData, 1)
'        sSalesCompany = Trim(arrData(lEachRow, dictColIndex("SalesCompanyName")))
'
'        If CZLOnly Then
'            If sSalesCompany <> sCZLCompName Then GoTo next_row
'        Else
'            If sSalesCompany = sCZLCompName Then GoTo next_row
'        End If
'
'        sProducer = Trim(arrData(lEachRow, dictColIndex("MatchedProductProducer")))
'        sProductName = Trim(arrData(lEachRow, dictColIndex("MatchedProductName")))
'        sProductSeries = Trim(arrData(lEachRow, dictColIndex("MatchedProductSeries")))
''        sLotNum = Trim(arrData(lEachRow, dictColIndex("LotNum")))
'
'        If CZLOnly Then
'            'sKey = sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries & DELIMITER & sLotNum
'            sKey = sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries
'        Else
'            'sKey = sSalesCompany & DELIMITER & sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries & DELIMITER & sLotNum
'            sKey = sSalesCompany & DELIMITER & sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries
'        End If
'
'        If Not dictOut.Exists(sKey) Then
'            dictOut.Add sKey, CDbl(arrData(lEachRow, dictColIndex("ConvertQuantity")))
'        Else
'            dictOut(sKey) = dictOut(sKey) + CDbl(arrData(lEachRow, dictColIndex("ConvertQuantity")))
'        End If
'        dictRowNoTmp(sKey) = (lEachRow + 1) & DELIMITER & arrData(lEachRow, dictColIndex("SellPrice"))
'next_row:
'    Next
'
'    Dim i As Long
'    For i = 0 To dictOut.Count - 1
'        sKey = dictOut.Keys(i)
'
'        dictOut(sKey) = dictOut(sKey) & DELIMITER & dictRowNoTmp(sKey)
'    Next
'
'    Set dictColIndex = Nothing
'    Set dictRowNoTmp = Nothing
'
'    Set fReadSheetUnifiedSalesInfoSumQuantityByCZL = dictOut
'    Set dictOut = Nothing
'End Function
Function fReadSCompUnifiedInvSumQuantityByCZL(Optional CZLOnly As Boolean = True) As Dictionary
    Dim dictOut As Dictionary
    
    Dim arrData()
    Dim sCZLCompName As String
    Dim lEachRow As Long
    Dim sSalesCompany As String
    Dim sProducer As String
    Dim sProductName As String
    Dim sProductSeries As String
   ' Dim sLotNum As String
    Dim sKey As String
    Set dictOut = New Dictionary
    
    sCZLCompName = fGetCompanyNameByID_Common("CZL")
    
    Call fRemoveFilterForSheet(shtSalesCompInvUnified)
    
    Call fCopyReadWholeSheetData2Array(shtSalesCompInvUnified, arrData)
    
    For lEachRow = LBound(arrData, 1) To UBound(arrData, 1)
        sSalesCompany = Trim(arrData(lEachRow, SCompUnifiedInv.SalesCompany))
        
        If CZLOnly Then
            If sSalesCompany <> sCZLCompName Then GoTo next_row
        Else
            If sSalesCompany = sCZLCompName Then GoTo next_row
        End If
        
        sProducer = Trim(arrData(lEachRow, SCompUnifiedInv.ProductProducer))
        sProductName = Trim(arrData(lEachRow, SCompUnifiedInv.ProductName))
        sProductSeries = Trim(arrData(lEachRow, SCompUnifiedInv.ProductSeries))
'        sLotNum = Trim(arrData(lEachRow, dictColIndex("LotNum")))
        
        If CZLOnly Then
            'sKey = sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries & DELIMITER & sLotNum
            sKey = sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries
        Else
            'sKey = sSalesCompany & DELIMITER & sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries & DELIMITER & sLotNum
            sKey = sSalesCompany & DELIMITER & sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries
        End If
        
        If Not dictOut.Exists(sKey) Then
            dictOut.Add sKey, CDbl(arrData(lEachRow, SCompUnifiedInv.InformedInventory))
        Else
            dictOut(sKey) = dictOut(sKey) + CDbl(arrData(lEachRow, SCompUnifiedInv.InformedInventory))
        End If
next_row:
    Next
    
    Erase arrData
    
    Set fReadSCompUnifiedInvSumQuantityByCZL = dictOut
    Set dictOut = Nothing
End Function

Sub subMain_CZLMonthEndInventoryRollOver()
    Dim response As VbMsgBoxResult
    response = MsgBox(Prompt:="该操作会用最新计算的库存覆盖更新期初库存，无法撤消，你确定要继续吗？" _
                        & vbCr & "继续，请点【Yes】" & vbCr & "否则，请点【No】" _
                        , Buttons:=vbCritical + vbYesNoCancel + vbDefaultButton2)
    If response <> vbYes Then Exit Sub
    
    fRemoveFilterForSheet shtCZLRolloverInv
    fRemoveFilterForSheet shtCZLInventory
    Call fDeleteRowsFromSheetLeaveHeader(shtCZLRolloverInv)
    
    Dim arrData()
    Call fCopyReadWholeSheetData2Array(shtCZLInventory, arrData)
    Call fWriteArray2Sheet(shtCZLRolloverInv, arrData)
    Erase arrData
    
    fMsgBox "采芝林的计算的库存成功转入，作为下一个月的期初库存。", vbInformation
    fShowAndActiveSheet shtCZLRolloverInv
End Sub

