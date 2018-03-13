Attribute VB_Name = "MD02_CZLInventoryCompare"
Option Explicit
Option Base 1

Sub subMain_CZLInventory()
    If fGetReplaceUnifyCZLSales2CompaniesErrorRowCount > 0 Then
        fMsgBox "采芝林的销售数据中有药品在系统中找不到，无法计算库存，请先处理这些错误。"
        shtSalesInfos.Visible = xlSheetVisible
        shtException.Visible = xlSheetVisible:         shtException.Activate
        End
    End If
    
    If Not fIsDev() Then On Error GoTo err_handle
    
    gsRptID = "CALCULATE_PROFIT"
    
    fVeryHideSheet shtException
    Call fCleanSheetOutputResetSheetOutput(shtException)
    fClearContentLeaveHeader shtCZLInventory
    
    fCalculateCZLInventory
    fActiveVisibleSwitchSheet shtCZLInventory, , False
    
    If fZero(gsBusinessErrorMsg) Then fMsgBox "【采芝林】库存计算完成！", vbInformation
err_handle:
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
    fEnableExcelOptionsAll
    End
End Sub


Sub subMain_CompareCZLInventory()
    If fGetReplaceUnifyCZLSales2CompaniesErrorRowCount > 0 Then
        fErr "采芝林的销售数据中有药品在系统中找不到，无法计算库存，请先处理这些错误。"
        shtSalesInfos.Visible = xlSheetVisible
        shtException.Visible = xlSheetVisible:         shtException.Activate
        End
    End If
    
    If Not fIsDev() Then On Error GoTo error_handling
    
    fRemoveFilterForSheet shtCZLInventory
    fRemoveFilterForSheet shtProductMaster
    fClearContentLeaveHeader shtCZLInvDiff
    
    fInitialization
    
    gsRptID = "COMPARE_CZL_INVENTORY"
    Call fReadSysConfig_InputTxtSheetFile
    
    gsRptFilePath = fReadSysConfig_Output(, gsRptType)
    
    Call fLoadFileByFileTag("PRODUCT_MASTER")
    Call fReadMasterSheetData("PRODUCT_MASTER", shtProductMaster)
    
    Dim dictCZLInformedInv As Dictionary
    Set dictCZLInformedInv = fReadArray2DictionaryWithMultipleKeyColsSingleItemCol(arrMaster _
                            , Array(dictMstColIndex("ProductProducer"), dictMstColIndex("ProductName"), dictMstColIndex("ProductSeries")) _
                            , CLng(dictMstColIndex("CZLInformedInventory")), DELIMITER)
    Erase arrMaster
    
    Dim dictCZLInv As Dictionary
    Dim arrCZLInv()
    Call fCopyReadWholeSheetData2Array(shtCZLInventory, arrCZLInv)
    'fReadArray2DictionaryWithMultipleKeyColsSingleItemCol
    Set dictCZLInv = fReadArray2DictionaryWithMultipleKeyColsSingleItemColSum(arrCZLInv _
                            , Array(1, 2, 3) _
                            , 6, DELIMITER)
    Erase arrCZLInv
    
    Dim dictInventoryDiff As Dictionary
    Set dictInventoryDiff = fCompare2Inventory(dictCZLInformedInv, dictCZLInv)
    
    arrOutput = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictInventoryDiff, , False)
    Call fAppendArray2Sheet(shtCZLInvDiff, arrOutput)
    arrOutput = fConvertDictionaryDelimiteredItemsTo2DimenArrayForPaste(dictInventoryDiff)
    shtCZLInvDiff.Cells(2, 5).Resize(UBound(arrOutput, 1), UBound(arrOutput, 2)).Value = arrOutput
     
    Call fFormatOutputSheet(shtCZLInvDiff)
    
    shtCZLInvDiff.Rows(1).RowHeight = 25
    shtCZLInvDiff.Visible = xlSheetVisible
    shtCZLInvDiff.Activate
    shtCZLInvDiff.Range("A1").Select
error_handling:
    If fCheckIfGotBusinessError Then
        Call fSetReneratedReport(, "-")
        GoTo reset_excel_options
    End If
    
    If fCheckIfUnCapturedExceptionAbnormalError Then
        Call fSetReneratedReport(, "-")
        GoTo reset_excel_options
    End If
    
    Call fSetReneratedReport(, shtCZLSales2CompRawData.Name)
    fMsgBox "采芝林库存差异计算结果在表：[" & shtCZLInvDiff.Name & "] 中，请检查！", vbInformation
    
reset_excel_options:
    Err.Clear
    fEnableExcelOptionsAll
    End
    
End Sub

Function fCompare2Inventory(dictCZLInformedInv As Dictionary, dictCZLInv As Dictionary) As Dictionary
    Dim dictOut As Dictionary
     
    Dim i As Long
    Dim sProdLotKey As String
    Dim dblInformedInv As Double
    Dim dblCalculatedInv As Double
    
    Set dictOut = New Dictionary
    
    For i = 0 To dictCZLInformedInv.Count - 1
        sProdLotKey = dictCZLInformedInv.Keys(i)
        dblInformedInv = CDbl(dictCZLInformedInv.Items(i))
        
        If Not dictCZLInv.Exists(sProdLotKey) Then
            dictOut.Add sProdLotKey, dblInformedInv & DELIMITER & "0" & DELIMITER & dblInformedInv
        Else
            dblCalculatedInv = dictCZLInv(sProdLotKey)
            dictOut.Add sProdLotKey, dblInformedInv & DELIMITER & dblCalculatedInv & DELIMITER & (dblInformedInv - dblCalculatedInv)
        End If
    Next
    
    Set fCompare2Inventory = dictOut
    Set dictOut = Nothing
End Function



