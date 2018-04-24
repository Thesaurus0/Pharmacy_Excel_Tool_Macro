Attribute VB_Name = "MF01_SalesCompInvCal"
Option Explicit
Option Base 1

Private dictSCompRolloverInv As Dictionary
Private dictCZLSales2SalesComp As Dictionary    'purchase
Private dictSCompSales2Hospital As Dictionary   'sales
Private sYearMonth As String

Sub subMain_CalculateSalesCompInventory()
    If Not fIsDev() Then On Error GoTo err_handle
    fCheckIfErrCountNotZero_CZLSales2Comp
    
    Call fValidaterngYearMonth(sYearMonth)
    
    If Not fPromptToConfirmToContinue("�㵱ǰ����������ǣ�" & sYearMonth & "," & vbCr & "��ȷ�����������ȷ��") Then fErr
    
    gsRptID = "CALCULATE_PROFIT"
    Call fReadSysConfig_InputTxtSheetFile
    
    fVeryHideSheet shtException
    Call fCleanSheetOutputResetSheetOutput(shtException)
    Call fRemoveFilterForSheet(shtSalesCompInvCalcd)
    Call fClearContentLeaveHeader(shtSalesCompInvCalcd)
    
    Call fCalculateSalesCompanyInventory
    fActiveVisibleSwitchSheet shtSalesCompInvCalcd, , False
    
    If fZero(gsBusinessErrorMsg) Then fMsgBox "���С���ҵ��˾������֥�ֳ��⣩�Ŀ�������ɣ�", vbInformation
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

Private Function fCalculateSalesCompanyInventory()
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
    
    Call fRemoveFilterForSheet(shtSalesCompRolloverInv)     'rollover inventory
    Call fRemoveFilterForSheet(shtCZLSales2Companies)       'purchase
    Call fRemoveFilterForSheet(shtSalesInfos)               'sales
    
    Call fSortDataInSheetSortSheetData(shtSalesCompRolloverInv, Array(SCompRollover.SalesCompany, SCompRollover.ProductProducer, SCompRollover.ProductName))
    Call fSortDataInSheetSortSheetData(shtCZLSales2Companies, Array(CZLSales2Comp.SalesCompany, CZLSales2Comp.ProductProducer, CZLSales2Comp.ProductName, CZLSales2Comp.ProductSeries))
    Call fSortDataInSheetSortSheetData(shtSalesInfos, Array(Sales2Hospital.SalesCompany, Sales2Hospital.ProductProducer, Sales2Hospital.ProductName, Sales2Hospital.ProductSeries))

    Set dictCZLSales2SalesComp = fReadSheetCZLSalesToCompaniesByCompaniesExceptCZL()  'CZL sales 2 comp = purchase    dictCZLSales2SalesComp
    Call fReadUnifiedSalesInfoToHospital2Dictionary    ' sales to hospital    dictSCompSales2Hospital
    Call fReadSalesCompanyRolloverInventory2Dictionary      'dictSCompRolloverInv
    
    '================= verify lot number ========================================
    Dim dictMissedLot As Dictionary
    Set dictMissedLot = New Dictionary
    
    For i = 0 To dictSCompSales2Hospital.Count - 1
        sKey = dictSCompSales2Hospital.Keys(i)
        
        If Not dictCZLSales2SalesComp.Exists(sKey) Then
           ' dictMissedLot.Add sKey, 0
            dictCZLSales2SalesComp.Add sKey, 0     'add for output
        End If
    Next
    
    'rollover
    For i = 0 To dictSCompRolloverInv.Count - 1
        sKey = dictSCompRolloverInv.Keys(i)
        
        If Not dictCZLSales2SalesComp.Exists(sKey) Then
            dictCZLSales2SalesComp.Add sKey, 0     'add for output
        End If
    Next
    
    If dictMissedLot.Count > 0 Then
        Call fAddMissedSelfSalesLotNumToSheetException(dictMissedLot)
    End If
    
    Set dictMissedLot = Nothing
    '---------------------------------------------------------------
    
    '================= calculate inventory of this month  ========================================
    ReDim arrOut(1 To dictCZLSales2SalesComp.Count, 7) 'CZL sales 2 comp = purchase
    
    For i = 0 To dictCZLSales2SalesComp.Count - 1  'CZL sales 2 comp = purchase
        sKey = dictCZLSales2SalesComp.Keys(i)
        
        dblPurchaseQty = CDbl(Split(dictCZLSales2SalesComp(sKey), DELIMITER)(0))
        dblSellQty = 0
        dblRolloverQty = 0
        
        If dictSCompSales2Hospital.Exists(sKey) Then
            dblSellQty = CDbl(Split(dictSCompSales2Hospital(sKey), DELIMITER)(0))
        End If
                
        If dictSCompRolloverInv.Exists(sKey) Then
            dblRolloverQty = dictSCompRolloverInv(sKey)
        End If

        arrOut(i + 1, 1) = Split(sKey, DELIMITER)(0)
        arrOut(i + 1, 2) = Split(sKey, DELIMITER)(1)
        arrOut(i + 1, 3) = Split(sKey, DELIMITER)(2)
        arrOut(i + 1, 4) = Split(sKey, DELIMITER)(3)
   '     arrOut(i + 1, 5) = fGetProductUnit(arrOut(i + 1, 2), arrOut(i + 1, 3), arrOut(i + 1, 4))
        arrOut(i + 1, 6) = "'" & Split(sKey, DELIMITER)(4)  'lot num
        
        arrOut(i + 1, 7) = dblPurchaseQty - dblSellQty + dblRolloverQty
        
'        If IsNumeric(Split(dictCZLSales2SalesComp(sKey), DELIMITER)(2)) Then
'            arrOut(i + 1, 7) = CDbl(Split(dictCZLSales2SalesComp(sKey), DELIMITER)(2))     'purcahse price
'        Else
'        End If
    Next
    '---------------------------------------------------------------
    
    Set dictCZLSales2SalesComp = Nothing
    Set dictSCompRolloverInv = Nothing
    Set dictSCompSales2Hospital = Nothing
    
    'fCalculateCZLInventory = arrOut
    Call fAppendArray2Sheet(shtSalesCompInvCalcd, arrOut)
    Erase arrOut
End Function

Private Function fReadUnifiedSalesInfoToHospital2Dictionary()
    Dim arrData()
    Dim dictColIndex As Dictionary
    
    Call fRemoveFilterForSheet(shtSelfSalesOrder)
    Call fReadSheetDataByConfig("SALES_COMPANY_SALES_TO_HOSPITAL", dictColIndex, arrData, , , , , shtSalesInfos)
    'Call fValidateDuplicateInArray(arrData, Array(dictColIndex("ProductProducer"), dictColIndex("ProductName"), dictColIndex("ProductSeries")), False, shtSelfSalesOrder, 1, 1, "���� + ���� + ���")
    
    Dim sCZLCompName As String
    Dim sSalesCompany As String
    Dim lEachRow As Long
    Dim sProducer As String
    Dim sProductName As String
    Dim sProductSeries As String
    Dim sLotNum As String
    Dim sKey As String
    Dim dictRowNoTmp As Dictionary
    
'    sCZLCompName = fGetCompany_CompanyName("CZL")
'    If sCZLCompName <> fGetCompanyNameByID_Common("CZL") Then fErr "CZL���������ò�һ�£� [Sales Company List - Common Importing - Sales File] �� [Sales Company List]"
    sCZLCompName = fGetCompanyNameByID_Common("CZL")
    
    Set dictSCompSales2Hospital = New Dictionary
    Set dictRowNoTmp = New Dictionary
    
    For lEachRow = LBound(arrData, 1) To UBound(arrData, 1)
        sSalesCompany = Trim(arrData(lEachRow, dictColIndex("SalesCompanyName")))
        
        If sSalesCompany = sCZLCompName Then GoTo next_row
        
        sProducer = Trim(arrData(lEachRow, dictColIndex("MatchedProductProducer")))
        sProductName = Trim(arrData(lEachRow, dictColIndex("MatchedProductName")))
        sProductSeries = Trim(arrData(lEachRow, dictColIndex("MatchedProductSeries")))
        sLotNum = Trim(arrData(lEachRow, dictColIndex("LotNum")))
        
        sKey = sSalesCompany & DELIMITER & sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries & DELIMITER & sLotNum
        
        If Not dictSCompSales2Hospital.Exists(sKey) Then
            dictSCompSales2Hospital.Add sKey, CDbl(arrData(lEachRow, dictColIndex("ConvertQuantity")))
        Else
            dictSCompSales2Hospital(sKey) = dictSCompSales2Hospital(sKey) + CDbl(arrData(lEachRow, dictColIndex("ConvertQuantity")))
        End If
        
        dictRowNoTmp(sKey) = (lEachRow + 1) & DELIMITER & arrData(lEachRow, dictColIndex("SellPrice"))
next_row:
    Next
    
    Dim i As Long
    For i = 0 To dictSCompSales2Hospital.Count - 1
        sKey = dictSCompSales2Hospital.Keys(i)
        
        dictSCompSales2Hospital(sKey) = dictSCompSales2Hospital(sKey) & DELIMITER & dictRowNoTmp(sKey)
    Next
    
    Set dictColIndex = Nothing
    Set dictRowNoTmp = Nothing
End Function



Private Function fReadSalesCompanyRolloverInventory2Dictionary()
    Dim arrData()
    
    Call fCopyReadWholeSheetData2Array(shtSalesCompRolloverInv, arrData)
    
    Set dictSCompRolloverInv = fReadArray2DictionaryWithMultipleKeyColsSingleItemColSum(arrData _
                            , Array(SCompRollover.SalesCompany, SCompRollover.ProductProducer, SCompRollover.ProductName, SCompRollover.ProductSeries, SCompRollover.LotNum) _
                            , CLng(SCompRollover.RolloverQty), DELIMITER)
    Erase arrData
End Function


Private Function fAddMissedSelfSalesLotNumToSheetException(dictMissedLotNum As Dictionary)
    Dim arrMissedLotNum()
    'Dim sErr As String
    Dim lRecCount As Long
    Dim lStartRow As Long
    
    lRecCount = fGetDictionayDelimiteredItemsCount(dictMissedLotNum)
    If lRecCount > 0 Then
        lStartRow = fGetshtExceptionNewRow
        
        arrMissedLotNum = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictMissedLotNum, , False)
        
        shtException.Cells(lStartRow, 1).Value = "��ҵ��˾���۵�ҽԺ�����Ų������ڲ�֥�ֵ������У����ǲ�Ӧ�ó���Ĵ������ݣ�ĳһ����ԭʼ���������е���������"
        lStartRow = lStartRow + 1
        Call fPrepareHeaderToSheet(shtException, Array("��ҵ��˾", "ҩƷ����", "ҩƷ����", "ҩƷ���", "�����ڵ�����"), lStartRow)
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

        gsBusinessErrorMsg = gsBusinessErrorMsg & lRecCount & "��ҩƷ�ĳ����ġ����š��ڲ�֥�ֵ����۱����Ҳ�����" & vbCr _
            & "���顾��֥�ֵ����۳�������"
    End If
End Function


