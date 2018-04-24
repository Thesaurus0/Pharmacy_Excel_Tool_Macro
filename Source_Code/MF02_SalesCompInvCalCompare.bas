Attribute VB_Name = "MF02_SalesCompInvCalCompare"
Option Explicit
Option Base 1

Sub subMain_CompareSalesCompanyInventory()

'    If fGetReplaceUnifyInventoryErrorRowCount > 0 Then
'        fErr "��ҵ��˾�Ŀ����������ҩƷ��ϵͳ���Ҳ������޷����п��˶ԣ����ȴ�����Щ����"
'        shtSalesCompInvUnified.Visible = xlSheetVisible
'        shtException.Visible = xlSheetVisible:         shtException.Activate
'        End
'    End If
    
    If Not fIsDev() Then On Error GoTo error_handling
    
    fCheckIfErrCountNotZero_SalesInventory
    
    fRemoveFilterForSheet shtSalesCompInvCalcd
    fRemoveFilterForSheet shtSalesCompInvUnified
    fRemoveFilterForSheet shtSalesCompInvDiff
    fClearContentLeaveHeader shtSalesCompInvDiff
    
    fInitialization
    
    gsRptID = "COMPARE_SALES_COMP_INVENTORY"
    Call fReadSysConfig_InputTxtSheetFile
    
    gsRptFilePath = fReadSysConfig_Output(, gsRptType)
    
    Call fCopyReadWholeSheetData2Array(shtSalesCompInvUnified, arrMaster)
    
    Dim dictSCompInformedInv As Dictionary
    Set dictSCompInformedInv = fReadArray2DictionaryWithMultipleKeyColsSingleItemColSum(arrMaster _
                            , Array(SCompUnifiedInv.SalesCompany, SCompUnifiedInv.ProductProducer, SCompUnifiedInv.ProductName, SCompUnifiedInv.ProductSeries, SCompUnifiedInv.LotNum) _
                            , CLng(SCompUnifiedInv.InformedInventory), DELIMITER)
    Erase arrMaster
    
    Dim dictSCompCalInv As Dictionary
    Dim arrSCompCalInv()
    Call fCopyReadWholeSheetData2Array(shtSalesCompInvCalcd, arrSCompCalInv)
    Set dictSCompCalInv = fReadArray2DictionaryWithMultipleKeyColsSingleItemColSum(arrSCompCalInv _
                            , Array(SCompInvCalcd.SalesCompany, SCompInvCalcd.ProductProducer, SCompInvCalcd.ProductName, SCompInvCalcd.ProductSeries, SCompInvCalcd.LotNum) _
                            , SCompInvCalcd.InventoryQty, DELIMITER)
    Erase arrSCompCalInv
    
    Dim dictInventoryDiff As Dictionary
    Set dictInventoryDiff = fCompare2Inventory(dictSCompInformedInv, dictSCompCalInv)
    Set dictSCompCalInv = Nothing
    Set dictSCompInformedInv = Nothing
    
    arrOutput = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictInventoryDiff, , False)
    Call fAppendArray2Sheet(shtSalesCompInvDiff, arrOutput)
    arrOutput = fConvertDictionaryDelimiteredItemsTo2DimenArrayForPaste(dictInventoryDiff)
    shtSalesCompInvDiff.Cells(2, SCompInvDiff.InformedQty).Resize(UBound(arrOutput, 1), UBound(arrOutput, 2)).Value = arrOutput
    
    Call fSortDataInSheetSortSheetData(shtSalesCompInvDiff, Array(SCompInvDiff.SalesCompany, SCompInvDiff.ProductProducer, SCompInvDiff.ProductName, SCompInvDiff.ProductSeries, SCompInvDiff.LotNum))
    
    Call fFormatOutputSheet(shtSalesCompInvDiff)
    
    shtSalesCompInvDiff.Rows(1).RowHeight = 25
    shtSalesCompInvDiff.Visible = xlSheetVisible
    shtSalesCompInvDiff.Activate
    Application.Goto shtSalesCompInvDiff.Range("A" & fGetValidMaxRow(shtSalesCompInvDiff)), True
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
    fMsgBox "��֥�ֿ�����������ڱ�[" & shtSalesCompInvDiff.Name & "] �У����飡", vbInformation
    
reset_excel_options:
    Err.Clear
    fEnableExcelOptionsAll
    End
    
End Sub
Private Function fCompare2Inventory(dictSCompInformedInv As Dictionary, dictSCompCalInv As Dictionary) As Dictionary
    Dim dictOut As Dictionary
     
    Dim i As Long
    Dim sProdLotKey As String
    Dim dblInformedInv As Double
    Dim dblCalculatedInv As Double
    
    Set dictOut = New Dictionary
    
    For i = 0 To dictSCompInformedInv.Count - 1
        sProdLotKey = dictSCompInformedInv.Keys(i)
        dblInformedInv = CDbl(dictSCompInformedInv.Items(i))
        
        If Not dictSCompCalInv.Exists(sProdLotKey) Then
            dictOut.Add sProdLotKey, dblInformedInv & DELIMITER & "0" & DELIMITER & dblInformedInv
        Else
            dblCalculatedInv = dictSCompCalInv(sProdLotKey)
            dictOut.Add sProdLotKey, dblInformedInv & DELIMITER & dblCalculatedInv & DELIMITER & (dblInformedInv - dblCalculatedInv)
        End If
    Next
    
    
    For i = 0 To dictSCompCalInv.Count - 1
        sProdLotKey = dictSCompCalInv.Keys(i)
        dblCalculatedInv = CDbl(dictSCompCalInv.Items(i))
        
        If Not dictSCompInformedInv.Exists(sProdLotKey) Then
            dictOut.Add sProdLotKey, "0" & DELIMITER & dblCalculatedInv & DELIMITER & dblCalculatedInv * -1
        End If
    Next
    
    Set fCompare2Inventory = dictOut
    Set dictOut = Nothing
End Function

Sub subMain_SalesCompanyMonthEndInventoryRollOver()
    Dim response As VbMsgBoxResult
    response = MsgBox(Prompt:="�ò����������¼���Ŀ�渲�Ǹ����ڳ���棬�޷���������ȷ��Ҫ������" _
                        & vbCr & "��������㡾Yes��" & vbCr & "������㡾No��" _
                        , Buttons:=vbCritical + vbYesNoCancel + vbDefaultButton2)
    If response <> vbYes Then Exit Sub
    
    Call fClearContentLeaveHeader(shtSalesCompRolloverInv)
    
    Dim arrData()
    Call fCopyReadWholeSheetData2Array(shtSalesCompInvCalcd, arrData)
    Call fWriteArray2Sheet(shtSalesCompRolloverInv, arrData)
    Erase arrData
    
    fMsgBox "��ҵ��˾�ļ������õĿ��ɹ�ת�룬��Ϊ��һ���µ��ڳ���档", vbInformation
    fShowAndActiveSheet shtSalesCompRolloverInv
End Sub



