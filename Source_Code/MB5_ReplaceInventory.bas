Attribute VB_Name = "MB5_ReplaceInventory"
Option Explicit
Option Base 1

Sub subMain_ReplaceInventory()
    If Not fIsDev Then On Error GoTo error_handling
    
    Call fSetReplaceUnifyErrorRowCount_SCompInventory(999)
    
    shtInventoryRawDataRpt.Visible = xlSheetVisible
    shtException.Visible = xlSheetVeryHidden
    Call fUnProtectSheet(shtSalesCompInvUnified)
    Call fCleanSheetOutputResetSheetOutput(shtSalesCompInvUnified)
    'Call fCleanSheetOutputResetSheetOutput(shtException)
    Call fDeleteRowsFromSheetLeaveHeader(shtException)
    'shtException.Cells.NumberFormat = "@"
    shtException.Cells.WrapText = True
    
    fInitialization

    gsRptID = "REPLACE_UNIFY_INVENTORY"

    Call fReadSysConfig_InputTxtSheetFile

    gsRptFilePath = fReadSysConfig_Output(, gsRptType)
    Call fLoadFilesAndRead2Variables

    Call fPrepareOutputSheetHeaderAndTextColumns(shtSalesCompInvUnified)

'    ReDim arrErrRows(1 To UBound(arrMaster, 1) * 2)
'    alErrCnt = 0
'    ReDim arrWarningRows(1 To UBound(arrMaster, 1) * 2)
'    alWarningCnt = 0
    Set dictErrorRows = New Dictionary
    Set dictWarningRows = New Dictionary
    
    Call fProcessData
'    If alErrCnt > 0 Then
'        ReDim Preserve arrErrRows(1 To alErrCnt)
'    Else
'        arrErrRows = Array()
'    End If
'    If alWarningCnt > 0 Then
'        ReDim Preserve arrWarningRows(1 To alWarningCnt)
'    Else
'        arrWarningRows = Array()
'    End If
    
    If dictErrorRows.count > 0 Or dictErrorRows.count > 0 Then shtException.Visible = xlSheetVisible
    
error_handling:
    'If shtException.Visible = xlSheetVisible Then
        Call fAppendArray2Sheet(shtSalesCompInvUnified, arrOutput)
    
        Call fFormatOutputSheet(shtSalesCompInvUnified)
    
       ' Call fProtectSheetAndAllowEdit(shtInventoryRawDataRpt, shtInventoryRawDataRpt.Columns(4), UBound(arrOutput, 1) + 1, UBound(arrOutput, 2), False)
        Call fPostProcess(shtSalesCompInvUnified)
        
        Call fSetFormatForExceptionCells(shtSalesCompInvUnified, dictErrorRows, "REPORT_ERROR_COLOR")
        Call fSetFormatForExceptionCells(shtSalesCompInvUnified, dictWarningRows, "REPORT_WARNING_COLOR")
        Call fSetReplaceUnifyErrorRowCount_SCompInventory(dictErrorRows.count)
    
        shtSalesCompInvUnified.Visible = xlSheetVisible
        shtSalesCompInvUnified.Activate
        shtSalesCompInvUnified.Range("A1").Select
        
    If shtException.Visible = xlSheetVisible Then
        Dim lExcepMaxCol As Long
        lExcepMaxCol = fGetValidMaxCol(shtException)
        Call fSetFormatBoldOrangeBorderForHeader(shtException, lExcepMaxCol)
        Call fSetBorderLineForSheet(shtException, lExcepMaxCol)
        Call fBasicCosmeticFormatSheet(shtException, lExcepMaxCol)
        Call fSetFormatForOddEvenLineByFixColor(shtException, lExcepMaxCol)
        shtException.Activate
    End If
    
    If fCheckIfGotBusinessError Then GoTo reset_excel_options
    If fCheckIfUnCapturedExceptionAbnormalError Then GoTo reset_excel_options
 
    fMsgBox "成功整合在工作表：[" & shtSalesCompInvUnified.Name & "] 中，请检查！", vbInformation

reset_excel_options:
    Err.Clear
    fClearRefVariables
    fEnableExcelOptionsAll
    'End
End Sub

Private Function fLoadFilesAndRead2Variables()
    'gsCompanyID
    Call fLoadFileByFileTag("IMPORTED_DATA")
    Call fReadMasterSheetData("IMPORTED_DATA", , , True)
End Function

Private Function fProcessData()
    Call fRedimArrOutputBaseArrMaster
    
    Dim lEachRow As Long
    Dim sCompanyLongID As String
    Dim sCompanyName As String
        
    Dim sProducer As String
    Dim sReplacedProducer As String
    
    Dim sProductName As String
    Dim sReplacedProductName As String
    
    Dim sProductSeries As String
    Dim sReplacedProductSeries As String
    
    Dim sProductUnit As String
    Dim sReplacedProductUnit As String
    Dim sProductMasterUnit As String
    
    Dim dblRatio As Double
    
    Dim dictNewProducer As Dictionary
    Dim dictNewProductName As Dictionary
    Dim dictNewProductSeries As Dictionary
    Dim dictNewProductUnit As Dictionary
    Dim dictNewProductUnitOrig As Dictionary
    
    Set dictNewProducer = New Dictionary
    Set dictNewProductName = New Dictionary
    Set dictNewProductSeries = New Dictionary
    Set dictNewProductUnit = New Dictionary
    Set dictNewProductUnitOrig = New Dictionary

    Dim sTmpKey As String
    For lEachRow = LBound(arrMaster, 1) To UBound(arrMaster, 1)
        If dictMstColIndex.Exists("OrigInventoryID") Then
            arrOutput(lEachRow, dictRptColIndex("OrigInventoryID")) = arrMaster(lEachRow, dictMstColIndex("OrigInventoryID"))
        End If
        
        If dictMstColIndex.Exists("SeqNo") Then
            arrOutput(lEachRow, dictRptColIndex("SeqNo")) = arrMaster(lEachRow, dictMstColIndex("SeqNo"))
        End If
        
        arrOutput(lEachRow, dictRptColIndex("SalesCompanyName")) = arrMaster(lEachRow, dictMstColIndex("SalesCompanyName"))
        arrOutput(lEachRow, dictRptColIndex("Quantity")) = arrMaster(lEachRow, dictMstColIndex("Quantity"))
        arrOutput(lEachRow, dictRptColIndex("LotNum")) = "'" & arrMaster(lEachRow, dictMstColIndex("LotNum"))
        sProducer = Trim(arrMaster(lEachRow, dictMstColIndex("ProductProducer")))
        arrOutput(lEachRow, dictRptColIndex("ProductProducer")) = sProducer
        sProductName = Trim(arrMaster(lEachRow, dictMstColIndex("ProductName")))
        arrOutput(lEachRow, dictRptColIndex("ProductName")) = sProductName
        sProductSeries = Trim(arrMaster(lEachRow, dictMstColIndex("ProductSeries")))
        arrOutput(lEachRow, dictRptColIndex("ProductSeries")) = sProductSeries
        sProductUnit = Trim(arrMaster(lEachRow, dictMstColIndex("ProductUnit")))
        arrOutput(lEachRow, dictRptColIndex("ProductUnit")) = sProductUnit
                
        ' Product producer replace -----------------
        If Not fReplaceAndValidateInProducerMaster(sProducer, sReplacedProducer) Then
            If Not dictNewProducer.Exists(sReplacedProducer) Then
                dictNewProducer.Add sReplacedProducer, "'" & (lEachRow + 1)
            Else
                dictNewProducer(sReplacedProducer) = dictNewProducer(sReplacedProducer) & "," & (lEachRow + 1)
            End If
            arrOutput(lEachRow, dictRptColIndex("MatchedProductProducer")) = ""
            
'            alErrCnt = alErrCnt + 1
'            arrErrRows(alErrCnt) = lEachRow + 1
'            alErrCnt = alErrCnt + 1
'            arrErrRows(alErrCnt) = dictRptColIndex("ProductProducer")
            Call fAddErrorColumnTodictErrorRows(lEachRow + 1, dictRptColIndex("ProductProducer"))
            GoTo next_sales
        Else
            arrOutput(lEachRow, dictRptColIndex("MatchedProductProducer")) = sReplacedProducer
        End If
        ' Product producer end -----------------
        
        ' Product Name replace -----------------
        If Not fReplaceAndValidateInProductNameMaster(sReplacedProducer, sProductName, sReplacedProductName) Then
            sTmpKey = sReplacedProducer & DELIMITER & sReplacedProductName
            If Not dictNewProductName.Exists(sTmpKey) Then
                dictNewProductName.Add sTmpKey, "'" & (lEachRow + 1)
            Else
                dictNewProductName(sTmpKey) = dictNewProductName(sTmpKey) & "," & lEachRow + 1
            End If
            arrOutput(lEachRow, dictRptColIndex("MatchedProductName")) = ""
            
'            alErrCnt = alErrCnt + 1
'            arrErrRows(alErrCnt) = lEachRow + 1
'            alErrCnt = alErrCnt + 1
'            arrErrRows(alErrCnt) = dictRptColIndex("ProductName")
            Call fAddErrorColumnTodictErrorRows(lEachRow + 1, dictRptColIndex("ProductName"))
            GoTo next_sales
        Else
            arrOutput(lEachRow, dictRptColIndex("MatchedProductName")) = sReplacedProductName
        End If
        ' Product Name end -----------------
        
        ' Product Series replace -----------------
        If Not fReplaceAndValidateInProductSeriesMaster(sReplacedProducer, sReplacedProductName, sProductSeries, sReplacedProductSeries) Then
            sTmpKey = sReplacedProducer & DELIMITER & sReplacedProductName & DELIMITER & sReplacedProductSeries
            If Not dictNewProductSeries.Exists(sTmpKey) Then
                dictNewProductSeries.Add sTmpKey, "'" & (lEachRow + 1)
            Else
                dictNewProductSeries(sTmpKey) = dictNewProductSeries(sTmpKey) & "," & (lEachRow + 1)
            End If
            arrOutput(lEachRow, dictRptColIndex("MatchedProductSeries")) = ""
            
'            alErrCnt = alErrCnt + 1
'            arrErrRows(alErrCnt) = lEachRow + 1
'            alErrCnt = alErrCnt + 1
'            arrErrRows(alErrCnt) = dictRptColIndex("ProductSeries")
            Call fAddErrorColumnTodictErrorRows(lEachRow + 1, dictRptColIndex("ProductSeries"))
            GoTo next_sales
        Else
            arrOutput(lEachRow, dictRptColIndex("MatchedProductSeries")) = sReplacedProductSeries
        End If
        ' Product Series end -----------------
        
        ' Product Unit ration -----------------
'        Call fGetConvertUnitAndUnitRatio
        sProductMasterUnit = fGetProductUnit(sReplacedProducer, sReplacedProductName, sReplacedProductSeries)

        dblRatio = 0
        If Len(Trim(sProductUnit)) <= 0 Then
            arrOutput(lEachRow, dictRptColIndex("MatchedProductUnit")) = sProductMasterUnit
            dblRatio = 1
        Else
'            Call fReplaceAndValidateInProductUnitMaster(sReplacedProducer, sReplacedProductName, sReplacedProductSeries, sProductUnit _
'                                                , sReplacedProductUnit, dblRatio)
            sReplacedProductUnit = fFindInConfigedReplaceProductUnit(sReplacedProducer, sReplacedProductName, sReplacedProductSeries _
                                                        , sProductUnit, dblRatio)
            If Len(Trim(sReplacedProductUnit)) <= 0 Then
                sReplacedProductUnit = sProductUnit
            Else
                If sReplacedProductUnit <> sProductMasterUnit Then
                    fErr "【" & shtProductUnitRatio.Name & "】会计单位和药品主表不一样，请检查【" & shtProductUnitRatio.Name & "】" _
                        & vbCr & sReplacedProducer _
                        & vbCr & sReplacedProductName _
                        & vbCr & sReplacedProductSeries & vbCr & sProductMasterUnit
                End If

                If sProductUnit = sProductMasterUnit Then
                    If dblRatio <> 1 Then
                        fErr "原始文件单位和会计单位一样，但是倍数却不是1，请检查【" & shtProductUnitRatio.Name & "】" _
                            & vbCr & sReplacedProducer _
                            & vbCr & sReplacedProductName _
                            & vbCr & sReplacedProductSeries & vbCr & sProductMasterUnit
                    End If
                End If
            End If

            If sReplacedProductUnit <> sProductMasterUnit Then
                sTmpKey = sReplacedProducer & DELIMITER & sReplacedProductName & DELIMITER & sReplacedProductSeries & DELIMITER _
                                & sProductMasterUnit '& DELIMITER & sReplacedProductUnit
                If Not dictNewProductUnit.Exists(sTmpKey) Then
                    dictNewProductUnit.Add sTmpKey, "'" & (lEachRow + 1)
                    dictNewProductUnitOrig.Add sTmpKey, sReplacedProductUnit
                Else
                    dictNewProductUnit(sTmpKey) = dictNewProductUnit(sTmpKey) & "," & (lEachRow + 1)
                    
                    If InStr(dictNewProductUnitOrig(sTmpKey), sReplacedProductUnit) <= 0 Then _
                    dictNewProductUnitOrig(sTmpKey) = dictNewProductUnitOrig(sTmpKey) & "," & sReplacedProductUnit
                End If

                arrOutput(lEachRow, dictRptColIndex("MatchedProductUnit")) = ""
            
'                alErrCnt = alErrCnt + 1
'                arrErrRows(alErrCnt) = lEachRow + 1
'                alErrCnt = alErrCnt + 1
'                arrErrRows(alErrCnt) = dictRptColIndex("ProductUnit")
                Call fAddErrorColumnTodictErrorRows(lEachRow + 1, dictRptColIndex("ProductUnit"))
                GoTo next_sales
            End If

            arrOutput(lEachRow, dictRptColIndex("MatchedProductUnit")) = sReplacedProductUnit
        End If
        ' Product Unit ration -----------------
        
        If dblRatio <> 1 Then
            Debug.Print lEachRow
        End If
        
        arrOutput(lEachRow, dictRptColIndex("ConvertQuantity")) = arrMaster(lEachRow, dictMstColIndex("Quantity")) * dblRatio
       ' arrOutput(lEachRow, dictRptColIndex("ConvertSellPrice")) = arrMaster(lEachRow, dictMstColIndex("SellPrice")) / dblRatio
        'arrOutput(lEachRow, dictRptColIndex("RecalSellAmount")) = arrOutput(lEachRow, dictRptColIndex("ConvertQuantity")) _
                                                                * arrOutput(lEachRow, dictRptColIndex("ConvertSellPrice"))
next_sales:
    Next
    
    'Call fAddNewFoundMissedHospitalToSheet(dictNewHospital, False)
    'Call fAddNewFoundHospitalToSheetException(dictNewHospital, True)
    Call fAddNewFoundMissedProducerToSheetException(dictNewProducer)
    Call fAddNewFoundMissedProductNameToSheetException(dictNewProductName)
    Call fAddNewFoundMissedProductSeriesToSheetException(dictNewProductSeries)
    Call fAddNewFoundMissedProductUnitToSheetException(dictNewProductUnit, dictNewProductUnitOrig)
End Function

Function fAddNewFoundMissedHospitalToSheet(dictNewHospital As Dictionary, Optional bSetDictToNothing As Boolean = True)
    Dim arrNewHospital()
    arrNewHospital = fConvertDictionaryKeysTo2DimenArrayForPaste(dictNewHospital, bSetDictToNothing)
    Call fAppendArray2Sheet(shtHospital, arrNewHospital)
    
    If fUbound(arrNewHospital, 1) > 0 Then
        fMsgBox fUbound(arrNewHospital, 1) & "个医院找不到，" & vbCr & "他们被自动加入到了表【" & shtHospital.Name & "】中了." _
                & vbCr & "该表的最后面的数据为本次新加的。" & vbCr _
                & ""
    End If
    Erase arrNewHospital
End Function
Function fAddNewFoundHospitalToSheetException(ByRef dictNewHospital As Dictionary, Optional bSetDictToNothing As Boolean = True)
    Dim arrNewHospital()
    'Dim sErr As String
    Dim lStartRow As Long
    Dim lRecCount As Long
    
    lRecCount = fGetDictionayDelimiteredItemsCount(dictNewHospital)
        
    If lRecCount > 0 Then
'        shtException.Cells.NumberFormat = "@"
'        shtException.Cells.WrapText = True
        
        lStartRow = fGetshtExceptionNewRow
        arrNewHospital = fConvertDictionaryKeysTo2DimenArrayForPaste(dictNewHospital, False)
'        lStartRow = fGetValidMaxRow(shtException)
'        If lStartRow = 0 Then
'            lStartRow = lStartRow + 1
'        Else
'            lStartRow = lStartRow + 5
'        End If
        
        Call fPrepareHeaderToSheet(shtException, Array("本系统中找不到的医院", "行号"), lStartRow)
        
        shtException.Rows(lStartRow).Font.Color = RGB(255, 0, 0)
        shtException.Rows(lStartRow).Font.Bold = True
        
        Call fAppendArray2Sheet(shtException, arrNewHospital)
        shtException.Cells(lStartRow + 1, 2).Resize(dictNewHospital.count, 1).Value = fConvertDictionaryItemsTo2DimenArrayForPaste(dictNewHospital, False)

        Call fFreezeSheet(shtException)
        
'        shtException.Visible = xlSheetVisible
'        shtException.Activate
        
        If fNzero(gsBusinessErrorMsg) Then gsBusinessErrorMsg = gsBusinessErrorMsg & vbCr & vbCr & vbCr & "===============================" & vbCr & vbCr
        
        gsBusinessErrorMsg = gsBusinessErrorMsg & lRecCount & "个【医院】在本系统中找不到，" & vbCr _
            & "但是它们都被自动加入到了【医院】主表中" & vbCr _
            & "您需要在【医院替换】表中增加替换，并尝试运行，这样医院就会趋向统一。" & vbCr & vbCr _
            & "并把医院主表中的不用的记录删除掉"
    End If
End Function

Function fAddNewFoundMissedProducerToSheetException(dictNewProducer As Dictionary)
    '======= Producer Validation ===============================================
    Dim arrNewProducer()
    'Dim sErr As String
    Dim lStartRow As Long
    Dim lRecCount As Long
    
    lRecCount = fGetDictionayDelimiteredItemsCount(dictNewProducer)
    If lRecCount > 0 Then
'        shtException.Cells.NumberFormat = "@"
'        shtException.Cells.WrapText = True
        
        lStartRow = fGetshtExceptionNewRow
        arrNewProducer = fConvertDictionaryKeysTo2DimenArrayForPaste(dictNewProducer, False)
'        lStartRow = fGetValidMaxRow(shtException)
'        If lStartRow = 0 Then
'            lStartRow = lStartRow + 1
'        Else
'            lStartRow = lStartRow + 5
'        End If
        
        Call fPrepareHeaderToSheet(shtException, Array("本系统中找不到的药品生产厂家", "行号"), lStartRow)
        
        shtException.Rows(lStartRow).Font.Color = RGB(255, 0, 0)
        shtException.Rows(lStartRow).Font.Bold = True
        
        Call fAppendArray2Sheet(shtException, arrNewProducer)
        'sErr = fUbound(arrNewProducer)
        shtException.Cells(lStartRow + 1, 2).Resize(dictNewProducer.count, 1).Value = fConvertDictionaryItemsTo2DimenArrayForPaste(dictNewProducer, False)
       ' Erase arrNewProducer
        Call fFreezeSheet(shtException)
        
        shtException.Visible = xlSheetVisible
        shtException.Activate
        
        If fNzero(gsBusinessErrorMsg) Then gsBusinessErrorMsg = gsBusinessErrorMsg & vbCr & vbCr & vbCr & "===============================" & vbCr & vbCr
        
        gsBusinessErrorMsg = gsBusinessErrorMsg & lRecCount & "个药品【生产厂家】在本系统中找不到，您可能要：" & vbCr _
            & "(1). 在【药品厂家替换表】中添加一条替换记录" & vbCr _
            & "(2). 在【药品厂家主表】中新增一个厂家" & vbCr & vbCr _
            & "本次导入失败，完善数据后，请再次点击按钮进行【匹配替换统一】"
    End If
    '======= Producer end ===============================================
End Function

Private Function fAddNewFoundMissedProductNameToSheetException(dictNewProductName As Dictionary)
    '======= ProductName Validation ===============================================
    Dim arrNewProductName()
    'Dim sErr As String
    Dim lStartRow As Long
    Dim lRecCount As Long
    
    lRecCount = fGetDictionayDelimiteredItemsCount(dictNewProductName)
        
    If lRecCount > 0 Then
'        shtException.Cells.NumberFormat = "@"
'        shtException.Cells.WrapText = True
        lStartRow = fGetshtExceptionNewRow
        
        arrNewProductName = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictNewProductName, , False)
'        lStartRow = fGetValidMaxRow(shtException)
'        If lStartRow = 0 Then
'            lStartRow = lStartRow + 1
'        Else
'            lStartRow = lStartRow + 5
'        End If
        
        Call fPrepareHeaderToSheet(shtException, Array("药品厂家", "本系统中找不到的药品名称", "行号"), lStartRow)
        shtException.Rows(lStartRow).Font.Color = RGB(255, 0, 0)
        shtException.Rows(lStartRow).Font.Bold = True
        Call fAppendArray2Sheet(shtException, arrNewProductName)
        'sErr = fUbound(arrNewProductName)
        shtException.Cells(lStartRow + 1, 3).Resize(dictNewProductName.count, 1).Value = fConvertDictionaryItemsTo2DimenArrayForPaste(dictNewProductName, False)
        'Erase arrNewProductName
        Call fFreezeSheet(shtException)
        
'        shtException.Visible = xlSheetVisible
'        shtException.Activate
        
        If fNzero(gsBusinessErrorMsg) Then gsBusinessErrorMsg = gsBusinessErrorMsg & vbCr & vbCr & vbCr & "===============================" & vbCr & vbCr
        
        gsBusinessErrorMsg = gsBusinessErrorMsg & lRecCount & "个药品【名称】在本系统中找不到，您可能要：" & vbCr _
            & "(1). 在【药品名称替换表】中添加一条替换记录" & vbCr _
            & "(2). 在【药品名称主表】中新增一个名称" & vbCr _
            & "** 请注意：药品厂家没有问题，都匹配到了。" & vbCr & vbCr _
            & "本次导入失败，完善数据后，请再次点击按钮进行【匹配替换统一】"
    End If
    '======= ProductName end ===============================================
End Function

Function fAddNewFoundMissedProductSeriesToSheetException(dictNewProductSeries As Dictionary)
    '======= ProductSeries Validation ===============================================
    Dim arrNewProductSeries()
    'Dim sErr As String
    Dim lRecCount As Long
        Dim lStartRow As Long
    
    lRecCount = fGetDictionayDelimiteredItemsCount(dictNewProductSeries)
    If lRecCount > 0 Then
        lStartRow = fGetshtExceptionNewRow
        
        arrNewProductSeries = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictNewProductSeries, , False)
        
        Call fPrepareHeaderToSheet(shtException, Array("药品厂家", "药品名称", "本系统中找不到的药品【规格】", "行号"), lStartRow)
        shtException.Rows(lStartRow).Font.Color = RGB(255, 0, 0)
        shtException.Rows(lStartRow).Font.Bold = True
        Call fAppendArray2Sheet(shtException, arrNewProductSeries)
        'sErr = fUbound(arrNewProductSeries)
        
        shtException.Cells(lStartRow + 1, 4).Resize(dictNewProductSeries.count, 1).Value = fConvertDictionaryItemsTo2DimenArrayForPaste(dictNewProductSeries, False)
       ' Erase arrNewProductSeries
        Call fFreezeSheet(shtException)
        
'        shtException.Visible = xlSheetVisible
'        shtException.Activate
        
        If fNzero(gsBusinessErrorMsg) Then gsBusinessErrorMsg = gsBusinessErrorMsg & vbCr & vbCr & vbCr & "===============================" & vbCr & vbCr

        gsBusinessErrorMsg = gsBusinessErrorMsg & lRecCount & "个药品【规格】在本系统中找不到，您可能要：" & vbCr _
            & "(1). 在【药品规格替换表】中添加一条替换记录" & vbCr _
            & "(2). 在【药品主表】中新增一个规格" & vbCr _
            & "** 请注意：药品厂家和药品名称没有问题，都匹配到了。" & vbCr & vbCr _
            & "本次导入失败，完善数据后，请再次点击按钮进行【匹配替换统一】"
    End If
End Function

Function fAddNewFoundMissedProductUnitToSheetException(dictNewProductUnit As Dictionary, dictNewProductUnitOrig As Dictionary)
    Dim arrNewProductUnit()
    'Dim sErr As String
    Dim lRecCount As Long
        Dim lStartRow As Long
    
    lRecCount = fGetDictionayDelimiteredItemsCount(dictNewProductUnit)
    
    If lRecCount > 0 Then
'        shtException.Cells.NumberFormat = "@"
'        shtException.Cells.WrapText = True
        lStartRow = fGetshtExceptionNewRow
        
        arrNewProductUnit = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictNewProductUnit, , False)
'        lStartRow = fGetValidMaxRow(shtException)
'        If lStartRow = 0 Then
'            lStartRow = lStartRow + 1
'        Else
'            lStartRow = lStartRow + 5
'        End If
        
        Call fPrepareHeaderToSheet(shtException, Array("药品厂家", "药品名称", "药品规格", "药品会计单位", "原始文件单位", "行号"), lStartRow)
        shtException.Rows(lStartRow).Font.Color = RGB(255, 0, 0)
        shtException.Rows(lStartRow).Font.Bold = True
        Call fAppendArray2Sheet(shtException, arrNewProductUnit)
        'sErr = fUbound(arrNewProductUnit)
            
        shtException.Cells(lStartRow + 1, 5).Resize(dictNewProductUnitOrig.count, 1).Value = fConvertDictionaryItemsTo2DimenArrayForPaste(dictNewProductUnitOrig, False)
        shtException.Cells(lStartRow + 1, 6).Resize(dictNewProductUnit.count, 1).Value = fConvertDictionaryItemsTo2DimenArrayForPaste(dictNewProductUnit, False)
       ' Erase arrNewProductUnit
        Call fFreezeSheet(shtException)
        
'        shtException.Visible = xlSheetVisible
'        shtException.Activate
        
        If fNzero(gsBusinessErrorMsg) Then gsBusinessErrorMsg = gsBusinessErrorMsg & vbCr & vbCr & vbCr & "===============================" & vbCr & vbCr

        gsBusinessErrorMsg = gsBusinessErrorMsg & lRecCount & "个药品【单位】和设定的会计单位不一致，您可能要：" & vbCr _
            & "(1). 在【药品单位倍数表】中添加一条记录" & vbCr _
            & "(2). 在【药品主表】中修改其单位" & vbCr _
            & "** 请注意：药品厂家、名称、规格没有问题，都匹配到了。" & vbCr & vbCr _
            & "本次导入失败，完善数据后，请再次点击按钮进行【匹配替换统一】"
    End If
    '======= ProductUnit end ===============================================
End Function
Function fReplaceAndValidateInHospitalMaster(sHospital As String, ByRef sReplacedHospital As String) As Boolean
    sReplacedHospital = fFindInConfigedReplaceHospital(sHospital)
    If fZero(sReplacedHospital) Then sReplacedHospital = sHospital
    
    fReplaceAndValidateInHospitalMaster = fHospitalExistsInHospitalMaster(sReplacedHospital)
End Function

Function fReplaceAndValidateInProducerMaster(sProducer As String, ByRef sReplacedProducer As String) As Boolean
    sReplacedProducer = fFindInConfigedReplaceProducer(sProducer)
    If fZero(sReplacedProducer) Then sReplacedProducer = sProducer
    
    fReplaceAndValidateInProducerMaster = fProducerExistsInProducerMaster(sReplacedProducer)
End Function

Function fReplaceAndValidateInProductNameMaster(sReplacedProducer As String, sProductName As String, ByRef sReplacedProductName As String) As Boolean
    sReplacedProductName = fFindInConfigedReplaceProductName(sReplacedProducer, sProductName)
    If fZero(sReplacedProductName) Then sReplacedProductName = sProductName
    
    fReplaceAndValidateInProductNameMaster = fProductNameExistsInProductNameMaster(sReplacedProducer, sReplacedProductName)
End Function
'ProductSeriesMaster = ProductMaster
Function fReplaceAndValidateInProductSeriesMaster(sReplacedProducer As String, sReplacedProductName As String _
                                        , sProductSeries As String, ByRef sReplacedProductSeries As String) As Boolean
    sReplacedProductSeries = fFindInConfigedReplaceProductSeries(sReplacedProducer, sReplacedProductName, sProductSeries)
    If fZero(sReplacedProductSeries) Then sReplacedProductSeries = sProductSeries
    
    fReplaceAndValidateInProductSeriesMaster = fProductSeriesExistsInProductMaster(sReplacedProducer, sReplacedProductName, sReplacedProductSeries)
End Function

''ProductUnitMaster = ProductMaster
'Function fReplaceAndValidateInProductUnitMaster(sReplacedProducer As String, sReplacedProductName As String, sReplacedProductSeries As String _
'                                        , sProductUnit As String _
'                                        , ByRef sReplacedProductUnit As String _
'                                        , ByRef dblRatio As Double) As Boolean
'    sReplacedProductUnit = fFindInConfigedReplaceProductUnit(sReplacedProducer, sReplacedProductName, sReplacedProductSeries _
'                                                        , sProductUnit, dblRatio)
'  '  sProductMasterUnit = fGetProductUnit(sReplacedProducer, sReplacedProductName, sReplacedProductSeries)
'
''    If Len(Trim(sReplacedProductUnit)) <= 0 Then
''        sReplacedProductUnit = sProductUnit
''    Else
''        If sReplacedProductUnit <> sProductMasterUnit Then
''            fErr "【" & shtProductUnitRatio.Name & "】会计单位和药品主表不一样，请检查【" & shtProductUnitRatio.Name & "】" _
''                & vbCr & sReplacedProducer _
''                & vbCr & sReplacedProductName _
''                & vbCr & sReplacedProductSeries & vbCr & sProductMasterUnit
''        End If
''
''        If sProductUnit = sProductMasterUnit Then
''            If dblRatio = 1 Then
''                fErr "原始文件单位和会计单位一样，但是倍数却不是1，请检查【" & shtProductUnitRatio.Name & "】" _
''                    & vbCr & sReplacedProducer _
''                    & vbCr & sReplacedProductName _
''                    & vbCr & sReplacedProductSeries & vbCr & sProductMasterUnit
''            End If
''        End If
'    End If
'
'    'fReplaceAndValidateInProductUnitMaster = (sReplacedProductUnit = sProductMasterUnit)
'End Function


'Function fGetConvertUnitAndUnitRatio() As Boolean
'    Dim bOut As Boolean
'
'        sProductMasterUnit = fGetProductUnit(sReplacedProducer, sReplacedProductName, sReplacedProductSeries)
'
'        dblRatio = 0
'        If Len(Trim(sProductUnit)) <= 0 Then
'            arrOutput(lEachRow, dictRptColIndex("MatchedProductUnit")) = sProductMasterUnit
'            dblRatio = 1
'        Else
''            Call fReplaceAndValidateInProductUnitMaster(sReplacedProducer, sReplacedProductName, sReplacedProductSeries, sProductUnit _
''                                                , sReplacedProductUnit, dblRatio)
'
'            sReplacedProductUnit = fFindInConfigedReplaceProductUnit(sReplacedProducer, sReplacedProductName, sReplacedProductSeries _
'                                                        , sProductUnit, dblRatio)
'            If Len(Trim(sReplacedProductUnit)) <= 0 Then
'                sReplacedProductUnit = sProductUnit
'            Else
'                If sReplacedProductUnit <> sProductMasterUnit Then
'                    fErr "【" & shtProductUnitRatio.Name & "】会计单位和药品主表不一样，请检查【" & shtProductUnitRatio.Name & "】" _
'                        & vbCr & sReplacedProducer _
'                        & vbCr & sReplacedProductName _
'                        & vbCr & sReplacedProductSeries & vbCr & sProductMasterUnit
'                End If
'
'                If sProductUnit = sProductMasterUnit Then
'                    If dblRatio = 1 Then
'                        fErr "原始文件单位和会计单位一样，但是倍数却不是1，请检查【" & shtProductUnitRatio.Name & "】" _
'                            & vbCr & sReplacedProducer _
'                            & vbCr & sReplacedProductName _
'                            & vbCr & sReplacedProductSeries & vbCr & sProductMasterUnit
'                    End If
'                End If
'            End If
'
'            If sReplacedProductUnit <> sProductMasterUnit Then
'                sTmpKey = sReplacedProducer & DELIMITER & sReplacedProductName & DELIMITER & sReplacedProductSeries & DELIMITER _
'                                & sProductMasterUnit & DELIMITER & sReplacedProductUnit
'                If Not dictNewProductUnit.Exists(sTmpKey) Then
'                    dictNewProductUnit.Add sTmpKey, lEachRow + 1
'                End If
'
'                arrOutput(lEachRow, dictRptColIndex("MatchedProductUnit")) = ""
'                GoTo next_sales
'            End If
'
'            arrOutput(lEachRow, dictRptColIndex("MatchedProductUnit")) = sReplacedProductUnit
'        End If
'End Function

