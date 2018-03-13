Attribute VB_Name = "Common_Oper_Excel"
Option Explicit
Option Base 1

Function fOpenFileSelectDialogAndSetToSheetRange(rngAddrOrName As String _
                            , Optional asFileFilters As String = "" _
                            , Optional asTitle As String = "" _
                            , Optional shtParam As Worksheet)
    Dim sFile As String
    
    If shtParam Is Nothing Then Set shtParam = shtMenu
    
    sFile = fSelectFileDialog(Trim(shtParam.Range(rngAddrOrName).Value), , asTitle)
    If Len(sFile) > 0 Then shtParam.Range(rngAddrOrName).Value = sFile
End Function

Function fFindInWorksheet(rngToFindIn As Range, sWhatToFind As String _
                    , Optional abNotFoundThenError As Boolean = True _
                    , Optional abAllowMultiple As Boolean = False) As Range
    If Len(Trim(sWhatToFind)) <= 0 Then fErr "Wrong param sWhatToFind to fFindInWorksheet " & sWhatToFind
    
    Dim rngOut  As Range
    Dim rngFound As Range
    Dim lFoundCnt As Long
    Dim sFirstAddress As String
    
    Set rngFound = rngToFindIn.Find(What:=sWhatToFind _
                                    , after:=rngToFindIn.Cells(rngToFindIn.Rows.Count, rngToFindIn.Columns.Count) _
                                    , LookIn:=xlValues _
                                    , LookAt:=xlWhole _
                                    , SearchOrder:=xlByRows _
                                    , SearchDirection:=xlNext _
                                    , MatchCase:=False _
                                    , MatchByte:=False)
    Set rngOut = rngFound
    
    If rngFound Is Nothing Then
        If abNotFoundThenError Then
            fErr """" & sWhatToFind & """ cannot be found in sheet " & rngToFindIn.Parent.Name & "[" & rngToFindIn.Address & "], pls check your program."
        Else
            GoTo exit_function
        End If
    Else
        If Not abAllowMultiple Then
            sFirstAddress = rngFound.Address
            lFoundCnt = 1
            
            Do While True
                Set rngFound = rngToFindIn.Find(What:=sWhatToFind _
                                            , after:=rngFound _
                                            , LookIn:=xlValues _
                                            , LookAt:=xlWhole _
                                            , SearchOrder:=xlByRows _
                                            , SearchDirection:=xlNext _
                                            , MatchCase:=False _
                                            , MatchByte:=False)
                If rngFound Is Nothing Then Exit Do
                If rngFound.Address = sFirstAddress Then Exit Do
                
                lFoundCnt = lFoundCnt + 1
            Loop
            
            If lFoundCnt > 1 Then
                fErr lFoundCnt & " copies of """ & sWhatToFind & """ were found in sheet " & rngToFindIn.Parent.Name & ", pls check your program."
            End If
        End If
    End If
exit_function:
    Set fFindInWorksheet = rngOut
    Set rngOut = Nothing
    Set rngFound = Nothing
End Function

Function fGetRangeByStartEndPos(shtParam As Worksheet, alStartRow As Long, alStartCol As Long, alEndRow As Long, alEndCol As Long) As Range
    With shtParam
        Set fGetRangeByStartEndPos = .Range(.Cells(alStartRow, alStartCol), .Cells(alEndRow, alEndCol))
    End With
End Function

Function fReadRangeDatatoArrayByStartEndPos(shtParam As Worksheet, alStartRow As Long, alStartCol As Long, alEndRow As Long, alEndCol As Long) As Variant
    fReadRangeDatatoArrayByStartEndPos = fReadRangeDataToArray(fGetRangeByStartEndPos(shtParam, alStartRow, alStartCol, alEndRow, alEndCol))
End Function

Function fReadRangeDataToArray(rngParam As Range) As Variant
    Dim arrOut()
    
    If fRangeIsSingleCell(rngParam) Then
        ReDim arrOut(1 To 1, 1 To 1)
        arrOut(1, 1) = rngParam.Value
    Else
        arrOut = rngParam.Value
    End If
    
    fReadRangeDataToArray = arrOut
    Erase arrOut
End Function

Function fSetSpecifiedConfigCellValue(shtConfig As Worksheet, asTag As String, asRtnCol As String, asCriteria As String _
                                , sValue As String _
                                , Optional bDevUatProd As Boolean = False _
                                )
    Dim sAddr As String
    sAddr = fGetSpecifiedConfigCellAddress(shtConfig, asTag, asRtnCol, asCriteria, False, bDevUatProd)
    shtConfig.Range(sAddr).Value = sValue
End Function
Function fGetSpecifiedConfigCellValue(shtConfig As Worksheet, asTag As String, asRtnCol As String, asCriteria As String _
                                , Optional bDevUatProd As Boolean = False _
                                )
    Dim sAddr As String
    sAddr = fGetSpecifiedConfigCellAddress(shtConfig, asTag, asRtnCol, asCriteria, False, bDevUatProd)
    fGetSpecifiedConfigCellValue = shtConfig.Range(sAddr).Value
End Function
Function fGetSpecifiedConfigCellAddress(shtConfig As Worksheet, asTag As String, asRtnCol As String _
                                , asCriteria As String _
                                , Optional bAllowMultiple As Boolean = False _
                                , Optional bDevUatProd As Boolean = False _
                                )
    'asCriteria: colA=Value01, colB=Value02
    Dim arrColNames()
    Dim arrColValues()
    Dim iRtnColIndex As Integer
    Dim iEnvColIndex As Integer
    
    Call fSplitDataCriteria(asCriteria, arrColNames, arrColValues)
    
    iRtnColIndex = fEnlargeArayWithValue(arrColNames, asRtnCol)
    
    'DEV/UAT/PROD must be put at the end of arrColsName, since fFindMatchDataInArrayWithCriteria will read it.
    If bDevUatProd Then
        iEnvColIndex = fEnlargeArayWithValue(arrColNames, "DEV/UAT/PROD")
    End If
    
    Dim lConfigStartRow As Long _
        , lConfigStartCol As Long _
        , lConfigEndRow As Long _
        , lOutConfigHeaderAtRow As Long _
        , bNetValues As Boolean
    Dim arrConfigData()
    Dim arrColsIndex()

    Call fReadConfigBlockToArray(asTag:=asTag, shtParam:=shtConfig, arrColsName:=arrColNames _
                                , arrConfigData:=arrConfigData _
                                , arrColsIndex:=arrColsIndex _
                                , lConfigStartRow:=lConfigStartRow _
                                , lConfigStartCol:=lConfigStartCol _
                                , lConfigEndRow:=lConfigEndRow _
                                , lOutConfigHeaderAtRow:=lOutConfigHeaderAtRow _
                                , abNoDataConfigThenError:=True _
                                )
    Dim lMatchRow As Long
    Dim sErr As String
    lMatchRow = fFindMatchDataInArrayWithCriteria(arrConfigData, arrColsIndex, arrColValues, bAllowMultiple, sErr, bDevUatProd)
    
    If lMatchRow < 0 Then
        fErr sErr & " with criteria " & vbCr & asCriteria & vbCr & "gsEnv: " & gsEnv & vbCr & shtConfig.Name & vbCr & asTag
    End If
    
    fGetSpecifiedConfigCellAddress = shtConfig.Cells(lOutConfigHeaderAtRow + lMatchRow, lConfigStartCol + arrColsIndex(iRtnColIndex) - 1).Address(external:=True)
End Function

Function fFindMatchDataInArrayWithCriteria(arr(), arrColsIndex(), arrColValues() _
                                        , bAllowMultiple As Boolean _
                                        , ByRef asErrmsg As String _
                                , Optional bDevUatProd As Boolean = False _
                                        ) As Long
'-1:
' -2: more than 1 matched
' -3: no match
    Dim lEachRow As Long
    Dim lEachCol As Long
    Dim i As Integer
    Dim bAllColAreSame As Boolean
    Dim lMatchCnt As Long
    Dim lOut As Long
    Dim sEachEnv As String
    
    If bDevUatProd And fZero(gsEnv) Then fErr "bDevUatProd is true but gsenv = blank"
    
    asErrmsg = ""
    lOut = -1
    lMatchCnt = 0
    For lEachRow = LBound(arr, 1) To UBound(arr, 1)
        If fArrayRowIsBlankHasNoData(arr, lEachRow) Then GoTo next_row
        
        If bDevUatProd Then
            sEachEnv = arr(lEachRow, arrColsIndex(UBound(arrColsIndex)))
            If Not (sEachEnv = gsEnv Or sEachEnv = "SHARED") Then GoTo next_row
        End If
        
        bAllColAreSame = True
        For i = LBound(arrColValues) To UBound(arrColValues)
            lEachCol = arrColsIndex(i)
            
            If Trim(CStr(arr(lEachRow, lEachCol))) <> arrColValues(i) Then
                bAllColAreSame = False
                GoTo next_row
            End If
        Next
        
        If bAllColAreSame Then
            lMatchCnt = lMatchCnt + 1
            lOut = lEachRow
            
            If bAllowMultiple Then GoTo exit_fun
        End If
next_row:
    Next
    
    If lMatchCnt > 1 Then
        If Not bAllowMultiple Then
            lOut = -2
            asErrmsg = lMatchCnt & " records were matched "
        End If
    ElseIf lMatchCnt <= 0 Then
        lOut = -3
        asErrmsg = "No record were matched "
    End If
exit_fun:
    fFindMatchDataInArrayWithCriteria = lOut
End Function

Function fSplitDataCriteria(asCriteria As String, ByRef arrColNames(), ByRef arrColValues())
    'asCriteria: colA=Value01, colB=Value02
    Dim arrCriteria
    Dim sCol As String
    Dim sValue As String
    Dim i As Integer
    Dim sEachCriteria As String
    
    arrCriteria = Split(asCriteria, ",")
    
    ReDim arrColNames(LBound(arrCriteria) To UBound(arrCriteria))
    ReDim arrColValues(LBound(arrCriteria) To UBound(arrCriteria))
    
    For i = LBound(arrCriteria) To UBound(arrCriteria)
        sEachCriteria = Trim(arrCriteria(i))    ' colA=Value01
        
        sCol = Trim(Split(sEachCriteria, "=")(0))
        sValue = Trim(Split(sEachCriteria, "=")(1))
        
        arrColNames(i) = sCol
        arrColValues(i) = sValue
    Next
    
    Erase arrCriteria
End Function

Function fWriteArray2Sheet(sht As Worksheet, arrData, Optional lStartRow As Long = 2, Optional lStartCol As Long = 1)
    If fArrayIsEmptyOrNoData(arrData) Then Exit Function
    
    If fGetArrayDimension(arrData) <> 2 Then
        fErr "Wrong array to paste to sheet: fGetArrayDimension(arrData) <> 2"
    End If
    
    sht.Cells(lStartRow, lStartCol).Resize(UBound(arrData, 1), UBound(arrData, 2)).Value = arrData
End Function

Function fAppendArray2Sheet(sht As Worksheet, ByRef arrData, Optional lStartCol As Long = 1, Optional bEraseArray As Boolean = True)
    If fArrayIsEmptyOrNoData(arrData) Then Exit Function
    
'    If fGetArrayDimension(arrData) <> 2 Then
'        fErr "Wrong array to paste to sheet: fGetArrayDimension(arrData) <> 2"
'    End If
    
    Dim lFromRow As Long
    lFromRow = fGetValidMaxRow(sht) + 1
    
    sht.Cells(lFromRow, lStartCol).Resize(UBound(arrData, 1), UBound(arrData, 2)).Value = arrData
    If bEraseArray Then Erase arrData
End Function

Function fAutoFilterAutoFitSheet(sht As Worksheet, Optional alMaxCol As Long = 0 _
                                , Optional ColumnWidthAuto As Boolean = True)

    Dim lMaxCol As Long
    
    If alMaxCol > 0 Then
        lMaxCol = alMaxCol
    Else
        lMaxCol = fGetValidMaxCol(sht)
    End If
    
    If sht.AutoFilterMode Then sht.AutoFilterMode = False
    
    fGetRangeByStartEndPos(sht, 1, 1, 1, lMaxCol).AutoFilter
    
    If ColumnWidthAuto Then sht.Cells.EntireColumn.AutoFit
    sht.Cells.EntireRow.AutoFit
End Function

Function fFreezeSheet(sht As Worksheet, Optional alSplitCol As Long = 0, Optional alSplitRow As Long = 1)
    sht.Activate
    ActiveWindow.FreezePanes = False
    ActiveWindow.SplitColumn = alSplitCol
    ActiveWindow.SplitRow = alSplitRow
    ActiveWindow.FreezePanes = True
End Function

Function fExcelFileIsOpen(sExcelFileName As String, Optional wbOut As Workbook) As Boolean
    On Error Resume Next
    Set wbOut = Workbooks(fGetFileBaseName(sExcelFileName))
    Err.Clear
    
    fExcelFileIsOpen = (Not wbOut Is Nothing)
End Function

Function fRemoveFilterForAllSheets(Optional wb As Workbook, Optional ByVal asDegree As String = "SHOW_ALL_DATA")
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    On Error GoTo error_handling
    Dim sht As Worksheet
    For Each sht In wb.Worksheets
        Call fRemoveFilterForSheet(sht, asDegree)
    Next
    
error_handling:
    If Err.Number <> 0 Then
        MsgBox sht.Name & vbCr & Err.Description
    End If
    Set sht = Nothing
End Function
Function fDeleteBlankRowsFromAllSheets(Optional wb As Workbook)
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    Dim sht As Worksheet
    For Each sht In wb.Worksheets
        Call fDeleteBlankRowsFromSheet(sht)
    Next
    
    Set sht = Nothing
End Function
Function fDeleteBlankRowsFromSheet(sht As Worksheet)
    Dim lUsedRangMaxRow As Long
    Dim lValidMaxRow As Long
    
    lUsedRangMaxRow = sht.UsedRange.Row + sht.UsedRange.Rows.Count - 1
    lValidMaxRow = fGetValidMaxRow(sht)
    
    If lUsedRangMaxRow > lValidMaxRow Then
        sht.Rows((lValidMaxRow + 1) & ":" & lUsedRangMaxRow).Delete shift:=xlUp
    End If
End Function
Function fRemoveFilterForSheet(sht As Worksheet, Optional ByVal asDegree As String = "SHOW_ALL_DATA")
    asDegree = UCase(Trim(asDegree))
    
    Dim rgActiveCell As Range
    Set rgActiveCell = ActiveCell
    
    If sht.FilterMode Then  'advanced filter
        sht.ShowAllData
    End If
    
    If sht.AutoFilterMode Then  'auto filter
        If fZero(asDegree) Or asDegree = "SHOW_ALL_DATA" Then
            sht.AutoFilter.ShowAllData
        Else
            sht.AutoFilterMode = False
        End If
    End If
    
    rgActiveCell.Select
    Set rgActiveCell = Nothing
End Function

Function fActiveXControlExistsInSheet(sht As Worksheet, asControlName As String, ByRef objOut As Object) As Boolean
    Dim bOut As Boolean
    bOut = False
    
    On Error GoTo err_exit
    Set objOut = CallByName(sht, asControlName, VbGet)
    bOut = True
    
err_exit:
    Err.Clear
    fActiveXControlExistsInSheet = bOut
End Function
Function fActiveXControlExistsInSheet2(sht As Worksheet, asControlName As String, ByRef objOut As Object) As Boolean
    Dim obj As Object
    Dim bOut As Boolean
    bOut = False
    
    For Each obj In sht.DrawingObjects
        If obj.Name = asControlName Then
            Set objOut = obj
            bOut = True
            Exit For
        End If
    Next
    
err_exit:
    Set obj = Nothing
    fActiveXControlExistsInSheet2 = bOut
End Function
 
Function fSheetExists(asShtName As String, Optional ByRef shtOut As Worksheet, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet
    Dim bOut As Boolean
    
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    bOut = False
    asShtName = UCase(Trim(asShtName))
    
    For Each sht In wb.Worksheets
        If UCase(sht.Name) = asShtName Then
            bOut = True
            Set shtOut = sht
            Exit For
        End If
    Next
    
    Set sht = Nothing
    fSheetExists = bOut
End Function

Function fIfExcelFileOpenedToCloseIt(sExcelFile As String)
    If fExcelFileIsOpen(sExcelFile) Then
        fErr "Excel File is open, pleae close it first." _
             & vbCr & fGetFileBaseName(sExcelFile)
    End If
End Function
 
Function fSortDataInSheetSortSheetData(sht As Worksheet, arrSortByCols, Optional arrAscend)
'arrAscendDescend : array(true, false, true)
    Dim lMaxRow As Long
    Dim lMaxCol As Long
    Dim rgToSort As Range
    Dim rgSortBy As Range
    Dim iSortOrder As XlSortOrder
    Dim i As Long
    Dim lSortCol As Long
    
    lMaxRow = fGetValidMaxRow(sht)
    lMaxCol = fGetValidMaxCol(sht)
    
    If lMaxRow <= 0 Or lMaxCol <= 0 Then Exit Function

    Set rgToSort = fGetRangeByStartEndPos(sht, 1, 1, lMaxRow, lMaxCol)
    
    sht.Sort.SortFields.Clear
    
    If IsArray(arrSortByCols) Then
        For i = LBound(arrSortByCols) To UBound(arrSortByCols)
            lSortCol = arrSortByCols(i)
            
            If Not IsMissing(arrAscend) Then
                If arrAscend(i) Then
                    iSortOrder = xlAscending
                Else
                    iSortOrder = xlDescending
                End If
            Else
                iSortOrder = xlAscending
            End If
            
            Set rgSortBy = fGetRangeByStartEndPos(sht, 2, lSortCol, lMaxRow, lSortCol)
            sht.Sort.SortFields.Add Key:=rgSortBy, SortOn:=xlSortOnValues _
                    , Order:=iSortOrder, DataOption:=xlSortNormal
        Next
    Else
        lSortCol = arrSortByCols
        
        If Not IsMissing(arrAscend) Then
            If arrAscend Then
                iSortOrder = xlAscending
            Else
                iSortOrder = xlDescending
            End If
        Else
            iSortOrder = xlAscending
        End If
            
        Set rgSortBy = fGetRangeByStartEndPos(sht, 2, lSortCol, lMaxRow, lSortCol)
        sht.Sort.SortFields.Add Key:=rgSortBy, SortOn:=xlSortOnValues _
                , Order:=iSortOrder, DataOption:=xlSortNormal
    End If
    
    sht.Sort.SetRange rgToSort
    sht.Sort.Header = xlYes
    sht.Sort.MatchCase = False
    sht.Sort.Orientation = xlTopToBottom
    sht.Sort.SortMethod = xlPinYin
    sht.Sort.Apply
    
    Set rgToSort = Nothing
    Set rgSortBy = Nothing
End Function

Function fProtectSheetAndAllowEdit(sht As Worksheet, Optional rngAllowEdit As Range _
                    , Optional lMaxRow As Long = 0, Optional lMaxCol As Long = 0 _
                    , Optional bLockColor As Boolean = True)

    If lMaxRow <= 0 Then lMaxRow = fGetValidMaxRow(sht)
    
    If lMaxRow < 2 Then Exit Function
    
    If lMaxCol <= 0 Then lMaxCol = fGetValidMaxCol(sht)
    
    Dim rgUsed As Range
    Set rgUsed = fGetRangeByStartEndPos(sht, 2, 1, lMaxRow, lMaxCol)
    
    If bLockColor Then
        rgUsed.Interior.Color = 13553360
        
        If Not rngAllowEdit Is Nothing Then
            'rngAllowEdit.Interior.Color = RGB(255, 255, 255)
            rngAllowEdit.Interior.Pattern = xlNone
            rngAllowEdit.Interior.TintAndShade = 0
            rngAllowEdit.Interior.PatternTintAndShade = 0
        End If
    End If
    
    sht.Cells.Locked = True
    
    If Not rngAllowEdit Is Nothing Then
        rngAllowEdit.Locked = False
        rngAllowEdit.FormulaHidden = False
    End If
    
    Set rgUsed = Nothing
    
    sht.Protect userinterfaceonly:=True, Password:=PW_PROTECT_SHEET _
        , DrawingObjects:=True _
        , Contents:=True _
        , Scenarios:=True _
        , AllowFormattingCells:=True _
        , AllowFormattingColumns:=True _
        , AllowFormattingRows:=True _
        , AllowInsertingColumns:=False _
        , AllowInsertingRows:=False _
        , AllowInsertingHyperlinks:=True _
        , AllowDeletingColumns:=False _
        , AllowDeletingRows:=False _
        , AllowSorting:=True _
        , AllowFiltering:=True _
        , AllowUsingPivotTables:=False

End Function

Function fUnProtectSheet(sht As Worksheet)
    sht.Unprotect Password:=PW_PROTECT_SHEET
End Function

Function fSetValidationListForRange(rngParam As Range, asValueListOrExternalAddr As String)
'asValueListOrExternalAddr
' 1) Formula1:="=$K$5:$K$21"    --> range().address( external:=true)
' 2) Formula1:="a,b,c,d,e,f,g"  --> "a, b, c, d, e, f"

    With rngParam.Validation
        .Delete
        .Add Type:=xlValidateList _
            , AlertStyle:=xlValidAlertStop _
            , Operator:=xlBetween _
            , Formula1:=asValueListOrExternalAddr
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
End Function

Function fModifyMoveActiveXButtonOnSheet(rngToPlaceTheButton As Range, sBtnTechName As String _
                                , Optional dblOffsetLeft As Double = 0, Optional dblOffsetTop As Double = 0 _
                                , Optional dblWidth As Double = 0, Optional dblHeight As Double = 0 _
                                , Optional lBackColor As Long = 0 _
                                , Optional lForeColor As Long = 0)
    Dim sht As Worksheet
    Set sht = rngToPlaceTheButton.Parent
    
    Dim obj As Object
    Dim eachObj As Object
    
    For Each eachObj In sht.OLEObjects
        If eachObj.Name = sBtnTechName Then
            Set obj = eachObj
            Exit For
        End If
    Next
    
    If obj Is Nothing Then fErr "Button " & sBtnTechName & " cannot be found in sheet " & sht.Name
    
    If dblWidth = 0 Then dblWidth = obj.Width
    If dblHeight = 0 Then dblHeight = obj.Height
    
    Dim offsetLeft As Double
    If dblOffsetLeft = 0 Then
        offsetLeft = (rngToPlaceTheButton.Width - dblWidth) / 2 - 2
    Else
        offsetLeft = dblOffsetLeft
    End If
    
    Dim offsetTop As Double
    If dblOffsetTop = 0 Then
        offsetTop = (rngToPlaceTheButton.Width - dblWidth) / 2 - 2
    Else
        offsetTop = dblOffsetTop
    End If
    
    obj.Left = rngToPlaceTheButton.Left + offsetLeft
    obj.Top = rngToPlaceTheButton.Top + offsetTop
    obj.Width = dblWidth
    obj.Height = dblHeight
    
    If lBackColor <> 0 Then obj.Object.BackColor = lBackColor
    If lForeColor <> 0 Then obj.Object.ForeColor = lForeColor
    
    Set obj = Nothing
    Set eachObj = Nothing
End Function

Function fSortDataInSheetSortSheetDataByFileSpec(asFileTag As String, arrSortByColNames, Optional arrAscend _
                                    , Optional shtData As Worksheet)
    Dim sFileSpecTag As String
    Dim shtToRead As Worksheet
    Dim dictColIndex As Dictionary
    Dim arrSortByCols()
    
    sFileSpecTag = fGetInputFileFileSpecTag(asFileTag)
    
    If shtData Is Nothing Then
        Set shtToRead = fGetInputFileSheetAfterLoadingToThisWorkBook(asFileTag)
    Else
        Set shtToRead = shtData
    End If
    
    Set dictColIndex = fReadInputFileSpecConfigItem(asFileTag, "LETTER_INDEX", shtToRead)
    
    ReDim arrSortByCols(LBound(arrSortByColNames) To UBound(arrSortByColNames))
    
    Dim i As Integer
    For i = LBound(arrSortByColNames) To UBound(arrSortByColNames)
        If Not dictColIndex.Exists(arrSortByColNames(i)) Then fErr "Column name does not exists " & arrSortByColNames(i)
        arrSortByCols(i) = dictColIndex(arrSortByColNames(i))
    Next
    
    Call fSortDataInSheetSortSheetData(shtToRead, arrSortByCols, arrAscend)
    
    Set dictColIndex = Nothing
    Set shtToRead = Nothing
End Function

Function fGetSelectedRowCount(rngParam As Range)
'    Dim eachArea As Range
'    Dim lRowCnt As Long
'
'    lRowCnt = 0
'    For Each eachArea In rngParam.Areas
'        lRowCnt = lRowCnt + eachArea.Rows.Count
'    Next
'    Set eachArea = Nothing
'
'    fGetSelectedRowCount = lRowCnt
End Function

Function fIfSelectedMoreThanOneRow(rngParam As Range) As Boolean
    Dim eachArea As Range
    Dim lRowNoSaved As Long
    Dim bOut As Boolean
    
    bOut = False
    lRowNoSaved = 0
    For Each eachArea In rngParam.Areas
        If eachArea.Rows.Count > 1 Then
            bOut = True
            Exit For
        Else
            If lRowNoSaved = 0 Then
                lRowNoSaved = eachArea.Row
            Else
                If eachArea.Row <> lRowNoSaved Then
                    bOut = True
                    Exit For
                End If
            End If
        End If
    Next
    Set eachArea = Nothing
    
    fIfSelectedMoreThanOneRow = bOut
End Function

Function fVeryHideSheet(ByRef sht As Worksheet)
    sht.Visible = xlSheetVeryHidden
End Function

Function fHideSheet(ByRef sht As Worksheet)
    sht.Visible = xlSheetHidden
End Function

Function fShowSheet(ByRef sht As Worksheet)
    sht.Visible = xlSheetVisible
End Function

Function fClearContentLeaveHeader(shtParam As Worksheet, Optional alHeaderByRow As Long = 1)
    Dim lMaxRow As Long
    lMaxRow = fGetValidMaxRow(shtParam)
    
    Dim lDataStartRow As Long
    lDataStartRow = alHeaderByRow + 1
    
    If lMaxRow > lDataStartRow Then
        With fGetRangeByStartEndPos(shtParam, lDataStartRow, 1, lMaxRow, fGetValidMaxCol(shtParam))
            .ClearContents
            '.ClearFormats
            .ClearComments
            .ClearNotes
            .ClearOutline
        End With
    End If
End Function
