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

Function fFindInWorksheet(rngToFindIn As Excel.Range, sWhatToFind As String _
                    , Optional abNotFoundThenError As Boolean = True _
                    , Optional abAllowMultiple As Boolean = False) As Range
    If Len(Trim(sWhatToFind)) <= 0 Then fMsgRaiseErr "Wrong param sWhatToFind to fFindInWorksheet " & sWhatToFind
    
    Dim rngOut  As Range
    Dim rngFound As Range
    Dim lFoundCnt As Long
    Dim sFirstAddress As String
    
    Set rngFound = rngToFindIn.Find(What:=sWhatToFind _
                                    , after:=rngToFindIn.Cells(rngToFindIn.Rows.Count, rngToFindIn.Columns.Count) _
                                    , LookIn:=xlValues _
                                    , Lookat:=xlWhole _
                                    , SearchOrder:=xlByRows _
                                    , SearchDirection:=xlNext _
                                    , MatchCase:=False _
                                    , MatchByte:=False)
    Set rngOut = rngFound
    
    If rngFound Is Nothing Then
        If abNotFoundThenError Then
            fMsgRaiseErr """" & sWhatToFind & """ cannot be found in sheet " & rngToFindIn.Parent.Name & "[" & rngToFindIn.Address & "], pls check your program."
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
                                            , Lookat:=xlWhole _
                                            , SearchOrder:=xlByRows _
                                            , SearchDirection:=xlNext _
                                            , MatchCase:=False _
                                            , MatchByte:=False)
                If rngFound Is Nothing Then Exit Do
                If rngFound.Address = sFirstAddress Then Exit Do
                
                lFoundCnt = lFoundCnt + 1
            Loop
            
            If lFoundCnt > 1 Then
                fMsgRaiseErr lFoundCnt & " copies of """ & sWhatToFind & """ were found in sheet " & rngToFindIn.Parent.Name & ", pls check your program."
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
    Dim arrOUt()
    
    If fRangeIsSingleCell(rngParam) Then
        ReDim arrOUt(1 To 1, 1 To 1)
        arrOUt(1, 1) = rngParam.Value
    Else
        arrOUt = rngParam.Value
    End If
    
    fReadRangeDataToArray = arrOUt
    Erase arrOUt
End Function

Function fSetSpecifiedConfigCellAddress(shtConfig As Worksheet, asTag As String, asRtnCol As String, asCriteria As String _
                                , sValue As String _
                                , Optional bAllowMultiple As Boolean = False _
                                )
    Dim sAddr As String
    sAddr = fGetSpecifiedConfigCellAddress(shtConfig, asTag, asRtnCol, asCriteria, False)
    shtConfig.Range(sAddr).Value = sValue
End Function
Function fGetSpecifiedConfigCellAddress(shtConfig As Worksheet, asTag As String, asRtnCol As String _
                                , asCriteria As String _
                                , Optional bAllowMultiple As Boolean = False _
                                )
'                                , Optional bExternalAddress As Boolean = True _
'                                , Optional bNoMatachedPromptError As Boolean = True
    'asCriteria: colA=Value01, colB=Value02
    Dim arrColNames()
    Dim arrColValues()
    Dim iRtnColIndex As Integer
    Call fSplitDataCriteria(asCriteria, arrColNames, arrColValues)

    ReDim Preserve arrColNames(LBound(arrColNames) To UBound(arrColNames) + 1)
    arrColNames(UBound(arrColNames)) = asRtnCol
    
    iRtnColIndex = UBound(arrColNames)
    
    Dim lConfigStartRow As Long _
                                , lConfigStartCol As Long _
                                , lConfigEndRow As Long _
                                , lOutConfigHeaderAtRow As Long _
                                , abNoDataConfigThenError As Boolean _
                                , bNetValues As Boolean
    Dim arrConfigData()
    Dim arrColsIndex()

    Call fReadConfigBlockToArray(asTag:=asTag, rngToFindIn:=shtConfig.Cells, arrColsName:=arrColNames _
                                , arrConfigData:=arrConfigData _
                                , arrColsIndex:=arrColsIndex _
                                , lConfigStartRow:=lConfigStartRow _
                                , lConfigStartCol:=lConfigStartCol _
                                , lConfigEndRow:=lConfigEndRow _
                                , lOutConfigHeaderAtRow:=lOutConfigHeaderAtRow _
                                , abNoDataConfigThenError:=abNoDataConfigThenError _
                                )
    
    Dim lMatchRow As Long
    Dim sErr As String
    lMatchRow = fFindMatchDataInArrayWithCriteria(arrConfigData, arrColsIndex, arrColValues, bAllowMultiple, sErr)
    
    If lMatchRow < 0 Then
        fMsgRaiseErr sErr & " with criteria " & vbCr & asCriteria
    End If
    
    fGetSpecifiedConfigCellAddress = shtConfig.Cells(lOutConfigHeaderAtRow + lMatchRow, lConfigStartCol + arrColsIndex(iRtnColIndex) - 1).Address(external:=True)
End Function

Function fFindMatchDataInArrayWithCriteria(arr(), arrColsIndex(), arrColValues() _
                                        , bAllowMultiple As Boolean _
                                        , ByRef asErrmsg As String _
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
    
    asErrmsg = ""
    lOut = -1
    lMatchCnt = 0
    For lEachRow = LBound(arr, 1) To UBound(arr, 1)
        If fArrayRowIsBlankHasNoData(arr, lEachRow) Then GoTo next_row
        
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
        fMsgRaiseErr "Wrong array to paste to sheet: fGetArrayDimension(arrData) <> 2"
    End If
    
    sht.Cells(lStartRow, lStartCol).Resize(UBound(arrData, 1), UBound(arrData, 2)).Value = arrData
End Function

Function fAppendArray2Sheet(sht As Worksheet, arrData, Optional lStartCol As Long = 1)
    If fArrayIsEmptyOrNoData(arrData) Then Exit Function
    
    If fGetArrayDimension(arrData) <> 2 Then
        fMsgRaiseErr "Wrong array to paste to sheet: fGetArrayDimension(arrData) <> 2"
    End If
    
    Dim lFromRow As Long
    lFromRow = fGetValidMaxRow(sht) + 1
    
    sht.Cells(lFromRow, lStartCol).Resize(UBound(arrData, 1), UBound(arrData, 2)).Value = arrData
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

Function fSortDataInSheetSortSheetData(sht As Worksheet, arrColsOrSingleCol, Optional arrAscendDescend)

    

End Function
 
