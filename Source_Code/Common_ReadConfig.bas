Attribute VB_Name = "Common_ReadConfig"
Option Explicit
Option Base 1

Enum InputFile
    [_first] = 1
    ReportID = 1
    FileTag = 2
    FilePath = 3
    Source = 4
    ReLoadOrNot = 5
    FileSpecTag = 6
    Env = 7
    DefaultSheet = 8
    PivotTableTag = 9
    RowNo = 10
    [_last] = 10
End Enum

Function fFindAllColumnsIndexByColNames(rngToFindIn As Range, arrColsName, ByRef arrColsIndex() _
                                , ByRef alHeaderAtRow As Long, Optional bReturnLetter As Boolean = False)
    If fArrayIsEmptyOrNoData(arrColsName) Then fErr "arrColsName is empty."
    If fArrayHasBlankValue(arrColsName) Then fErr "arrColsName has blank element." & vbCr & Join(arrColsName, vbCr)
    If fArrayHasDuplicateElement(arrColsName) Then fErr "arrColsName has duplicate element."
    
    ReDim arrColsIndex(LBound(arrColsName) To UBound(arrColsName))
    
    Dim lColAtRow As Long
    Dim lEachCol As Long
    Dim sEachColName As String
    Dim rngFound As Range
    
    lColAtRow = 0
    For lEachCol = LBound(arrColsName) To UBound(arrColsName)
        sEachColName = Trim(arrColsName(lEachCol))
        sEachColName = Replace(sEachColName, "*", "~*")
        
        Set rngFound = fFindInWorksheet(rngToFindIn, sEachColName)
        
        If lColAtRow <> 0 Then
            If lColAtRow <> rngFound.Row Then
                fErr "Columns are not at the same row."
            End If
        Else
            lColAtRow = rngFound.Row
        End If
        
        If bReturnLetter Then
            arrColsIndex(lEachCol) = fNum2Letter(rngFound.Column)
        Else
            arrColsIndex(lEachCol) = rngFound.Column
        End If
    Next
    
    alHeaderAtRow = lColAtRow
    Set rngFound = Nothing
End Function
'
'Function fValidateDuplicateKeys(arrConfigData(), arrColsIndex(), arrKeyCols, lHeaderAtRow As Long, lStartCol As Long)
'    If fArrayIsEmptyOrNoData(arrKeyCols) Then Exit Function
'
'    Dim lEachRow As Long
'    Dim lEachCol As Long
'    Dim i As Long
'    Dim sKeyStr As String
'    Dim dict As New Dictionary
'
'    For lEachRow = LBound(arrConfigData, 1) To UBound(arrConfigData, 1)
'        sKeyStr = ""
'
'        For i = LBound(arrKeyCols) To UBound(arrKeyCols)
'            'lEachCol = arrColsIndex(arrKeyCols(i) - 1)
'            lEachCol = arrColsIndex(arrKeyCols(i))
'            sKeyStr = sKeyStr & Trim(CStr(arrConfigData(lEachRow, lEachCol)))
'        Next
'
'        If dict.Exists(sKeyStr) Then
'            fErr "Duplicate key " & sKeyStr & " was found " & vbCr & "at row: " & (lHeaderAtRow + lEachRow) _
'                     & ", column: " & fNum2Letter((lStartCol + lEachCol))
'        Else
'            dict.Add sKeyStr, 0
'        End If
'    Next
'
'    Set dict = Nothing
'End Function

Function fValidateDuplicateKeysForConfigBlock(arrConfigData(), arrColsIndex(), arrKeyCols, lHeaderAtRow As Long, lStartCol As Long)
    If fArrayIsEmptyOrNoData(arrKeyCols) Then Exit Function
    
    Dim lEachRow As Long
    Dim lEachCol As Long
    Dim i As Long
    Dim sKeyStr As String
    Dim dict As New Dictionary
    
    For lEachRow = LBound(arrConfigData, 1) To UBound(arrConfigData, 1)
        sKeyStr = ""
        
        For i = LBound(arrKeyCols) To UBound(arrKeyCols)
            'lEachCol = arrColsIndex(arrKeyCols(i) - 1)
            lEachCol = arrColsIndex(arrKeyCols(i))
            sKeyStr = sKeyStr & Trim(CStr(arrConfigData(lEachRow, lEachCol)))
        Next
        
        If dict.Exists(sKeyStr) Then
            fErr "Duplicate key " & sKeyStr & " was found " & vbCr & "at row: " & (lHeaderAtRow + lEachRow) _
                     & ", column: " & fNum2Letter((lStartCol + lEachCol))
        Else
            dict.Add sKeyStr, 0
        End If
    Next
    
    Set dict = Nothing
End Function

Function fReadConfigBlockToArrayValidated(asTag As String, rngToFindIn As Range, arrColsName _
                                , Optional arrKeyCols _
                                , Optional ByRef lConfigStartRow As Long _
                                , Optional ByRef lConfigStartCol As Long _
                                , Optional ByRef lConfigEndRow As Long _
                                , Optional ByRef lOutConfigHeaderAtRow As Long _
                                , Optional abNoDataConfigThenError As Boolean = False _
                                , Optional bNetValues As Boolean = True) As Variant
    'arrKeyCols:  array(1, 2, 3, 5), or unnecessary: array()
    Dim arrConfigData()
    Dim arrColsIndex()
    Dim arrOUt()

    Call fReadConfigBlockToArray(asTag:=asTag, rngToFindIn:=rngToFindIn, arrColsName:=arrColsName _
                                , arrConfigData:=arrConfigData _
                                , arrColsIndex:=arrColsIndex _
                                , lConfigStartRow:=lConfigStartRow _
                                , lConfigStartCol:=lConfigStartCol _
                                , lConfigEndRow:=lConfigEndRow _
                                , lOutConfigHeaderAtRow:=lOutConfigHeaderAtRow _
                                , abNoDataConfigThenError:=abNoDataConfigThenError _
                                )
    
    If fArrayIsEmptyOrNoData(arrConfigData) Then GoTo exit_fun
    
    'Call fValidateDuplicateKeys(arrConfigData, arrColsIndex, arrKeyCols, lOutConfigHeaderAtRow, lConfigStartCol)
    
    If bNetValues Then
        ReDim arrOUt(LBound(arrConfigData, 1) To UBound(arrConfigData, 1), 1 To UBound(arrColsIndex) - LBound(arrColsIndex) + 1)
        
        Dim lEachRow As Long
        Dim lEachCol As Long
        Dim i As Long
        
        For lEachRow = LBound(arrConfigData, 1) To UBound(arrConfigData, 1)
            'i = LBound(arrColsIndex) + 1
            i = LBound(arrColsIndex)
            For lEachCol = LBound(arrColsIndex) To UBound(arrColsIndex)
                arrOUt(lEachRow, i) = arrConfigData(lEachRow, arrColsIndex(lEachCol))
                i = i + 1
            Next
        Next
    End If
exit_fun:
    Erase arrColsIndex
    
    If bNetValues Then
        fReadConfigBlockToArrayValidated = arrOUt
    Else
        fReadConfigBlockToArrayValidated = arrConfigData
    End If
    
    Erase arrConfigData
    Erase arrOUt
End Function
Function fReadConfigBlockToArrayNet(asTag As String, rngToFindIn As Range, arrColsName() _
                                , Optional ByRef lConfigStartRow As Long _
                                , Optional ByRef lConfigStartCol As Long _
                                , Optional ByRef lConfigEndRow As Long _
                                , Optional ByRef lOutConfigHeaderAtRow As Long _
                                , Optional abNoDataConfigThenError As Boolean = False _
                                ) As Variant
    Dim arrOUt()
    Dim arrColsIndex()
    Dim arrConfigData()

    Call fReadConfigBlockToArray(asTag:=asTag, rngToFindIn:=rngToFindIn, arrColsName:=arrColsName _
                                , arrConfigData:=arrConfigData _
                                , arrColsIndex:=arrColsIndex _
                                , lConfigStartRow:=lConfigStartRow _
                                , lConfigStartCol:=lConfigStartCol _
                                , lConfigEndRow:=lConfigEndRow _
                                , lOutConfigHeaderAtRow:=lOutConfigHeaderAtRow _
                                , abNoDataConfigThenError:=abNoDataConfigThenError _
                                )
    If fArrayIsEmptyOrNoData(arrConfigData) Then GoTo exit_fun
    
    ReDim arrOUt(LBound(arrConfigData, 1) To UBound(arrConfigData, 1), LBound(arrColsIndex) To UBound(arrColsIndex))
    
    Dim lEachRow As Long
    Dim lEachCol As Long
    Dim i As Long
    
    For lEachRow = LBound(arrConfigData, 1) To UBound(arrConfigData, 1)
        i = LBound(arrColsIndex)
        For lEachCol = LBound(arrColsIndex) To UBound(arrColsIndex)
            arrOUt(lEachRow, i) = arrConfigData(lEachRow, arrColsIndex(lEachCol))
            i = i + 1
        Next
    Next
exit_fun:
    Erase arrColsIndex
    Erase arrConfigData
    fReadConfigBlockToArrayNet = arrOUt
    Erase arrOUt
End Function
Function fReadConfigBlockToArray(asTag As String, rngToFindIn As Range, arrColsName _
                                , ByRef arrConfigData() _
                                , ByRef arrColsIndex() _
                                , Optional ByRef lConfigStartRow As Long _
                                , Optional ByRef lConfigStartCol As Long _
                                , Optional ByRef lConfigEndRow As Long _
                                , Optional ByRef lOutConfigHeaderAtRow As Long _
                                , Optional abNoDataConfigThenError As Boolean = False _
                                )
    arrConfigData = Array()
    
    Dim shtConfig As Worksheet
    Set shtConfig = rngToFindIn.Parent
    
    Call fReadConfigBlockStartEnd(asTag, rngToFindIn, lConfigStartRow, lConfigStartCol, lConfigEndRow)
    
    If lConfigEndRow < lConfigStartRow + 1 Then
        If abNoDataConfigThenError Then
            fErr "No data is configured under tag " & asTag & " in sheet " & shtConfig.Name & vbCr _
                    & "You must leave at least one blank line after the tag."
        End If
    End If
    
    Set rngToFindIn = fGetRangeByStartEndPos(shtConfig, lConfigStartRow, lConfigStartCol, lConfigEndRow, Columns.Count)
    Call fFindAllColumnsIndexByColNames(rngToFindIn, arrColsName, arrColsIndex, lOutConfigHeaderAtRow)
    
    Dim lColsMinCol As Long
    Dim lColsMaxCol As Long
    
    lColsMinCol = Application.WorksheetFunction.Min(arrColsIndex)
    lColsMaxCol = Application.WorksheetFunction.Max(arrColsIndex)
    
    lConfigEndRow = fGetValidMaxRowOfRange(fGetRangeByStartEndPos(shtConfig, lConfigStartRow, lConfigStartCol, lConfigEndRow, lColsMaxCol))
    
    If lConfigEndRow > lOutConfigHeaderAtRow Then
        arrConfigData = fReadRangeDatatoArrayByStartEndPos(shtConfig, lOutConfigHeaderAtRow + 1, lColsMinCol, lConfigEndRow, lColsMaxCol)
    End If
    
    lConfigStartCol = lColsMinCol
    
    Dim lEachCol As Long
    'change 10, 15, 20, to 1, 6, 11
    For lEachCol = UBound(arrColsIndex) To LBound(arrColsIndex) Step -1
        arrColsIndex(lEachCol) = arrColsIndex(lEachCol) - lColsMinCol + 1
    Next
    
    Set shtConfig = Nothing
End Function

Function fReadConfigBlockStartEnd(asTag As String, rngToFindIn As Range _
                                , ByRef lOutBlockStartRow As Long _
                                , ByRef lOutBlockStartCol As Long _
                                , ByRef lOutBlockEndRow As Long)
    
    Dim shtSource As Worksheet
    Dim lMaxRow As Long
    Dim rngTagFound As Range
    Dim lTagRow As Long
    Dim lTagCol As Long
    
    Set shtSource = rngToFindIn.Parent
    lMaxRow = fGetValidMaxRow(shtSource)
    
    Set rngTagFound = fFindInWorksheet(rngToFindIn, asTag)
    lTagRow = rngTagFound.Row
    lTagCol = rngTagFound.Column
    
    Set rngTagFound = fFindInWorksheet(fGetRangeByStartEndPos(shtSource, lTagRow + 1, lTagCol, lMaxRow, lTagCol) _
                                    , "[*]", False, True)
    If rngTagFound Is Nothing Then
        lOutBlockEndRow = lMaxRow
    Else
        lOutBlockEndRow = rngTagFound.Row - 1
    End If
    
    lOutBlockStartRow = lTagRow + 1
    lOutBlockStartCol = lTagCol
    
    Set shtSource = Nothing
    Set rngTagFound = Nothing
End Function

Function fReadConfigInputFiles(Optional asReportID As String = "")
    If asReportID = "" Then asReportID = gsRptID
    
    Dim asTag As String
    Dim arrColsName()
    Dim rngToFindIn As Range
    Dim arrConfigData()
    Dim lConfigStartRow As Long
    Dim lConfigStartCol As Long
    Dim lConfigEndRow As Long
    Dim lConfigHeaderAtRow As Long
                                
    asTag = "[Input Files]"
    ReDim arrColsName(InputFile.ReportID To InputFile.PivotTableTag)
    
    arrColsName(InputFile.ReportID) = "Report ID"
    arrColsName(InputFile.FileTag) = "File Tag"
    arrColsName(InputFile.FilePath) = "File Full Path"
    arrColsName(InputFile.Source) = "Source"
    arrColsName(InputFile.ReLoadOrNot) = "When Data Already Loaded To Sheet"
    arrColsName(InputFile.FileSpecTag) = "File Spec Tag"
    arrColsName(InputFile.Env) = "DEV/UAT/PROD"
    arrColsName(InputFile.DefaultSheet) = "Which Sheet To Import"
    arrColsName(InputFile.PivotTableTag) = "Pivot Table Tag To Be Created From This Data Source"
     
    arrConfigData = fReadConfigBlockToArrayNet(asTag:=asTag, rngToFindIn:=shtSysConf.Cells _
                                , arrColsName:=arrColsName _
                                , lConfigStartRow:=lConfigStartRow _
                                , lConfigStartCol:=lConfigStartCol _
                                , lConfigEndRow:=lConfigEndRow _
                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
                                , abNoDataConfigThenError:=True _
                                )
    Erase arrColsName
    
    Call fValidateDuplicateInArray(arrConfigData, Array(InputFile.ReportID, InputFile.FileTag), False _
        , shtSysConf, lConfigHeaderAtRow, lConfigStartCol, "Report ID + File Tag")
    Call fValidateBlankInArray(arrConfigData, InputFile.ReportID, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, "Report ID")
    Call fValidateBlankInArray(arrConfigData, InputFile.FileTag, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, "File Tag")
    
    Set gDictInputFiles = New Dictionary
    
    Dim lEachRow As Long
    Dim lActualRow As Long
    Dim sRptNameStr As String
    
    For lEachRow = LBound(arrConfigData, 1) To UBound(arrConfigData, 1)
        If fArrayRowIsBlankHasNoData(arrConfigData, lEachRow) Then GoTo next_row
        
        sRptNameStr = DELIMITER & arrConfigData(lEachRow, 1) & DELIMITER
        If InStr(sRptNameStr, DELIMITER & asReportID & DELIMITER) <= 0 Then GoTo next_row
        
        lActualRow = lConfigHeaderAtRow + lEachRow
        
        
next_row:
    Next
    
    Erase arrConfigData
End Function

Function fReadSysConfig_InputOutputTxt(Optional asReportID As String = "")
    If asReportID = "" Then asReportID = gsRptID
    Call fReadConfigInputFiles(sReportID)
End Function

Function fReadConfigWholeColsToDictionary(shtConfig As Worksheet, asTag As String, asKeyNotNullCol As String, asRtnCol As String) As Dictionary
    If fZero(asTag) Or fZero(asKeyNotNullCol) Or fZero(asRtnCol) Then fErr "Wrong param"

    Dim bRtnColIsKeyCol As Boolean
    bRtnColIsKeyCol = (Trim(asKeyNotNullCol) = Trim(asRtnCol))

    Dim arrColNames()
    ReDim arrColNames(0 To 1)
    arrColNames(0) = asKeyNotNullCol
    arrColNames(1) = asRtnCol

    Dim rngToFindIn As Range
    Dim arrConfigData()
    Dim lConfigStartRow As Long
    Dim lConfigStartCol As Long
    Dim lConfigEndRow As Long
    Dim lConfigHeaderAtRow As Long

    arrKeyColsForValidation = Array(1, 2)

    arrConfigData = fReadConfigBlockToArray(asTag:=asTag, rngToFindIn:=shtConfig.Cells _
                                , arrColsName:=arrColsName _
                                , lConfigStartRow:=lConfigStartRow _
                                , lConfigStartCol:=lConfigStartCol _
                                , lConfigEndRow:=lConfigEndRow _
                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
                                , abNoDataConfigThenError:=True _
                                )
End Function

Function fReadConfigWholeColsToArray(shtConfig As Worksheet, asTag As String, arrFetchCols, Optional arrKeyColsForValidation) As Variant
'arrKeyColsForValidation : Array(1, 2, 5)
    Dim rngToFindIn As Range
    Dim arrConfigData()
    Dim lConfigStartRow As Long
    Dim lConfigStartCol As Long
    Dim lConfigEndRow As Long
    Dim lConfigHeaderAtRow As Long
     
    arrConfigData = fReadConfigBlockToArrayValidated(asTag:=asTag, rngToFindIn:=shtConfig.Cells _
                                , arrColsName:=arrFetchCols _
                                , lConfigStartRow:=lConfigStartRow _
                                , lConfigStartCol:=lConfigStartCol _
                                , lConfigEndRow:=lConfigEndRow _
                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
                                , abNoDataConfigThenError:=True _
                                , arrKeyCols:=arrKeyColsForValidation _
                                , bNetValues:=True _
                                )
    fReadConfigWholeColsToArray = arrConfigData
    Erase arrConfigData
End Function


Function fOverWriteGDictWithUserInputOnMenuScreen()
    Call fOverWriteGDictVariables_gDictInputfiles
End Function
