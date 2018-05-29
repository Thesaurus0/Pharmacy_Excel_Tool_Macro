Attribute VB_Name = "MZ_UT"
Option Explicit
Option Base 1

Sub AllUnitTest()
'    Dim asTag As String, rngToFindIn As Range _
'                                , arrConfigData() _
'                                , lConfigStartRow As Long _
'                                , lConfigStartCol As Long _
'                                , lConfigEndRow As Long _
'                                , lOutBlockHeaderAtRow As Long
'    Dim arrColsName()
'    Dim arrColsIndex()
'    Dim lConfigHeaderAtRow As Long
'
'    asTag = "[Input Files]"
'    arrColsName = Array("xxa", "Company ID", "Company Name")
'
'    Set rngToFindIn = ActiveSheet.Cells
'
'Call fReadConfigBlockToArray(asTag:=asTag, rngToFindIn:=activeshet.Cells, arrColsName:=arrColsName _
'                                , arrConfigData:=arrConfigData _
'                                , arrColsIndex:=arrColsIndex _
'                                , lConfigStartRow:=lConfigStartRow _
'                                , lConfigStartCol:=lConfigStartCol _
'                                , lConfigEndRow:=lConfigEndRow _
'                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
'                                , abNoDataConfigThenError:=True _
'                                )
 
'Call fReadConfigBlockToArray(asTag:=asTag, rngToFindIn:=ActiveSheet.Cells, arrColsName:=arrColsName _
'                                , arrConfigData:=arrConfigData _
'                                , arrColsIndex:=arrColsIndex _
'                                , lConfigStartRow:=lConfigStartRow _
'                                , lConfigStartCol:=lConfigStartCol _
'                                , lConfigEndRow:=lConfigEndRow _
'                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
'                                , abNoDataConfigThenError:=True _
'                                )
                       
'arrConfigData = fReadConfigBlockToArrayNet(asTag:=asTag, rngToFindIn:=ActiveSheet.Cells, arrColsName:=arrColsName _
'                                , lConfigStartRow:=lConfigStartRow _
'                                , lConfigStartCol:=lConfigStartCol _
'                                , lConfigEndRow:=lConfigEndRow _
'                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
'                                , abNoDataConfigThenError:=True _
'                                )
'arrConfigData = fReadConfigBlockToArrayValidated(asTag:=asTag, rngToFindIn:=rngToFindIn, arrColsName:=arrColsName _
'                                , lConfigStartRow:=lConfigStartRow _
'                                , lConfigStartCol:=lConfigStartCol _
'                                , lConfigEndRow:=lConfigEndRow _
'                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
'                                , abNoDataConfigThenError:=True _
'                                , arrKeyCols:=Array(2) _
'                                , bNetValues:=False _
'                                )
'    Dim arr()
'
'    'Debug.Print UBound(arr) & "-" & LBound(arr)
'   ' Debug.Print fArrayIsEmpty(arr)
'    'Debug.Print fGetArrayDimension(arr)
'    Dim a
'    Set a = ActiveCell.MergeArea
'
'    'Dim a
'    Set a = Selection
'
'    Debug.Print fGetValidMaxRowOfRange(Selection, True)
'
'    Dim bbb()
'    'bbb = fReadRangeDataToArray(Selection)

    'Debug.Print fGetSpecifiedConfigCellAddress(shtSysConf, "[Input Files]", "File Full Path", "Company ID = PW")
    'Debug.Print fGenRandomUniqueString
    'Debug.Assert fTrim(vbLf & " abcd " & vbCr) = "abcd"
    'Debug.Print fJoin(Selection.Value)
    
'    Dim arr
'    arr = fReadConfigWholeColsToArray(shtSysConf, "[Sales Company List]", Array("Company ID", "Company Name"), Array(1))
    
    'Call fReadConfigInputFiles
    
    'Call ThisWorkbook.fReadConfigGetAllCommandBars
'    Dim sAddr As String
'    sAddr = Range("A12:Z34").Address(ReferenceStyle:=xlR1C1, external:=True)
'    sAddr = Range("A12:Z34").Address(external:=True)
''    Debug.Print sAddr
''    Debug.Print fReplaceConvertR1C1ToA1(sAddr)
'
'    Dim rng As Range
'    Set rng = fGetRangeFromExternalAddress(sAddr)
'    Debug.Print rng.Address
    
    'Debug.Print fGetFileExtension("abce\ef\a\c\aaa.txt")
    
   ' Call fConvertFomulaToValueForSheetIfAny(ActiveSheet)
   fDeleteAllConditionFormatFromSheet ActiveSheet
   Call fSetConditionFormatForOddEvenLine(ActiveSheet, , , , Array(1), True)
   Call fSetConditionFormatForBorders(ActiveSheet, , , , Array(1), True)
End Sub

Sub testa()
'    Debug.Print Asc(" ")
'    Debug.Print Asc(vbCr)
'    Debug.Print Asc(vbLf)
'    Debug.Print Asc(vbCrLf)
'    Debug.Print Asc(vbNewLine)
'    Debug.Print Asc(vbTab)
    Dim aa
    aa = ActiveSheet.Range("c10:f20")
    
'    Dim bb(2, 4)
'    bb(0, 0) = "a"
'
'    Dim cc()
'    cc = Array()
'    Debug.Print LBound(aa, 1) & " - " & UBound(aa, 1)
'    Debug.Print LBound(aa, 2) & " - " & UBound(aa, 2)
'    Debug.Print LBound(bb, 1) & " - " & UBound(bb, 1)
'    Debug.Print LBound(bb, 2) & " - " & UBound(bb, 2)
'    Debug.Print LBound(cc, 1) & " - " & UBound(cc, 1)
'    Debug.Print LBound(cc, 2) & " - " & UBound(cc, 2)
    
'    Const DELI = " " & DELIMITER & " "
'    Dim f
'    'f = aa(0)
'    'Debug.Print Join(aa(3), "")
'    Dim s As String
'    Debug.Print fArrayIsEmptyOrNoData(s)
'    Dim a As String
'
'    a = "c"
'    'Debug.Print Switch(a = "a", 1, a = "b", 2, a = "c", 3, a = "e", 4)
'    Debug.Print Switch("c", 3, "e", 4)
    
'    Dim a
'    a = "[xxx]"
'    Debug.Print (a Like "[*]")

    Dim arr(1000)
    
    Dim i As Long
    
    For i = 1 To 1000
        arr(i) = Rnd() * 1000
    Next
    
    Call fSortArrayQuickSortDesc(arr)
    Call fSortArrayQuickSort(arr)
End Sub

Sub aaaa()
'    Dim a
'    'a = shtSalesRawDataRpt
'
'       'shtSalesRawDataRpt.Close
'
''    Call fHideAllSheetExcept("1", "2", "6", "24")
''    Dim rngAddr As String
''    rngAddr = fGetRangeByStartEndPos(shtProductMaster, 2, 1, 800, 1).Address(external:=True)
''    rngAddr = "=" & rngAddr
''    Call fSetValidationListForRange(fGetRangeByStartEndPos(shtProductProducerReplace, 2, 1, 1000, 1), rngAddr)
'
''    Dim sht ' As Worksheet
''     sht = Evaluate("shtSelfSalesCal")
''
''     fKeepCopyContent
''     Application.CutCopyMode = 0
''     fCopyFromKept
'    'End
'    fGetFSO
'    Debug.Print fCheckPath("F:\Github_Local_Repository\\Pharmacy_Excel_Tool_Macro\历史表\a.txt")
'    'a = Dir("F:\Github_Local_Repository\Pharmacy_Excel_Tool_Macro\历史表\a.txt", )
'   Debug.Print a
    End
End Sub
