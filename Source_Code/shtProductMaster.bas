VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtProductMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Option Base 1

Dim dictProductNameMaster As Dictionary

Private Sub btnValidateProductMaster_Click()
    On Error GoTo exit_sub
    
    Dim arrData()
    Dim dictColIndex As Dictionary
    
    fInitialization
    gsRptID = "REPLACE_UNIFY_SALES_INFO"
    Call fReadSysConfig_InputTxtSheetFile
    
    Call fReadSheetDataByConfig("PRODUCT_MASTER", dictColIndex, arrData, , , , , shtProductMaster)
    
    Call fValidateDuplicateInArray(arrData, Array(dictColIndex("ProductProducer") _
                                                , dictColIndex("ProductName") _
                                                , dictColIndex("ProductSeries")) _
                , False, shtProductProducerMaster, 1, 1, "����+����+���")
    Call fValidateBlankInArray(arrData, dictColIndex("ProductProducer"), shtProductMaster, 1, 1, "ҩƷ����")
    Call fValidateBlankInArray(arrData, dictColIndex("ProductName"), shtProductMaster, 1, 1, "ҩƷ����")
    Call fValidateBlankInArray(arrData, dictColIndex("ProductSeries"), shtProductMaster, 1, 1, "ҩƷ���")
    Call fValidateBlankInArray(arrData, dictColIndex("ProductUnit"), shtProductMaster, 1, 1, "ҩƷ��λ")
    
    Dim lEachRow As Long
    For lEachRow = LBound(arrData, 1) To UBound(arrData, 1)
        If Not fProductNameExistsInProductNameMaster(CStr(arrData(lEachRow, dictColIndex("ProductProducer"))) _
                                            , CStr(arrData(lEachRow, dictColIndex("ProductName")))) Then
            fErr "��ҩƷ���� + ҩƷ���ơ���������ҩƷ���������С�" & vbCr & "�кţ�" & lEachRow + 1
        End If
    Next
    
    fMsgBox "û�з��ִ���", vbInformation
exit_sub:
    Set dictColIndex = Nothing
    fEnableExcelOptionsAll
    End
End Sub

Function fReadSheetProductNameMaster2Dictionary()
    Dim arrData()
    Dim dictColIndex As Dictionary
    
    Call fReadSheetDataByConfig("PRODUCT_NAME_MASTER", dictColIndex, arrData, , , , , shtProductNameMaster)
    Set dictProductNameMaster = fReadArray2DictionaryMultipleKeysWithKeysOnly(arrData _
                                    , Array(dictColIndex("ProductProducer"), dictColIndex("ProductName")) _
                                    , DELIMITER)
    Set dictColIndex = Nothing
End Function
Function fProductNameExistsInProductNameMaster(sProductProducer As String, sProductName As String) As Boolean
    If dictProductNameMaster Is Nothing Then Call fReadSheetProductNameMaster2Dictionary
    
    fProductNameExistsInProductNameMaster = dictProductNameMaster.Exists(sProductProducer & DELIMITER & sProductName)
End Function

Private Sub Worksheet_Change(ByVal Target As Range)
'    Dim sProducerCol As String
'    Dim rgIntersect As Range
'
'    sProducerCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - HOSPITAL_MASTER]", "Column Index", "Column Tech Name=ProductProducer")
'
'    Set rgIntersect = Intersect(Target, Me.Columns(sProducerCol))
'
'    If Not rgIntersect Is Nothing Then
'        If rgIntersect.Areas.Count > 1 Then fErr "Please select only one cell."
'
'        Dim lProducerColMaxRow As Long
'        lProducerColMaxRow = Me.Range(sProducerCol & 1).End(xlDown).Row
'        rngAddr = Me.Range(sProducerCol & 2 & ":" & sProducerCol & lProducerColMaxRow).Address(external:=True)
'        rngAddr = "=" & rngAddr
'
''        Call fSetValidationListForshtProductProducerReplace_Producer(rngAddr)
''
''        sOthersCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File -PRODUCER_REPLACE_SHEET]", "Column Index" _
''                    , "Column Tech Name=ToProducer")
''        lOthersColMaxRow = shtProductProducerReplace.Columns(sOthersCol).End(xlDown).Row + 1000
''        Call fSetValidationListForRange(shtProductProducerReplace.Range(sOthersCol & 2 & ":" & sOthersCol & lOthersColMaxRow), rngAddr)
''        'Call fSetValidationListForRange(fGetRangeByStartEndPos(shtProductProducerReplace, 2, 1, 1000, 1), rngAddr)
'    End If
'
'    Set rgIntersect = Nothing
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
