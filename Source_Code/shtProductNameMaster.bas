VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtProductNameMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub btnProductNameMasterValid_Click()
    On Error GoTo exit_sub
    
    Dim arrData()
    Dim dictColIndex As Dictionary
    
    fInitialization
    gsRptID = "REPLACE_UNIFY_SALES_INFO"
    Call fReadSysConfig_InputTxtSheetFile
    
    Call fReadSheetDataByConfig("PRODUCER_NAME_MASTER", dictColIndex, arrData, , , , , shtProductNameMaster)
    
    Call fValidateDuplicateInArray(arrData, Array(dictColIndex("ProductProducer"), dictColIndex("ProductName")) _
                    , False, shtProductNameMaster, 1, 1, "生产厂家+药品名称")
    
    Call fValidateBlankInArray(arrData, dictColIndex("ProductProducer"), shtProductNameMaster, 1, 1, "生产厂家")
    Call fValidateBlankInArray(arrData, dictColIndex("ProductName"), shtProductNameMaster, 1, 1, "药品名称")
    
    fMsgBox "没有发现错误", vbInformation
exit_sub:
    fEnableExcelOptionsAll
    Set dictColIndex = Nothing
End Sub

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
