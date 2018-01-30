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

Private Sub btnValidateProductMaster_Click()
    Call fValidateSheet
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Call fResetdictProductMaster
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
    On Error GoTo Exit_Sub
    Application.ScreenUpdating = False
    
    Const ProducerCol = 1
    Const ProductNameCol = 2
    Const ProductSeriesCol = 4
    
    Dim rgIntersect As Range
    Dim sProducer As String
    Dim sProductName As String
    Dim sValidationListAddr As String
        
    'product name
    Set rgIntersect = Intersect(Target, Me.Columns(ProductNameCol))
    
    If Not rgIntersect Is Nothing Then
        If rgIntersect.Areas.Count > 1 Then GoTo Exit_Sub   'fErr "不能选多个"
        If rgIntersect.Rows.Count <> 1 Then GoTo Exit_Sub

        sProducer = rgIntersect.Offset(0, ProducerCol - ProductNameCol).Value
        
        If fNzero(sProducer) Then
            Call fSetFilterForSheet(shtProductNameMaster, 1, sProducer)
            Call fCopyFilteredDataToRange(shtProductNameMaster, 2)
            
            sValidationListAddr = "=" & shtDataStage.Columns("A").Address(external:=True)
            'Call fSetValidationListForshtProductNameReplace_ProductName(sValidationListAddr, 3)
            Call fSetValidationListForRange(rgIntersect, sValidationListAddr)
        End If
    Else
    End If
    
Exit_Sub:
    fEnableExcelOptionsAll
    Application.ScreenUpdating = True
    
    If Err.Number <> 0 Then fMsgBox Err.Description
End Sub

Function fValidateSheet(Optional bErrMsgBox As Boolean = True) As Boolean
    On Error GoTo Exit_Sub
    
    Call fTrimAllCellsForSheet(Me)
    
    Dim arrData()
    Dim dictColIndex As Dictionary
    
    fInitialization
    gsRptID = "REPLACE_UNIFY_SALES_INFO"
    Call fReadSysConfig_InputTxtSheetFile
    
    Call fReadSheetDataByConfig("PRODUCT_MASTER", dictColIndex, arrData, , , , , Me)
    
    Call fValidateDuplicateInArray(arrData, Array(dictColIndex("ProductProducer") _
                                                , dictColIndex("ProductName") _
                                                , dictColIndex("ProductSeries")) _
                , False, Me, 1, 1, "厂家+名称+规格")
    Call fValidateBlankInArray(arrData, dictColIndex("ProductProducer"), Me, 1, 1, "药品厂家")
    Call fValidateBlankInArray(arrData, dictColIndex("ProductName"), Me, 1, 1, "药品名称")
    Call fValidateBlankInArray(arrData, dictColIndex("ProductSeries"), Me, 1, 1, "药品规格")
    Call fValidateBlankInArray(arrData, dictColIndex("ProductUnit"), Me, 1, 1, "药品单位")
    
    Dim lEachRow As Long
    For lEachRow = LBound(arrData, 1) To UBound(arrData, 1)
        If Not fProductNameExistsInProductNameMaster(CStr(arrData(lEachRow, dictColIndex("ProductProducer"))) _
                                            , CStr(arrData(lEachRow, dictColIndex("ProductName")))) Then
            fErr "【药品厂家 + 药品名称】不存在于药品名称主表中。" & vbCr & "行号：" & lEachRow + 1
        End If
    Next
    
    If bErrMsgBox Then fMsgBox "[" & Me.Name & "]表 没有发现错误", vbInformation
Exit_Sub:
    Set dictColIndex = Nothing
    fEnableExcelOptionsAll
    Erase arrData
    
    If Err.Number <> 0 Then
        fShowAndActiveSheet Me
        fValidateSheet = False
    Else
        fValidateSheet = True
    End If
End Function
