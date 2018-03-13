VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtSelfPurchaseOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub btnValidate_Click()
    Call fValidateSheet
End Sub

Function fValidateSheet(Optional bErrMsgBox As Boolean = True) As Boolean
    On Error GoTo Exit_Sub
    
    Dim lErrRowNo As Long, lErrColNo As Long
    
    Call fTrimAllCellsForSheet(Me)
    
    Dim arrData()
    Dim dictColIndex As Dictionary
    
    fInitialization
    gsRptID = "CALCULATE_PROFIT"
    Call fReadSysConfig_InputTxtSheetFile
    
    Call fReadSheetDataByConfig("SELF_PURCHASE_ORDER", dictColIndex, arrData, , , , , Me)
    
'    Call fValidateDuplicateInArray(arrData, Array(dictColIndex("ProductProducer"), dictColIndex("ProductName")) _
'                    , False, me, 1, 1, "生产厂家+药品名称")
    
    Call fValidateBlankInArray(arrData, dictColIndex("ProductProducer"), Me, 1, 1, "生产厂家")
    Call fValidateBlankInArray(arrData, dictColIndex("ProductName"), Me, 1, 1, "药品名称")
    Call fValidateBlankInArray(arrData, dictColIndex("ProductSeries"), Me, 1, 1, "药品规格")
    Call fValidateBlankInArray(arrData, dictColIndex("ProductUnit"), Me, 1, 1, "药品单位")
    Call fValidateDateColInArray(arrData, dictColIndex("PurchaseDate"), Me, 1, 1, "销售出货日期")
    'Call fValidateBlankInArray(arrData, dictColIndex("PurchasePrice"), Me, 1, 1, "出货单价")
    'Call fValidateBlankInArray(arrData, dictColIndex("LotNum"), Me, 1, 1, "批号")
    
    Call fSortDataInSheetSortSheetData(Me, Array(dictColIndex("PurchaseDate") _
                                                , dictColIndex("ProductProducer") _
                                                , dictColIndex("ProductName"), dictColIndex("ProductUnit")))

    Call fCheckIfProductExistsInProductMaster(arrData, dictColIndex("ProductProducer"), dictColIndex("ProductName"), dictColIndex("ProductSeries"), lErrRowNo, lErrColNo)

    If bErrMsgBox Then fMsgBox "[" & Me.Name & "]表 没有发现错误", vbInformation
Exit_Sub:
    fEnableExcelOptionsAll
    Set dictColIndex = Nothing
    Erase arrData
    
    If fCheckIfGotBusinessError Then
        fValidateSheet = False
    Else
        If fCheckIfUnCapturedExceptionAbnormalError Then
            fValidateSheet = False
        Else
            fShowAndActiveSheet Me
            fValidateSheet = True
        End If
    End If
    
    If lErrRowNo > 0 Then
        fShowAndActiveSheet Me
        Application.GoTo Me.Cells(lErrRowNo, lErrColNo) ', True
    End If
'    If Err.Number <> 0 Then
'        fShowAndActiveSheet Me
'        fValidateSheet = False
'    Else
'        fValidateSheet = True
'    End If
End Function

Private Sub Worksheet_Change(ByVal Target As Range)
    fResetdictSelfPurchaseOD
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error GoTo Exit_Sub
    Application.ScreenUpdating = False
    
    Const ProducerCol = 1
    Const ProductNameCol = 2
    Const ProductSeriesCol = 3
    Const ProductUnitCol = 4
    
    Dim rgIntersect As Range
    Set rgIntersect = Intersect(Target, Me.Columns(ProductNameCol))
    
    If Not rgIntersect Is Nothing Then
        If rgIntersect.Areas.Count > 1 Then GoTo Exit_Sub    'fErr "不能选多个"
        If rgIntersect.Rows.Count <> 1 Then GoTo Exit_Sub
            
        Dim sProducer As String
        Dim sValidationListAddr As String
        
        sProducer = rgIntersect.Offset(0, ProducerCol - ProductNameCol).Value
        
        If fNzero(sProducer) Then
            Call fSetFilterForSheet(shtProductNameMaster, 1, sProducer)
            Call fCopyFilteredDataToRange(shtProductNameMaster, 2)
            
            sValidationListAddr = "=" & shtDataStage.Columns("A").Address(external:=True)
            'Call fSetValidationListForshtProductNameReplace_ProductName(sValidationListAddr, 3)
            Call fSetValidationListForRange(rgIntersect, sValidationListAddr)
        End If
    Else
        'product SeriesCol
        Set rgIntersect = Intersect(Target, Me.Columns(ProductSeriesCol))
        
        If Not rgIntersect Is Nothing Then
            If rgIntersect.Areas.Count > 1 Then GoTo Exit_Sub    'fErr "不能选多个"
            If rgIntersect.Rows.Count <> 1 Then GoTo Exit_Sub
            
            sProducer = rgIntersect.Offset(0, ProducerCol - ProductSeriesCol).Value
            sProductName = rgIntersect.Offset(0, ProductNameCol - ProductSeriesCol).Value
            
            If fNzero(sProducer) And fNzero(sProductName) Then
                Call fSetFilterForSheet(shtProductMaster, Array(1, 2), Array(sProducer, sProductName))
                Call fCopyFilteredDataToRange(shtProductMaster, 3)
                
                sValidationListAddr = "=" & shtDataStage.Columns("A").Address(external:=True)
                'Call fSetValidationListForshtProductNameReplace_ProductName(sValidationListAddr, 3)
                Call fSetValidationListForRange(rgIntersect, sValidationListAddr)
            End If
        Else
            
            'product SeriesCol
            Set rgIntersect = Intersect(Target, Me.Columns(ProductUnitCol))
            
            If Not rgIntersect Is Nothing Then
                If rgIntersect.Areas.Count > 1 Then GoTo Exit_Sub    'fErr "不能选多个"
                If rgIntersect.Rows.Count <> 1 Then GoTo Exit_Sub
                
                sProducer = rgIntersect.Offset(0, ProducerCol - ProductUnitCol).Value
                sProductName = rgIntersect.Offset(0, ProductNameCol - ProductUnitCol).Value
                sProductSeries = rgIntersect.Offset(0, ProductSeriesCol - ProductUnitCol).Value
                
                If fNzero(sProducer) And fNzero(sProductName) Then
                    Call fSetFilterForSheet(shtProductMaster, Array(1, 2, 3), Array(sProducer, sProductName, sProductSeries))
                    Call fCopyFilteredDataToRange(shtProductMaster, 4)
                    
                    sValidationListAddr = "=" & shtDataStage.Columns("A").Address(external:=True)
                    'Call fSetValidationListForshtProductNameReplace_ProductName(sValidationListAddr, 3)
                    Call fSetValidationListForRange(rgIntersect, sValidationListAddr)
                End If
            Else
            
            End If
        End If
    End If
    
Exit_Sub:
    fEnableExcelOptionsAll
    Application.ScreenUpdating = True
    
    If Err.Number <> 0 Then fMsgBox Err.Description
End Sub
