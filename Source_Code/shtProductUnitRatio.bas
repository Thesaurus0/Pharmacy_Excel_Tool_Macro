VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtProductUnitRatio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub btnValidate_Click()
    Call sub_Validate
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

Function fValidateSheet()
    On Error GoTo Exit_Sub
    
    Call fTrimAllCellsForSheet(Me)
    
    Dim arrData()
    Dim dictColIndex As Dictionary
    
    fInitialization
    gsRptID = "REPLACE_UNIFY_SALES_INFO"
    Call fReadSysConfig_InputTxtSheetFile
    
    Call fReadSheetDataByConfig("PRODUCT_UNIT_RATIO_SHEET", dictColIndex, arrData, , , , , shtProductUnitRatio)
    
    Call fValidateBlankInArray(arrData, dictColIndex("ProductProducer"), shtProductNameMaster, 1, 1, "生产厂家")
    Call fValidateBlankInArray(arrData, dictColIndex("ProductName"), shtProductNameMaster, 1, 1, "药品名称")
    Call fValidateBlankInArray(arrData, dictColIndex("ProductSeries"), shtProductNameMaster, 1, 1, "规格")
    Call fValidateBlankInArray(arrData, dictColIndex("FromUnit"), shtProductNameMaster, 1, 1, "统一计量单位")
    Call fValidateBlankInArray(arrData, dictColIndex("ProductUnit"), shtProductNameMaster, 1, 1, "原始文件药品单位")
    
    Call fValidateDuplicateInArray(arrData, Array(dictColIndex("ProductProducer") _
                                  , dictColIndex("ProductName") _
                                  , dictColIndex("ProductSeries") _
                                  , dictColIndex("FromUnit") _
                                  , dictColIndex("ProductUnit")) _
                                , False, shtProductUnitRatio, 1, 1, "生产厂家+药品名称+规格+统一计量单位+原始文件药品单位")
    
    Call fCheckIfProductExistsInProductMaster(arrData, dictColIndex("ProductProducer"), dictColIndex("ProductName"), dictColIndex("ProductSeries"))

    fMsgBox "[" & Me.Name & "]表 没有发现错误", vbInformation
Exit_Sub:
    fEnableExcelOptionsAll
    Set dictColIndex = Nothing
    Erase arrData
    
    If Err.Number <> 0 Then
        fShowAndActiveSheet Me
        fValidateSheet = False
    Else
        fValidateSheet = True
    End If
End Function

