VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtProductTaxRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Option Base 1

Enum ProdTaxRate
    ProductProducer = 1
    ProductName = 2
    ProductSeries = 3
    TaxRate = 4
End Enum

Private Sub btnValidate_Click()
    Call fValidateSheet
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error GoTo exit_sub
    Application.ScreenUpdating = False
    
    Const ProducerCol = 1
    Const ProductNameCol = 2
    Const ProductSeriesCol = 3
    'Const ProductUnitCol = 4
    Dim sProductName As String
    Dim rgIntersect As Range
    Set rgIntersect = Intersect(Target, Me.Columns(ProductNameCol))
    
    If Not rgIntersect Is Nothing Then
        If rgIntersect.Areas.count > 1 Then GoTo exit_sub    'fErr "不能选多个"
        If rgIntersect.Rows.count <> 1 Then GoTo exit_sub
            
        Dim sProducer As String
        Dim sValidationListAddr As String
        
        sProducer = rgIntersect.Offset(0, ProducerCol - ProductNameCol).Value
        Call fGetProductNameValidationListAndSetToCell(rgIntersect, sProducer)
        
'        If fNzero(sProducer) Then
'            Call fSetFilterForSheet(shtProductNameMaster, ProductNameMst.ProdProducer, sProducer)
'            Call fCopyFilteredDataToRange(shtProductNameMaster, 2)
'
'            sValidationListAddr = "=" & shtDataStage.Columns("A").Address(external:=True)
'            'Call fSetValidationListForshtProductNameReplace_ProductName(sValidationListAddr, 3)
'            Call fSetValidationListForRange(rgIntersect, sValidationListAddr)
'        End If
    Else
        'product SeriesCol
        Set rgIntersect = Intersect(Target, Me.Columns(ProductSeriesCol))
        
        If Not rgIntersect Is Nothing Then
            If rgIntersect.Areas.count > 1 Then GoTo exit_sub    'fErr "不能选多个"
            If rgIntersect.Rows.count <> 1 Then GoTo exit_sub
            
            sProducer = rgIntersect.Offset(0, ProducerCol - ProductSeriesCol).Value
            sProductName = rgIntersect.Offset(0, ProductNameCol - ProductSeriesCol).Value
            Call fGetProductSeriesValidationListAndSetToCell(rgIntersect, sProducer, sProductName)
            
'            If fNzero(sProducer) And fNzero(sProductName) Then
'                Call fSetFilterForSheet(shtProductMaster, Array(ProductMst.ProductProducer, ProductMst.ProductName), Array(sProducer, sProductName))
'                Call fCopyFilteredDataToRange(shtProductMaster, 3)
'
'                sValidationListAddr = "=" & shtDataStage.Columns("A").Address(external:=True)
'                'Call fSetValidationListForshtProductNameReplace_ProductName(sValidationListAddr, 3)
'                Call fSetValidationListForRange(rgIntersect, sValidationListAddr)
'            End If
        Else
'
'            'product UnitCol
'            Set rgIntersect = Intersect(Target, Me.Columns(ProductUnitCol))
'
'            If Not rgIntersect Is Nothing Then
'                If rgIntersect.Areas.Count > 1 Then GoTo exit_sub    'fErr "不能选多个"
'                If rgIntersect.Rows.Count <> 1 Then GoTo exit_sub
'
'                sProducer = rgIntersect.Offset(0, ProducerCol - ProductUnitCol).Value
'                sProductName = rgIntersect.Offset(0, ProductNameCol - ProductUnitCol).Value
'                sProductSeries = rgIntersect.Offset(0, ProductSeriesCol - ProductUnitCol).Value
'
'                If fNzero(sProducer) And fNzero(sProductName) Then
'                    Call fSetFilterForSheet(shtProductMaster, Array(ProductMst.ProductProducer, ProductMst.ProductName, ProductMst.ProductSeries), Array(sProducer, sProductName, sProductSeries))
'                    Call fCopyFilteredDataToRange(shtProductMaster, 4)
'
'                    sValidationListAddr = "=" & shtDataStage.Columns("A").Address(external:=True)
'                    'Call fSetValidationListForshtProductNameReplace_ProductName(sValidationListAddr, 3)
'                    Call fSetValidationListForRange(rgIntersect, sValidationListAddr)
'                End If
'            Else
            
'            End If
        End If
    End If
    
exit_sub:
    fEnableExcelOptionsAll
    Application.ScreenUpdating = True
    
    If Err.Number <> 0 Then fMsgBox Err.Description
End Sub

Function fValidateSheet(Optional bErrMsgBox As Boolean = True) As Boolean
    On Error GoTo exit_sub
    
    Dim lErrRowNo As Long, lErrColNo As Long
    Call fTrimAllCellsForSheet(Me)
    
    Dim arrData()
    
    fInitialization
    gsRptID = "CALCULATE_PROFIT"
    Call fReadSysConfig_InputTxtSheetFile
    
    Call fCopyReadWholeSheetData2Array(Me, arrData)
    
    Call fValidateBlankInArray(arrData, ProdTaxRate.ProductProducer, Me, 1, 1, "生产厂家")
    Call fValidateBlankInArray(arrData, ProdTaxRate.ProductName, Me, 1, 1, "药品名称")
    Call fValidateBlankInArray(arrData, ProdTaxRate.ProductSeries, Me, 1, 1, "规格")
    
    Call fValidateDuplicateInArray(arrData, Array(ProdTaxRate.ProductProducer _
                                  , ProdTaxRate.ProductName _
                                  , ProdTaxRate.ProductSeries) _
                                , False, Me, 1, 1, "生产厂家+药品名称+规格 ")
    
    Call fCheckIfProductExistsInProductMaster(arrData, ProdTaxRate.ProductProducer _
                                  , ProdTaxRate.ProductName, ProdTaxRate.ProductSeries, lErrRowNo, lErrColNo)

    If bErrMsgBox Then fMsgBox "[" & Me.Name & "]表 保存成功", vbInformation: ThisWorkbook.Save
exit_sub:
    fEnableExcelOptionsAll
    'Set dictColIndex = Nothing
    Erase arrData
    
    If Err.Number <> 0 Then
        fShowAndActiveSheet Me
        fValidateSheet = False
    Else
        fValidateSheet = True
    End If
    If lErrRowNo > 0 Then
        fShowAndActiveSheet Me
        Application.Goto Me.Cells(lErrRowNo, lErrColNo) ', True
    End If
End Function

