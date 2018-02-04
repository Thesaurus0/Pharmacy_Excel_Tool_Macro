VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtSalesManCommConfig"
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
    
    Call fReadSheetDataByConfig("SALESMAN_COMMISSION_CONFIG", dictColIndex, arrData, , , , , Me)
    
    Call fValidateNumericColInArray(arrData, dictColIndex("BidPrice"), Me, 1, 1, "中标价")
    Call fValidateNumericColInArray(arrData, dictColIndex("Commission1"), Me, 1, 1, "佣金1")
    Call fValidateNumericColInArray(arrData, dictColIndex("Commission2"), Me, 1, 1, "佣金2")
    Call fValidateNumericColInArray(arrData, dictColIndex("Commission3"), Me, 1, 1, "佣金3")
    
    Call fValidateBlankInArray(arrData, dictColIndex("ProductProducer"), Me, 1, 1, "生产厂家")
    Call fValidateBlankInArray(arrData, dictColIndex("ProductName"), Me, 1, 1, "药品名称")
    Call fValidateBlankInArray(arrData, dictColIndex("ProductSeries"), Me, 1, 1, "原始规格")
    
    Call fValidateDuplicateInArray(arrData, Array(dictColIndex("SalesCompany") _
                                  , dictColIndex("Hospital") _
                                  , dictColIndex("ProductProducer") _
                                  , dictColIndex("ProductName") _
                                  , dictColIndex("ProductSeries") _
                                  , dictColIndex("BidPrice")) _
                                , False, Me, 1, 1, "商业公司+医院+生产厂家+药品名称+规格+中标价")
                                
    Call fCheckIfProducerExistsInProducerMaster(arrData, dictColIndex("ProductProducer"), "[药品生产厂家]", lErrRowNo, lErrColNo)
    Call fCheckIfProductNameExistsInProductNameMaster(arrData, dictColIndex("ProductProducer"), dictColIndex("ProductName"), "药品名称", lErrRowNo, lErrColNo)
    Call fCheckIfProductExistsInProductMaster(arrData, dictColIndex("ProductProducer"), dictColIndex("ProductName"), dictColIndex("ProductSeries"), lErrRowNo, lErrColNo)

    Call fCheckIfSalesManExistsInSalesManMaster(arrData, dictColIndex("SalesMan1"), "[业务员1]", lErrRowNo, lErrColNo)
    Call fCheckIfSalesManExistsInSalesManMaster(arrData, dictColIndex("SalesMan2"), "[业务员2]", lErrRowNo, lErrColNo)
    Call fCheckIfSalesManExistsInSalesManMaster(arrData, dictColIndex("SalesMan3"), "[业务员3]", lErrRowNo, lErrColNo)

    If bErrMsgBox Then fMsgBox "[" & Me.Name & "]表 没有发现错误", vbInformation
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
    
    If lErrRowNo > 0 Then
        fShowAndActiveSheet Me
        Application.GoTo Me.Cells(lErrRowNo, lErrColNo) ', True
    End If
End Function

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error GoTo Exit_Sub
    Application.ScreenUpdating = False
    
    Const ProducerCol = 3
    Const ProductNameCol = 4
    Const ProductSeriesCol = 5
    
    Dim rgIntersect As Range
    Dim sProducer As String
    Dim sProductName As String
    Dim sValidationListAddr As String
        
    'product name
    Set rgIntersect = Intersect(Target, Me.Columns(ProductNameCol))
    
    If Not rgIntersect Is Nothing Then
        If rgIntersect.Areas.Count > 1 Then GoTo Exit_Sub    'fErr "不能选多个"
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
        End If
    End If
    
Exit_Sub:
    fEnableExcelOptionsAll
    Application.ScreenUpdating = True
    
    If Err.Number <> 0 Then fMsgBox Err.Description
End Sub
