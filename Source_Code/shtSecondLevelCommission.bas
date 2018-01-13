VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtSecondLevelCommission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub btnShtSecondLevelValidation_Click()
    Call sub_Validate
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error GoTo Exit_Sub
    Application.ScreenUpdating = False
    
    Const ProducerCol = 3
    Const ProductNameCol = 4
    Const ProductSeriesCol = 5
    'Const ProductUnitCol = 5
    
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
'            'product SeriesCol
'            Set rgIntersect = Intersect(Target, Me.Columns(ProductUnitCol))
'
'            If Not rgIntersect Is Nothing Then
'                If rgIntersect.Areas.Count > 1 Then fErr "不能选多个"
'                If rgIntersect.Rows.Count <> 1 Then GoTo exit_sub
'
'                sProducer = rgIntersect.Offset(0, ProducerCol - ProductUnitCol).Value
'                sProductName = rgIntersect.Offset(0, ProductNameCol - ProductUnitCol).Value
'                sProductSeries = rgIntersect.Offset(0, ProductSeriesCol - ProductUnitCol).Value
'
'                If fNzero(sProducer) And fNzero(sProductName) Then
'                    Call fSetFilterForSheet(shtProductMaster, Array(1, 2, 3), Array(sProducer, sProductName, sProductSeries))
'                    Call fCopyFilteredDataToRange(shtProductMaster, 4)
'
'                    sValidationListAddr = "=" & shtDataStage.Columns("A").Address(external:=True)
'                    'Call fSetValidationListForshtProductNameReplace_ProductName(sValidationListAddr, 3)
'                    Call fSetValidationListForRange(rgIntersect, sValidationListAddr)
'                End If
'            Else
'
'            End If
        End If
    End If
    
Exit_Sub:
    fEnableExcelOptionsAll
    Application.ScreenUpdating = True
    
    If Err.Number <> 0 Then fMsgBox Err.Description
End Sub

Function fValidateSheet()
    On Error GoTo Exit_Sub
    
    Dim arrData()
    Dim dictColIndex As Dictionary
    
    fInitialization
    gsRptID = "CALCULATE_PROFIT"
    Call fReadSysConfig_InputTxtSheetFile
    
    Call fRemoveFilterForSheet(shtSecondLevelCommission)
    Call fReadSheetDataByConfig("SECOND_LEVEL_COMMISSION", dictColIndex, arrData, , , , , shtSecondLevelCommission)
    
    Call fValidateDuplicateInArray(arrData, Array(dictColIndex("SalesCompany") _
                                                , dictColIndex("Hospital") _
                                                , dictColIndex("ProductProducer") _
                                                , dictColIndex("ProductName") _
                                                , dictColIndex("ProductSeries")) _
                , False, shtProductProducerMaster, 1, 1, "商业公司+医院+厂家+名称+规格")
                
    Call fValidateBlankInArray(arrData, dictColIndex("SalesCompany"), shtProductMaster, 1, 1, "商业公司")
    Call fValidateBlankInArray(arrData, dictColIndex("Hospital"), shtProductMaster, 1, 1, "医院")
    Call fValidateBlankInArray(arrData, dictColIndex("ProductProducer"), shtProductMaster, 1, 1, "药品厂家")
    Call fValidateBlankInArray(arrData, dictColIndex("ProductName"), shtProductMaster, 1, 1, "药品名称")
    Call fValidateBlankInArray(arrData, dictColIndex("ProductSeries"), shtProductMaster, 1, 1, "药品规格")
    
    Call fCheckIfHospitalExistsInHospitalMaster(arrData, dictColIndex("Hospital"))
    Call fCheckIfProducerExistsInProducerMaster(arrData, dictColIndex("ProductProducer"))
    Call fCheckIfProductNameExistsInProductNameMaster(arrData, dictColIndex("ProductProducer"), dictColIndex("ProductName"))
    Call fCheckIfProductExistsInProductMaster(arrData, dictColIndex("ProductProducer"), dictColIndex("ProductName"), dictColIndex("ProductSeries"))
    
    fMsgBox "[" & Me.Name & "]表 没有发现错误", vbInformation
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

