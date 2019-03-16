VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtSecondLevelCommission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Option Base 1

Enum SecondLevelComm
    SalesCompany = 1
    Hospital = 2
    ProductProducer = 3
    ProductName = 4
    ProductSeries = 5
    Commission = 6
    [_first] = SalesCompany
    [_last] = Commission
End Enum

Private Sub btnShtSecondLevelValidation_Click()
    Call fValidateSheet
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error GoTo exit_sub
    Application.ScreenUpdating = False
    
    Const ProducerCol = 3
    Const ProductNameCol = 4
    Const ProductSeriesCol = 5
    'Const ProductUnitCol = 5
    
    Dim sProductName As String
    Dim rgIntersect As Range
    Set rgIntersect = Intersect(Target, Me.Columns(ProductNameCol))
    
    If Not rgIntersect Is Nothing Then
        If rgIntersect.Areas.Count > 1 Then GoTo exit_sub    'fErr "不能选多个"
        If rgIntersect.Rows.Count <> 1 Then GoTo exit_sub
            
        Dim sProducer As String
        Dim sValidationListAddr As String
        
        sProducer = Me.Cells(rgIntersect.Row, SecondLevelComm.ProductProducer).Value
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
            If rgIntersect.Areas.Count > 1 Then GoTo exit_sub    'fErr "不能选多个"
            If rgIntersect.Rows.Count <> 1 Then GoTo exit_sub
            
            sProducer = Me.Cells(rgIntersect.Row, SecondLevelComm.ProductProducer).Value
            sProductName = Me.Cells(rgIntersect.Row, SecondLevelComm.ProductName).Value
            Call fGetProductSeriesValidationListAndSetToCell(rgIntersect, sProducer, sProductName)
            
'            If fNzero(sProducer) And fNzero(sProductName) Then
'                Call fSetFilterForSheet(shtProductMaster, Array(1, 2), Array(sProducer, sProductName))
'                Call fCopyFilteredDataToRange(shtProductMaster, 3)
'
'                sValidationListAddr = "=" & shtDataStage.Columns("A").Address(external:=True)
'                'Call fSetValidationListForshtProductNameReplace_ProductName(sValidationListAddr, 3)
'                Call fSetValidationListForRange(rgIntersect, sValidationListAddr)
'            End If
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
    
exit_sub:
    fEnableExcelOptionsAll
    Application.ScreenUpdating = True
    
    If Err.Number <> 0 Then fMsgBox Err.Description
End Sub

Function fValidateSheet(Optional bErrMsgBox As Boolean = True) As Boolean
    On Error GoTo exit_sub
    
    Dim lErrRowNo As Long, lErrColNo As Long
    
    Dim arrData()
    'Dim dictColIndex As Dictionary
    
    fInitialization
    gsRptID = "CALCULATE_PROFIT"
    Call fReadSysConfig_InputTxtSheetFile
    
    Call fRemoveFilterForSheet(Me)
    'Call fReadSheetDataByConfig("SECOND_LEVEL_COMMISSION", dictColIndex, arrData, , , , , Me)
'    Call fValidateDuplicateInArray(arrData, Array(dictColIndex("SalesCompany") _
'                                                , SecondLevelComm.Hospital _
'                                                , SecondLevelComm.ProductProducer _
'                                                , SecondLevelComm.ProductName _
'                                                , SecondLevelComm.ProductSeries) _
'                , False, Me, 1, 1, "商业公司+医院+厂家+名称+规格")
                
    Call fCopyReadWholeSheetData2Array(Me, arrData)
    Call fValidateDuplicateInArray(arrData, Array(SecondLevelComm.SalesCompany, SecondLevelComm.Hospital _
                  , SecondLevelComm.ProductProducer, SecondLevelComm.ProductName, SecondLevelComm.ProductSeries) _
                , False, Me, 1, 1, "商业公司+医院+厂家+名称+规格")
            
    Call fValidateBlankInArray(arrData, SecondLevelComm.SalesCompany, Me, 1, 1, "商业公司")
    Call fValidateBlankInArray(arrData, SecondLevelComm.Hospital, Me, 1, 1, "医院")
    Call fValidateBlankInArray(arrData, SecondLevelComm.ProductProducer, Me, 1, 1, "药品厂家")
    Call fValidateBlankInArray(arrData, SecondLevelComm.ProductName, Me, 1, 1, "药品名称")
    Call fValidateBlankInArray(arrData, SecondLevelComm.ProductSeries, Me, 1, 1, "药品规格")
    
    Call fCheckIfHospitalExistsInHospitalMaster(arrData, SecondLevelComm.Hospital)
    Call fCheckIfProducerExistsInProducerMaster(arrData, SecondLevelComm.ProductProducer, , lErrRowNo, lErrColNo)
    Call fCheckIfProductNameExistsInProductNameMaster(arrData, SecondLevelComm.ProductProducer, SecondLevelComm.ProductName, "", lErrRowNo, lErrColNo)
    Call fCheckIfProductExistsInProductMaster(arrData, SecondLevelComm.ProductProducer, SecondLevelComm.ProductName, SecondLevelComm.ProductSeries, lErrRowNo, lErrColNo)
    
    If bErrMsgBox Then fMsgBox "[" & Me.Name & "]表 保存成功", vbInformation: ThisWorkbook.Save
exit_sub:
    'Set dictColIndex = Nothing
    fEnableExcelOptionsAll
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
    
    If Err.Number <> 0 And Err.Number <> gErrNum Then fMsgBox Err.Description
End Function

