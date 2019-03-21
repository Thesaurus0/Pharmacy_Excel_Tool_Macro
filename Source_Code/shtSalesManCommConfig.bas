VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtSalesManCommConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Option Base 1
 
Enum SalesManComm
    SalesCompany = 1                        '   A
    Hospital = 2                        '   B
    ProductProducer = 3                     '   C
    ProductName = 4                     '   D
    ProductSeries = 5                       '   E
    BidPrice = 6                        '   F
    SalesMan1 = 7                       '   G
    Commission1 = 8                     '   H
    SalesMan2 = 9                       '   I
    Commission2 = 10                    '   J
    SalesMan3 = 11                      '   K
    Commission3 = 12                    '   L
    SalesMan4 = 13                      '   M
    Commission4 = 14                    '   N
    SalesMan5 = 15                      '   O
    Commission5 = 16                    '   P
    SalesMan6 = 17                      '   Q
    Commission6 = 18                    '   R
    SalesManager = 19                       '   S
    ManagerCommRatio = 20                     '   T
    [_first] = SalesCompany
    [_last] = 20
End Enum

Private Sub btnValidate_Click()
    Call fValidateSheet
End Sub

Function fValidateSheet(Optional bErrMsgBox As Boolean = True) As Boolean
    On Error GoTo exit_sub
    
    Dim lErrRowNo As Long, lErrColNo As Long
    
    Call fTrimAllCellsForSheet(Me)
    
    Dim arrData()
    Dim dictColIndex As Dictionary
    
    fInitialization
    gsRptID = "CALCULATE_PROFIT"
    Call fReadSysConfig_InputTxtSheetFile
    
    'Call fReadSheetDataByConfig("SALESMAN_COMMISSION_CONFIG", dictColIndex, arrData, , , , , Me)
    Call fCopyReadWholeSheetData2Array(Me, arrData)
    
    Call fValidateNumericColInArray(arrData, SalesManComm.BidPrice, Me, 1, 1, "中标价")
    Call fValidateNumericColInArray(arrData, SalesManComm.Commission1, Me, 1, 1, "佣金1")
    Call fValidateNumericColInArray(arrData, SalesManComm.Commission2, Me, 1, 1, "佣金2")
    Call fValidateNumericColInArray(arrData, SalesManComm.Commission3, Me, 1, 1, "佣金3")
    Call fValidateNumericColInArray(arrData, SalesManComm.Commission4, Me, 1, 1, "佣金4")
    Call fValidateNumericColInArray(arrData, SalesManComm.Commission5, Me, 1, 1, "佣金5")
    Call fValidateNumericColInArray(arrData, SalesManComm.Commission6, Me, 1, 1, "佣金6")
    
    Call fValidateBlankInArray(arrData, SalesManComm.ProductProducer, Me, 1, 1, "生产厂家")
    Call fValidateBlankInArray(arrData, SalesManComm.ProductName, Me, 1, 1, "药品名称")
    Call fValidateBlankInArray(arrData, SalesManComm.ProductSeries, Me, 1, 1, "原始规格")
    
    Call fValidateDuplicateInArray(arrData, Array(SalesManComm.SalesCompany _
                                  , SalesManComm.Hospital _
                                  , SalesManComm.ProductProducer _
                                  , SalesManComm.ProductName _
                                  , SalesManComm.ProductSeries _
                                  , SalesManComm.BidPrice) _
                                , False, Me, 1, 1, "商业公司+医院+生产厂家+药品名称+规格+中标价")
                                
    Call fCheckIfProducerExistsInProducerMaster(arrData, SalesManComm.ProductProducer, "[药品生产厂家]", lErrRowNo, lErrColNo)
    Call fCheckIfProductNameExistsInProductNameMaster(arrData, SalesManComm.ProductProducer, SalesManComm.ProductName, "药品名称", lErrRowNo, lErrColNo)
    Call fCheckIfProductExistsInProductMaster(arrData, SalesManComm.ProductProducer, SalesManComm.ProductName, SalesManComm.ProductSeries, lErrRowNo, lErrColNo)

    Call fCheckIfSalesManExistsInSalesManMaster(arrData, SalesManComm.SalesMan1, "[业务员1]", lErrRowNo, lErrColNo)
    Call fCheckIfSalesManExistsInSalesManMaster(arrData, SalesManComm.SalesMan2, "[业务员2]", lErrRowNo, lErrColNo)
    Call fCheckIfSalesManExistsInSalesManMaster(arrData, SalesManComm.SalesMan3, "[业务员3]", lErrRowNo, lErrColNo)

    If bErrMsgBox Then fMsgBox "[" & Me.Name & "]表 保存成功", vbInformation: ThisWorkbook.Save
exit_sub:
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
    On Error GoTo exit_sub
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
        If rgIntersect.Areas.Count > 1 Then GoTo exit_sub    'fErr "不能选多个"
        If rgIntersect.Rows.Count <> 1 Then GoTo exit_sub

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
            If rgIntersect.Areas.Count > 1 Then GoTo exit_sub    'fErr "不能选多个"
            If rgIntersect.Rows.Count <> 1 Then GoTo exit_sub
            
            sProducer = rgIntersect.Offset(0, ProducerCol - ProductSeriesCol).Value
            sProductName = rgIntersect.Offset(0, ProductNameCol - ProductSeriesCol).Value
            Call fGetProductSeriesValidationListAndSetToCell(rgIntersect, sProducer, sProductName)
            
'            If fNzero(sProducer) And fNzero(sProductName) Then
'                Call fSetFilterForSheet(shtProductMaster, Array(1, 2), Array(sProducer, sProductName))
'                Call fCopyFilteredDataToRange(shtProductMaster, 3)
'
'                sValidationListAddr = "=" & shtDataStage.Columns("A").Address(external:=True)
'                'Call fSetValidationListForshtProductNameReplace_ProductName(sValidationListAddr, 3)
'                Call fSetValidationListForRange(rgIntersect, sValidationListAddr)
'            End If
        End If
    End If
    
exit_sub:
    fEnableExcelOptionsAll
    Application.ScreenUpdating = True
    
    If Err.Number <> 0 Then fMsgBox Err.Description
End Sub
