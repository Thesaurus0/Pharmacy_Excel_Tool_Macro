VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtSellPriceInAdv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Option Base 1

Enum SellPriceInAdv
    [_first] = 1
    SalesCompany = 1
    ProductProducer = 2
    ProductName = 3
    ProductSeries = 4
    SellPrice = 5
    [_last] = SellPrice
End Enum

Private Sub btnShtSecondLevelValidation_Click()
    Call fValidateSheet
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error GoTo exit_sub
    Application.ScreenUpdating = False
    
    Dim sProducer As String
    Dim sProductName As String
    Dim sValidationListAddr As String
        
    Dim rgIntersect As Range
    Set rgIntersect = Intersect(Target, Me.Columns(SellPriceInAdv.ProductName))
    
    If Not rgIntersect Is Nothing Then
        If rgIntersect.Areas.Count > 1 Then GoTo exit_sub    'fErr "不能选多个"
        If rgIntersect.Rows.Count <> 1 Then GoTo exit_sub

        sProducer = Me.Cells(rgIntersect.Row, SellPriceInAdv.ProductProducer).Value
        Call fGetProductNameValidationListAndSetToCell(rgIntersect, sProducer)
    Else
        'product SeriesCol
        Set rgIntersect = Intersect(Target, Me.Columns(SellPriceInAdv.ProductSeries))
        
        If Not rgIntersect Is Nothing Then
            If rgIntersect.Areas.Count > 1 Then GoTo exit_sub    'fErr "不能选多个"
            If rgIntersect.Rows.Count <> 1 Then GoTo exit_sub
            
            sProducer = Me.Cells(rgIntersect.Row, SellPriceInAdv.ProductProducer).Value
            sProductName = Me.Cells(rgIntersect.Row, SellPriceInAdv.ProductName).Value
            
            Call fGetProductSeriesValidationListAndSetToCell(rgIntersect, sProducer, sProductName)
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
    
    Call fRemoveFilterForSheet(Me)
    
    Call fCopyReadWholeSheetData2Array(Me, arrData)
    
    Call fValidateDuplicateInArray(arrData, Array(SellPriceInAdv.SalesCompany, SellPriceInAdv.ProductProducer, SellPriceInAdv.ProductName, SellPriceInAdv.ProductSeries) _
                , False, Me, 1, 1, "商业公司+药品厂家+名称+规格")
                
    Call fValidateBlankInArray(arrData, SellPriceInAdv.SalesCompany, Me, 1, 1, "商业公司")
    Call fValidateBlankInArray(arrData, SellPriceInAdv.ProductProducer, Me, 1, 1, "药品厂家")
    Call fValidateBlankInArray(arrData, SellPriceInAdv.ProductName, Me, 1, 1, "药品名称")
    Call fValidateBlankInArray(arrData, SellPriceInAdv.ProductSeries, Me, 1, 1, "药品规格")
    
    Call fCheckIfProducerExistsInProducerMaster(arrData, SellPriceInAdv.ProductProducer, , lErrRowNo, lErrColNo)
    Call fCheckIfProductNameExistsInProductNameMaster(arrData, SellPriceInAdv.ProductProducer, SellPriceInAdv.ProductName, "", lErrRowNo, lErrColNo)
    Call fCheckIfProductExistsInProductMaster(arrData, SellPriceInAdv.ProductProducer, SellPriceInAdv.ProductName, SellPriceInAdv.ProductSeries, lErrRowNo, lErrColNo)
    
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

