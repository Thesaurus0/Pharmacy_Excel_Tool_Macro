VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtSalesManCommConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Enum SalesManComm
    SalesCompany = 1
    Hospital = 2
    ProductProducer = 3
    ProductName = 4
    ProductSeries = 5
    BidPrice = 6
    Commission1 = 8
    Commission2 = 10
    Commission3 = 12
    Commission4 = 14
    Commission5 = 16
    Commission6 = 18
    ManagerCommRatio = 20
    [_first] = SalesCompany
    [_last] = 100
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
    
    Call fReadSheetDataByConfig("SALESMAN_COMMISSION_CONFIG", dictColIndex, arrData, , , , , Me)
    
    Call fValidateNumericColInArray(arrData, dictColIndex("BidPrice"), Me, 1, 1, "�б��")
    Call fValidateNumericColInArray(arrData, dictColIndex("Commission1"), Me, 1, 1, "Ӷ��1")
    Call fValidateNumericColInArray(arrData, dictColIndex("Commission2"), Me, 1, 1, "Ӷ��2")
    Call fValidateNumericColInArray(arrData, dictColIndex("Commission3"), Me, 1, 1, "Ӷ��3")
    Call fValidateNumericColInArray(arrData, dictColIndex("Commission4"), Me, 1, 1, "Ӷ��4")
    Call fValidateNumericColInArray(arrData, dictColIndex("Commission5"), Me, 1, 1, "Ӷ��5")
    Call fValidateNumericColInArray(arrData, dictColIndex("Commission6"), Me, 1, 1, "Ӷ��6")
    
    Call fValidateBlankInArray(arrData, dictColIndex("ProductProducer"), Me, 1, 1, "��������")
    Call fValidateBlankInArray(arrData, dictColIndex("ProductName"), Me, 1, 1, "ҩƷ����")
    Call fValidateBlankInArray(arrData, dictColIndex("ProductSeries"), Me, 1, 1, "ԭʼ���")
    
    Call fValidateDuplicateInArray(arrData, Array(dictColIndex("SalesCompany") _
                                  , dictColIndex("Hospital") _
                                  , dictColIndex("ProductProducer") _
                                  , dictColIndex("ProductName") _
                                  , dictColIndex("ProductSeries") _
                                  , dictColIndex("BidPrice")) _
                                , False, Me, 1, 1, "��ҵ��˾+ҽԺ+��������+ҩƷ����+���+�б��")
                                
    Call fCheckIfProducerExistsInProducerMaster(arrData, dictColIndex("ProductProducer"), "[ҩƷ��������]", lErrRowNo, lErrColNo)
    Call fCheckIfProductNameExistsInProductNameMaster(arrData, dictColIndex("ProductProducer"), dictColIndex("ProductName"), "ҩƷ����", lErrRowNo, lErrColNo)
    Call fCheckIfProductExistsInProductMaster(arrData, dictColIndex("ProductProducer"), dictColIndex("ProductName"), dictColIndex("ProductSeries"), lErrRowNo, lErrColNo)

    Call fCheckIfSalesManExistsInSalesManMaster(arrData, dictColIndex("SalesMan1"), "[ҵ��Ա1]", lErrRowNo, lErrColNo)
    Call fCheckIfSalesManExistsInSalesManMaster(arrData, dictColIndex("SalesMan2"), "[ҵ��Ա2]", lErrRowNo, lErrColNo)
    Call fCheckIfSalesManExistsInSalesManMaster(arrData, dictColIndex("SalesMan3"), "[ҵ��Ա3]", lErrRowNo, lErrColNo)

    If bErrMsgBox Then fMsgBox "[" & Me.Name & "]�� û�з��ִ���", vbInformation
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
        Application.Goto Me.Cells(lErrRowNo, lErrColNo) ', True
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
        If rgIntersect.Areas.Count > 1 Then GoTo exit_sub    'fErr "����ѡ���"
        If rgIntersect.Rows.Count <> 1 Then GoTo exit_sub

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
            If rgIntersect.Areas.Count > 1 Then GoTo exit_sub    'fErr "����ѡ���"
            If rgIntersect.Rows.Count <> 1 Then GoTo exit_sub
            
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
    
exit_sub:
    fEnableExcelOptionsAll
    Application.ScreenUpdating = True
    
    If Err.Number <> 0 Then fMsgBox Err.Description
End Sub
