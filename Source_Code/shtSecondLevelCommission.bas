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
    On Error GoTo exit_sub
    
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
    
    '?????????????????????  to do
    
    fMsgBox "没有发现错误", vbInformation
exit_sub:
    Set dictColIndex = Nothing
    fEnableExcelOptionsAll
End Sub
