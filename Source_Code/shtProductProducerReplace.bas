VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtProductProducerReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub btnValidate_Click()
    Call sub_Validate
End Sub

Function fValidateSheet()
    On Error GoTo Exit_Sub
    
    Call fTrimAllCellsForSheet(Me)
    
    Dim arrData()
    Dim dictColIndex As Dictionary
    
    fInitialization
    gsRptID = "REPLACE_UNIFY_SALES_INFO"
    Call fReadSysConfig_InputTxtSheetFile
    
    Call fReadSheetDataByConfig("PRODUCER_REPLACE_SHEET", dictColIndex, arrData, , , , , Me)
    
    Call fValidateBlankInArray(arrData, dictColIndex("FromProducer"), Me, 1, 1, "生产厂家")
    Call fValidateBlankInArray(arrData, dictColIndex("ToProducer"), Me, 1, 1, "替换为")
    
    Call fValidateDuplicateInArray(arrData, Array(dictColIndex("FromProducer") _
                                  , dictColIndex("ToProducer")) _
                                , False, Me, 1, 1, "生产厂家+替换为")

    Call fCheckIfProducerExistsInProducerMaster(arrData, dictColIndex("ToProducer"), "[替换为]")
    
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


