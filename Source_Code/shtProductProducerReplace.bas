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
    Call fValidateSheet
End Sub

Function fValidateSheet(Optional bErrMsgBox As Boolean = True) As Boolean
    On Error GoTo exit_sub
    Dim lErrRowNo As Long, lErrColNo As Long
    
    Call fTrimAllCellsForSheet(Me)
    
    Dim arrData()
    Dim dictColIndex As Dictionary
    
    fInitialization
    gsRptID = "REPLACE_UNIFY_SALES_INFO"
    Call fReadSysConfig_InputTxtSheetFile
    
    Call fReadSheetDataByConfig("PRODUCER_REPLACE_SHEET", dictColIndex, arrData, , , , , Me)
    
    Call fValidateBlankInArray(arrData, dictColIndex("FromProducer"), Me, 1, 1, "��������")
    Call fValidateBlankInArray(arrData, dictColIndex("ToProducer"), Me, 1, 1, "�滻Ϊ")
    
    Call fValidateDuplicateInArray(arrData, Array(dictColIndex("FromProducer") _
                                  , dictColIndex("ToProducer")) _
                                , False, Me, 1, 1, "��������+�滻Ϊ")

    Call fCheckIfProducerExistsInProducerMaster(arrData, dictColIndex("ToProducer"), "[�滻Ϊ]", lErrRowNo, lErrColNo)
    
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


