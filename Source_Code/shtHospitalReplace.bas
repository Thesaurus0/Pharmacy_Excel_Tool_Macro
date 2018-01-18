VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtHospitalReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub btnValidate_Click()
    Call sub_Validate
End Sub

Function fValidateSheet()
    On Error GoTo exit_sub
    
    Call fTrimAllCellsForSheet(Me)
    
    Dim arrData()
    Dim dictColIndex As Dictionary
    
    fInitialization
    gsRptID = "REPLACE_UNIFY_SALES_INFO"
    Call fReadSysConfig_InputTxtSheetFile
    
    Call fReadSheetDataByConfig("HOSPITAL_REPLACE_SHEET", dictColIndex, arrData, , , , , Me)
    
    Call fValidateBlankInArray(arrData, dictColIndex("FromHospital"), Me, 1, 1, "医院")
    Call fValidateBlankInArray(arrData, dictColIndex("ToHospital"), Me, 1, 1, "替换为")
    
    Call fValidateDuplicateInArray(arrData, Array(dictColIndex("FromHospital") _
                                  , dictColIndex("ToHospital")) _
                                , False, Me, 1, 1, "医院+替换为")

    Call fCheckIfHospitalExistsInHospitalMaster(arrData, dictColIndex("ToHospital"))

    fMsgBox "[" & Me.Name & "]表 没有发现错误", vbInformation
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
End Function

