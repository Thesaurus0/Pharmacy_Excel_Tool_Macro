VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtHospitalReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Option Base 1

Private Sub btnValidate_Click()
    Call fValidateSheet
End Sub

Function fValidateSheet(Optional bErrMsgBox As Boolean = True, Optional ByRef alErrRowNo As Long, Optional ByRef alErrColNo As Long) As Boolean
    On Error GoTo exit_sub
    Dim lErrRowNo As Long, lErrColNo As Long
    
    Call fTrimAllCellsForSheet(Me)
    
    Dim arrData()
    Dim dictColIndex As Dictionary
    
    fInitialization
    gsRptID = "REPLACE_UNIFY_SALES_INFO"
    Call fReadSysConfig_InputTxtSheetFile
    
    Call fReadSheetDataByConfig("HOSPITAL_REPLACE_SHEET", dictColIndex, arrData, , , , , Me)
    
    Call fValidateBlankInArray(arrData, dictColIndex("FromHospital"), Me, 1, 1, "ҽԺ")
    Call fValidateBlankInArray(arrData, dictColIndex("ToHospital"), Me, 1, 1, "�滻Ϊ")
    
    Call fValidateDuplicateInArray(arrData, Array(dictColIndex("FromHospital") _
                                  , dictColIndex("ToHospital")) _
                                , False, Me, 1, 1, "ҽԺ+�滻Ϊ")

    Call fCheckIfHospitalExistsInHospitalMaster(arrData, dictColIndex("ToHospital"), , lErrRowNo, lErrColNo)

    If bErrMsgBox Then fMsgBox "[" & Me.Name & "]�� ����ɹ�", vbInformation: ThisWorkbook.Save
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

