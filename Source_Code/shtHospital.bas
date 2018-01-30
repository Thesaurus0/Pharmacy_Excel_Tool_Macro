VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtHospital"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub btnValidate_Click()
    Call fValidateSheet
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Call fResetdictHospitalMaster
End Sub

Function fValidateSheet(Optional bErrMsgBox As Boolean = True) As Boolean
    On Error GoTo Exit_Sub
    
    Call fTrimAllCellsForSheet(Me)
    
    Dim arrData()
    Dim dictColIndex As Dictionary
    
    fInitialization
    gsRptID = "REPLACE_UNIFY_SALES_INFO"
    Call fReadSysConfig_InputTxtSheetFile
    
    Call fReadSheetDataByConfig("HOSPITAL_MASTER", dictColIndex, arrData, , , , , Me)
    
    Call fValidateDuplicateInArray(arrData, dictColIndex("Hospital"), False, Me, 1, 1, "医院")
    
    Call fSortDataInSheetSortSheetData(Me, dictColIndex("Hospital"))
    
    If bErrMsgBox Then fMsgBox "[" & Me.Name & "]主表 没有发现错误", vbInformation
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
