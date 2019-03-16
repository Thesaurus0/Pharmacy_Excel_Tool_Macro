VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtHospital"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Option Base 1

Enum enHospital
    HospitalName = 1
    Address = 2
End Enum

Private Sub btnValidate_Click()
    Call fValidateSheet
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Call fResetdictHospitalMaster
End Sub

Function fValidateSheet(Optional bErrMsgBox As Boolean = True) As Boolean
    On Error GoTo exit_sub
    
    Call fTrimAllCellsForSheet(Me)
    
    Dim arrData()
    Dim dictColIndex As Dictionary
    
    fInitialization
    gsRptID = "REPLACE_UNIFY_SALES_INFO"
    Call fReadSysConfig_InputTxtSheetFile
    
    Call fReadSheetDataByConfig("HOSPITAL_MASTER", dictColIndex, arrData, , , , , Me)
    
    Call fValidateDuplicateInArray(arrData, dictColIndex("Hospital"), False, Me, 1, 1, "医院")
    
    Call fSortDataInSheetSortSheetData(Me, dictColIndex("Hospital"))
    
    If bErrMsgBox Then fMsgBox "[" & Me.Name & "]主表 保存成功", vbInformation: ThisWorkbook.Save
    'If bErrMsgBox Then fMsgBox "[医院]主表 没有发现错误", vbInformation
exit_sub:
    fEnableExcelOptionsAll
    Set dictColIndex = Nothing
    Erase arrData
    
    If Err.Number <> 0 Then
        alErrRowNo = (lEachRow + 1)
        alErrColNo = iColProductSeries
        fShowAndActiveSheet Me
        fValidateSheet = False
    Else
        fValidateSheet = True
    End If
End Function

Sub aaa()
    Debug.Print Me.Name
End Sub
