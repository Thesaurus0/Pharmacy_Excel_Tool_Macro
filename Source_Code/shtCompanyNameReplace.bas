VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtCompanyNameReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Option Base 1

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
    gsRptID = "REPLACE_UNIFY_CZL_SALES_TO_COMP"
    Call fReadSysConfig_InputTxtSheetFile
    
    Call fReadSheetDataByConfig("COMPANY_NAME_REPLACE_SHEET", dictColIndex, arrData, , , , , Me)
    
    Call fValidateBlankInArray(arrData, dictColIndex("FromCompanyName"), Me, 1, 1, "原始文件商业公司名称")
    Call fValidateBlankInArray(arrData, dictColIndex("ToCompanyName"), Me, 1, 1, "替换为")
    
    Call fValidateDuplicateInArray(arrData, Array(dictColIndex("FromCompanyName") _
                                  , dictColIndex("ToCompanyName")) _
                                , False, Me, 1, 1, "原始文件商业公司名称+替换为")

    Call fCheckIfCompanyNameExistsInrngStaticSalesCompanyNames(arrData, dictColIndex("ToCompanyName"), "[替换为]", lErrRowNo, lErrColNo)
    
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
        Application.Goto Me.Cells(lErrRowNo, lErrColNo) ', True
    End If
End Function


