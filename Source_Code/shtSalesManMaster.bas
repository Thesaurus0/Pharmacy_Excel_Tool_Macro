VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtSalesManMaster"
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
    
    Call fTrimAllCellsForSheet(Me)
    
    
    Dim arrData()
    Dim dictColIndex As Dictionary
    
    fInitialization
    gsRptID = "CALCULATE_PROFIT"
    Call fReadSysConfig_InputTxtSheetFile
    
    Call fReadSheetDataByConfig("SALESMAN_MASTER_SHEET", dictColIndex, arrData, , , , , Me)
    
    Call fValidateDuplicateInArray(arrData, dictColIndex("SalesManName"), False, Me, 1, 1, "业务员姓名")
'    Call fValidateBlankInArray(arrData, dictColIndex("ProductProducer"), Me, 1, 1, "药品厂家")
'
'    Dim lEachRow As Long
'    For lEachRow = LBound(arrData, 1) To UBound(arrData, 1)
'        If Not fProductNameExistsInProductNameMaster(CStr(arrData(lEachRow, dictColIndex("ProductProducer"))) _
'                                            , CStr(arrData(lEachRow, dictColIndex("ProductName")))) Then
'            fErr "【药品厂家 + 药品名称】不存在于药品名称主表中。" & vbCr & "行号：" & lEachRow + 1
'        End If
'    Next
    
    Call fSortDataInSheetSortSheetData(Me, dictColIndex("SalesManName"))
                                                
    If bErrMsgBox Then fMsgBox "[" & Me.Name & "]表 保存成功", vbInformation: ThisWorkbook.Save
exit_sub:
    Set dictColIndex = Nothing
    fEnableExcelOptionsAll
    Erase arrData
    
    If Err.Number <> 0 Then
        fShowAndActiveSheet Me
        fValidateSheet = False
    Else
        fValidateSheet = True
    End If
End Function

Private Sub Worksheet_Change(ByVal Target As Range)
    Call fResetdictSalesManMaster
End Sub
 
