VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtSalesManMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub btnValidate_Click()
    Call fValidateSheet
End Sub

Function fValidateSheet(Optional bErrMsgBox As Boolean = True) As Boolean
    On Error GoTo Exit_Sub
    
    Call fTrimAllCellsForSheet(Me)
    
    
    Dim arrData()
    Dim dictColIndex As Dictionary
    
    fInitialization
    gsRptID = "CALCULATE_PROFIT"
    Call fReadSysConfig_InputTxtSheetFile
    
    Call fReadSheetDataByConfig("SALESMAN_MASTER_SHEET", dictColIndex, arrData, , , , , Me)
    
    Call fValidateDuplicateInArray(arrData, dictColIndex("SalesManName"), False, Me, 1, 1, "ҵ��Ա����")
'    Call fValidateBlankInArray(arrData, dictColIndex("ProductProducer"), Me, 1, 1, "ҩƷ����")
'
'    Dim lEachRow As Long
'    For lEachRow = LBound(arrData, 1) To UBound(arrData, 1)
'        If Not fProductNameExistsInProductNameMaster(CStr(arrData(lEachRow, dictColIndex("ProductProducer"))) _
'                                            , CStr(arrData(lEachRow, dictColIndex("ProductName")))) Then
'            fErr "��ҩƷ���� + ҩƷ���ơ���������ҩƷ���������С�" & vbCr & "�кţ�" & lEachRow + 1
'        End If
'    Next
    
    Call fSortDataInSheetSortSheetData(Me, dictColIndex("SalesManName"))
                                                
    If bErrMsgBox Then fMsgBox "[" & Me.Name & "]�� û�з��ִ���", vbInformation
Exit_Sub:
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
 
