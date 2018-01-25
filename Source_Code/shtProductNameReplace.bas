VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtProductNameReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Option Base 1

Private Sub btnProductNameReplaceValid_Click()
    Call fValidateSheet
End Sub


Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error GoTo exit_sub
    Application.ScreenUpdating = False
    
    Const ProducerCol = 1
    Const ProductNameCol = 3
    
    Dim rgIntersect As Range
    Set rgIntersect = Intersect(Target, Me.Columns(ProductNameCol))
    
    If Not rgIntersect Is Nothing Then
        If rgIntersect.Areas.Count > 1 Then GoTo exit_sub    'fErr "����ѡ���"
        If rgIntersect.Rows.Count <> 1 Then GoTo exit_sub
            
        Dim sProducer As String
        Dim sValidationListAddr As String
        
        sProducer = rgIntersect.Offset(0, -2).Value
        
        If fNzero(sProducer) Then
            Call fSetFilterForSheet(shtProductNameMaster, 1, sProducer)
            Call fCopyFilteredDataToRange(shtProductNameMaster, 2)
            
            sValidationListAddr = "=" & shtDataStage.Columns("A").Address(external:=True)
            'Call fSetValidationListForshtProductNameReplace_ProductName(sValidationListAddr, 3)
            Call fSetValidationListForRange(rgIntersect, sValidationListAddr)
        End If
    End If
    
exit_sub:
    fEnableExcelOptionsAll
    Application.ScreenUpdating = True
    
    If Err.Number <> 0 Then fMsgBox Err.Description
   ' End
End Sub

Function fValidateSheet()
    On Error GoTo exit_sub
    
    Dim lErrRowNo As Long, lErrColNo As Long
    Call fTrimAllCellsForSheet(Me)
    
    Dim arrData()
    Dim dictColIndex As Dictionary
    
    fInitialization
    gsRptID = "REPLACE_UNIFY_SALES_INFO"
    Call fReadSysConfig_InputTxtSheetFile
    
    Call fReadSheetDataByConfig("PRODUCT_NAME_REPLACE_SHEET", dictColIndex, arrData, , , , , shtProductNameReplace)
    
    Call fValidateDuplicateInArray(arrData, Array(dictColIndex("ProductProducer") _
                                                , dictColIndex("FromProductName") _
                                                , dictColIndex("ToProductName")) _
                , False, shtProductProducerMaster, 1, 1, "ҩƷ�������� + ԭʼ�ļ�ҩƷ���� + �滻Ϊ")
    Call fValidateBlankInArray(arrData, dictColIndex("ProductProducer"), shtProductNameReplace, 1, 1, "ҩƷ����")
    Call fValidateBlankInArray(arrData, dictColIndex("FromProductName"), shtProductNameReplace, 1, 1, "ҩƷ����")
    Call fValidateBlankInArray(arrData, dictColIndex("ToProductName"), shtProductNameReplace, 1, 1, "ҩƷ���")
    
    Call fCheckIfProductNameExistsInProductNameMaster(arrData, dictColIndex("ProductProducer"), dictColIndex("ToProductName"), "�滻Ϊ", lErrRowNo, lErrColNo)
    
    fMsgBox "[" & Me.Name & "]�� û�з��ִ���", vbInformation
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
    If lErrRowNo > 0 Then
        Application.Goto Me.Cells(lErrRowNo, lErrColNo) ', True
    End If

End Function

