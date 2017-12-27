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
    On Error GoTo exit_sub
    
    Dim arrData()
    Dim dictColIndex As Dictionary
    
    fInitialization
    gsRptID = "REPLACE_UNIFY_SALES_INFO"
    Call fReadSysConfig_InputTxtSheetFile
    
    Call fReadSheetDataByConfig("PRODUCT_NAME_REPLACE_SHEET", dictColIndex, arrData, , , , , shtProductNameReplace)
    
    Call fValidateDuplicateInArray(arrData, Array(dictColIndex("ProductProducer") _
                                                , dictColIndex("FromProductName") _
                                                , dictColIndex("ToProductName")) _
                , False, shtProductProducerMaster, 1, 1, "药品生产厂家 + 原始文件药品名称 + 替换为")
    Call fValidateBlankInArray(arrData, dictColIndex("ProductProducer"), shtProductNameReplace, 1, 1, "药品厂家")
    Call fValidateBlankInArray(arrData, dictColIndex("FromProductName"), shtProductNameReplace, 1, 1, "药品名称")
    Call fValidateBlankInArray(arrData, dictColIndex("ToProductName"), shtProductNameReplace, 1, 1, "药品规格")
    
    Dim lEachRow As Long
    For lEachRow = LBound(arrData, 1) To UBound(arrData, 1)
        If Not fProductNameExistsInProductNameMaster(CStr(arrData(lEachRow, dictColIndex("ProductProducer"))) _
                                            , CStr(arrData(lEachRow, dictColIndex("ToProductName")))) Then
            fErr "【药品厂家 + 药品名称】不存在于药品名称主表中。" & vbCr & "行号：" & lEachRow + 1
        End If
    Next
    
    fMsgBox "没有发现错误", vbInformation
exit_sub:
    Set dictColIndex = Nothing
    fEnableExcelOptionsAll
    End
End Sub


Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error GoTo exit_sub
    Application.ScreenUpdating = False
    
    Const ProducerCol = 1
    Const ProductNameCol = 3
    
    Dim rgIntersect As Range
    Set rgIntersect = Intersect(Target, Me.Columns(ProductNameCol))
    
    If Not rgIntersect Is Nothing Then
        If rgIntersect.Areas.Count > 1 Then fErr "不能选多个"
            
        Dim sProducer As String
        Dim sValidationListAddr As String
        
        sProducer = rgIntersect.Offset(0, -2).Value
        
        If fNzero(sProducer) Then
            Call fSetFilterForSheet(shtProductNameMaster, 1, sProducer)
            Call fCopyFilteredDataToRange(shtProductNameMaster, 2)
            
            sValidationListAddr = "=" & shtDataStage.Columns("A").Address(external:=True)
            Call fSetValidationListForshtProductNameReplace_ProductName(sValidationListAddr, 3)
        End If
    End If
    
exit_sub:
    fEnableExcelOptionsAll
    Application.ScreenUpdating = True
   ' End
End Sub
