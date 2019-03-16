VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtMenuCompInvt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Option Base 1

Private Sub btnBatchImportSaleInfoFiles_Click()
    'Call fImportAllSalesInfoFiles
    Call subMain_ImportInventoryFiles
    
End Sub

Private Sub btnImportCommon_Click()
    subMain_ImportInventoryFiles_Common
End Sub

Private Sub btnSelect_CZL_Click()
    Call fOpenFileSelectDialogAndSetToSheetRangeForCompany("CZL")
End Sub

Private Sub btnSelect_GKYX_Click()
    Call fOpenFileSelectDialogAndSetToSheetRangeForCompany("GKYX")
End Sub

Private Sub btnSelect_FSGK_Click()
    Call fOpenFileSelectDialogAndSetToSheetRangeForCompany("FSGK")
End Sub

Private Sub btnSelect_GZGK_Click()
    Call fOpenFileSelectDialogAndSetToSheetRangeForCompany("GZGK")
End Sub

Private Sub btnSelect_GZHR_Click()
    Call fOpenFileSelectDialogAndSetToSheetRangeForCompany("GZHR")
End Sub

Private Sub btnSelect_HR_Click()
    Call fOpenFileSelectDialogAndSetToSheetRangeForCompany("HR")
End Sub

Private Sub btnSelect_PW_Click()
    Call fOpenFileSelectDialogAndSetToSheetRangeForCompany("PW")
End Sub

Private Sub btnSelect_SYY_Click()
    Call fOpenFileSelectDialogAndSetToSheetRangeForCompany("SYY")
End Sub

Private Sub btnSelect_SYYDZ_Click()
    Call fOpenFileSelectDialogAndSetToSheetRangeForCompany("SYYDZ")
End Sub

Private Sub btnSelect_TY_Click()
    Call fOpenFileSelectDialogAndSetToSheetRangeForCompany("TY")
End Sub

Private Sub btnSelect_XT_Click()
    Call fOpenFileSelectDialogAndSetToSheetRangeForCompany("XT")
End Sub

Private Sub btnSelect_ZHHR_Click()
    Call fOpenFileSelectDialogAndSetToSheetRangeForCompany("ZHHR")
End Sub

Private Sub btnSelect_ZSY_Click()
    Call fOpenFileSelectDialogAndSetToSheetRangeForCompany("ZSY")
End Sub

Private Sub btnSelectFile_Common_Click()
    Dim sFile As String
    
    sFile = fSelectFileDialog(Trim(Me.Range("rngInventoryFilePathComm").Value), , "请选择销售文件：" & Me.cbbCompanyList.Value)
    If Len(sFile) > 0 Then Me.Range("rngInventoryFilePathComm").Value = sFile
End Sub

Private Sub btnSelectGY_Click()
    Call fOpenFileSelectDialogAndSetToSheetRangeForCompany("GY")
End Sub

Function fOpenFileSelectDialogAndSetToSheetRangeForCompany(sCompany As String)
    Dim sHeader As String
    
    sHeader = LeftB(Me.Range("rngHeader_" & sCompany).Value, LenB(Me.Range("rngHeader_" & sCompany).Value) - 2)
    Call fOpenFileSelectDialogAndSetToSheetRange("rngSalesFilePath_" & sCompany, , sHeader, Me)
End Function


Sub sub_Initialize_CompanyListCombobox_Inventory()
    'Me.Visible = xlSheetVisible
    
    Me.cbbCompanyList.Clear
    
    Dim dictComp As Dictionary
    Set dictComp = fGetCompanyListCommon
    
    Dim i As Integer
    For i = 0 To dictComp.Count - 1
        Me.cbbCompanyList.AddItem dictComp.Keys(i)
    Next
    
    Set dictComp = Nothing
    
   ' Me.cbbCompanyList.Activate
End Sub

