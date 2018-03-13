VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtImportCZL2SalesCompSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub btnImportCZLSales2Comp_Click()
    Call subMain_ImportCZL2CompanySalesFile
End Sub

Private Sub btnSelectCZLFile2Comp_Click()
    Dim sHeader As String
    Dim sFile As String
    
    sHeader = Trim(Me.Range("rngHeader").Value)
    
    sFile = fSelectFileDialog(Trim(Me.Range("rngCZL2CompSalesFile").Value), , sHeader)
    If Len(sFile) > 0 Then Me.Range("rngCZL2CompSalesFile").Value = sFile
End Sub
