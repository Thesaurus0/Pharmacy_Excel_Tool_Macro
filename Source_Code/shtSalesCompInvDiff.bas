VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtSalesCompInvDiff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Enum SCompInvDiff
    SalesCompany = 1
    ProductProducer = 2
    ProductName = 3
    ProductSeries = 4
    LotNum = 5
    InformedQty = 6
    CalculatedQty = 7
    DiffQty = 8
    ProductUnit = 9
End Enum

'Function fProdProducerCol() As Integer
'    fProdProducerCol = SCompInvDiff.ProductProducer
'End Function
