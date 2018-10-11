VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtRefund"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Enum Refund
    SalesCompany = 1
    SalesDate = 2
    ProductProducer = 3
    ProductName = 4
    ProductSeries = 5
    ProductUnit = 6
    Hospital = 7
    LotNum = 8
    Quantity = 9
    BidPrice = 10
    DueNetPrice = 11
    ActualNetPrice = 12
    PriceDeviation = 13
    AmountDeviation = 14
    [_first] = SalesCompany
    [_last] = AmountDeviation
End Enum

