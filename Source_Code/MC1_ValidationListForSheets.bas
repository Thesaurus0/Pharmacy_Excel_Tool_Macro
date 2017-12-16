Attribute VB_Name = "MC1_ValidationListForSheets"
Option Explicit
Option Base 1



Function fSetValidationListForAllSheets()
    Dim sProducerAddr As String
    sProducerAddr = fGetProducerMasterColumnAddress_Producer
    
    Call fSetValidationListForshtProductMaster_Producer(sProducerAddr)
    Call fSetValidationListForshtProductProducerReplace_Producer(sProducerAddr)

    
End Function

Function fGetProducerMasterColumnAddress_Producer() As String
    Dim sProducerCol As String
    Dim lProducerColMaxRow As Long
    Dim sSourceAddr As String
    
    lProducerColMaxRow = Rows.Count
    
    sProducerCol = fGetSpecifiedConfigCellValue(shtFileSpec _
                                        , "[Input File - PRODUCT_PRODUCER_MASTER]" _
                                        , "Column Index" _
                                        , "Column Tech Name=ProductProducer")

    sSourceAddr = "=" & shtProductProducerMaster.Range(sProducerCol & 2 & ":" & sProducerCol & lProducerColMaxRow).Address(external:=True)
    fGetProducerMasterColumnAddress_Producer = sSourceAddr
End Function

Function fSetValidationListForshtProductMaster_Producer(sValidationListAddr As String)
    Dim sProducerCol As String
    Dim lMaxRow As Long
    
    sProducerCol = fGetSpecifiedConfigCellValue(shtFileSpec _
                                            , "[Input File - PRODUCT_MASTER]" _
                                            , "Column Index" _
                                            , "Column Tech Name=ProductProducer")
    
    lMaxRow = shtProductMaster.Columns(sProducerCol).End(xlDown).Row + 100000
    
    Call fSetValidationListForRange(shtProductMaster.Range(sProducerCol & 2 & ":" & sProducerCol & lMaxRow) _
                                    , sValidationListAddr)
End Function

Function fSetValidationListForshtProductProducerReplace_Producer(sValidationListAddr As String)
    Dim sProducerCol As String
    Dim lMaxRow As Long
    
    sProducerCol = fGetSpecifiedConfigCellValue(shtFileSpec _
                                            , "[Input File - PRODUCER_REPLACE_SHEET]" _
                                            , "Column Index" _
                                            , "Column Tech Name=ToProducer")
    
    lMaxRow = shtProductProducerReplace.Columns(sProducerCol).End(xlDown).Row + 100000
    
    Call fSetValidationListForRange(shtProductProducerReplace.Range(sProducerCol & 2 & ":" & sProducerCol & lMaxRow) _
                                    , sValidationListAddr)
End Function

