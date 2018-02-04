Attribute VB_Name = "MC1_ValidationListForSheets"
Option Explicit
Option Base 1

Function fSetValidationListForAllSheets()
    '============== SalesCompany ========================================
    Call fSetValidationListForshtFirstLevelCommission_SalesCompany("=rngStaticSalesCompanyNames")
    Call fSetValidationListForshtSecondLevelCommission_SalesCompany("=rngStaticSalesCompanyNames")
    Call fSetValidationListForshtSalesManCommConfig_SalesCompany("=rngStaticSalesCompanyNames")
    Call fSetValidationListForshtCompanyNameReplace_SalesCompany("=rngStaticSalesCompanyNames")
    '----------------------------------------------------------------------------------------

    '============== Hospital ========================================
    Dim sHospitalAddr As String
    sHospitalAddr = fGetHospitalMasterColumnAddress_Hospital

    Call fSetValidationListForshtHospitalReplace_Hospital(sHospitalAddr)
    Call fSetValidationListForshtSalesManCommConfig_Hospital(sHospitalAddr)
    Call fSetValidationListForshtSecondLevelCommission_Hospital(sHospitalAddr)
    '----------------------------------------------------------------------------------------

    '============== producer ========================================
    Dim sProducerAddr As String
    sProducerAddr = fGetProducerMasterColumnAddress_Producer

    Call fSetValidationListForshtProductMaster_Producer(sProducerAddr)
    Call fSetValidationListForshtProductNameMaster_Producer(sProducerAddr)
    Call fSetValidationListForshtSalesManCommConfig_Producer(sProducerAddr)
    Call fSetValidationListForshtFirstLevelCommission_Producer(sProducerAddr)
    Call fSetValidationListForshtSecondLevelCommission_Producer(sProducerAddr)
    Call fSetValidationListForshtProductProducerReplace_Producer(sProducerAddr)
    Call fSetValidationListForshtProductNameReplace_Producer(sProducerAddr)
    Call fSetValidationListForshtProductSeriesReplace_Producer(sProducerAddr)
    Call fSetValidationListForshtProductUnitRatio_Producer(sProducerAddr)
    Call fSetValidationListForshtSelfPurchaseOrder_Producer(sProducerAddr)
    Call fSetValidationListForshtSelfSalesOrder_Producer(sProducerAddr)
    Call fSetValidationListForshtSelfInventory_Producer(sProducerAddr)
    '----------------------------------------------------------------------------------------

    '============== productName ========================================
    Dim sProductNameAddr As String
    sProductNameAddr = fGetProductNameMasterColumnAddress_ProductName

    Call fSetValidationListForshtProductMaster_ProductName(sProductNameAddr)
    Call fSetValidationListForshtSalesManCommConfig_ProductName(sProductNameAddr)
    Call fSetValidationListForshtFirstLevelCommission_ProductName(sProductNameAddr)
    Call fSetValidationListForshtSecondLevelCommission_ProductName(sProductNameAddr)
    Call fSetValidationListForshtProductNameReplace_ProductName(sProductNameAddr)
    Call fSetValidationListForshtProductSeriesReplace_ProductName(sProductNameAddr)
    Call fSetValidationListForshtProductUnitRatio_ProductName(sProductNameAddr)
    Call fSetValidationListForshtSelfSalesOrder_ProductName(sProductNameAddr)
    '----------------------------------------------------------------------------------------

    '============== ProductSeries ========================================
    Dim sProductSeriesAddr As String
    sProductSeriesAddr = fGetProductSeriesMasterColumnAddress_ProductSeries

    Call fSetValidationListForshtSalesManCommConfig_ProductSeries(sProductSeriesAddr)
    Call fSetValidationListForshtFirstLevelCommission_ProductSeries(sProductSeriesAddr)
    Call fSetValidationListForshtSecondLevelCommission_ProductSeries(sProductSeriesAddr)
    Call fSetValidationListForshtProductSeriesReplace_ProductSeries(sProductSeriesAddr)
    Call fSetValidationListForshtProductUnitRatio_ProductSeries(sProductSeriesAddr)
    Call fSetValidationListForshtSelfSalesOrder_ProductSeries(sProductSeriesAddr)
    '----------------------------------------------------------------------------------------

    '============== ProductUnit ========================================
    Dim sProductUnitAddr As String
    sProductUnitAddr = fGetProductUnitMasterColumnAddress_ProductUnit

    Call fSetValidationListForshtProductUnitRatio_ProductUnit(sProductUnitAddr)
    Call fSetValidationListForshtSelfSalesOrder_ProductUnit(sProductUnitAddr)
    '----------------------------------------------------------------------------------------

    '============== salesman ========================================
    Dim sSalesManAddr As String
    sSalesManAddr = fGetSalesManMasterColumnAddress_SalesMan
    Call fSetValidationListForshtSalesManCommConfig_SalesMan(sSalesManAddr)
    '----------------------------------------------------------------------------------------
    
End Function

'============== producer ========================================
Function fGetProducerMasterColumnAddress_Producer() As String
    Dim sProducerCol As String
    Dim lProducerColMaxRow As Long
    Dim sSourceAddr As String
    
    lProducerColMaxRow = Rows.Count
    
    sProducerCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - PRODUCT_PRODUCER_MASTER]" _
                                        , "Column Index", "Column Tech Name=ProductProducer")

    sSourceAddr = "=" & shtProductProducerMaster.Range(sProducerCol & 2 & ":" & sProducerCol & lProducerColMaxRow).Address(external:=True)
    fGetProducerMasterColumnAddress_Producer = sSourceAddr
End Function

Function fSetValidationListForshtProductNameMaster_Producer(sValidationListAddr As String)
    Dim sProducerCol As String
    Dim lMaxRow As Long
    
    sProducerCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - PRODUCT_NAME_MASTER]" _
                                            , "Column Index", "Column Tech Name=ProductProducer")
    
    lMaxRow = shtProductNameMaster.Columns(sProducerCol).End(xlDown).Row + 100000
    If lMaxRow > Rows.Count Then lMaxRow = 100000
    Call fSetValidationListForRange(shtProductNameMaster.Range(sProducerCol & 2 & ":" & sProducerCol & lMaxRow) _
                                    , sValidationListAddr)
End Function
Function fSetValidationListForshtProductMaster_Producer(sValidationListAddr As String)
    Dim sProducerCol As String
    Dim lMaxRow As Long
    
    sProducerCol = fGetSpecifiedConfigCellValue(shtFileSpec _
                                            , "[Input File - PRODUCT_MASTER]" _
                                            , "Column Index" _
                                            , "Column Tech Name=ProductProducer")
    
    lMaxRow = shtProductMaster.Columns(sProducerCol).End(xlDown).Row + 100000
    If lMaxRow > Rows.Count Then lMaxRow = 100000
    Call fSetValidationListForRange(shtProductMaster.Range(sProducerCol & 2 & ":" & sProducerCol & lMaxRow) _
                                    , sValidationListAddr)
End Function


Function fSetValidationListForshtSalesManCommConfig_Producer(sValidationListAddr As String)
    Dim sTargetCol As String
    Dim lMaxRow As Long
    
    sTargetCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - SALESMAN_COMMISSION_CONFIG]" _
                                            , "Column Index", "Column Tech Name=ProductProducer")
    
    lMaxRow = shtSalesManCommConfig.Columns(sTargetCol).End(xlDown).Row + 100000
    If lMaxRow > Rows.Count Then lMaxRow = 100000
    Call fSetValidationListForRange(shtSalesManCommConfig.Range(sTargetCol & 2 & ":" & sTargetCol & lMaxRow) _
                                    , sValidationListAddr)
End Function

Function fSetValidationListForshtFirstLevelCommission_Producer(sValidationListAddr As String)
    Dim sTargetCol As String
    Dim lMaxRow As Long

    sTargetCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - FIRST_LEVEL_COMMISSION]" _
                                            , "Column Index", "Column Tech Name=ProductProducer")

    lMaxRow = shtFirstLevelCommission.Columns(sTargetCol).End(xlDown).Row + 100000
    If lMaxRow > Rows.Count Then lMaxRow = 100000
    Call fSetValidationListForRange(shtFirstLevelCommission.Range(sTargetCol & 2 & ":" & sTargetCol & lMaxRow) _
                                    , sValidationListAddr)
End Function
Function fSetValidationListForshtSecondLevelCommission_Producer(sValidationListAddr As String)
    Dim sTargetCol As String
    Dim lMaxRow As Long
    
    sTargetCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - SECOND_LEVEL_COMMISSION]" _
                                            , "Column Index", "Column Tech Name=ProductProducer")
    
    lMaxRow = shtSecondLevelCommission.Columns(sTargetCol).End(xlDown).Row + 100000
    If lMaxRow > Rows.Count Then lMaxRow = 100000
    Call fSetValidationListForRange(shtSecondLevelCommission.Range(sTargetCol & 2 & ":" & sTargetCol & lMaxRow) _
                                    , sValidationListAddr)
End Function


Function fSetValidationListForshtProductProducerReplace_Producer(sValidationListAddr As String)
    Dim sProducerCol As String
    Dim lMaxRow As Long
    
    sProducerCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - PRODUCER_REPLACE_SHEET]" _
                                            , "Column Index", "Column Tech Name=ToProducer")
    
    lMaxRow = shtProductProducerReplace.Columns(sProducerCol).End(xlDown).Row + 100000
    If lMaxRow > Rows.Count Then lMaxRow = 100000
    Call fSetValidationListForRange(shtProductProducerReplace.Range(sProducerCol & 2 & ":" & sProducerCol & lMaxRow) _
                                    , sValidationListAddr)
End Function
Function fSetValidationListForshtProductNameReplace_Producer(sValidationListAddr As String)
    Dim sTargetCol As String
    Dim lMaxRow As Long
    
    sTargetCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - PRODUCT_NAME_REPLACE_SHEET]" _
                                            , "Column Index", "Column Tech Name=ProductProducer")
    
    lMaxRow = shtProductNameReplace.Columns(sTargetCol).End(xlDown).Row + 100000
    If lMaxRow > Rows.Count Then lMaxRow = 100000
    Call fSetValidationListForRange(shtProductNameReplace.Range(sTargetCol & 2 & ":" & sTargetCol & lMaxRow) _
                                    , sValidationListAddr)
End Function
Function fSetValidationListForshtProductSeriesReplace_Producer(sValidationListAddr As String)
    Dim sTargetCol As String
    Dim lMaxRow As Long
    
    sTargetCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - PRODUCT_SERIES_REPLACE_SHEET]" _
                                            , "Column Index", "Column Tech Name=ProductProducer")
    
    lMaxRow = shtProductSeriesReplace.Columns(sTargetCol).End(xlDown).Row + 100000
    If lMaxRow > Rows.Count Then lMaxRow = 100000
    Call fSetValidationListForRange(shtProductSeriesReplace.Range(sTargetCol & 2 & ":" & sTargetCol & lMaxRow) _
                                    , sValidationListAddr)
End Function
Function fSetValidationListForshtProductUnitRatio_Producer(sValidationListAddr As String)
    Dim sTargetCol As String
    Dim lMaxRow As Long
    
    sTargetCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - PRODUCT_UNIT_RATIO_SHEET]" _
                                            , "Column Index", "Column Tech Name=ProductProducer")
    
    lMaxRow = shtProductUnitRatio.Columns(sTargetCol).End(xlDown).Row + 100000
    If lMaxRow > Rows.Count Then lMaxRow = 100000
    Call fSetValidationListForRange(shtProductUnitRatio.Range(sTargetCol & 2 & ":" & sTargetCol & lMaxRow) _
                                    , sValidationListAddr)
End Function
Function fSetValidationListForshtSelfSalesOrder_Producer(sValidationListAddr As String)
    Dim sTargetCol As String
    Dim lMaxRow As Long
    
    sTargetCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - SELF_SALES_ORDER]" _
                                            , "Column Index", "Column Tech Name=ProductProducer")
    
    lMaxRow = shtSelfSalesOrder.Columns(sTargetCol).End(xlDown).Row + 100000
    If lMaxRow > Rows.Count Then lMaxRow = 100000
    Call fSetValidationListForRange(shtSelfSalesOrder.Range(sTargetCol & 2 & ":" & sTargetCol & lMaxRow) _
                                    , sValidationListAddr)
End Function
Function fSetValidationListForshtSelfPurchaseOrder_Producer(sValidationListAddr As String)
    Dim sTargetCol As String
    Dim lMaxRow As Long
    
    sTargetCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - SELF_PURCHASE_ORDER]" _
                                            , "Column Index", "Column Tech Name=ProductProducer")
    
    lMaxRow = shtSelfPurchaseOrder.Columns(sTargetCol).End(xlDown).Row + 100000
    If lMaxRow > Rows.Count Then lMaxRow = 100000
    Call fSetValidationListForRange(shtSelfPurchaseOrder.Range(sTargetCol & 2 & ":" & sTargetCol & lMaxRow) _
                                    , sValidationListAddr)
End Function
Function fSetValidationListForshtSelfInventory_Producer(sValidationListAddr As String)
    Dim sTargetCol As String
    Dim lMaxRow As Long
    
    sTargetCol = "A"
    
    lMaxRow = shtSelfInventory.Columns(sTargetCol).End(xlDown).Row + 100000
    If lMaxRow > Rows.Count Then lMaxRow = 100000
    Call fSetValidationListForRange(shtSelfInventory.Range(sTargetCol & 2 & ":" & sTargetCol & lMaxRow) _
                                    , sValidationListAddr)
End Function
'----------------------------------------------------------------------------------------

'============== SalesCompany ========================================
Function fSetValidationListForshtFirstLevelCommission_SalesCompany(sValidationListAddr As String)
    Dim sTargetCol As String
    Dim lMaxRow As Long
    
    sTargetCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - FIRST_LEVEL_COMMISSION]" _
                                            , "Column Index", "Column Tech Name=SalesCompany")
    
    lMaxRow = shtFirstLevelCommission.Columns(sTargetCol).End(xlDown).Row + 100000
    If lMaxRow > Rows.Count Then lMaxRow = 100000
    Call fSetValidationListForRange(shtFirstLevelCommission.Range(sTargetCol & 2 & ":" & sTargetCol & lMaxRow) _
                                    , sValidationListAddr)
End Function

Function fSetValidationListForshtSecondLevelCommission_SalesCompany(sValidationListAddr As String)
    Dim sTargetCol As String
    Dim lMaxRow As Long
    
    sTargetCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - SECOND_LEVEL_COMMISSION]" _
                                            , "Column Index", "Column Tech Name=SalesCompany")
    
    lMaxRow = shtSecondLevelCommission.Columns(sTargetCol).End(xlDown).Row + 100000
    If lMaxRow > Rows.Count Then lMaxRow = 100000
    Call fSetValidationListForRange(shtSecondLevelCommission.Range(sTargetCol & 2 & ":" & sTargetCol & lMaxRow) _
                                    , sValidationListAddr)
End Function

Function fSetValidationListForshtSalesManCommConfig_SalesCompany(sValidationListAddr As String)
    Dim sTargetCol As String
    Dim lMaxRow As Long
    
    sTargetCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - SALESMAN_COMMISSION_CONFIG]" _
                                            , "Column Index", "Column Tech Name=SalesCompany")
    
    lMaxRow = shtSalesManCommConfig.Columns(sTargetCol).End(xlDown).Row + 100000
    If lMaxRow > Rows.Count Then lMaxRow = 100000
    Call fSetValidationListForRange(shtSalesManCommConfig.Range(sTargetCol & 2 & ":" & sTargetCol & lMaxRow) _
                                    , sValidationListAddr)
End Function

Function fSetValidationListForshtCompanyNameReplace_SalesCompany(sValidationListAddr As String)
    Dim sTargetCol As String
    Dim lMaxRow As Long
    
    sTargetCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - COMPANY_NAME_REPLACE_SHEET]" _
                                            , "Column Index", "Column Tech Name=ToCompanyName")
    
    lMaxRow = shtCompanyNameReplace.Columns(sTargetCol).End(xlDown).Row + 10000
    If lMaxRow > Rows.Count Then lMaxRow = 100000
    Call fSetValidationListForRange(shtCompanyNameReplace.Range(sTargetCol & 2 & ":" & sTargetCol & lMaxRow) _
                                    , sValidationListAddr)
End Function
'----------------------------------------------------------------------------------------


'============== SalesCompany ========================================
Function fGetHospitalMasterColumnAddress_Hospital() As String
    Dim sSourceCol As String
    Dim lMaxRow As Long
    Dim sSourceAddr As String
    
    lMaxRow = Rows.Count
    sSourceCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - HOSPITAL_MASTER]" _
                                        , "Column Index", "Column Tech Name=Hospital")

    sSourceAddr = "=" & shtHospital.Range(sSourceCol & 2 & ":" & sSourceCol & lMaxRow).Address(external:=True)
    fGetHospitalMasterColumnAddress_Hospital = sSourceAddr
End Function

Function fSetValidationListForshtHospitalReplace_Hospital(sValidationListAddr As String)
    Dim sTargetCol As String
    Dim lMaxRow As Long
    
    sTargetCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - HOSPITAL_REPLACE_SHEET]" _
                                            , "Column Index", "Column Tech Name=ToHospital")
    
    lMaxRow = shtHospitalReplace.Columns(sTargetCol).End(xlDown).Row + 100000
    If lMaxRow > Rows.Count Then lMaxRow = 100000
    Call fSetValidationListForRange(shtHospitalReplace.Range(sTargetCol & 2 & ":" & sTargetCol & lMaxRow) _
                                    , sValidationListAddr)
End Function
Function fSetValidationListForshtSalesManCommConfig_Hospital(sValidationListAddr As String)
    Dim sTargetCol As String
    Dim lMaxRow As Long
    
    sTargetCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - SALESMAN_COMMISSION_CONFIG]" _
                                            , "Column Index", "Column Tech Name=Hospital")
    
    lMaxRow = shtSalesManCommConfig.Columns(sTargetCol).End(xlDown).Row + 100000
    If lMaxRow > Rows.Count Then lMaxRow = 100000
    Call fSetValidationListForRange(shtSalesManCommConfig.Range(sTargetCol & 2 & ":" & sTargetCol & lMaxRow) _
                                    , sValidationListAddr)
End Function

'Function fSetValidationListForshtFirstLevelCommission_Hospital(sValidationListAddr As String)
'    Dim sTargetCol As String
'    Dim lMaxRow As Long
'
'    sTargetCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - FIRST_LEVEL_COMMISSION]" _
'                                            , "Column Index", "Column Tech Name=Hospital")
'
'    lMaxRow = shtFirstLevelCommission.Columns(sTargetCol).End(xlDown).Row + 100000
'
'    Call fSetValidationListForRange(shtFirstLevelCommission.Range(sTargetCol & 2 & ":" & sTargetCol & lMaxRow) _
'                                    , sValidationListAddr)
'End Function
Function fSetValidationListForshtSecondLevelCommission_Hospital(sValidationListAddr As String)
    Dim sTargetCol As String
    Dim lMaxRow As Long
    
    sTargetCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - SECOND_LEVEL_COMMISSION]" _
                                            , "Column Index", "Column Tech Name=Hospital")
    
    lMaxRow = shtSecondLevelCommission.Columns(sTargetCol).End(xlDown).Row + 100000
    If lMaxRow > Rows.Count Then lMaxRow = 100000
    Call fSetValidationListForRange(shtSecondLevelCommission.Range(sTargetCol & 2 & ":" & sTargetCol & lMaxRow) _
                                    , sValidationListAddr)
End Function
'----------------------------------------------------------------------------------------


'============== productName ========================================
Function fGetProductNameMasterColumnAddress_ProductName() As String
    Dim sSourceCol As String
    Dim lColMaxRow As Long
    Dim sSourceAddr As String
    
    lColMaxRow = Rows.Count
    sSourceCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - PRODUCT_NAME_MASTER]" _
                                        , "Column Index", "Column Tech Name=ProductName")

    sSourceAddr = "=" & shtProductNameMaster.Range(sSourceCol & 2 & ":" & sSourceCol & lColMaxRow).Address(external:=True)
    fGetProductNameMasterColumnAddress_ProductName = sSourceAddr
End Function

Function fSetValidationListForshtProductMaster_ProductName(sValidationListAddr As String)
    Dim sProducerCol As String
    Dim lMaxRow As Long
    
    sProducerCol = fGetSpecifiedConfigCellValue(shtFileSpec _
                                            , "[Input File - PRODUCT_MASTER]" _
                                            , "Column Index" _
                                            , "Column Tech Name=ProductName")
    
    lMaxRow = shtProductMaster.Columns(sProducerCol).End(xlDown).Row + 100000
    If lMaxRow > Rows.Count Then lMaxRow = 100000
    Call fSetValidationListForRange(shtProductMaster.Range(sProducerCol & 2 & ":" & sProducerCol & lMaxRow) _
                                    , sValidationListAddr)
End Function


Function fSetValidationListForshtSalesManCommConfig_ProductName(sValidationListAddr As String)
    Dim sTargetCol As String
    Dim lMaxRow As Long
    
    sTargetCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - SALESMAN_COMMISSION_CONFIG]" _
                                            , "Column Index", "Column Tech Name=ProductName")
    
    lMaxRow = shtSalesManCommConfig.Columns(sTargetCol).End(xlDown).Row + 100000
    If lMaxRow > Rows.Count Then lMaxRow = 100000
    Call fSetValidationListForRange(shtSalesManCommConfig.Range(sTargetCol & 2 & ":" & sTargetCol & lMaxRow) _
                                    , sValidationListAddr)
End Function

Function fSetValidationListForshtFirstLevelCommission_ProductName(sValidationListAddr As String)
    Dim sTargetCol As String
    Dim lMaxRow As Long

    sTargetCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - FIRST_LEVEL_COMMISSION]" _
                                            , "Column Index", "Column Tech Name=ProductName")

    lMaxRow = shtFirstLevelCommission.Columns(sTargetCol).End(xlDown).Row + 100000
    If lMaxRow > Rows.Count Then lMaxRow = 100000
    Call fSetValidationListForRange(shtFirstLevelCommission.Range(sTargetCol & 2 & ":" & sTargetCol & lMaxRow) _
                                    , sValidationListAddr)
End Function
Function fSetValidationListForshtSecondLevelCommission_ProductName(sValidationListAddr As String)
    Dim sTargetCol As String
    Dim lMaxRow As Long
    
    sTargetCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - SECOND_LEVEL_COMMISSION]" _
                                            , "Column Index", "Column Tech Name=ProductName")
    
    lMaxRow = shtSecondLevelCommission.Columns(sTargetCol).End(xlDown).Row + 100000
    If lMaxRow > Rows.Count Then lMaxRow = 100000
    Call fSetValidationListForRange(shtSecondLevelCommission.Range(sTargetCol & 2 & ":" & sTargetCol & lMaxRow) _
                                    , sValidationListAddr)
End Function

Function fSetValidationListForshtProductNameReplace_ProductName(sValidationListAddr As String, Optional iCol As Integer = 0)
    Dim sTargetCol As String
    Dim lMaxRow As Long
    
    If iCol = 0 Then
        sTargetCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - PRODUCT_NAME_REPLACE_SHEET]" _
                                            , "Column Index", "Column Tech Name=ToProductName")
    Else
        sTargetCol = fNum2Letter(iCol)
    End If
    
    lMaxRow = shtProductNameReplace.Columns(sTargetCol).End(xlDown).Row + 100000
    If lMaxRow > Rows.Count Then lMaxRow = 100000
    Call fSetValidationListForRange(shtProductNameReplace.Range(sTargetCol & 2 & ":" & sTargetCol & lMaxRow) _
                                    , sValidationListAddr)
End Function
Function fSetValidationListForshtProductSeriesReplace_ProductName(sValidationListAddr As String)
    Dim sTargetCol As String
    Dim lMaxRow As Long
    
    sTargetCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - PRODUCT_SERIES_REPLACE_SHEET]" _
                                            , "Column Index", "Column Tech Name=ProductName")
    
    lMaxRow = shtProductSeriesReplace.Columns(sTargetCol).End(xlDown).Row + 100000
    If lMaxRow > Rows.Count Then lMaxRow = 100000
    Call fSetValidationListForRange(shtProductSeriesReplace.Range(sTargetCol & 2 & ":" & sTargetCol & lMaxRow) _
                                    , sValidationListAddr)
End Function
Function fSetValidationListForshtProductUnitRatio_ProductName(sValidationListAddr As String)
    Dim sTargetCol As String
    Dim lMaxRow As Long
    
    sTargetCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - PRODUCT_UNIT_RATIO_SHEET]" _
                                            , "Column Index", "Column Tech Name=ProductName")
    
    lMaxRow = shtProductUnitRatio.Columns(sTargetCol).End(xlDown).Row + 100000
    If lMaxRow > Rows.Count Then lMaxRow = 100000
    Call fSetValidationListForRange(shtProductUnitRatio.Range(sTargetCol & 2 & ":" & sTargetCol & lMaxRow) _
                                    , sValidationListAddr)
End Function
Function fSetValidationListForshtSelfSalesOrder_ProductName(sValidationListAddr As String)
    Dim sTargetCol As String
    Dim lMaxRow As Long
    
    sTargetCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - SELF_SALES_ORDER]" _
                                            , "Column Index", "Column Tech Name=ProductName")
    
    lMaxRow = shtSelfSalesOrder.Columns(sTargetCol).End(xlDown).Row + 100000
    If lMaxRow > Rows.Count Then lMaxRow = 100000
    Call fSetValidationListForRange(shtSelfSalesOrder.Range(sTargetCol & 2 & ":" & sTargetCol & lMaxRow) _
                                    , sValidationListAddr)
End Function

'----------------------------------------------------------------------------------------

'============== ProductSeries ========================================
Function fGetProductSeriesMasterColumnAddress_ProductSeries() As String
    Dim sSourceCol As String
    Dim lColMaxRow As Long
    Dim sSourceAddr As String
    
    lColMaxRow = Rows.Count
    sSourceCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - PRODUCT_MASTER]" _
                                        , "Column Index", "Column Tech Name=ProductSeries")

    sSourceAddr = "=" & shtProductMaster.Range(sSourceCol & 2 & ":" & sSourceCol & lColMaxRow).Address(external:=True)
    fGetProductSeriesMasterColumnAddress_ProductSeries = sSourceAddr
End Function

Function fSetValidationListForshtSalesManCommConfig_ProductSeries(sValidationListAddr As String)
    Dim sTargetCol As String
    Dim lMaxRow As Long
    
    sTargetCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - SALESMAN_COMMISSION_CONFIG]" _
                                            , "Column Index", "Column Tech Name=ProductSeries")
    
    lMaxRow = shtSalesManCommConfig.Columns(sTargetCol).End(xlDown).Row + 100000
    If lMaxRow > Rows.Count Then lMaxRow = 100000
    Call fSetValidationListForRange(shtSalesManCommConfig.Range(sTargetCol & 2 & ":" & sTargetCol & lMaxRow) _
                                    , sValidationListAddr)
End Function

Function fSetValidationListForshtFirstLevelCommission_ProductSeries(sValidationListAddr As String)
    Dim sTargetCol As String
    Dim lMaxRow As Long

    sTargetCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - FIRST_LEVEL_COMMISSION]" _
                                            , "Column Index", "Column Tech Name=ProductSeries")

    lMaxRow = shtFirstLevelCommission.Columns(sTargetCol).End(xlDown).Row + 100000
    If lMaxRow > Rows.Count Then lMaxRow = 100000
    Call fSetValidationListForRange(shtFirstLevelCommission.Range(sTargetCol & 2 & ":" & sTargetCol & lMaxRow) _
                                    , sValidationListAddr)
End Function
Function fSetValidationListForshtSecondLevelCommission_ProductSeries(sValidationListAddr As String)
    Dim sTargetCol As String
    Dim lMaxRow As Long
    
    sTargetCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - SECOND_LEVEL_COMMISSION]" _
                                            , "Column Index", "Column Tech Name=ProductSeries")
    
    lMaxRow = shtSecondLevelCommission.Columns(sTargetCol).End(xlDown).Row + 100000
    If lMaxRow > Rows.Count Then lMaxRow = 100000
    Call fSetValidationListForRange(shtSecondLevelCommission.Range(sTargetCol & 2 & ":" & sTargetCol & lMaxRow) _
                                    , sValidationListAddr)
End Function

Function fSetValidationListForshtProductSeriesReplace_ProductSeries(sValidationListAddr As String)
    Dim sTargetCol As String
    Dim lMaxRow As Long
    
    sTargetCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - PRODUCT_SERIES_REPLACE_SHEET]" _
                                            , "Column Index", "Column Tech Name=ToProductSeries")
    
    lMaxRow = shtProductSeriesReplace.Columns(sTargetCol).End(xlDown).Row + 100000
    If lMaxRow > Rows.Count Then lMaxRow = 100000
    Call fSetValidationListForRange(shtProductSeriesReplace.Range(sTargetCol & 2 & ":" & sTargetCol & lMaxRow) _
                                    , sValidationListAddr)
End Function
Function fSetValidationListForshtProductUnitRatio_ProductSeries(sValidationListAddr As String)
    Dim sTargetCol As String
    Dim lMaxRow As Long
    
    sTargetCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - PRODUCT_UNIT_RATIO_SHEET]" _
                                            , "Column Index", "Column Tech Name=ProductSeries")
    
    lMaxRow = shtProductUnitRatio.Columns(sTargetCol).End(xlDown).Row + 100000
    If lMaxRow > Rows.Count Then lMaxRow = 100000
    Call fSetValidationListForRange(shtProductUnitRatio.Range(sTargetCol & 2 & ":" & sTargetCol & lMaxRow) _
                                    , sValidationListAddr)
End Function
Function fSetValidationListForshtSelfSalesOrder_ProductSeries(sValidationListAddr As String)
    Dim sTargetCol As String
    Dim lMaxRow As Long
    
    sTargetCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - SELF_SALES_ORDER]" _
                                            , "Column Index", "Column Tech Name=ProductSeries")
    
    lMaxRow = shtSelfSalesOrder.Columns(sTargetCol).End(xlDown).Row + 100000
    If lMaxRow > Rows.Count Then lMaxRow = 100000
    Call fSetValidationListForRange(shtSelfSalesOrder.Range(sTargetCol & 2 & ":" & sTargetCol & lMaxRow) _
                                    , sValidationListAddr)
End Function

'----------------------------------------------------------------------------------------


'============== ProductUnit ========================================
Function fGetProductUnitMasterColumnAddress_ProductUnit() As String
    Dim sSourceCol As String
    Dim lColMaxRow As Long
    Dim sSourceAddr As String
    
    lColMaxRow = Rows.Count
    sSourceCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - PRODUCT_MASTER]" _
                                        , "Column Index", "Column Tech Name=ProductUnit")

    sSourceAddr = "=" & shtProductMaster.Range(sSourceCol & 2 & ":" & sSourceCol & lColMaxRow).Address(external:=True)
    fGetProductUnitMasterColumnAddress_ProductUnit = sSourceAddr
End Function

Function fSetValidationListForshtProductUnitRatio_ProductUnit(sValidationListAddr As String)
    Dim sTargetCol As String
    Dim lMaxRow As Long
    
    sTargetCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - PRODUCT_UNIT_RATIO_SHEET]" _
                                            , "Column Index", "Column Tech Name=ProductUnit")
    
    lMaxRow = shtProductUnitRatio.Columns(sTargetCol).End(xlDown).Row + 100000
    If lMaxRow > Rows.Count Then lMaxRow = 100000
    Call fSetValidationListForRange(shtProductUnitRatio.Range(sTargetCol & 2 & ":" & sTargetCol & lMaxRow) _
                                    , sValidationListAddr)
End Function
Function fSetValidationListForshtSelfSalesOrder_ProductUnit(sValidationListAddr As String)
    Dim sTargetCol As String
    Dim lMaxRow As Long
    
    sTargetCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - SELF_SALES_ORDER]" _
                                            , "Column Index", "Column Tech Name=ProductUnit")
    
    lMaxRow = shtSelfSalesOrder.Columns(sTargetCol).End(xlDown).Row + 100000
    If lMaxRow > Rows.Count Then lMaxRow = 100000
    Call fSetValidationListForRange(shtSelfSalesOrder.Range(sTargetCol & 2 & ":" & sTargetCol & lMaxRow) _
                                    , sValidationListAddr)
End Function
'----------------------------------------------------------------------------------------

'
'Function fSetValidationListForSingleCell(sValidationListAddr As String, rngCell As Range)
'
'    Call fSetValidationListForRange(rngCell, sValidationListAddr)
'End Function


'============== salesman ========================================
Function fGetSalesManMasterColumnAddress_SalesMan() As String
    Dim sSalesManCol As String
    Dim lSalesManColMaxRow As Long
    Dim sSourceAddr As String
    
    lSalesManColMaxRow = Rows.Count
    
    sSalesManCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - SALESMAN_MASTER]" _
                                        , "Column Index", "Column Tech Name=SalesManName")

    sSourceAddr = "=" & shtSalesManMaster.Range(sSalesManCol & 2 & ":" & sSalesManCol & lSalesManColMaxRow).Address(external:=True)
    fGetSalesManMasterColumnAddress_SalesMan = sSourceAddr
End Function

Function fSetValidationListForshtSalesManCommConfig_SalesMan(sValidationListAddr As String)
    Dim sSalesManCol As String
    Dim lMaxRow As Long
    
    '1
    sSalesManCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - SALESMAN_COMMISSION_CONFIG]" _
                                            , "Column Index", "Column Tech Name=SalesMan1")
    
    lMaxRow = shtSalesManCommConfig.Columns(sSalesManCol).End(xlDown).Row + 100000
    If lMaxRow > Rows.Count Then lMaxRow = 100000
    Call fSetValidationListForRange(shtSalesManCommConfig.Range(sSalesManCol & 2 & ":" & sSalesManCol & lMaxRow), sValidationListAddr)
    
    '2
    sSalesManCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - SALESMAN_COMMISSION_CONFIG]" _
                                            , "Column Index", "Column Tech Name=SalesMan2")
    
    lMaxRow = shtSalesManCommConfig.Columns(sSalesManCol).End(xlDown).Row + 100000
    If lMaxRow > Rows.Count Then lMaxRow = 100000
    Call fSetValidationListForRange(shtSalesManCommConfig.Range(sSalesManCol & 2 & ":" & sSalesManCol & lMaxRow), sValidationListAddr)
    
    '3
    sSalesManCol = fGetSpecifiedConfigCellValue(shtFileSpec, "[Input File - SALESMAN_COMMISSION_CONFIG]" _
                                            , "Column Index", "Column Tech Name=SalesMan3")
    
    lMaxRow = shtSalesManCommConfig.Columns(sSalesManCol).End(xlDown).Row + 100000
    If lMaxRow > Rows.Count Then lMaxRow = 100000
    Call fSetValidationListForRange(shtSalesManCommConfig.Range(sSalesManCol & 2 & ":" & sSalesManCol & lMaxRow), sValidationListAddr)
End Function

'----------------------------------------------------------------------------------------
