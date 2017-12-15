Attribute VB_Name = "MB_0_RibbonButton"
Option Explicit
Option Base 1
  
Sub subMain_Ribbon_ImportSalesInfoFiles()
    On Error Resume Next
    
    If shtMenu.Visible = xlSheetVisible Then
        If ActiveSheet.Name <> shtMenu.Name Then
            shtMenu.Visible = xlSheetVisible
            shtMenu.Activate
            Range("A63").Select
        Else
            shtMenu.Visible = xlSheetVeryHidden
        End If
    Else
        shtMenu.Visible = xlSheetVisible
        shtMenu.Activate
        Range("A63").Select
    End If
    
    err.Clear
End Sub

Sub subMain_Hospital()
    If shtHospital.Visible = xlSheetVisible Then
        If ActiveSheet.Name <> shtHospital.Name Then
            shtHospital.Visible = xlSheetVisible
            shtHospital.Activate
            Range("a1").Select
        Else
            shtHospital.Visible = xlSheetVeryHidden
        End If
    Else
        shtHospital.Visible = xlSheetVisible
        shtHospital.Activate
        Range("a1").Select
    End If
End Sub

Sub subMain_HospitalReplacement()
    If shtHospitalReplace.Visible = xlSheetVisible Then
        If ActiveSheet.Name <> shtHospitalReplace.Name Then
            shtHospitalReplace.Visible = xlSheetVisible
            shtHospitalReplace.Activate
            Range("a1").Select
        Else
            shtHospitalReplace.Visible = xlSheetVeryHidden
        End If
    Else
        shtHospitalReplace.Visible = xlSheetVisible
        shtHospitalReplace.Activate
        Range("a1").Select
    End If
End Sub

Sub subMain_RawSalesInfos()
    If shtSalesRawDataRpt.Visible = xlSheetVisible Then
        If ActiveSheet.Name <> shtSalesRawDataRpt.Name Then
            shtSalesRawDataRpt.Visible = xlSheetVisible
            shtSalesRawDataRpt.Activate
            Range("a1").Select
        Else
            shtSalesRawDataRpt.Visible = xlSheetVeryHidden
        End If
    Else
        shtSalesRawDataRpt.Visible = xlSheetVisible
        shtSalesRawDataRpt.Activate
        Range("a1").Select
    End If
End Sub

Sub subMain_SalesInfos()
    If shtSalesInfos.Visible = xlSheetVisible Then
        If ActiveSheet.Name <> shtSalesInfos.Name Then
            shtSalesInfos.Visible = xlSheetVisible
            shtSalesInfos.Activate
            Range("a1").Select
        Else
            shtSalesInfos.Visible = xlSheetVeryHidden
        End If
    Else
        shtSalesInfos.Visible = xlSheetVisible
        shtSalesInfos.Activate
        Range("a1").Select
    End If
End Sub

Sub subMain_ProductMaster()
    If shtProductMaster.Visible = xlSheetVisible Then
        If ActiveSheet.Name <> shtProductMaster.Name Then
            shtProductMaster.Visible = xlSheetVisible
            shtProductMaster.Activate
            Range("a1").Select
        Else
            shtProductMaster.Visible = xlSheetVeryHidden
        End If
    Else
        shtProductMaster.Visible = xlSheetVisible
        shtProductMaster.Activate
        Range("a1").Select
    End If
End Sub
Sub subMain_ProductProducerReplace()

End Sub
Sub subMain_ProductNameReplace()

End Sub
Sub subMain_ProductSeriesReplace()

End Sub
Sub subMain_ProductUnitRatio()

End Sub


Sub subMain_InvisibleAllBusinessSheets()
    shtMenu.Visible = xlSheetVisible
    shtHospital.Visible = xlSheetVeryHidden
    shtHospitalReplace.Visible = xlSheetVeryHidden
    shtSalesRawDataRpt.Visible = xlSheetVeryHidden
    shtSalesInfos.Visible = xlSheetVeryHidden
End Sub
