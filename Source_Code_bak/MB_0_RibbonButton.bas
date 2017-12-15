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

Sub subMain_InvisibleAllBusinessSheets()
    shtMenu.Visible = xlSheetVisible
    shtHospital.Visible = xlSheetVeryHidden
    shtHospitalReplace.Visible = xlSheetVeryHidden
    shtSalesRawDataRpt.Visible = xlSheetVeryHidden
    shtSalesInfos.Visible = xlSheetVeryHidden
End Sub
