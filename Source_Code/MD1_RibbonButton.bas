Attribute VB_Name = "MD1_RibbonButton"
Option Explicit
Option Base 1
   
Sub subMain_Ribbon_ImportSalesInfoFiles()
    If shtMenu.Visible = xlSheetVisible Then
        shtMenu.Visible = xlSheetVeryHidden
    Else
        shtMenu.Visible = xlSheetVisible
        shtMenu.Activate
        Range("A63").Select
    End If
End Sub

