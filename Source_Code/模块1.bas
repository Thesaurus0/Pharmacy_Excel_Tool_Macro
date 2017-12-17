Attribute VB_Name = "친욥1"
Sub 브1()
Attribute 브1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 브1 브
'

'
'    Call fModifyMoveActiveXButtonOnSheet(shtSalesRawDataRpt.Cells(1, fGetValidMaxCol(shtSalesRawDataRpt) + 1) _
'                                        , "btnReplaceUnify", 1, 1, , 25, RGB(255, 20, 134), RGB(255, 255, 255))

    Debug.Print ActiveCell.Address(external:=True)
    
    Dim a As Range
    
    Set a = fGetRangeFromExternalAddress(ActiveCell.Address(external:=True))
End Sub
