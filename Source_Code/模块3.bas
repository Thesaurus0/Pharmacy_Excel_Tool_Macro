Attribute VB_Name = "模块3"
'#If Win64 Then
'   Declare PtrSafe Function MyMathFunc Lib "User32" (ByVal N As LongLong) As LongLong
'#Else
'   Declare Function MyMathFunc Lib "User32" (ByVal N As Long) As Long
'#End If
'#If VBA7 Then
'   Declare PtrSafe Sub MessageBeep Lib "User32" (ByVal N As Long)
'#Else
'   Declare Sub MessageBeep Lib "User32" (ByVal N AS Long)
'#End If
'
'Sub 宏4()
''
'' 宏4 宏
''
'
''
'    ChDir "F:\Github_Local_Repository\Pharmacy_Excel_Tool_Macro"
'    Workbooks.Open Filename:= _
'        "F:\Github_Local_Repository\Pharmacy_Excel_Tool_Macro\Pharmacy_Excel_Tool_Macro_V0.6.xlsm" _
'        , Editable:=True
'    Workbooks.Open Filename:= _
'        "F:\Github_Local_Repository\Pharmacy_Excel_Tool_Macro\Pharmacy_Excel_Tool_Macro_V0.5.xlsm" _
'        , Editable:=True
'End Sub
''Sub DisableDelSheet()
''    Application.CommandBars("edit").Controls(12).Enabled = False
''    Application.CommandBars("ply").Controls(2).Enabled = False
''    Application.CommandBars("ply").Controls(3).Enabled = False
''End Sub
''Sub EnableDelSheet()
''Application.CommandBars("edit").Controls(12).Enabled = True
''Application.CommandBars("ply").Controls(2).Enabled = True
''Application.CommandBars("ply").Controls(3).Enabled = True
''End Sub
'
'Sub mydelesh()
'    On Error Resume Next
'    Application.DisplayAlerts = False '关闭警告
'    For Each Sh In Sheets '在这个工作簿内循环
'        If Sh.CodeName = "Sheet1" Then '工作表代码名为sheet1,工作表名为"sheet8"
'            MsgBox "禁止删除" & Sh.Name '
'        Else
'            Sh.Delete '删除
'        End If
'    Next
'    Application.DisplayAlerts = True '打开 警告
'End Sub
Sub testaaa()
    Range("A1") = "津金世"
    Range("A2") = "金世力德(匹多莫德颗粒)"
    Range("A3") = "2g:0.4g*6袋"
End Sub
