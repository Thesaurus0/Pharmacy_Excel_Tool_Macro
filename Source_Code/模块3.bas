Attribute VB_Name = "ģ��3"
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
'Sub ��4()
''
'' ��4 ��
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
'    Application.DisplayAlerts = False '�رվ���
'    For Each Sh In Sheets '�������������ѭ��
'        If Sh.CodeName = "Sheet1" Then '�����������Ϊsheet1,��������Ϊ"sheet8"
'            MsgBox "��ֹɾ��" & Sh.Name '
'        Else
'            Sh.Delete 'ɾ��
'        End If
'    Next
'    Application.DisplayAlerts = True '�� ����
'End Sub
Sub testaaa()
    Range("A1") = "�����"
    Range("A2") = "��������(ƥ��Ī�¿���)"
    Range("A3") = "2g:0.4g*6��"
End Sub
