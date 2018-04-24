Attribute VB_Name = "xx1"
Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long


 Sub test()
    'MsgBox frmProgressBar.Handle
    
'    Const aa = 1
'    Const bb = 1&
'
'    Dim a
'    Dim b
'
'    a = aa
'    b = bb
'
'    fOpenFile "F:\Github_Local_Repository\Pharmacy_Excel_Tool_Macro\code.vb"
'

    Debug.Print fNum2Letter(Columns.Count)
 End Sub

Function f(t)
    Do
        f = Chr((t - 1) Mod 26 + 65) & f
        t = (t - 1) \ 26
    Loop While t
End Function

