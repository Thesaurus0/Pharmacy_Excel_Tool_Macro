Attribute VB_Name = "ģ��1"
Sub ��1()
Attribute ��1.VB_ProcData.VB_Invoke_Func = " \n14"
    'Clipboard.Clear
    Dim myData As DataObject
    Dim sOriginText As String
    Set myData = New DataObject
    myData.GetFromClipboard
    
     myData.StartDrag
    sOriginText = myData.GetText()
    
    myData.Clear
    myData.SetText ""
    myData.PutInClipboard
End Sub
Sub ��2()
Attribute ��2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ��2 ��
'

    'Dim format As Data
'
    Dim s
    Dim sProductProducer As String
    Dim sProductName As String
    
    sProductName = "2g:0.4g*6��"
    
    If sProductName Like "��������(*" Then
         MsgBox "a:"
    End If
    
    If sProductName Like "2g:0.4g^*6��*" Then
         MsgBox "b:"
    End If
End Sub

Sub cccc()
    Dim lLastMaxRow As Long
    lLastMaxRow = shtDataStage.Cells(Rows.Count, 2).End(xlUp).Row
Debug.Print lLastMaxRow
End Sub

