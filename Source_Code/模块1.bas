Attribute VB_Name = "Ä£¿é1"
Sub ºê1()
Attribute ºê1.VB_ProcData.VB_Invoke_Func = " \n14"
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
Sub ºê2()
Attribute ºê2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ºê2 ºê
'

    'Dim format As Data
'
    Dim s
    Dim sProductProducer As String
    Dim sProductName As String
    
    sProductName = "2g:0.4g*6´ü"
    
    If sProductName Like "½ðÊÀÁ¦µÂ(*" Then
         MsgBox "a:"
    End If
    
    If sProductName Like "2g:0.4g^*6´ü*" Then
         MsgBox "b:"
    End If
End Sub

Sub cccc()
    Debug.Print ActiveCell.Value
End Sub

