Attribute VB_Name = "ģ��4"
Sub ��4()
Attribute ��4.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ��4 ��
'

'
    Sheets("ҩƷ����").Select
    ActiveSheet.Range("$A$1:$B$33").AutoFilter Field:=1, Criteria1:="���ݰ���ɽ�����"
    Range("B2:B7").Select
    Selection.Copy
    Sheets("ForFilter").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("ҩƷ����").Select
End Sub
