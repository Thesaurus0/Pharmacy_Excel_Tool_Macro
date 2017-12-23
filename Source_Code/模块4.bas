Attribute VB_Name = "模块4"
Sub 宏4()
Attribute 宏4.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 宏4 宏
'

'
    Sheets("药品名称").Select
    ActiveSheet.Range("$A$1:$B$33").AutoFilter Field:=1, Criteria1:="广州白云山陈李济"
    Range("B2:B7").Select
    Selection.Copy
    Sheets("ForFilter").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("药品名称").Select
End Sub
