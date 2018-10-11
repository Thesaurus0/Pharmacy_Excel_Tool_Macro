Attribute VB_Name = "친욥1"
Sub 브1()
Attribute 브1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 브1 브
'

'
    Cells.Select
    ActiveWorkbook.Worksheets("TMPOUTPUT").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("TMPOUTPUT").Sort.SortFields.Add Key:=Range( _
        "C1:C909"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("TMPOUTPUT").Sort.SortFields.Add Key:=Range( _
        "A1:A909"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("TMPOUTPUT").Sort
        .SetRange Range("A1:D909")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
