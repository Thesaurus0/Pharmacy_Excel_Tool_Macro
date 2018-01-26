Attribute VB_Name = "Ä£¿é3"
Sub adcaf()
    Dim a As New ProgressBar
    
    a.ShowBar
    a.ChangeProcessBarValue 0.1
'    Application.Wait (Now() + TimeSerial(0, 0, 1))
'    a.ChangeProcessBarValue 0.2
'    a.SleepBar 1000
'    a.ChangeProcessBarValue 0.5
'
'    Application.Wait (Now() + TimeSerial(0, 0, 1))
'    a.ChangeProcessBarValue 1, format(1 / 5, "@")
'    Application.Wait (Now() + TimeSerial(0, 0, 1))
    a.ChangeProcessBarValue 1, "okokok"
    Application.Wait (Now() + TimeSerial(0, 0, 3))
    a.DestroyBar
    
    a.ShowBar
    a.ChangeProcessBarValue 0.1
    Application.Wait (Now() + TimeSerial(0, 0, 1))
    a.ChangeProcessBarValue 0.2
    a.SleepBar 1000
    a.ChangeProcessBarValue 0.5
    a.DestroyBar
    Set a = Nothing
    
End Sub

Sub azxcvadfs()
    Debug.Print format(1 / 5, "@")
End Sub
