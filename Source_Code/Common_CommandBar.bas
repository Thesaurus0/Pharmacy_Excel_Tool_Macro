Attribute VB_Name = "Common_CommandBar"
Option Explicit
Option Base 1

Sub sub_RemoveAllToolBars()
    Dim tmpBar As CommandBar
    
    For Each tmpBar In Application.CommandBars
        If Not tmpBar.BuiltIn Then
            tmpBar.Delete
        End If
    Next
End Sub


Option Explicit

Sub sub_add_new_bar(as_bar_name As String)
    Dim lcb_new_commdbar As CommandBar
        
    Call sub_RemoveToolBar(as_bar_name)
    
    Set lcb_new_commdbar = Application.CommandBars.Add(as_bar_name, msoBarTop)
    lcb_new_commdbar.Visible = True
End Sub

Public Sub sub_RemoveToolBar(as_toolbar As String)
    On Error Resume Next

    Dim lcb_commdbar As CommandBar
    
    Set lcb_commdbar = Nothing
    
    Application.CommandBars(as_toolbar).Delete
    Application.CommandBars("Custom 1").Delete
End Sub

Sub sub_remove_all_bars()
    On Error Resume Next
    Dim tempbar As CommandBar

    For Each tempbar In Application.CommandBars
        'If tempbar.Name Like "my_bar*" Then
            tempbar.Delete
        'End If
    Next

End Sub

Public Sub sub_add_new_button(as_bar_name As String, as_btn_caption As String, _
                    as_on_action As String, ai_face_id As Integer, _
                    Optional as_tip_text As String)

    Dim lcb_commdbar As CommandBar
    Dim lbtn_new_button As CommandBarButton
    
    Set lcb_commdbar = Application.CommandBars(as_bar_name)
        
    Set lbtn_new_button = lcb_commdbar.Controls.Add(msoControlButton)
    With lbtn_new_button
        .Caption = as_btn_caption
        .Style = msoButtonIconAndCaptionBelow
        '.OnAction = "sub_RemoveToolBar"
        .OnAction = as_on_action
        .FaceId = ai_face_id
        .TooltipText = as_tip_text
        .BeginGroup = True
    End With
    
    'Set lcb_commdbar = Nothing
    'Set lbtn_new_button = Nothing
    
End Sub




