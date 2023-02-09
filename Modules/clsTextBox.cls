VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents myForm As MSForms.UserForm
Attribute myForm.VB_VarHelpID = -1
Public WithEvents mFrame As MSForms.Frame
Attribute mFrame.VB_VarHelpID = -1
Public WithEvents mPage As MSForms.MultiPage
Attribute mPage.VB_VarHelpID = -1
Public WithEvents Textbox As MSForms.Textbox
Attribute Textbox.VB_VarHelpID = -1
Public WithEvents ComboBox As MSForms.ComboBox
Attribute ComboBox.VB_VarHelpID = -1
Public WithEvents barBottom As MSForms.Label
Attribute barBottom.VB_VarHelpID = -1
Public WithEvents barTitle As MSForms.Label
Attribute barTitle.VB_VarHelpID = -1
Public WithEvents BackGround As MSForms.Label
Attribute BackGround.VB_VarHelpID = -1
Public WithEvents barIcon As MSForms.Label
Attribute barIcon.VB_VarHelpID = -1
Public WithEvents DropArrow As MSForms.Label
Attribute DropArrow.VB_VarHelpID = -1

Public ctBox        As New Collection
'adaptação Organic Sheets
Public Sub clasBoxInvisibleAll(form As MSForms.UserForm)
    
    On Error Resume Next
    
    For Each ctrl In myForm.Controls
        
        myForm.Controls.Remove ctrl.Name & "barBottom"
        myForm.Controls.Remove ctrl.Name & "barTitle"
        myForm.Controls.Remove ctrl.Name & "barIcon"
        myForm.Controls.Remove ctrl.Name & "BackGround"
        myForm.Controls.Remove ctrl.Name & "DropArrow"
        
    Next ctrl
    
End Sub
Public Sub clasBoxInvisible(form As MSForms.UserForm, Controle As control)
    
    For Each ctrl In myForm.Controls
        On Error Resume Next
        
        If TypeName(ctrl) = "TextBox" And ctrl.Name = Controle.Name Then        'ctrl.Name = Controle.Name Then
        
        myForm.Controls.Remove ctrl.Name & "barBottom"
        myForm.Controls.Remove ctrl.Name & "barTitle"
        myForm.Controls.Remove ctrl.Name & "barIcon"
        myForm.Controls.Remove ctrl.Name & "BackGround"
        
        ElseIf TypeName(ctrl) = "ComboBox" And ctrl.Name = Controle.Name Then
        
        myForm.Controls.Remove ctrl.Name & "barBottom"
        myForm.Controls.Remove ctrl.Name & "barTitle"
        myForm.Controls.Remove ctrl.Name & "barIcon"
        myForm.Controls.Remove ctrl.Name & "BackGround"
        myForm.Controls.Remove ctrl.Name & "DropArrow"
        
        End If
    
    Next ctrl

End Sub
Public Sub clasBox(form As MSForms.UserForm)
    
    FontSize = 10: FontName = "MontSerrat"
    Dim ctrl        As control
    Set myForm = form
    
    For Each ctrl In myForm.Controls
        If TypeName(ctrl) = "TextBox" And ctrl.Visible = True Then
            With ctrl
                Call TextBoxStyle(ctrl)
                
                Set barBottom = ctrl.Parent.Controls.Add("Forms.Label.1", .Name & "barBottom")
                Call AddBarBottom(ctrl)
                
                Set barTitle = ctrl.Parent.Controls.Add("Forms.Label.1", .Name & "barTitle")
                Call AddBarTitle(ctrl)
                
                Set barIcon = ctrl.Parent.Controls.Add("Forms.Label.1", .Name & "barIcon")
                AddBarIcon ctrl
                
                Set BackGround = ctrl.Parent.Controls.Add("Forms.Label.1", .Name & "BackGround")
                AddBackGround ctrl
                
                .ZOrder (0)
                
            End With
            
            Set txt = New clsTextBox
            Set txt.Textbox = ctrl
            Set txt.myForm = myForm
            ctBox.Add txt
            
        ElseIf TypeName(ctrl) = "ComboBox" And ctrl.Visible = True Then
            With ctrl
                Call ComboBoxStyle(ctrl)
                
                Set barBottom = ctrl.Parent.Controls.Add("Forms.Label.1", .Name & "barBottom")
                Call AddBarBottom(ctrl)
                
                Set barTitle = ctrl.Parent.Controls.Add("Forms.Label.1", .Name & "barTitle")
                Call AddBarTitle(ctrl)
                
                Set barIcon = ctrl.Parent.Controls.Add("Forms.Label.1", .Name & "barIcon")
                AddBarIcon ctrl
                
                Set BackGround = ctrl.Parent.Controls.Add("Forms.Label.1", .Name & "BackGround")
                AddBackGround ctrl
                
                .ZOrder (0)
                
                Set DropArrow = ctrl.Parent.Controls.Add("Forms.Label.1", .Name & "DropArrow")
                Call AddDropArrow(ctrl)
                
            End With
            
            Set txt = New clsTextBox
            Set txt.ComboBox = ctrl
            Set txt.DropArrow = DropArrow
            Set txt.myForm = myForm
            ctBox.Add txt
            
        ElseIf TypeName(ctrl) = "MultiPage" Then
            Set txt = New clsTextBox
            Set txt.mPage = ctrl
            Set txt.myForm = myForm
            ctBox.Add txt
            
        ElseIf TypeName(ctrl) = "Frame" Then
            Set txt = New clsTextBox
            Set txt.mFrame = ctrl
            Set txt.myForm = myForm
            ctBox.Add txt
        End If
        
    Next ctrl
    
End Sub
Sub TextBoxStyle(ctrl As control)
    
    With ctrl
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleSingle
        .BorderStyle = fmBorderStyleNone
        .SelectionMargin = False
        .Font.Name = FontName
        .ForeColor = fColor
        .Font.Size = FontSize
        
        If ctrl.Height > 20 Then
            .Height = ctrl.Height
        Else
            .Height = 20
        End If
        '        .ControlTipText = "cBox"
    End With
End Sub
Sub ComboBoxStyle(ctrl As control)
    With ctrl
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleSingle
        .BorderStyle = fmBorderStyleNone
        .SelectionMargin = False
        .Font.Name = FontName
        .ForeColor = fColor
        .ShowDropButtonWhen = fmShowDropButtonWhenNever
        .Font.Size = FontSize
        .Height = 24
        '        .ControlTipText = "cBox"
    End With
End Sub
Sub AddBarBottom(ctrl As control)
    With barBottom
        .Height = 1
        .Width = ctrl.Width
        .Left = ctrl.Left
        .BackColor = tColor
        .Top = ctrl.Top + ctrl.Height
        .ZOrder (0)
    End With
End Sub
Sub AddBarTitle(ctrl As control)
    With barTitle
        .Height = ctrl.Height
        .Width = ctrl.Width
        .Left = ctrl.Left + 2
        .Caption = ctrl.ControlTipText
        .ForeColor = tColor
        .Font.Name = FontName
        .Font.Size = FontSize
        .BackStyle = fmBackStyleTransparent
        .Top = ctrl.Top
        .ZOrder (0)
    End With
End Sub
Sub AddBackGround(ctrl As control)
    With BackGround
        .Height = ctrl.Height + 4
        .Width = ctrl.Width
        .BackColor = bColor
        .Font.Size = FontSize
        '        .BackStyle = fmBackStyleTransparent
        .Top = ctrl.Top - 4
        .Left = ctrl.Left
        .ZOrder (1)
    End With
End Sub
Sub AddBarIcon(ctrl As control)
    
    'On Error Resume Next
    
    With barIcon
        .Font.Name = "Segoe MDL2 Assets" '"icon-font" '
        .Font.Size = 14
        If ctrl.Tag <> "" Then .Caption = ChrW(ctrl.Tag)
        .Width = 18
        .Height = 18
        .ForeColor = tColor
        .Top = ctrl.Top + 2
        .Left = ctrl.Left - .Width
        .BackStyle = fmBackStyleTransparent
        .ZOrder (1)
        '        .BorderStyle = fmBorderStyleSingle
    End With
End Sub

Sub AddDropArrow(ctrl As control)
    With DropArrow
        .Font.Name = "Webdings"
        .Caption = "6"
        .Height = ctrl.Height
        .Width = .Height
        .Left = ctrl.Width + ctrl.Left - .Width + 8
        .BackStyle = fmBackStyleTransparent
        .ForeColor = tColor
        .BorderStyle = fmBorderStyleNone
        .Top = ctrl.Top + 1
        .Font.Size = FontSize
        .ZOrder (0)
        
    End With
End Sub

Sub SetForm(form As MSForms.UserForm)
    Set txt = New clsTextBox
    Set txt.myForm = form
    ctBox.Add txt
End Sub
Private Sub mFrame_Click()
    BoxExit
End Sub

Private Sub mPage_Click(ByVal Index As Long)
    BoxExit
End Sub

Private Sub mPage_MouseDown(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    BoxExit
End Sub

Private Sub TextBox_Change()
    TextBoxEnter
End Sub
Private Sub Textbox_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    TextBoxEnter
End Sub
Private Sub Textbox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    TextBoxEnter
End Sub

Private Sub txtCaption_Click()
    BoxExit
End Sub
Sub TextBoxEnter()
    BoxExit
    Dim control     As control
    
    On Error Resume Next
    
    For Each control In myForm.Controls
        If TypeName(control) = "TextBox" And control.Name = Textbox.Name And control.Visible = True Then
            With myForm.Controls(Textbox.Name & "barBottom")
                .BackColor = eColor
                .Height = 1.2
            End With
            With myForm.Controls(Textbox.Name & "barTitle")
                .ForeColor = eColor
                .Top = Textbox.Top - 17
                .Font.Size = 9
                .Left = Textbox.Left
            End With
            
            With myForm.Controls(Textbox.Name & "barIcon")
                .ForeColor = eColor
            End With
            
        End If
    Next control
End Sub
Sub BoxExit()
    
    Dim control     As control
    On Error Resume Next
    For Each control In myForm.Controls
        If TypeName(control) = "TextBox" And control.Visible = True Or TypeName(control) = "ComboBox" And control.Visible = True Then
            If Len(control) = 0 Then
                With myForm.Controls(control.Name & "barBottom")
                    .BackColor = tColor
                    .Height = 1
                End With
                With myForm.Controls(control.Name & "barTitle")
                    .ForeColor = tColor
                    .Top = control.Top
                    .Left = control.Left + 2
                    .Font.Size = FontSize
                End With
                With myForm.Controls(control.Name & "barIcon")
                    .ForeColor = tColor
                End With
                
            Else
            
                With myForm.Controls(control.Name & "barBottom")
                    .BackColor = tColor
                    .Height = 1
                End With
                With myForm.Controls(control.Name & "barTitle")
                    .ForeColor = tColor
                    .Top = control.Top - 17
                    .Left = control.Left
                    .Font.Size = 9
                End With
                
                With myForm.Controls(control.Name & "barIcon")
                    .ForeColor = tColor
                End With
            End If
        End If
        
        If TypeName(control) = "ComboBox" And control.Visible = True Then
            With myForm.Controls(control.Name & "DropArrow")
                .ForeColor = tColor
            End With
        End If
    Next
End Sub
Sub ComboBoxEnter()
    BoxExit
    Dim control     As control
    On Error Resume Next
    For Each control In myForm.Controls
        
        If TypeName(control) = "ComboBox" And control.Name = ComboBox.Name And control.Visible = True Then
            
            With myForm.Controls(ComboBox.Name & "barBottom")
                .BackColor = eColor
                .Height = 1.2
            End With
            With myForm.Controls(ComboBox.Name & "barTitle")
                .ForeColor = eColor
                .Top = ComboBox.Top - 17
                .Font.Size = 9
                .Left = ComboBox.Left
            End With
            
            With myForm.Controls(ComboBox.Name & "barIcon")
                .ForeColor = eColor
            End With
            
            With myForm.Controls(ComboBox.Name & "DropArrow")
                .ForeColor = eColor
            End With
            '                Control.DropDown
        End If
    Next
    
End Sub

Private Sub ComboBox_Change()
    ComboBoxEnter
End Sub

Private Sub ComboBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ComboBoxEnter
End Sub

Private Sub ComboBox_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ComboBoxEnter
End Sub

Private Sub DropArrow_Click()
    BoxExit
    Dim ComboName   As String
    Dim ctrl        As control
    For Each ctrl In myForm.Controls
        If TypeName(ctrl) = "ComboBox" Then
            ComboName = Mid(DropArrow.Name, 1, Len(DropArrow.Name) - 9)
            If ComboName = ctrl.Name Then
                With myForm.Controls(ComboName)
                    .DropDown
                    With myForm.Controls(ComboName & "barBottom")
                        .BackColor = eColor
                        .Height = 1.2
                    End With
                    
                    With myForm.Controls(ComboName & "barTitle")
                        .ForeColor = eColor
                        .Top = myForm.Controls(ComboName).Top - 17
                        .Font.Size = 9
                        .Left = myForm.Controls(ComboName).Left
                    End With
                End With
                With DropArrow
                    .ForeColor = eColor
                End With
            End If
        End If
    Next ctrl
End Sub

Private Sub DropArrow_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    '    Call MouseMoveIcon
End Sub

Private Sub myForm_Click()
    BoxExit
End Sub







