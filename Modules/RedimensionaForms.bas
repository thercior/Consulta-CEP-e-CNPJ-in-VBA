Attribute VB_Name = "RedimensionaForms"
Option Explicit

Private dLargura    As Single
Private dAltura     As Single
Private Ufrm        As Object

'Irá armazenar as dimensões de todos os controles dentro do formulário
'em suas respectivas Tags, para que posteriomentes sejam usadas para cálculo
Public Sub ArmazenaDimIN(ByVal uf As Object)

    Dim oCtrl As control
    Dim dFontSize As Double

    Set Ufrm = uf

    dLargura = Ufrm.InsideWidth
    dAltura = Ufrm.InsideHeight

    For Each oCtrl In Ufrm.Controls
        With oCtrl
            On Error Resume Next
                dFontSize = IIf(HasFont(oCtrl), .Font.Size, 0)
            On Error GoTo 0
            .Tag = .Width & "*" & .Left & "*" & .Height & "*" & .Top & "*" & dFontSize
        End With
    Next
    '200*10*80*100*12
End Sub

Private Function HasFont(ByVal oCtrl As control) As Boolean

    Dim oFont As Object

    On Error Resume Next
    Set oFont = CallByName(oCtrl, "Font", VbGet)
    HasFont = Not oFont Is Nothing
    Set oFont = Nothing

End Function

'Este procedimento deve ser colocado dentro do Evento Rezize do Formulário
'Irá realizar a multiplicação entre os valores inicialmente armazenados nas Tags
'e a taxa de proporção do resultado entre as dimensões iniciais e finais
Public Sub AjustaDimControle(Optional ByVal Dummey As Boolean)

    Dim oCtrl As control

    On Error Resume Next

    For Each oCtrl In Ufrm.Controls
        With oCtrl
            If .Tag <> "" Then  '200*10*80*100*12     'largura * distancia esquerda * altura * dist topo * tam fonte
                .Width = Split(.Tag, "*")(0) * ((Ufrm.InsideWidth) / dLargura)
                .Left = Split(.Tag, "*")(1) * (Ufrm.InsideWidth) / dLargura
                .Height = Split(.Tag, "*")(2) * (Ufrm.InsideHeight) / dAltura
                .Top = Split(.Tag, "*")(3) * (Ufrm.InsideHeight) / dAltura
                If HasFont(oCtrl) Then
                    .Font.Size = Split(.Tag, "*")(4) * (Ufrm.InsideWidth) / dLargura
                End If
            End If
        End With
    Next

End Sub



