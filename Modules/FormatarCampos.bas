Attribute VB_Name = "FormatarCampos"
Option Explicit
Sub Formatar_CEP(ByVal KeyAscii As MSForms.ReturnInteger, Controle As control)
'Sub/ M�scara para bloquear o campo CEP somente para n�meros e inserir h�fen automaticamente.
'XX.XXX-XXX -> 10 c
Controle.MaxLength = 10
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0 'Bloqueia e permite apenas n�meros no campo CNPJ

If Len(Controle.Text) = 2 Then Controle.Text = Controle.Text & "."

If Len(Controle.Text) = 6 Then Controle.Text = Controle.Text & "-"

End Sub

Sub Formatar_CPF(ByVal KeyAscii As MSForms.ReturnInteger, Controle As control)
'Sub/ M�scara para bloquear o campo CEP somente para n�meros e inserir h�fen automaticamente.
'XXX.XXX.XXX-XX -> 14 c
Controle.MaxLength = 14
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0 'Bloqueia e permite apenas n�meros no campo CNPJ

If Len(Controle.Text) = 3 Or Len(Controle.Text) = 7 Then Controle.Text = Controle.Text & "."

If Len(Controle.Text) = 11 Then Controle.Text = Controle.Text & "-"

End Sub

Sub Formatar_CNPJ(ByVal KeyAscii As MSForms.ReturnInteger, Controle As control)
'Sub/ M�scara para bloquear o campo CEP somente para n�meros e inserir h�fen automaticamente.
'XX.XXX.XXX/XXXX-XX -> 18 c
Controle.MaxLength = 18
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0 'Bloqueia e permite apenas n�meros no campo CNPJ

If Len(Controle.Text) = 2 Or Len(Controle.Text) = 6 Then Controle.Text = Controle.Text & "."

If Len(Controle.Text) = 10 Then Controle.Text = Controle.Text & "/"

If Len(Controle.Text) = 15 Then Controle.Text = Controle.Text & "-"

End Sub

Sub Formatar_Cel(ByVal KeyAscii As MSForms.ReturnInteger, Controle As control)
'Sub/ M�scara para bloquear o campo CEP somente para n�meros e inserir h�fen automaticamente.
'(XX) X.XXXX-XXXX -> 16 c
Controle.MaxLength = 16
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0 'Bloqueia e permite apenas n�meros no campo CNPJ

If Len(Controle.Text) = 1 Then Controle.Text = "(" & Controle.Text

If Len(Controle.Text) = 3 Then Controle.Text = Controle.Text & ")"

If Len(Controle.Text) = 4 Then Controle.Text = Controle.Text & " "

If Len(Controle.Text) = 6 Then Controle.Text = Controle.Text & "."

If Len(Controle.Text) = 11 Then Controle.Text = Controle.Text & "-"

End Sub
