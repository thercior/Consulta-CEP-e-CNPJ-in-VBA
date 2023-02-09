VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Consulta_Cep 
   Caption         =   "Consulta de CEP Via API"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6105
   OleObjectBlob   =   "Consulta_Cep.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Consulta_Cep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public tbox As New clsTextBox

Private Sub btConsultar_Click()
Dim cep     As String

cep = txtCEP.Value

Call BuscarCep(cep)

End Sub

Private Sub txtCEP_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'Sub/ M�scara para bloquear o campo CEP somente para n�meros e inserir h�fen automaticamente.
 Call FormatarCampos.Formatar_CEP(KeyAscii, txtCEP)
End Sub

Private Sub UserForm_Activate()
    PersonalizaForm
    ArmazenaDimIN Me
End Sub

Private Sub UserForm_Initialize()
Call textperson
End Sub

Sub textperson()
 '1� cor: da fonte (textbox forecolor)
 '2� cor: de destaque quando selecionado (When TextBox Enter)
 '3� cor: do R�tulo (T�tulo e cor do bot�o linha)
 '4� cor: de fundo
 
TxtColor 1512210, 15874686, 10395294, 15856371

tbox.clasBoxInvisibleAll Me
tbox.clasBox Me
tbox.BoxExit

End Sub
