VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} consultaCNPJ 
   Caption         =   "Consulta de CNPJ via API"
   ClientHeight    =   8760.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12300
   OleObjectBlob   =   "consultaCNPJ.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "consultaCNPJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public tbox As New clsTextBox

Private Sub txtCNPJ_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    '---------------------- Para formatar o campo de digita��o do CNPJ -----------------------------'
    Call FormatarCampos.Formatar_CNPJ(KeyAscii, Me.txtCNPJ)
End Sub

Private Sub UserForm_Activate()
    ''------------------- Para personalizar o userform e armazenar as dimens�es atuais -------------'
    PersonalizaForm
    ArmazenaDimIN Me
End Sub

Private Sub btConsulta_Click()
Dim Cnpj$, RazaoSocial$, NomeFantasia$, Endereco$, numero$, complemento$, bairro$, _
cidade$, uf$, cep$, tel$, Email$, porte$, DtAbertura$, Situacao$, NaturJuridica$, _
AtivPrincipal$, CNAEPrincipal$, Capital$, AtivSecundarias, Socios, ArrSecundarias _
, ArrSocios
Dim ws As Object, wb As Object
Dim Lista As Range
Dim lin%

ThisWorkbook.Activate
Set ws = Planilha2
Set Lista = ws.Range("A2:E100")
Cnpj = Me.txtCNPJ.Value

'------------------------- Verifica��o se o campo do CNPJ est� vazio --------------------------'
If Cnpj = "" Then
    MsgBox "O campo do CNPJ est� vazio" & vbCr _
     & vbCr & "Por favor, insira um CNPJ v�lido.", vbInformation, "Registro n�o encontrado!"
     Exit Sub
Else
    
    Call BuscarCNPJ(Cnpj, RazaoSocial, NomeFantasia, Endereco, numero, complemento, bairro, _
                    cidade, uf, cep, tel, Email, porte, DtAbertura, Situacao, NaturJuridica, _
                    AtivPrincipal, CNAEPrincipal, Capital, AtivSecundarias, Socios, _
                    ArrSecundarias, ArrSocios)
End If

'---------------------- Verifica se foi encontrado algum registro v�lido -----------------------'
If RazaoSocial = "" And NomeFantasia = "" And Endereco = "" And AtivPrincipal = "" Then
    Cnpj = Me.txtCNPJ
    MsgBox "A Consulta CNPJ n�o encontrou um registro v�lido para o CNPJ " & Cnpj & vbCr _
     & vbCr & "Por favor, insira um CNPJ v�lido.", vbInformation, "Registro n�o encontrado!"
     Exit Sub
End If
    '---------------------------- Sa�da de Dados B�sico da empresa -----------------------------'
    Me.txtRazSocial = RazaoSocial
    Me.txtNomeFantasia = NomeFantasia
    Me.txtEndereco = Endereco
    Me.txtNumero = numero
    Me.txtBairro = bairro
    Me.txtComplemento = complemento
    Me.txtCidade = cidade
    Me.txtUF = uf
    Me.txtTel = tel
    Me.txtEmail = Email
    Me.txtCEP = cep
    
    '------------------------------ Sa�da de Informa��es da empresa ----------------------------'
    Me.txtPorte = porte
    Me.txtDtAbertura = DtAbertura
    Me.txtSituacao = Situacao
    Me.txtNatJuridica = NaturJuridica
    Me.txtAtivPrincipal = AtivPrincipal
    Me.txtCNAEPrincipal = CNAEPrincipal
    
    ArrSecundarias = ws.Range("A2").CurrentRegion.Value
    With Me.LtAtivSecundarias
        .Clear
        .ColumnCount = 2
        .ColumnWidths = "100;300"
        For lin = LBound(ArrSecundarias) To UBound(ArrSecundarias)
            .AddItem ArrSecundarias(lin, 1)
            .List(.ListCount - 1, 1) = ArrSecundarias(lin, 2)
            .List(.ListCount - 1, 2) = ArrSecundarias(lin, 2)
        Next lin
    End With
    
'-------------------------- Sa�da de dados do Quadro social da empresa -------------------------'
    ArrSocios = ws.Range("D2").CurrentRegion.Value
    With Me.LtSocios
        .Clear
        .ColumnCount = 2
        .ColumnWidths = "230;100"
        For lin = LBound(ArrSocios) To UBound(ArrSocios)
            .AddItem ArrSocios(lin, 1)
            .List(.ListCount - 1, 1) = ArrSocios(lin, 2)
            .List(.ListCount - 1, 2) = ArrSocios(lin, 2)
        Next lin
    End With
   
    Me.txtCapital = Format(Replace(Capital, ".", ","), "R$ #,000.00")
    
'-------------------------- Boas Pr�ticas de VBA: Destruir objetos -----------------------------'
Set ws = Nothing
Set Lista = Nothing

End Sub

Private Sub UserForm_Initialize()
Call textperson
End Sub

Private Sub UserForm_Resize()
    AjustaDimControle
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
