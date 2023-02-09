Attribute VB_Name = "BuscaCep"
Option Explicit

Sub BuscarCep(cep As String)
Dim api     As New MSXML2.ServerXMLHTTP60
Dim html    As New MSHTML.HTMLDocument
Dim url     As String

'------------------- Formata a vari�vel CEP para o padr�o de busca da API VIA CEP -----------------'
cep = Replace(Replace(cep, ".", ""), "-", "")

'--------------------- Defini��o da API VIACEP ----------------------------------------------------'
url = "https://viacep.com.br/ws/" & cep & "/xml/"

'--------------------- Abertura da Conex�o e chamada da requisi��o da API VIACEP ------------------'
With api
        .Open "GET", url
        .send
End With

'--------------------- Resposta da resquisi��o da API ---------------------------------------------'
html.body.innerHTML = api.responseText

On Error GoTo Trata_erro:
    
    Consulta_Cep.txtEndereco.Value = html.getElementsByTagName("logradouro")(0).innerText
    Consulta_Cep.txtComplemento.Value = html.getElementsByTagName("complemento")(0).innerText
    Consulta_Cep.txtBairro.Value = html.getElementsByTagName("bairro")(0).innerText
    Consulta_Cep.txtCidade.Value = html.getElementsByTagName("localidade")(0).innerText
    Consulta_Cep.txtUF.Value = html.getElementsByTagName("uf")(0).innerText

Exit Sub

'---------------------------- Tratamento de Erros - CEP inv�lido / n�o encontrado ------------------'
Trata_erro:
    Consulta_Cep.txtLogradouro.Value = ""
    Consulta_Cep.txtEndereco.Value = ""
    Consulta_Cep.txtComplemento.Value = ""
    Consulta_Cep.txtBairro.Value = ""
    Consulta_Cep.txtCidade.Value = ""
    Consulta_Cep.txtUF.Value = ""
    MsgBox "Cep inv�lido. Por favor, digite um cep v�lido!", vbCritical, "Aten��o!"

End Sub

Sub Exibir_Userform()

Consulta_Cep.Show

End Sub


