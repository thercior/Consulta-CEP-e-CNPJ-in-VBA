Attribute VB_Name = "BuscaCNPJ"
Option Explicit

''---------------------- M�DULO PARA CONSULTAR CNPJ VIA API ------------------------------------'
'------ Declara��o de vari�veis j� dentro o m�dulo, sendo trazidas do Formul�rio. Onde _
 optional � pq s�o vari�vels opcionais ---------------------------------------------------------'
 
Public Sub BuscarCNPJ(Cnpj$, RazaoSocial$, NomeFantasia$, _
    Endereco$, numero$, complemento$, bairro$, _
    cidade$, uf$, cep$, tel$, Email$ _
    , Optional porte$, Optional DtAbertura$, Optional Situacao$ _
    , Optional NaturJuridica$, Optional AtivPrincipal$ _
    , Optional CNAEPrincipal$, Optional Capital$, Optional AtivSecundarias As Variant, _
    Optional Socios As Variant, Optional ArrSecundarias As Variant, Optional ArrSocios As Variant)

Dim url As String
Dim json As Object, ws As Object, Lista As Range
Dim http As New MSXML2.XMLHTTP60
Dim i%

ThisWorkbook.Activate
Set ws = Planilha2
Set Lista = ws.Range("A2:E100")

Lista.Clear
'--------------------------- Arrumar o campo CNPJ ----------------------------------------------'

Cnpj = Replace(Replace(Replace(Cnpj, ".", ""), "/", ""), "-", "")


'--------------------------- Defini��o do site fornecedor da API para CNPJ ---------------------'
url = "https://receitaws.com.br/v1/cnpj/" & Cnpj

'--------------------------- Abertura de conex�o com o site/API para CNPJ ----------------------'
With http
    .Open "GET", url, False
    .send
End With

'----------- Variavel/fun��o objeto que converte o JSON para funcionar no VBA ------------------'
Set json = JsonConverter.ParseJson(http.responseText)

''-------------------------------- Dados B�sico da empresa --------------------------------------'
On Error Resume Next
RazaoSocial = json("nome")
NomeFantasia = json("fantasia")
Endereco = json("logradouro")
numero = json("numero")
complemento = json("complemento")
bairro = json("bairro")
cidade = json("municipio")
uf = json("uf")
tel = json("telefone")
Email = json("email")
cep = json("cep")

'-------------------------------- Informa��es da empresa ---------------------------------------'
porte = json("porte")
DtAbertura = json("abertura")
Situacao = json("status")
NaturJuridica = json("natureza_juridica")
AtivPrincipal = json("atividade_principal")(1)("text")
CNAEPrincipal = json("atividade_principal")(1)("code")
With ws
    If .Range("A1") = "" And .Range("B1") = "" Then
    .Range("A1") = "CNAE Secund�rio"
    .Range("B1") = "Atividades Secund�rias"
    End If
    i = 2
    For Each AtivSecundarias In json("atividades_secundarias")
        ws.Cells(i, 1) = AtivSecundarias("code")
        ws.Cells(i, 2) = AtivSecundarias("text")
        i = i + 1
    Next
'------------------------------- Quadro social da empresa --------------------------------------'
     If .Range("D1") = "" And .Range("D1") = "" Then
    .Range("D1") = "S�cio"
    .Range("D1") = "Cargo do S�cio"
    End If
    
    i = 2
    For Each Socios In json("qsa")
        .Cells(i, 4) = Socios("nome")
        .Cells(i, 5) = Socios("qual")
        i = i + 1
    Next
    .Range("A:E").EntireColumn.AutoFit
End With
Capital = json("capital_social")

''------------------------ Boas pr�ticas de VBA: destruir os objetos ---------------------------'
Set ws = Nothing
Set Lista = Nothing
Set json = Nothing
End Sub

Sub ChamaCNPJ()
ConsultaCNPJ.Show
End Sub

