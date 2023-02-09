# **CONSULTA DE CEP E CNPJ VIA API UTILIZANDO VBA**
Macro para Consulta de CEP e CNPJ via API utilizando o Visual Basic For Applications (VBA)

## **Pré-requisitos**
Api Windows para ajustar dimensionamento, transparência, adição de botões maximizar/minimizar do MSForms, ativação de Referências/Bibliotecas, utilização de Módulos de Classes para efeito place holder.

## **Ativação de referências**
Ativar as seguintes bibliotecas/referências da Guia Ferramentas > Referências:

  - *Microsoft HTML Object Library*
  - *Microsoft XML, v6*
  - *Microsoft Scripting RunTime*

## **Conversor formato JSON para dicionário lido pelo VBA**
Foi utilizado a biblioteca VBA-JSON, no módulo JsonConverter, para converter o formato JSON da resposta da requisição para um formato de dicionário que fosse lido pelo VBA.
A biblioteca encontra-se no módulo JsonConverter.
O VBA-JSON foi desenvolvido e disponibilizado por Tim Hall no seguinte repositório:
<https://github.com/VBA-tools/VBA-JSON>

## **Módulos para efeito placeholder nas textbox**
Foi utilizado os módulos de classe *clsTextBox* e módulo *mdltextbox* para colocar o efeito de placeholder nos textbox.
Módulos desenvolvidos por Ricado Camisa e disponível em: 
<https://github.com/ricardocamisa/clsTextBox>

## **Módulos para dimensionar, ajustar e modificar o MSForms**
  - *ModificaForms: adiciona botões maximizar e minimizar, e transparência ao formulário*
  - *RedimensionaForms: ajusta as dimensões do MSForm e todos os componentes, seguindo proporção do ajuste e do tamanho da tela*

Módulos desenvolvidos em conjunto com aula do curso Programando Excel em VBA de Marcelo do Nascimento, disponível em:
<https://hotmart.com/pt-br/marketplace/produtos/programando-o-excel-com-vba/S70500759S>

## **Utilização**
Tanto o consulta CEP como consulta CNPJ utiliza-se o número respectivo para realizar a buscar e retornar as informações especificadas.
Nesta versão, a busca ocorre apenas verificando os respectivos números de CEP e CNPJ.
Basta apenas digitar os números, sem formatos. Os campos possuem máscaras para formatação automaticamente.
