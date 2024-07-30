#

# Google Sheets

Esse script foi projetado para ser usado como um gatilho `onEdit()` no `Planilhas Google`. 

Quando um usuário faz uma edição em uma célula em uma planilha com o nome `Página1`, o script fará o seguinte:


- Obtenha o índice de linha ativa da célula editada.

- Se o índice de linha ativa for maior que 1 (o que significa que não é a linha de cabeçalho), ele obterá o valor da célula na coluna 3 `CelulaData` da linha ativa.

- Se `CelulaData` for uma cadeia de caracteres vazia, ele obterá o valor da célula na coluna 2 `Email` da linha ativa.

- Se `Email` não for uma cadeia de caracteres vazia, ele gerará uma nova data e a definirá como o valor da célula na coluna 3 da linha ativa.

- Em seguida, ele entrará em um loop, a partir da próxima linha abaixo da linha ativa, e repetirá o processo até encontrar uma linha em que os valores das colunas `1, 2 e 3` estejam todos vazios.

#

## Criando o Evento

Navegue até `Acionadores` 

Clique em `Crie um novo acionador` 

- Escolha a função que será executada `onEdit`

- Escolha qual implantação deve ser executada `teste`

- Selecione a origem do evento `Da planilha`

- Selecione o tipo de evento `Ao alterar`

- Clique em `Salvar`

##

Espero que isso ajude a esclarecer o propósito do código!
