# Automação de Organização de Horários por Instrutor

Este projeto automatiza a organização de horários de instrutores a partir de uma planilha Excel. O script lê uma planilha com os dados dos instrutores, cria ou atualiza abas separadas para cada instrutor e preenche as informações relevantes de acordo com os dados fornecidos.

## Funcionalidades

- **Leitura de dados de instrutores**: O script lê a coluna de instrutores da planilha original.
- **Criação de abas individuais**: Para cada instrutor encontrado, é criada uma aba dedicada (caso ainda não exista).
- **Preenchimento de informações**: Cada aba contém os dados de "Data", "Horário de Início", "Horário de Fim" e "Turma", extraídos da planilha original.
- **Evita duplicidade de abas**: Se uma aba já existe para um instrutor, o script adiciona novas informações na mesma aba, sem duplicá-la.

## Estrutura do Código

O código é dividido em funções modulares, facilitando a manutenção e reutilização:

1. **`carregar_planilha(caminho_arquivo)`**: Carrega a planilha Excel original.
2. **`limpar_nome_aba(nome)`**: Sanitiza os nomes dos instrutores, removendo caracteres inválidos e limitando o tamanho do nome da aba a 31 caracteres.
3. **`obter_ou_criar_aba(workbook, nome_aba)`**: Cria uma aba para o instrutor caso ainda não exista, ou retorna a aba já existente.
4. **`adicionar_titulos(aba)`**: Adiciona os títulos das colunas na aba do instrutor (Data, Horário de Início, Horário de Fim, Turma).
5. **`preencher_aba(aba, dados)`**: Preenche a aba com os dados do instrutor.
6. **`extrair_dados_linha(sheet, linha)`**: Extrai os dados de uma linha específica na planilha original.
7. **`processar_planilha(nome_arquivo)`**: Função principal que executa todo o fluxo, de carregar a planilha até salvar e abrir o arquivo atualizado.

## Como Usar

1. **Pré-requisitos**:
   - Python 3.x
   - Biblioteca `openpyxl`
   
   Para instalar o `openpyxl`, use o seguinte comando:
   ```bash
   pip install openpyxl

## Passos para Execução:

Coloque o arquivo da planilha Excel no diretório especificado no código.
Execute o script, que irá automaticamente organizar as informações de acordo com o nome dos instrutores.
A planilha será salva com as abas criadas/atualizadas para cada instrutor.

## Modificando o Caminho da Planilha:
Certifique-se de alterar o valor da variável NOME_ARQUIVO no script para o caminho correto do seu arquivo Excel.
