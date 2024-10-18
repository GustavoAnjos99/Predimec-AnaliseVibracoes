# Análise de vibrações Predimec®️
Projeto que automatiza o desenvolvimento e formatação dos relatórios de análise de vibrações da [Predimec®️](https://predimec.com.br)
<br><br>
![predimecAV-inicialização](https://github.com/user-attachments/assets/f7916903-a2fa-4b44-b728-f08af2a7836e)
## Sobre o projeto
A Predimec realiza o trabalho de análise de vibrações em seus clientes, porém, a realização dos relatórios desses trabalhos acaba sendo demorada, atrasando todos que fazem este relatório. Portanto, foi demandado a criação de uma aplicação para automatizar o desenvolvimento e formatação destes relatórios surgindo este programa.

Os relatórios são gerados por um outro programa onde o mesmo adiciona os dados das análise de vibrações em um arquivo Word, a partir disso é onde começa a formatação para criar os relatórios Predimec. Colocando o arquivo Word com os dados e o Excel para criação de tabelas e gráficos, é só executar a aplicação para ter os relatórios prontos.
### Funcionamento
O projeto todo foi realizado em Python, utilizando as bibliotecas:
- [python-docx](https://python-docx.readthedocs.io/en/latest/) para manipulação dos arquivos Word.
- [openpyxl](https://openpyxl.readthedocs.io/en/stable/) para manipulação dos arquivos Excel.
- [pywin32](https://pypi.org/project/pywin32/) para funções específicas e gerais do sistema.
- [pyinstaller](https://pyinstaller.org/en/stable/) para criação do arquivo .exe da aplicação.
Para o funcionamento do projeto, primeiramente é obrigatório ser colocado o arquivo WORD na mesma pasta do arquivo EXCEL, onde será executado o `main.exe` analisando e fazendo alterações nos arquivos conforme o padrão de relatórios Predimec. Após o programa executar, na pasta raiz do projeto vai ser criada a pasta "RELATÓRIOS FORMATADOS", onde conterá os arquivos finalizados.
> Para mais informações de uso, basta consultar o `manual da aplicação.html`.

Quais são as alterações feitas?
> Para fins de explicação e didática foi usado nomes genéricos de equipamentos/nome da empresa/área etc.

**1° Passo:** Primeiro, é corrigida a tabela de Status dos equipamentos, onde mostra o status da última visita feita pela Predimec:
![tabelaStatus-np-p](https://github.com/user-attachments/assets/8a2eac86-64fc-4cdc-9191-76458ce67da4)
Funções do programa nesse passo:
- Formata coluna "status", onde são trocados pro exemplo, N -> Normal e assim em diante;
- Formata primeira coluna como cabeçalho;
- Exclui colunas desnecessárias;
- Identifica "TAG" do equipamento e adiciona o nome dele ao lado;
- Adiciona coluna OS onde são contados as ordens de serviços;

**2° Passo:** Após isso, o próximo passo do sistema é consertar e formatar as tabelas de Ordem de Serviço:

[no image]

**3° Passo:** Seguindo, agora no Excel, é pegado os dados das tabelas anteriores e colocados nos campos desejados para criação dos gráficos:
![graficosExcel-np-p](https://github.com/user-attachments/assets/22cf74a7-6e64-4296-a804-7559a4bd643b)
Funções do programa nesse passo:
- Pega os valores da coluna "Status" do **1° Passo** e adiciona a fonte de dados do primeiro gráfico;
- Identifica as falhas encontradas e o Status em cada Ordem de serviço e soma os mesmos para criação do segundo gráfico;
- Adiciona a data e dados da nova análise de vibrações e incrementa no gráfico de tendência;
Por fim, a aplicação copia os gráficos e coloca-os no arquivo Word, seguindo as marcações:
![graficosWord-np-p](https://github.com/user-attachments/assets/5fc68a54-73cb-4326-a8a5-000a5b48c05f)
Funções do sistema nesse passo:
- Captura todos os gráficos gerados no Excel e adiciona-os no Word;
### Pré-formatação de arquivos
Os arquivos, tanto Word quanto Excel terão que ser "pré-formatados" para análise e captura de pontos específicos ao sistema, como por exemplo, aonde o sistema identifica a posição de uma tabela/gráfico.
As marcações são:
- "[grafico1], [grafico2], [grafico3]"
	- Usado no WORD.
	- Posições onde os respectivos gráficos do EXCEL serão inseridos no WORD.
## Código e Lógica
O sistema inteiro é dividido em três arquivos que contém todas as funções necessárias para executação do sistema.
### Arquivo: `main.py`
Este é o arquivo principal que engloba todas as funções dos outros arquivos, este é o arquivo que é executado ao iniciar o programa, nele contém:
- Importações de todas as funções de `functions_WORD.py` e `functions_EXCEL.py`.
- Funções para abrir e ler os arquivos .docx e .xlsm ou .xlsx .
- Funções para lógica de fotos dos gráficos.
- Funções para salvamento dos arquivos .docx e .xlsm ou .xlsx .
#### Funções definidas
| Função | parâmetros | Finalidade |
|--------|--------|------------|
| `pegarGraficosExcel()` |`app`,<br> `workbook_file_name`,<br> `workbook` | Pega os gráficos do excel, tranforma-os em imagens e coloca no diretório do projeto. 
| `excluirImagensPATH()` | `imagem`| Excluir imagens do diretório do projeto após serem inseridas no arquivo WORD.
#### Principais variáveis definidas
| Nome | Valor |
|------|-------|
|`documentoWord`| Documento .docx que será formatado.|
|`documentoExcel`| Documento excel que será formatado.|
|`tabela`| tabela de listagem e status dos equipamentos (arquivo word)|
|`totLinhas`| quantidade de linhas da `tabela`|
|`totColunas`| quantidade de colunas da `tabela`|
|`tabelasCount`| quantidade de tabelas no `documentoWord`|
|`planilhaListagem`| Planilha entitulada "Listagem" do `documentoExcel`|
|`planilhaGraficos`| Planilha entitulada "Gráficos do `documentoExcel`|
### Arquivo `funcitons_EXCEL.py`
Este é o arquivo que contém todas as funções necessárias para manipulação do arquivo EXCEL, sobre suas responsabilidades ele:
- Atualiza todas as fonte de dados dos gráficos EXCEL.
- Formula os gráficos para manda-los ao arquivo WORD.
#### Funções definidas
| Função | parâmetros | Finalidade |
|--------|--------|------------|
| `addColunaListagem()` | `valores`,<br> `pagina`,<br> `pagina2` | Atualiza a fonte de dados para geração do gráfico de status (Gráfico de pizza) |
| `arrumarTabela_2()` | `pagina`,<br> `valores` | Soma a contagem das falhas e seu grau de estabilidade (Gráfico de Colunas) |
| `arrumarTabela_3()` | `pagina` | Adiciona mais uma data ao gráfico de tendência.|
|`substCelulaTBL3()` | `pag`,<br> `coluna` | adiciona valores as colunas especificadas em `arrumarTabela_3()`|
|`retornarMesAno()` | `data_original` | Recebe uma data no formato dd/mm/aaaa e retorna a data formatada em MM/AAAA ex.: Out/2024. |

> Este arquivo não possui nenhuma variável definida.

### Arquivo `functions_WORD`
Neste arquivo está contido todas as funções e variáveis para formatação e manipulação do arquivo WORD, sua finalidade é:
- Adicionar data correta a capa do relatório.
- Adicionar ao documento WORD todos os 3 gráficos do arquivo EXCEL.
- Formatar a tabela de listagem dos itens.
- Formatar as tabelas de ordem de serviço.
#### Funções definidas
| Função | parâmetros | Finalidade |
|--------|--------|------------|
|`WORD_arrumarAbreviacoes()`|`tabela`,<br> `index`| Corrige as abreviações dos status da tabela de listagem ex.: N -> Normal |
|`WORD_arrumarOS()`|`tabela`,<br> `totLinhas`| Contabiliza em ordem as OS da tabela de listagem. |
|`WORD_arrumarCounts()`| `count`| Corrige a formatação de números de apenas 1 dígito ex.: 2 -> 02|
|`WORD_retornarData()`| - | Retorna a data (hoje) no formato dd/mm/aaaa|
|`WORD_arrumarTabelOS_equipamento()`|`documento`, <br> `tabelasCount`,<br>| Corrige abreviações, número de OS e formatação de todas as tabelas OS |
|`WORD_deletarColuna()`|`documento`,<br> `table`,<br> `columns`,<br>| Deleta a coluna da tabela especificada nos parâmetros.|
|`WORD_formatarCelula()`|`celula`| Adiciona negrito e alinhamento da celula especificada no parâmetro.|
|`WORD_formatarCabecalho()`|`celula`|Adiciona negrito, alinhamento e preenchimento na celula especificada no parâmetro.|
|`WORD_formatarData()`|`celula`|Formata a celula da data na capa do documento.|
|`WORD_addCabecalhoVertical()`|`tabela`,<br> `totLinhas`|Formata cabeçalho da primeira coluna da tabela de listagem utilizando o  `WORD_formatarCabecalho()`|
|`WORD_colunaValores()`|`tabela`,<br> `index`| Retorna um array de todos os itens da coluna da tabela especificada nos parâmetros.|
|`WORD_arrumarEquipamentoTabela()`|`tabela`,<br> `totLinhas`| Adiciona à tabela de listagem o nome dos equipamentos referentes a TAG do equipamento.|
|`WORD_identifiarDefeito()`|`documento`,<br> `tabelasCount`| Retorna um array com os defeitos e o grau de estabilidade encontrados nas tabelas OS. Utilizado em `arrumarTabela_2()`|
|`WORD_addGraficos()`|`paragrafo`,<br> `nm`| Adiciona no arquivo WORD os gráficos recebidos do `pegarGraficosExcel()`|
#### Variáveis deifinidas
| Nome | Valor |
|------|-------|
|`equipamentos`| Dict que contém a lista de TAG e nome para correção dos nomes do equipamento|
|`defeitos`| Array que contém a lista de falhas|
---
<br>

