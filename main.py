from docx import Document
from openpyxl import load_workbook
from functions_WORD import *
from functions_EXCEL import *
import os
import pathlib
from win32com.client import Dispatch
from docx.shared import Inches

print(r"""
________            _____________                           __________  
___  __ \_________________  /__(_)______ _______________    ___  ___  \ 
__  /_/ /_  ___/  _ \  __  /__  /__  __ `__ \  _ \  ___/    __  / _ \  |
_  ____/_  /   /  __/ /_/ / _  / _  / / / / /  __/ /__      _  / , _/ / 
/_/     /_/    \___/\__,_/  /_/  /_/ /_/ /_/\___/\___/      | /_/|_| /  
                                                             \______/   
Iniciando processo de formatação...
      """)
ARQUIVO_WORD = ''
ARQUIVO_EXCEL = ''

arquivos = os.listdir('./')
for arquivo in arquivos:
    if arquivo.endswith(".docx"):
        ARQUIVO_WORD = arquivo
    if arquivo.endswith(".xlsm") or arquivo.endswith(".xlsx"):
        ARQUIVO_EXCEL = arquivo

f = open(ARQUIVO_WORD, 'rb')
g = open(ARQUIVO_EXCEL, 'rb')
documentoWord = Document(f)
documentoExcel = load_workbook(g)

## WORD ===================================================================================================================
tabela = documentoWord.tables[1]
totLinhas = len(tabela.rows)
totColunas = int(len(tabela.columns) - (len(tabela.columns) / 2))
tabelasCount = len(documentoWord.tables)


WORD_formatarData(documentoWord.tables[0].columns[0].cells[0])
WORD_deletarColuna(documentoWord,1,0)
WORD_deletarColuna(documentoWord,1,2)
WORD_deletarColuna(documentoWord,1,6)
WORD_deletarColuna(documentoWord,1,6)
WORD_arrumarTabelaOS_equipamento(documentoWord, tabelasCount)
WORD_arrumarEquipamentoTabela(tabela, totLinhas)
WORD_addCabecalhoVertical(tabela, totLinhas)
WORD_verificarOS_titulo(documentoWord)
for i in range(0, totColunas):
    WORD_arrumarAbreviacoes(tabela, i)
WORD_arrumarOS(tabela, totLinhas)

## EXCEL ==================================================================================================================
planilhaListagem = documentoExcel['Listagem']
planilhaGraficos = documentoExcel['Gráficos']

addColunaListagem(WORD_colunaValores(tabela, 4), planilhaListagem, planilhaGraficos)
arrumarTabela_2(planilhaGraficos, WORD_indentificarDefeito(documentoWord, tabelasCount))
arrumarTabela_3(planilhaGraficos)


## SALVAMENTO DO ARQUIVO ==================================================================================================
PASTA_RESULTADOS = "RELATÓRIOS FORMATADOS"
os.makedirs(PASTA_RESULTADOS, exist_ok=True)

teste = ARQUIVO_EXCEL.split('.')
ARQUIVO_EXCEL = f"{teste[0]}.xlsx"
caminhoWord = os.path.join(PASTA_RESULTADOS, ARQUIVO_WORD)
caminhoExcel = os.path.join(PASTA_RESULTADOS, ARQUIVO_EXCEL)
documentoExcel.save(caminhoExcel)


app = Dispatch("Excel.Application")
workbook_file_name = rf"{str(pathlib.Path().resolve())}\RELATÓRIOS FORMATADOS\{ARQUIVO_EXCEL}"
workbook = app.Workbooks.Open(Filename=workbook_file_name)

app.DisplayAlerts = False

for i, sheet in enumerate(workbook.Worksheets):
    for chartObject in sheet.ChartObjects():
        chartObject.Chart.Export(rf"{str(pathlib.Path().resolve())}\chart{str(i+1)}.png")
        i +=1
workbook.Close(SaveChanges=False, Filename=workbook_file_name)

for i in documentoWord.paragraphs:
    if i.text == "[gráfico1]":
        i.text = ''
        i.alignment = WD_ALIGN_PARAGRAPH.CENTER
        img = i.add_run()
        img.add_picture(rf"{str(pathlib.Path().resolve())}\chart1.png", width=Inches(4))
    if i.text == "[gráfico2]":
        i.text = ''
        i.alignment = WD_ALIGN_PARAGRAPH.CENTER
        img = i.add_run()
        img.add_picture(rf"{str(pathlib.Path().resolve())}\chart2.png", width=Inches(4))
    if i.text == "[gráfico3]":
        i.text = ''
        i.alignment = WD_ALIGN_PARAGRAPH.CENTER
        img = i.add_run()
        img.add_picture(rf"{str(pathlib.Path().resolve())}\chart3.png", width=Inches(4))
documentoWord.save(caminhoWord)

if os.path.exists("chart1.png"):
  os.remove("chart1.png")
if os.path.exists("chart2.png"):
  os.remove("chart2.png")
if os.path.exists("chart3.png"):
  os.remove("chart3.png")

print("\nArquivos formatados com sucesso!\n")