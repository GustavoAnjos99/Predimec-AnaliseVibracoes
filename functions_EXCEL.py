from functions_WORD import WORD_retornarData
import datetime

def addColunaListagem(valores, pagina, pagina2):
    print(f"🛠 -> Adicionando listagem ao EXCEL...")
    pagina.delete_cols(5)
    totalValores = len(valores)
    for i in range(0, totalValores):
        pagina[f"E{i+1}"].value = valores[i]
    pagina2['N12'] = ''
    print(f"✔ -> Listagem adicionada ao EXCEL!")


def EXCEL_arrumarTabela_2(pagina, valores):
    print(f"🛠 -> Adicionando falhas identificadas ao gráfico...")
    count = 1
    for i in range(0, len(valores)):
        for j in valores[i][1]:
            pagina[f'J{count}'].value = ""
            pagina[f'K{count}'].value = ""
            pagina[f'J{count}'].value = valores[i][0]
            pagina[f'K{count}'].value = j.capitalize()
            count += 1
    print(f"🛠 -> Falhas identificadas adicionadas ao gráfico!")

def EXCEL_arrumarTabela_3(pagina):
    print(f"🛠 -> Adicionando gráfico de tendência...")
    if not pagina['N47'].value:
        substCelulaTBL3(pagina, "N")
        return
    elif not pagina['O47'].value:
        substCelulaTBL3(pagina, "O")
        return
    elif not pagina['P47'].value:
        substCelulaTBL3(pagina, "P")
        return
    elif not pagina['Q47'].value:
        substCelulaTBL3(pagina, "Q")
        return
    elif not pagina['R47'].value:
        substCelulaTBL3(pagina, "R")
        return
    elif pagina['R47'].value:
        pagina.move_range("O47:R52", rows=0, cols=-1)
        substCelulaTBL3(pagina, "R")
        return

def substCelulaTBL3(pag, coluna):
    pag[f"{coluna}47"] = retornarMesAno(WORD_retornarData())
    pag[f"{coluna}48"] = pag['N8'].value
    pag[f"{coluna}49"] = pag['N9'].value
    pag[f"{coluna}50"] = pag['N10'].value
    pag[f"{coluna}51"] = pag['N11'].value
    pag[f"{coluna}52"] = pag['N12'].value

def retornarMesAno(data_original):
    data_obj = datetime.datetime.strptime(data_original, "%d/%m/%Y")
    data_formatada = data_obj.strftime("%b/%Y")
    return data_formatada

def EXCEL_corrigirFormulas(planilhagraficos):
    print(f"🛠 -> Corrigindo formulas EXCEL...")
    planilhagraficos['N8'] = '=COUNTIF(Listagem!$E:$E, Gráficos!M8)'
    planilhagraficos['N9'] = '=COUNTIF(Listagem!$E:$E, Gráficos!M9)'
    planilhagraficos['N10'] = '=COUNTIF(Listagem!$E:$E, Gráficos!M10)'
    planilhagraficos['N11'] = '=COUNTIF(Listagem!$E:$E, Gráficos!M11)'
    planilhagraficos['N12'] = '=COUNTIF(Listagem!$E:$E, Gráficos!M12)'
    
    colunas = ["N", "O", "P"]
    for col in colunas:
        for i in range(28, 34):
            planilhagraficos[f"{col}{i}"] = f"=COUNTIFS(Listagem!$J:$J, Gráficos!${col}$27, Listagem!$K:$K, Gráficos!$M{i})"
    
    print(f"✔ -> Formulas corrigidas!")