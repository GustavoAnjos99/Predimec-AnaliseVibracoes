from functions_WORD import WORD_retornarData

def addColunaListagem(valores, pagina, pagina2):
    pagina.delete_cols(5)
    totalValores = len(valores)
    for i in range(0, totalValores):
        pagina[f"E{i+1}"].value = valores[i]
    pagina2['N12'] = ''

def arrumarTabela_2(pagina, valores):
    for rows in pagina['N28:P34']:
        for cell in rows:
            cell.value = 0
    for statusDefeito in valores:
        linha = ''
        coluna = ''
        if statusDefeito[0] == "Aceitável":
            coluna = "N"
        elif statusDefeito[0] == "Alerta":
            coluna = "O"
        elif statusDefeito[0] == "Crítico":
            coluna = "P"

        for i in statusDefeito[1]:
            if i == "DESBALANCEAMENTO":
                linha = "28"
            elif i == "DESALINHAMENTO":
                linha = "29"
            elif i == "FOLGAS" or i == "FOLGA":
                linha = "30"
            elif i == "BASE":
                linha = "31"
            elif i == "LUBRIFICAÇÃO":
                linha = "32"
            elif i == "ROLAMENTO":
                linha = "33"
            elif i == "OUTROS":
                linha = "34"
        pagina[f"{coluna}{linha}"].value += 1

def arrumarTabela_3(pagina):
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
    pag[f"{coluna}47"] = retornarDataMes(WORD_retornarData())
    pag[f"{coluna}48"] = pag['N8'].value
    pag[f"{coluna}49"] = pag['N9'].value
    pag[f"{coluna}50"] = pag['N10'].value
    pag[f"{coluna}51"] = pag['N11'].value
    pag[f"{coluna}52"] = pag['N12'].value

def retornarDataMes(data):
    dataArray = data.split("/")
    mesData = ''
    if dataArray[1] == "01":
        mesData = f"jan/{dataArray[2]}"
    if dataArray[1] == "02":
        mesData = f"fev/{dataArray[2]}"
    if dataArray[1] == "03":
        mesData = f"mar/{dataArray[2]}"
    if dataArray[1] == "04":
        mesData = f"abr/{dataArray[2]}"
    if dataArray[1] == "05":
        mesData = f"mai/{dataArray[2]}"
    if dataArray[1] == "06":
        mesData = f"jun/{dataArray[2]}"
    if dataArray[1] == "07":
        mesData = f"jul/{dataArray[2]}"
    if dataArray[1] == "08":
        mesData = f"ago/{dataArray[2]}"
    if dataArray[1] == "09":
        mesData = f"set/{dataArray[2]}"
    if dataArray[1] == "10":
        mesData = f"out/{dataArray[2]}"
    if dataArray[1] == "11":
        mesData = f"nov/{dataArray[2]}"
    if dataArray[1] == "12":
        mesData = f"dez/{dataArray[2]}"
    return mesData