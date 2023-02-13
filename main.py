import requests
from bs4 import BeautifulSoup
import locale  #Transforma o sistema de metricas utilizado

from openpyxl.workbook import workbook, Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from tabulate import tabulate
from modelos import FundoImobiliario, Estrategia
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8') #Transforma todo o sistema no sistema americano



def trata_porcentagem(porcentagem_str):
    """
    Funcao para retirar o sinal de porcentagem do valor ex: 7,04% a str sera quebrada no % criando duas posicoes
    onde usaremos a posicao 0 do array posicao 0 7.04 posicao 1 > None por quebrar no final da str
    :param porcentagem_str: Recebe uma string ex: 7,04% vai quebrar a string separando os numeros do %
    :return: Retorna apenas os numeros da str passada atribuindo na posicao 0 do valor. ex: 7,04
    """
    return locale.atof(porcentagem_str.split('%')[0])


def trata_decimal(decimal_str):
    """
    Funcao usera o metodo atof para realizar a transformacao
    :param decimal_str: Recebe os dados e transofmra para o sistema americano
    :return: retorna sem o R$ e espaco
    """

    return locale.atof(decimal_str)


headers = {'User-Agent': 'Mozila/5.0'} #usado para nao apresentar erro no debbuger
resposta = requests.get('https://www.fundamentus.com.br/fii_resultado.php', headers=headers)

#A beautiul soup pega o retorno do requests, variavel resposta. Dentro dela tem todo conteudo HTML e dentro dele uma variavel
#text que armazena o conteudo que precisamos para fazer o webscraping
soup = BeautifulSoup(resposta.text, 'html.parser') #html parser eh um padrao da documentacao do beautifulsoup (ver)
linhas = soup.find(id='tabelaResultado').find('tbody').find_all('tr') #tabelaResultado esta contido dentro text

resultado = []

#Objeto de classe para ser usado como parametro para pesquisar no fundo
estrategia = Estrategia( #Criando objeto do tipo estrategia
    cotacao_atual_minima=20.0,
    dividend_yield_minimo=8,
    p_vp_minimo=0.50,
    valor_mercado_minimo=50000000,
    liquidez_minima=40000,
    qt_minima_imoveis=5,
    maxima_vacancia_media=5
)
for linha in linhas: #Loop criado para iterar sobre as linhas e poder fazer a 'raspagem' dos dados desejados
    dados_fundo = linha.find_all('td') #buscando pelas linhas onde contem os dados desejados para raspagem
    codigo = dados_fundo[0].text #Atribuindo para cada linha do cabealho uma variavel
    segmento = dados_fundo[1].text
    cotacao = trata_decimal(dados_fundo[2].text)
    ffo_yield = trata_porcentagem(dados_fundo[3].text)
    dividend_yield = trata_porcentagem(dados_fundo[4].text)
    p_vp = trata_decimal(dados_fundo[5].text)
    valor_mercado = trata_decimal(dados_fundo[6].text)
    liquidez = trata_decimal(dados_fundo[7].text)
    qt_imoveis = int(dados_fundo[8].text)
    preco_m2 = trata_decimal(dados_fundo[9].text)
    aluguel_m2 = trata_decimal(dados_fundo[10].text)
    cap_rate = trata_porcentagem(dados_fundo[11].text)
    vacancia = trata_porcentagem(dados_fundo[12].text)
#Criando objeto do tipo FundoImobiliario
    fundo_imobiliario = FundoImobiliario(codigo, segmento, cotacao, ffo_yield, dividend_yield, p_vp, valor_mercado,
                                         liquidez, qt_imoveis, preco_m2, aluguel_m2, cap_rate, vacancia)

#Usando o objeto estrategia para aplicar o metodo aplica_estrategia passando como parametro o obj fundo_imobi
    if estrategia.aplica_estrategia(fundo_imobiliario): #Se o resultado for True aplica append
        resultado.append(fundo_imobiliario)

#Usando Tabulate
cabecalho = ['CODIGO', 'SEGMENTO', 'COTACAO ATUAL', 'DIVIDEND YIELD']
tabela = []
#For para iterar sobre a tabela criada em resultado pegando cada elemento escolhido no cabecalho
for elemento in resultado:
    tabela.append([
        elemento.codigo,
        elemento.segmento,
        locale.currency(elemento.cotacao_atual),
        f'{locale.str(elemento.dividend_yield)}%'
    ])

resultado_tabulate = (tabulate(tabela, headers=cabecalho, tablefmt='rounded_outline', numalign="right"))
print(f'{soup.title.string:>47}')
print(resultado_tabulate)

#Adicionar dados para uma planilha
excel = Workbook()
planilha = excel.active
planilha.title = 'Tabela'
planilha.append(cabecalho)
indice = 2
for l in tabela:
    planilha.append(l)
    indice += 1

#Define onde ficara a tabela
tab = Table(displayName='Tabela', ref='A1:D14')

#Define estilo da tabela
estilo = TableStyleInfo(name='TableStyleMedium9', showFirstColumn=False,
                        showLastColumn=False, showRowStripes=True, showColumnStripes=True)

#Atribui o estilo ao obj criado
tab.tableStyleInfo = estilo

#Adiciona a tabela na planilha
planilha.add_table(tab)

excel.save('./planilha/Investimento.xlsx')
