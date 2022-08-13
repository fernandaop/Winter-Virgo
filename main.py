import urllib.request
import PySimpleGUI as gui
import win32com.client as win32
import json
import sqlalchemy as sql
import datetime

ssl_args = {'ssl_ca': "DigiCertGlobalRootCA.crt.pem"}
engine =sql.create_engine('mysql+pymysql://fernanda.pereira:2uwHbaCukUl&yWYUXxkC@vdwh.mysql.database.azure.com:3306/dwh', connect_args=ssl_args)
conn = engine.connect()
response = engine.execute('select cra_sch.ticker_symbol, cra_inst.corporation_name, cra_sch.event_name, cra_sch.event_unit_price, cra_sch.payment_date from dwh.vw_up2data_fixed_income_cra_schedule as cra_sch join dwh.vw_up2data_fixed_income_cra_instrument as cra_inst on (cra_sch.ticker_symbol = cra_inst.ticker_symbol) where (cra_inst.corporation_name = "VIRGO COMPANHIA DE SECURITIZACAO") or (cra_inst.corporation_name = "VIRGO II COMPANHIA DE SECURITIZACAO") or (cra_inst.corporation_name = "ISEC SECURITIZADORA S.A"); ')
response1 = engine.execute('select cri_sch.ticker_symbol, cri_inst.corporation_name, cri_sch.event_name, cri_sch.event_unit_price, cri_sch.payment_date from dwh.vw_up2data_fixed_income_cri_schedule as cri_sch join dwh.vw_up2data_fixed_income_cri_instrument as cri_inst on (cri_sch.ticker_symbol = cri_inst.ticker_symbol) where (cri_inst.corporation_name = "VIRGO COMPANHIA DE SECURITIZACAO") or (cri_inst.corporation_name = "VIRGO II COMPANHIA DE SECURITIZACAO") or (cri_inst.corporation_name = "ISEC SECURITIZADORA S.A");')
def crai(response):
    crai= []
    for row in response:
        if row[4] == datetime.date.today():
            if row not in crai:
                crai.append(row)
    return crai
cra = crai(response)
cri = crai(response1)
b3 = cri + cra

def b3_sep(b3):
    dicionario = {}
    v_juros = []
    v_amort = []
    v_amex = []
    for dic in b3:
        id = dic[0]
        preco = str(round(dic[3],3))
        dicionario["id"] = id
        dicionario["tipo"] = dic[2]
        dicionario["pu"] = preco
        if dic[2] == 'PAGAMENTO DE JUROS':
            v_juros.append(dicionario)
            dicionario = {}
        elif dic[2] == "AMORTIZACAO":
            v_amort.append(dicionario)
            dicionario = {}
        elif dic[2] == "AMORTIZACAO EXTRAORDINARIA":
            v_amex.append(dicionario)
            dicionario = {}
    return v_juros, v_amort, v_amex
v_juros, v_amort, v_amex = b3_sep(b3)
v_b3 = v_juros + v_amort + v_amex

def url_galaxia():
    url = "https://redash.virgo.inc/api/queries/112/results.json?api_key=8vPmO96cK7hQ8mahDr6C4LMleuYLBBeZhi7fnwSP"
    response = urllib.request.urlopen(url)
    data = json.loads(response.read())
    return data

galaxia = url_galaxia()
resultg = galaxia["query_result"]["data"]["rows"]

def filter(g, b):
    filter= []
    for i in b:
        for j in g:
            if (i['id'] == j['id']):
                if j not in filter:
                    filter.append(j)
    return filter
filtro = filter(resultg, v_b3)

def juros(resultg, resultb):
    lista = []
    lista.append(resultg[0]["id"])
    if resultg[0]['juros'] == None:
        lista.append("Valor não adicionado")
    else:
        valorg = str(round(resultg[0]['juros'], 3))
        lista.append(valorg)
    lista.append(resultb[0]["pu"])
    return lista

juros = juros(filtro, v_juros)

def amorta(resultg, resultb):
    lista = []
    lista.append(resultg[0]["id"])
    if resultg[0]["amortizacao"] != None:
        valorg = str(round(resultg[0]["amortizacao"],3))
    else:
        valorg = "Valor não adicionado"
    if resultb != []:
        valorb = str(resultb[0]["pu"])
    else:
        valorb = "Valor não adicionado"
    lista.append(valorg)
    lista.append(valorb)
    return lista
amort = amorta(filtro, v_amort)

def amex(resultg, resultb):
    lista = []
    lista.append(resultg[0]["id"])
    if resultg[0]["amex"] != None:
        valorg = str(round(resultg[0]["amex"],3))
    else:
        valorg = "Valor não adicionado"
    if resultb != []:
        valorb = str(resultb[0]["pu"])
    else:
        valorb = "Valor não adicionado"
    lista.append(valorg)
    lista.append(valorb)
    return lista
amex = amex(filtro, v_amex)

headings = ['ID', 'VALOR GALÁXIA', 'VALOR B3']

def valuesf(g):
    listona = []
    lista = []
    for i in g:
        lista.append(i['titulo'])
        lista.append(i['id'])
        if i['juros'] == None:
            lista.append('Valor não aplicado no Galáxia')
        else:
            lista.append(str(i['juros']))
        if i['amortizacao']== None:
            lista.append('Valor não aplicado no Galáxia')
        else:
            lista.append(str(i['juros']))
        if i['amex'] == None:
            lista.append('Valor não aplicado no Galáxia')
        else:
            lista.append(str(i['amex']))
        lista.append(i['responsavel'])
        listona.append(lista)
    return listona
def valuesb(b3):
    lista = []
    for i in b3:
        lista.append(list(i))
    return lista
listag = valuesf(filtro)
listab = valuesb(b3)
def comp(filtro, v_b3, string):
    listona = []
    lista = []
    lista.append(v_b3[0]['id'])
    lista.append(str(v_b3[0]['pu']))
    lista.append(str(filtro[0][string]))
    listona.append(lista)
    return listona
compj = comp(filtro, v_juros, 'juros')

if v_amort != []:
    compm = comp(filtro, v_amort, 'amortizacao')
else:
    compm = [[filtro[0]['id'],'Sem alterações', 'Sem alterações' ]]

if v_amex != []:
    compa = comp(filtro, v_amex, 'amex')
else:
    compa = [[filtro[0]['id'],'Sem alterações', 'Sem alterações']]

def mensagem( compj, compm, compa):
    m1 = ''
    m2= ''
    m3 = ''
    if compj[0][1] != compj[0][2]:
        mensagem = ("Olá, {}!\n A PU de ID igual a {} se encontra divergente com a B3.".format(filtro[0]['responsavel'], filtro[0]["id"]))
        m1= 'JUROS DIVERGENTE/FALTANTE'
    if compm[0][1] != compm[0][2]:
        mensagem = ("Olá, {}!\n A PU de ID igual a {} se encontra divergente com a B3.".format(filtro[0]['responsavel'],filtro[0]["id"]))
        m2 = 'AMORTIZAÇÃO DIVERGENTE/FALTANTE'
    if compa[0][1] != compa[0][2]:
        mensagem = ("Olá, {}!\n A PU de ID igual a {} se encontra divergente com a B3.".format(filtro[0]['responsavel'],filtro[0]["id"]))
        m3 = "AMEX DIVERGENTE/FALTANTE"
    if filtro[0]['juros'] == "None":
        mensagem = ("Olá, {}!\n A PU de ID igual a {} não foi aplicada no galáxia, impossibilitando a validação dos valores da B3.Aguardo o lançamento, Obrigada!".format(filtro[0]['responsavel'],filtro[0]["id"]))
    if compj[0][1] == "None":
        mensagem = ("Olá, {}!\n A PU de ID igual a {} não foi aplicada na B3, impossibilitando a validação dos valores da B3.Aguardo o lançamento, Obrigada!".format(filtro[0]['responsavel'], filtro[0]["id"]))
    return mensagem, m1, m2, m3

mensagem, m1, m2, m3= mensagem(compj,compm, compa)

gui.theme('DarkPurple')
gui.SetOptions(text_color="#e4e4e4", font='Any 11')
instructions_layout = [
    [gui.Table(values=listag, headings=['Título', 'Id', 'Juros', 'Amortizacao', 'Amex', 'Responsável'],  vertical_scroll_only=False, max_col_width=35,
               auto_size_columns=False,
               display_row_numbers=False,
               justification='left',
               num_rows=10,
               key='-TABLE-',
               col_widths=[13, 11, 15,15,15, 14],
               row_height=35,
               tooltip='GALÁXIA')]
]

form_layout = [
    [gui.Table(values=listab, headings=['ID', 'Corporação', 'Tipo', 'Valor', 'Data'], vertical_scroll_only=False, max_col_width=35,
               auto_size_columns=False,
               display_row_numbers=False,
               justification='left',
               num_rows=10,
               key='-TABLE-',
               col_widths=[11, 25, 20, 11, 12],
               row_height=35,
               tooltip='B3')]
]

table_layout = [[gui.Text('Juros')],
    [gui.Table(values=compj, headings=headings, max_col_width=38,
               auto_size_columns=True,
               display_row_numbers=False,
               justification='left',
               num_rows=10,
               key='-TABLE-',
               row_height=35,
               tooltip='GALÁXIA')],
                [gui.Text(m1, font=('Any', '12', 'bold'), text_color='red')]
]
table2_layout = [[gui.Text('Amortização')],
    [gui.Table(values=compm, headings=headings, max_col_width=38,
               auto_size_columns=True,
               display_row_numbers=False,
               justification='left',
               num_rows=10,
               key='-TABLE-',
               row_height=35,
               tooltip='GALÁXIA')],
                [gui.Text(m2, font=('Any', '12', 'bold'), text_color='red')]
]
table3_layout = [[gui.Text('Amex')],
    [gui.Table(values=compa, headings=headings, max_col_width=38,
               auto_size_columns=True,
               display_row_numbers=False,
               justification='left',
               num_rows=10,
               key='-TABLE-',
               row_height=35,
               tooltip='GALÁXIA')],
                [gui.Text(m3, font=('Any', '12', 'bold'), text_color='red')]]

t = [[gui.Column(table_layout, element_justification='c' ), gui.Column(table2_layout, element_justification='c'), gui.Column(table3_layout, element_justification='c')]]

gui.SetOptions(text_color="#e4e4e4", font='Any 11')
tab_group = [[gui.Push(), gui.Image(filename='img.png'), gui.Push(), gui.Text('Extração de Relatório de Liquidação Diária',size=(40, 1), font=('Any 18')),gui.Push()],

    [gui.TabGroup(
        [[gui.Tab('Relatório B3', form_layout, title_color='Purple', background_color='Pink',
                 tooltip='B3', element_justification='c'),

          gui.Tab('Relatório Galáxia', instructions_layout, title_color='Purple', background_color='Pink',
                tooltip='B3', element_justification='c'),

          gui.Tab('Comparação', t, title_color='Blue', background_color='Pink',
          tooltip='Comparar', element_justification='c')]],
        tab_location='centertop',
        title_color='Pink', tab_background_color='Purple', selected_title_color='Purple',
        selected_background_color='Pink', border_width=5), gui.Button('Fechar')
    ]
]
# Define Window
window = gui.Window("Relatório Virgo", tab_group, icon="galaxia.ico")

while True:
    event, values = window.read()
    if event == "Fechar" or event == gui.WIN_CLOSED:
        break
window.close()

# enviar email:
email = filtro[0]['email']

if mensagem != None:
    outlook = win32.Dispatch('outlook.application')
    email = outlook.CreateItem(0)
    email.To = email
    email.Subject = 'Validação de PUs imputadas na B3'
    email.HTMLBody = mensagem
    email.Send()
    print('email enviado')
