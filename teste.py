import urllib.request
import json
import PySimpleGUI as gui
from pathlib import Path
import win32com.client as win32
import pandas as pd
import sys
import csv
import json
import sqlalchemy as sql
import datetime
headings=['ID', 'VALOR GALÁXIA', 'VALOR B3']
gui.theme('DarkPurple')
gui.SetOptions(text_color="#e4e4e4", font='Any 11')
instructions_layout = [
    [gui.Table(values=['22F0284570', '6.497', '0.0', '0.0','Kaio Teixeira Ortiz'], headings=['id', 'Juros', 'Amortizacao', 'Amex', 'Responsavel'],  vertical_scroll_only=False, max_col_width=35,
               auto_size_columns=False,
               display_row_numbers=False,
               justification='left',
               num_rows=10,
               key='-TABLE-',
               col_widths=[15, 10, 15,9,9, 9, 15],
               row_height=35,
               tooltip='GALÁXIA')]
]

form_layout = [
    [gui.Table(values=['22F0284570', '6.497', 'PAGAMENTO DE JUROS'], headings=['ID', 'Valor', 'Tipo'], vertical_scroll_only=False, max_col_width=35,
               auto_size_columns=False,
               display_row_numbers=False,
               justification='left',
               num_rows=10,
               key='-TABLE-',
               col_widths=[10, 20, 10],
               row_height=35,
               tooltip='B3')]
]

table_layout = [[gui.Text('Divergência de juros')],
    [gui.Table(values=['22F0284570', '6.497', '6.497'], headings=headings, max_col_width=38,
               auto_size_columns=True,
               display_row_numbers=False,
               justification='left',
               num_rows=10,
               key='-TABLE-',
               row_height=35,
               tooltip='GALÁXIA')]
]
table2_layout = [[gui.Text('Divergência de amortização')],
    [gui.Table(values=[None, None, None], headings=headings, max_col_width=38,
               auto_size_columns=True,
               display_row_numbers=False,
               justification='left',
               num_rows=10,
               key='-TABLE-',
               row_height=35,
               tooltip='GALÁXIA')]
]
table3_layout = [[gui.Text('Divergência de amex')],
    [gui.Table(values=[None, None, None], headings=headings, max_col_width=38,
               auto_size_columns=True,
               display_row_numbers=False,
               justification='left',
               num_rows=10,
               key='-TABLE-',
               row_height=35,
               tooltip='GALÁXIA')]
]
t = [[gui.Column(table_layout, element_justification='c' ), gui.Column(table2_layout, element_justification='c')
          , gui.Column(table3_layout, element_justification='c')]]

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



import urllib.request
# import json
# import PySimpleGUI as gui
# from pathlib import Path
# import win32com.client as win32
# import pandas as pd
# import sys
# import csv
# import json
# import sqlalchemy as sql
# import datetime
#
# ssl_args = {'ssl_ca': "DigiCertGlobalRootCA.crt.pem"}
# engine =sql.create_engine('mysql+pymysql://fernanda.pereira:2uwHbaCukUl&yWYUXxkC@vdwh.mysql.database.azure.com:3306/dwh', connect_args=ssl_args)
# conn = engine.connect()
# response = engine.execute('select cra_sch.ticker_symbol, cra_inst.corporation_name, cra_sch.event_name, cra_sch.event_unit_price, cra_sch.payment_date from dwh.vw_up2data_fixed_income_cra_schedule as cra_sch join dwh.vw_up2data_fixed_income_cra_instrument as cra_inst on (cra_sch.ticker_symbol = cra_inst.ticker_symbol)')
# def cra(response):
#     cra = []
#     dic = {}
#     for row in response:
#         if row[1] == 'ISEC SECURITIZADORA S.A' or row[1] == 'VIRGO II COMPANHIA DE SECURITIZACAO' or row[1] == 'VIRGO COMPANHIA DE SECURITIZACAO':
#             if row[4] == datetime.date.today():
#                 dic["id"] = row[0]
#                 dic["tipo"] = row[2]
#                 dic["preço"] = round(row[3],3)
#                 if dic not in cra:
#                     cra.append(dic)
#     return cra
# cra = cra(response)
# response1 = engine.execute('select cri_sch.ticker_symbol, cri_inst.corporation_name, cri_sch.event_name, cri_sch.event_unit_price, cri_sch.payment_date from dwh.vw_up2data_fixed_income_cri_schedule as cri_sch join dwh.vw_up2data_fixed_income_cri_instrument as cri_inst on (cri_sch.ticker_symbol = cri_inst.ticker_symbol)')
# def cri(response):
#     cri = []
#     dic = {}
#     for row in response:
#         if row[1] == 'ISEC SECURITIZADORA S.A' or row[1] == 'VIRGO II COMPANHIA DE SECURITIZACAO' or row[1] == 'VIRGO COMPANHIA DE SECURITIZACAO':
#             if row[4] == datetime.date.today():
#                 dic["id"] = row[0]
#                 dic["tipo"] = row[2]
#                 dic["preço"] = round(row[3],3)
#                 if dic not in cri:
#                     cri.append(dic)
#     return cri
# cri = cri(response1)
# b3 = cri + cra
# print(b3)
# def b3_sep(b3):
#     dicionario = {}
#     v_juros = []
#     v_amort = []
#     v_amex = []
#     for dic in b3:
#         id = dic['id']
#         preco = dic['preço']
#         if dic['tipo'] == 'PAGAMENTO DE JUROS':
#             dicionario["id"] = id
#             dicionario["pu"] = preco
#             v_juros.append(dicionario)
#         elif dic['tipo'] == 'AMORTIZACAO':
#             dicionario["id"] = id
#             dicionario["pu"] = preco
#             v_amort.append(dicionario)
#         elif dic['tipo'] == 'AMORTIZACAO EXTRAORDINARIA':
#             dicionario["id"] = id
#             dicionario["pu"] = preco
#             v_amex.append(dicionario)
#     return v_juros, v_amort, v_amex
# v_juros, v_amort, v_amex = b3_sep(b3)
# v_b3 = v_juros + v_amort + v_amex
# print(v_juros)
# print(v_amort)
# def url_galaxia():
#     url = "https://redash.virgo.inc/api/queries/112/results.json?api_key=8vPmO96cK7hQ8mahDr6C4LMleuYLBBeZhi7fnwSP"
#     response = urllib.request.urlopen(url)
#     data = json.loads(response.read())
#     return data
#
# galaxia = url_galaxia()
# resultg = galaxia["query_result"]["data"]["rows"]
#
# # def filter(g, b):
# #     filter= []
# #     for i in b:
# #         for j in g:
# #             if (j['id'] == i['id']):
# #                 if j not in filter:
# #                     filter.append(j)
# #     return filter
# # filtro = filter(resultg, v_b3)
# #
# # def juros(resultg, resultb):
# #     lista = []
# #     for i in resultb:
# #         lista.append(resultb["id"])
# #         valorg = resultg['juros']
# #         lista.append(valorg)
# #         lista.append(resultb["pu"])
# #     return lista
#
# headings=['ID', 'VALOR GALÁXIA', 'VALOR B3']
#
# def listas(dict):
#     lista_cont = []
#     if dict is None:
#         return None
#     else:
#         for d in dict:
#             lista_cont.append(list(d.values()))
#         lista_title = list(dict[0].keys())
#         for i in lista_cont:
#             for m in i:
#                 if type(m) is float:
#                     m = '{}'.format(m)
#         return lista_cont, lista_title
#
# filtro, lista_titleg = listas([{'titulo':'ESTR BRASILATA','id': '22F0284570', 'responsavel': 'Kaio Teixeira Ortiz', 'juros':'6.497', 'amortizacao':None, 'amex':None}])
# lista_contb, lista_titleb = listas(b3)
#
# gui.theme('DarkPurple')
# gui.SetOptions(text_color="#e4e4e4", font='Any 11')
# instructions_layout = [
#     [gui.Table(values=filtro, headings=lista_titleg,  vertical_scroll_only=False, max_col_width=35,
#                auto_size_columns=False,
#                display_row_numbers=False,
#                justification='left',
#                num_rows=10,
#                key='-TABLE-',
#                col_widths=[15, 10, 15,9,9, 9, 15],
#                row_height=35,
#                tooltip='GALÁXIA')]
# ]
#
# form_layout = [
#     [gui.Table(values=lista_contb, headings=lista_titleb, vertical_scroll_only=False, max_col_width=35,
#                auto_size_columns=False,
#                display_row_numbers=False,
#                justification='left',
#                num_rows=10,
#                key='-TABLE-',
#                col_widths=[10, 20, 10],
#                row_height=35,
#                tooltip='B3')]
# ]
#
# table_layout = [[gui.Text('Divergência de juros')],
#     [gui.Table(values=['22F0284570', '6.497', '6.497'], headings=headings, max_col_width=38,
#                auto_size_columns=True,
#                display_row_numbers=False,
#                justification='left',
#                num_rows=10,
#                key='-TABLE-',
#                row_height=35,
#                tooltip='GALÁXIA')]
# ]
# table2_layout = [[gui.Text('Divergência de amortização')],
#     [gui.Table(values=[None, None, None], headings=headings, max_col_width=38,
#                auto_size_columns=True,
#                display_row_numbers=False,
#                justification='left',
#                num_rows=10,
#                key='-TABLE-',
#                row_height=35,
#                tooltip='GALÁXIA')]
# ]
# table3_layout = [[gui.Text('Divergência de amex')],
#     [gui.Table(values=[None, None, None], headings=headings, max_col_width=38,
#                auto_size_columns=True,
#                display_row_numbers=False,
#                justification='left',
#                num_rows=10,
#                key='-TABLE-',
#                row_height=35,
#                tooltip='GALÁXIA')]
# ]
# t = [[gui.Column(table_layout, element_justification='c' ), gui.Column(table2_layout, element_justification='c')
#           , gui.Column(table3_layout, element_justification='c')]]
#
# gui.SetOptions(text_color="#e4e4e4", font='Any 11')
# tab_group = [[gui.Push(), gui.Image(filename='img.png'), gui.Push(), gui.Text('Extração de Relatório de Liquidação Diária',size=(40, 1), font=('Any 18')),gui.Push()],
#     [gui.TabGroup(
#         [[gui.Tab('Relatório B3', form_layout, title_color='Purple', background_color='Pink',
#                  tooltip='B3', element_justification='c'),
#
#           gui.Tab('Relatório Galáxia', instructions_layout, title_color='Purple', background_color='Pink',
#                 tooltip='B3', element_justification='c'),
#
#           gui.Tab('Comparação', t, title_color='Blue', background_color='Pink',
#           tooltip='Comparar', element_justification='c')]],
#         tab_location='centertop',
#         title_color='Pink', tab_background_color='Purple', selected_title_color='Purple',
#         selected_background_color='Pink', border_width=5), gui.Button('Fechar')
#     ]
# ]
# # Define Window
# window = gui.Window("Relatório Virgo", tab_group, icon="galaxia.ico")
#
# while True:
#     event, values = window.read()
#     if event == "Fechar" or event == gui.WIN_CLOSED:
#         break
# window.close()
# #
# # # enviar email:
# # def email(resultg, lista):
# #     mensagem = []
# #     email = []
# #     for y in lista:
# #         for x in resultg:
# #             if x["id"] == y:
# #                 email.append(x["email"])
# #                 mensagem.append("A PU do título que possui ID igual a {} se encontra divergente com a B3.".format(y))
# #     return email, mensagem
# # #email, mensagem = email(resultg, lista)
# # # if email != []:
# # #     for contato in sys.getsizeof(email):
# # #         outlook = win32.Dispatch('outlook.application')
# # #         email = outlook.CreateItem(0)
# # #         email.To = email(contato)
# # #         email.HTMLBody = mensagem(contato)
# # #         email.Send()
# # # else:
# # #     outlook = win32.Dispatch('outlook.application')
# # #     email = outlook.CreateItem(0)
# # #     email.To = 'fernanda.pereira@virgo.inc'
# # #     email.Subject = 'PU divergente'
# # #     email.HTMLBody = "A PU do título que possui ID igual a {} se encontra divergente com a B3."
# # #     email.Send()
# # #     print('e-mail enviado com sucesso!')