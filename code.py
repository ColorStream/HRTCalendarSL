import PySimpleGUI as sg
import os
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from datetime import datetime, timedelta
import time

def gera_calendario():

    sg.theme('DarkPurple')

    # Janela principal
    layout = [  [sg.Text('Criador de calendário para TH')],
                [sg.Text('Seu nome: '), sg.InputText(key='nome')],
                [sg.Text('Seu éster: '), sg.InputText(key='ester')],
                [sg.Text('Sua dose em mg: '), sg.InputText(key='dose')],
                [sg.Text('Seu volume total em mL: '), sg.InputText(key='vol_total')],
                [sg.Text('Concentração do frasco em mg/ml: '), sg.InputText(key='concentracao')],
                [sg.Text('Intervalo das doses em dias: '), sg.InputText(key='intervalo')],
                [sg.Text('Data de início: '), sg.InputText(key='inicio')],
                [sg.Button('Gerar'), sg.Button('Limpar'), sg.Button('Cancelar')],
                [sg.Output(s=(75, 10))]
    ]

    window = sg.Window('Criador de calendário para TH injetável', layout, font=("Bookman Old Style", 16))
         
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Cancelar':
            break
        if event == 'Limpar':
            for key in values:
                window[key]('')

        if event == 'Gerar':
            try:
            
                # Entrada do usuário
                nome_usr = values['nome']
                ester = values['ester']
                dose = float(values['dose'])
                ml_total = float(values['vol_total'])
                concentracao = float(values['concentracao'])
                intervalo = float(values['intervalo'])
                data_inicio_str = values['inicio'] 

                # Converter data de início para datetime
                data_inicio = datetime.strptime(data_inicio_str, "%d/%m/%Y")
            except ValueError:
                sg.popup("Erro: Verifique os valores inseridos. Certifique-se de que são números válidos e a data está no formato dd/mm/yyyy.")
                continue

            # Criar planilha
            wb = Workbook()
            ws = wb.active
            ws.title = "Calendário"
            ws.sheet_view.showGridLines = False

            # Estilos para o Excel
            titulo_fill = PatternFill(start_color="FF00FF", end_color="FF00FF", fill_type="solid")
            header_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
            font_row = Font(name="Bookman Old Style", size=12)
            bold_font = Font(bold=True, color="FFFFFF")
            center_align = Alignment(horizontal="center", vertical="center")
            thin_border = Border(
                left=Side(style="thin"), right=Side(style="thin"),
                top=Side(style="thin"), bottom=Side(style="thin")
            )

            # Estrutura da planilha
            ws['B2'] = f"{nome_usr} - {ester} {dose}mg / {intervalo} dias"
            ws.merge_cells('B2:G2')
            ws['B2'].fill = titulo_fill
            ws['B2'].font = bold_font
            ws['B2'].alignment = center_align

            # Primeiras linhas

            ws['A4'] = "Hoje:"
            ws['B4'] = "=today()"
            ws['B4'].number_format = "DD/MM/YYYY"
            ws['C4'] = ml_total
            ws['D4'] = "mL"
            ws['E4'] = concentracao
            ws['F4'] = "mg/mL"

            ws['A5'] = "Dose (mg)"
            ws['B5'] = dose
            ws['C5'] = dose/concentracao
            ws['D5'] = "mL"

            ws['A6'] = "Intervalo:"
            ws['B6'] = intervalo

            ws['A7'] = "Data de início"
            ws['B7'] = data_inicio
            ws['C7'] = "=b4-b7"
            ws['A8'] = "Vai durar:"
            ws['B8'] = "=$C$4/($B$5/$E$4)"

            # Estilos
            header_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")  # Azul claro
            font_header = Font(bold=True, color="000000", size=12)
            font_row = Font(name="Arial", size=10)  # Fonte das linhas
            center_align = Alignment(horizontal="center", vertical="center")
            thin_border = Border(
                left=Side(style="thin"), right=Side(style="thin"),
                top=Side(style="thin"), bottom=Side(style="thin")
            )

            # Cabeçalhos
            headers = ["Dia", "mL", "Dias corridos", "Meses", "Lado", "Aplicações"]
            for col_num, header in enumerate(headers, start=2):
                cell = ws.cell(row=12, column=col_num, value=header)
                cell.fill = header_fill
                cell.font = font_header
                cell.alignment = center_align
                cell.border = thin_border

            # Imprimir cabeçalhos no terminal
            print(f"{'Aplicação':<12}{'Dia':<15}{'Data':<12}{'mL Restante':<15}{'Dias Corridos':<15}{'Meses':<8}{'Lado':<10}")
            print("-" * 80)

            # Variáveis de controle
            data_inicio = datetime.strptime("16/06/2023", "%d/%m/%Y")
            dose = 5.6
            concentracao = 40.0
            intervalo = 5  # Intervalo de dias
            ml_total = 10.0  # Total de mL inicial
            linha_inicial = 13
            aplicacao = 1
            lado_atual = "Esquerdo"
            dias_corridos = 0

            # Loop para preencher as linhas
            while ml_total > 0:
                # Calcular valores
                data_atual = data_inicio + timedelta(days=dias_corridos)
                ml_dose = dose / concentracao
                ml_restante = round(ml_total, 2) - (dose / concentracao) # mL restante
                ml_total -= ml_dose  # Atualizar mL restante

                # Linha no Excel
                linha_atual = linha_inicial + aplicacao - 1
                ws[f"A{linha_atual}"] = data_atual.strftime("%A")  # Dia da semana
                ws[f"B{linha_atual}"] = data_atual.strftime("%d/%m/%Y")  # Data
                ws[f"D{linha_atual}"] = dias_corridos  # Dias corridos
                ws[f"C{linha_atual}"] = ml_restante  # mL Restante
                ws[f"E{linha_atual}"] = round(dias_corridos / 30, 2)  # Meses
                ws[f"F{linha_atual}"] = lado_atual  # Lado
                ws[f"G{linha_atual}"] = aplicacao  # Aplicações

                # Aplicar estilo na linha
                row_fill = PatternFill(start_color="F0F8FF", end_color="F0F8FF", fill_type="solid") if aplicacao % 2 == 0 else PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
                for col in "ABCDEFG":
                    cell = ws[f"{col}{linha_atual}"]
                    cell.font = font_row
                    cell.alignment = center_align
                    cell.fill = row_fill
                    cell.border = thin_border

                # Imprimir no terminal
                print(
                    f"{aplicacao:<12}{data_atual.strftime('%A'):<15}{data_atual.strftime('%d/%m/%Y'):<12}"
                    f"{ml_restante:<15}{dias_corridos:<15}{round(dias_corridos / 30, 2):<8}{lado_atual:<10}"
                )

                # Atualizar variáveis
                lado_atual = "Direito" if lado_atual == "Esquerdo" else "Esquerdo"
                ml_restante -= (dose / concentracao)
                dias_corridos += intervalo
                aplicacao += 1
                linha_atual += 1

            # Ajustar larguras das colunas
            for column in ws.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                ws.column_dimensions[column_letter].width = max_length + 2

            # Salvar o arquivo
            wb.save(f"TH {nome_usr}.xlsx")
            print("Calendário criado com sucesso!")
    
gera_calendario()
