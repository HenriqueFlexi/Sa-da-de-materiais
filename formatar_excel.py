import os
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

EXCEL_ARQUIVO = "saida_materiais.xlsx"

def formatar_planilha():
    if not os.path.exists(EXCEL_ARQUIVO):
        print("Arquivo Excel não encontrado.")
        return

    wb = load_workbook(EXCEL_ARQUIVO)
    if "Saídas" not in wb.sheetnames:
        print("Aba 'Saídas' não encontrada.")
        return

    ws = wb["Saídas"]

    fonte_negrito = Font(bold=True)
    preenchimento_cabecalho = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
    alinhamento_centralizado = Alignment(horizontal="center", vertical="center")
    borda_fina = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # Cabeçalho
    for cell in ws[1]:
        cell.font = fonte_negrito
        cell.fill = preenchimento_cabecalho
        cell.alignment = alinhamento_centralizado
        cell.border = borda_fina

    # Demais linhas
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = alinhamento_centralizado
            cell.border = borda_fina

    # Largura das colunas
    colunas = {
        "A": 20,  # Data/Hora
        "B": 20,  # Nome
        "C": 25,  # Projeto
        "D": 30,  # Material
        "E": 12,  # Quantidade
    }

    for col, largura in colunas.items():
        ws.column_dimensions[col].width = largura

    wb.save(EXCEL_ARQUIVO)
    print("Formatação aplicada com sucesso!")

# Chamada da função
formatar_planilha()
