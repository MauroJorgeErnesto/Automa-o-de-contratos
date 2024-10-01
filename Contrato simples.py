from docx import Document
from datetime import datetime
import pandas as pd

tabela = pd.read_excel("Dadosothers.xlsx", parse_dates=['Inicio', 'Fim'])

contratos_empresas = {
    "HAPAG": r'C:caminho de pasta\MODELO CONTRATO HAPAG.docx',
    "COSCO": r'C:caminho de pasta\MODELO CONTRATO COSCO.docx',
    "ALIANÇA": r'C:caminho de pasta\MODELO CONTRATO ALIANÇA.docx',
    "EVERGREEN": r'C:caminho de pasta\MODELO CONTRATO EVERGREEN.docx',
    "ONE": r'C:caminho de pasta\MODELO CONTRATO ONE.docx'}

caminho_pasta = r'C:caminho de pasta\Contratos feitos'  # Substitua por seu caminho de pasta desejado

for linha in tabela.index:
    referencias = {
        "VVVV": tabela.loc[linha, "Vinte"],
        "QQQQ": tabela.loc[linha, "Quarenta"],
        "OOOO": tabela.loc[linha, "POL"],
        "PPPP": tabela.loc[linha, "POD"],
        "NNNN": tabela.loc[linha, "Navio"],
        "YYYY": tabela.loc[linha, "Viagem"],
        "IIII": tabela.loc[linha, "Inicio"].strftime("%b %d"),
        "FFFF": tabela.loc[linha, "Fim"].strftime("%b %d"),
        "RRRR": tabela.loc[linha, "Ano"],
        "DDDD": datetime.now().strftime("%d de %B de %Y"),
        "TTTT": tabela.loc[linha, "Taxa"]
    }
    nome_arquivo = ("Contrato " + str(tabela.loc[linha, 'Solicitante']) + " - " + str(tabela.loc[linha, 'Navio']) + ' VG. ' + str(tabela.loc[linha, 'Viagem']) + ' - ' + str(tabela.loc[linha, 'POL']) + '.docx')

    documento = Document(contratos_empresas[tabela.loc[linha, 'Empresa']])

    for key in referencias:
        for paragraph in documento.paragraphs:
            if key in paragraph.text:
                for run in paragraph.runs:
                    run.text = run.text.replace(key, str(referencias[key]))

    documento.save(f"{caminho_pasta}\\{nome_arquivo}.docx")
