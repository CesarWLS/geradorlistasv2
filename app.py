import streamlit as st
import pandas as pd
import docx
from datetime import date, timedelta, datetime
import re
import io
import os
import openpyxl
from openpyxl.cell.cell import MergedCell
from copy import copy
import holidays
from PIL import Image as PILImage

# Nome exato do seu arquivo de modelo que deve estar na mesma pasta
MODELO_PADRAO_NOME = "MODELOLISTA.xlsx"

# ==========================================
# FUNÇÕES DE LÓGICA E FORMATAÇÃO
# ==========================================

def limpar_especialidade(texto):
    texto_limpo = re.sub(r'\(.*?\)', '', texto).strip()
    abreviacoes = {
        "TO": "Terapia Ocupacional",
        "FISIO": "Fisioterapia",
        "FONO": "Fonoaudiologia",
        "NUTRI": "Nutrição",
        "MUSICA": "Musicoterapia",
        "MÚSICA": "Musicoterapia",
        "ABA": "Psicologia ABA"
    }
    chave = texto_limpo.upper()
    return abreviacoes.get(chave, texto_limpo)

def limpar_nome_arquivo(nome):
    nome_limpo = re.sub(r'[\\/*?:_\|\[\]]', '', nome)
    return re.sub(r'\s+', ' ', nome_limpo).strip()

def remover_conselho(nome_profissional):
    if not nome_profissional:
        return ""
    padrao_conselho = r'\b(CRP|CREFITO|CRM|COREN|COFFITO|OAB|CREF|CRFa|CRFª|CRN|CRT|CRBM|CBO)\b.*'
    nome_limpo = re.sub(padrao_conselho, '', nome_profissional, flags=re.IGNORECASE)
    return re.sub(r'[-\s,/|:_]+$', '', nome_limpo).strip()

def copiar_estilos_celula(origem, destino):
    if origem.font: destino.font = copy(origem.font)
    if origem.border: destino.border = copy(origem.border)
    if origem.fill: destino.fill = copy(origem.fill)
    if origem.number_format: destino.number_format = copy(origem.number_format)
    if origem.alignment: destino.alignment = copy(origem.alignment)

def limpar_style_celula(celula):
    celula.font = None
    celula.border = openpyxl.styles.Border()
    celula.fill = openpyxl.styles.PatternFill(fill_type=None)
    celula.alignment = None

def clonar_aba_com_estilo(wb_modelo, wb_destino, nome_nova_aba, imagens_cache=None, nova_posicao_imagem=None):
    aba_origem = wb_modelo.worksheets[0]
    aba_destino = wb_destino.create_sheet(title=nome_nova_aba)
    aba_destino.views.sheetView[0].showGridLines = True

    for col, col_dim in aba_origem.column_dimensions.items():
        aba_destino.column_dimensions[col].width = col_dim.width

    for row, row_dim in aba_origem.row_dimensions.items():
        aba_destino.row_dimensions[row].height = row_dim.height

    for mesclada in aba_origem.merged_cells.ranges:
        aba_destino.merge_cells(str(mesclada))

    for r in range(1, aba_origem.max_row + 1):
        for c in range(1, aba_origem.max_column + 1):
            cel_origem = aba_origem.cell(row=r, column=c)
            cel_destino = aba_destino.cell(row=r, column
