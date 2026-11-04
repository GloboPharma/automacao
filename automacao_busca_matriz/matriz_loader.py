"""Leitura da planilha INF-000007 (Matriz de Acessos x Cargos).

Converte a grade (sistemas nas linhas, cargos nas colunas) para uma tabela
no formato longo: uma linha por (departamento, cargo, sistema, perfil).
A planilha é aberta somente para leitura.
"""
from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl.utils import get_column_letter

ARQUIVO_PADRAO = (
    Path(__file__).parent / "dados" / "INF-000007-MATRIZ DE ACESSOS X CARGOS.xlsx"
)

LINHA_DEPARTAMENTO = 5
LINHA_CARGO = 6
PRIMEIRA_LINHA_SISTEMA = 7
COLUNA_SISTEMA = 2
COLUNA_IMPACTO = 3
COLUNA_TIPO = 4
COLUNA_DONO = 5
COLUNA_POPS = 6
PRIMEIRA_COLUNA_CARGO = 7

VAZIOS = {"", "N/A", "NA", "-", "NONE"}

COLUNAS = [
    "departamento", "cargo", "sistema", "impacto", "tipo_sistema",
    "area_dono", "pops", "perfil",
]


def _texto(valor) -> str:
    if valor is None:
        return ""
    return " ".join(str(valor).split())


def _perfis(valor) -> list[str]:
    """Quebra o conteúdo da célula em perfis; descarta N/A e vazios."""
    if valor is None:
        return []
    saida = []
    for parte in str(valor).splitlines():
        parte = " ".join(parte.split())
        if parte.upper() not in VAZIOS:
            saida.append(parte)
    return saida


def _rotular_repetidos(cols: dict[int, tuple[str, str]]) -> dict[int, str]:
    """Devolve {coluna: cargo}. Se o mesmo cargo aparece em mais de uma coluna
    do mesmo departamento, acrescenta '[col. XX]' para não misturar acessos."""
    contagem: dict[tuple[str, str], int] = {}
    for chave in cols.values():
        contagem[chave] = contagem.get(chave, 0) + 1
    return {
        col: cargo if contagem[(dep, cargo)] == 1
        else f"{cargo} [col. {get_column_letter(col)}]"
        for col, (dep, cargo) in cols.items()
    }


def carregar_matriz(caminho: str | Path = ARQUIVO_PADRAO) -> pd.DataFrame:
    wb = openpyxl.load_workbook(caminho, read_only=False, data_only=True)
    ws = wb.worksheets[0]

    # Departamento: valor só na 1ª célula da mesclagem; propaga até o próximo.
    departamentos: dict[int, str] = {}
    atual = ""
    for col in range(PRIMEIRA_COLUNA_CARGO, ws.max_column + 1):
        valor = _texto(ws.cell(LINHA_DEPARTAMENTO, col).value)
        if valor:
            atual = valor
        departamentos[col] = atual

    cargos = _rotular_repetidos(
        {
            col: (departamentos[col], _texto(ws.cell(LINHA_CARGO, col).value))
            for col in range(PRIMEIRA_COLUNA_CARGO, ws.max_column + 1)
            if _texto(ws.cell(LINHA_CARGO, col).value)
        }
    )

    registros = []
    for lin in range(PRIMEIRA_LINHA_SISTEMA, ws.max_row + 1):
        sistema = _texto(ws.cell(lin, COLUNA_SISTEMA).value)
        if not sistema:
            continue
        base = {
            "sistema": sistema,
            "impacto": _texto(ws.cell(lin, COLUNA_IMPACTO).value),
            "tipo_sistema": _texto(ws.cell(lin, COLUNA_TIPO).value),
            "area_dono": _texto(ws.cell(lin, COLUNA_DONO).value),
            "pops": _texto(ws.cell(lin, COLUNA_POPS).value),
        }
        for col, cargo in cargos.items():
            for perfil in _perfis(ws.cell(lin, col).value):
                registros.append({
                    **base,
                    "departamento": departamentos[col],
                    "cargo": cargo,
                    "perfil": perfil,
                })
    wb.close()
    return pd.DataFrame(registros, columns=COLUNAS)


def listar_cargos(caminho: str | Path = ARQUIVO_PADRAO) -> pd.DataFrame:
    """Todos os cargos (inclusive os sem nenhum acesso) com seu departamento."""
    wb = openpyxl.load_workbook(caminho, data_only=True)
    ws = wb.worksheets[0]
    atual, cols = "", {}
    for col in range(PRIMEIRA_COLUNA_CARGO, ws.max_column + 1):
        dep = _texto(ws.cell(LINHA_DEPARTAMENTO, col).value)
        if dep:
            atual = dep
        cargo = _texto(ws.cell(LINHA_CARGO, col).value)
        if cargo:
            cols[col] = (atual, cargo)
    wb.close()
    rotulos = _rotular_repetidos(cols)
    return pd.DataFrame(
        [{"departamento": cols[c][0], "cargo": rotulos[c]} for c in cols],
        columns=["departamento", "cargo"],
    )
