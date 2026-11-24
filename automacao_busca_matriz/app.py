"""Consulta da Matriz de Acessos x Cargos (INF-000007).

Executar:  streamlit run app.py
Lê somente a cópia local em dados/. Para atualizar, substitua a cópia.
"""
import io
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from matriz_loader import ARQUIVO_PADRAO, carregar_matriz, listar_cargos

LOGO = Path(__file__).parent / "assets" / "logo.png"
MODOS = ["Cargo", "Sistema", "Comparar cargos", "Busca livre"]

# Paleta oficial Globo Pharma.
LARANJA, LARANJA_3, LARANJA_4, LARANJA_5 = "#FA4616", "#F69679", "#FCD5C0", "#FEF1EB"
GRAFITE, CINZA_CLARO = "#2D2D2D", "#F2F2F2"

COR_IMPACTO = {
    "IMPACTO DIRETO": f"background-color:{LARANJA};color:white;font-weight:600",
    "COM IMPACTO": f"background-color:{LARANJA_4};color:{GRAFITE}",
    "SEM IMPACTO": f"background-color:{CINZA_CLARO};color:{GRAFITE}",
}
COR_SITUACAO = {
    "igual": f"background-color:{CINZA_CLARO};color:{GRAFITE}",
    "só A": f"background-color:{LARANJA_4};color:{GRAFITE}",
    "só B": f"background-color:{LARANJA_4};color:{GRAFITE}",
    "perfil diferente": f"background-color:{LARANJA};color:white;font-weight:600",
}

ROTULOS = {
    "departamento": st.column_config.TextColumn("Departamento"),
    "cargo": st.column_config.TextColumn("Cargo"),
    "sistema": st.column_config.TextColumn("Sistema"),
    "perfil": st.column_config.TextColumn("Perfil de acesso"),
    "impacto": st.column_config.TextColumn("Impacto", width="small"),
    "tipo_sistema": st.column_config.TextColumn("Tipo de sistema", width="small"),
    "area_dono": st.column_config.TextColumn("Área dona"),
    "perfis": st.column_config.TextColumn("Perfis de acesso", width="large"),
    "situação": st.column_config.TextColumn("Situação", width="small"),
}

st.set_page_config(page_title="Matriz de Acessos x Cargos", page_icon="🔐", layout="wide")

st.markdown(
    f"""
    <style>
      .block-container {{ padding-top: 2rem; }}
      h1 {{ font-weight: 700; letter-spacing: -0.01em; margin-bottom: 0; }}
      [data-testid="stMetricValue"] {{ color: {LARANJA}; font-weight: 700; }}
      [data-testid="stMetric"] {{
        background: {LARANJA_5}; border-left: 4px solid {LARANJA};
        padding: .6rem 1rem; border-radius: 0 6px 6px 0;
      }}
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_data(show_spinner="Lendo a planilha...")
def _dados(mtime: float):
    return carregar_matriz(), listar_cargos()


def _para_arquivo(df: pd.DataFrame) -> pd.DataFrame:
    """Listas de perfis viram texto único para CSV/Excel."""
    saida = df.copy()
    for c in saida.columns:
        if saida[c].map(lambda v: isinstance(v, list)).any():
            saida[c] = saida[c].map(lambda v: " | ".join(v) if isinstance(v, list) else v)
    return saida


def _excel(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    return buf.getvalue()


def _download(df: pd.DataFrame, nome: str):
    nome = re.sub(r"[^\w\-]+", "_", nome).strip("_")
    plano = _para_arquivo(df)
    c1, c2, _ = st.columns([1, 1, 6])
    c1.download_button("Baixar CSV", plano.to_csv(index=False).encode("utf-8-sig"),
                       f"{nome}.csv", "text/csv", key=f"csv_{nome}")
    c2.download_button("Baixar Excel", _excel(plano), f"{nome}.xlsx", key=f"xls_{nome}")


def _quebrar_listas(df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    """Listas de perfis viram um perfil por linha dentro da célula; devolve a altura de linha."""
    saida, maior = df.copy(), 1
    for c in saida.columns:
        if saida[c].map(lambda v: isinstance(v, list)).any():
            maior = max(maior, saida[c].map(lambda v: len(v) if isinstance(v, list) else 1).max())
            saida[c] = saida[c].map(lambda v: "\n".join(v) if isinstance(v, list) else v)
    return saida, max(35, min(maior * 22 + 10, 400))


def _tabela(df: pd.DataFrame, destaque: str | None = None, mapa: dict | None = None,
            cfg: dict | None = None):
    """Exibe a tabela com rótulos legíveis e cor na coluna de destaque."""
    dado, altura = _quebrar_listas(df)
    if destaque and destaque in dado.columns:
        dado = dado.style.map(lambda v: (mapa or {}).get(v, ""), subset=[destaque])
    st.dataframe(dado, width="stretch", hide_index=True, column_config=cfg or ROTULOS,
                 row_height=altura)


def _perfis_por(df: pd.DataFrame, chaves: list[str]) -> pd.DataFrame:
    return (df.groupby(chaves, sort=False)["perfil"].agg(list)
            .reset_index().rename(columns={"perfil": "perfis"}))


def _cargo_por_departamento(cargos: pd.DataFrame, chave: str, rotulo: str = "",
                            cargo_inicial: int = 0) -> tuple[str, str]:
    c1, c2 = st.columns(2)
    dep = c1.selectbox(f"Departamento{rotulo}", sorted(cargos.departamento.unique()),
                       key=f"dep_{chave}")
    lista = cargos[cargos.departamento == dep].cargo.tolist()
    cargo = c2.selectbox(f"Cargo{rotulo}", lista, key=f"cargo_{chave}",
                         index=min(cargo_inicial, len(lista) - 1))
    return dep, cargo


# ---------- cabeçalho ----------
if not ARQUIVO_PADRAO.exists():
    st.error("Planilha não encontrada. Coloque o arquivo da matriz (.xlsx) na pasta "
             "`dados/` do projeto e recarregue a página.")
    st.stop()

df, cargos = _dados(ARQUIVO_PADRAO.stat().st_mtime)

topo_logo, topo_titulo = st.columns([1, 6], vertical_alignment="center")
if LOGO.exists():
    topo_logo.image(str(LOGO), width=150)
topo_titulo.title("Matriz de Acessos x Cargos")
st.caption(
    f"INF-000007 — cópia local, atualizada em "
    f"{datetime.fromtimestamp(ARQUIVO_PADRAO.stat().st_mtime):%d/%m/%Y às %H:%M}"
)

modo = st.segmented_control("Consultar por", MODOS, default="Cargo",
                            label_visibility="collapsed", key="modo") or "Cargo"
st.divider()

# ---------- consultas ----------
if modo == "Cargo":
    dep, cargo = _cargo_por_departamento(cargos, "consulta")
    sel = df[(df.departamento == dep) & (df.cargo == cargo)]
    st.subheader(f"{cargo} — {dep}")
    if sel.empty:
        st.info("Este cargo não possui acesso cadastrado na matriz.")
    else:
        res = _perfis_por(sel, ["sistema", "impacto", "tipo_sistema", "area_dono"])
        m1, m2, m3 = st.columns(3)
        m1.metric("Sistemas com acesso", len(res))
        m2.metric("Com impacto direto", int((res.impacto == "IMPACTO DIRETO").sum()))
        m3.metric("Perfis atribuídos", len(sel))
        _tabela(res, "impacto", COR_IMPACTO)
        _download(res, f"acessos_{cargo}")

elif modo == "Sistema":
    sistema = st.selectbox("Sistema", sorted(df.sistema.unique()))
    sel = df[df.sistema == sistema]
    info = sel.iloc[0]
    st.subheader(sistema)
    st.caption(f"{info.tipo_sistema.capitalize()} — {info.impacto.lower()} — "
               f"área dona: {info.area_dono or 'não informada'}")
    res = _perfis_por(sel, ["departamento", "cargo"])
    m1, m2 = st.columns(2)
    m1.metric("Cargos com acesso", len(res))
    m2.metric("Departamentos", res.departamento.nunique())
    _tabela(res)
    _download(res, f"cargos_{sistema}")

elif modo == "Comparar cargos":
    ca, cb = st.columns(2)
    with ca:
        st.markdown("**Cargo A**")
        dep_a, cargo_a = _cargo_por_departamento(cargos, "a", " A")
    with cb:
        st.markdown("**Cargo B**")
        dep_b, cargo_b = _cargo_por_departamento(cargos, "b", " B", cargo_inicial=1)

    def _perfis(dep, cargo):
        s = df[(df.departamento == dep) & (df.cargo == cargo)]
        return s.groupby("sistema", sort=False)["perfil"].agg(list).to_dict()

    pa, pb = _perfis(dep_a, cargo_a), _perfis(dep_b, cargo_b)
    col_a, col_b = f"A: {cargo_a}", f"B: {cargo_b}"
    linhas = []
    for sis in sorted(set(pa) | set(pb)):
        va, vb = pa.get(sis, []), pb.get(sis, [])
        if not va:
            sit = "só B"
        elif not vb:
            sit = "só A"
        elif sorted(va) == sorted(vb):
            sit = "igual"
        else:
            sit = "perfil diferente"
        linhas.append({"sistema": sis, col_a: va, col_b: vb, "situação": sit})
    comp = pd.DataFrame(linhas, columns=["sistema", col_a, col_b, "situação"])

    st.subheader("Comparação")
    resumo = comp["situação"].value_counts()
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Só no A", int(resumo.get("só A", 0)))
    m2.metric("Só no B", int(resumo.get("só B", 0)))
    m3.metric("Perfil diferente", int(resumo.get("perfil diferente", 0)))
    m4.metric("Iguais", int(resumo.get("igual", 0)))

    padrao = ["só A", "só B", "perfil diferente"]
    st.session_state.setdefault("comp_filtro", padrao)

    def _restaurar_comparacao():
        st.session_state["comp_filtro"] = padrao

    f1, f2 = st.columns([5, 2], vertical_alignment="bottom")
    filtro = f1.multiselect("Mostrar", ["só A", "só B", "perfil diferente", "igual"],
                            key="comp_filtro")
    f2.button("Limpar filtros", on_click=_restaurar_comparacao, width="stretch",
              icon=":material/filter_alt_off:", type="primary", key="limpar_comp")
    vista = comp[comp["situação"].isin(filtro)]
    if vista.empty:
        st.info("Nenhum sistema na situação escolhida. Inclua outra situação em Mostrar.")
    else:
        cfg = dict(ROTULOS)
        cfg[col_a] = cfg[col_b] = st.column_config.TextColumn(width="large")
        _tabela(vista, "situação", COR_SITUACAO, cfg)
        _download(vista, "comparacao_cargos")

else:
    def _limpar_busca():
        st.session_state.update(busca_termo="", busca_tipo=[], busca_impacto=[])

    termo = st.text_input("Buscar por cargo, sistema, perfil ou departamento",
                          placeholder="Ex.: Protheus, analista, qualidade", key="busca_termo")
    f1, f2, f3 = st.columns([3, 3, 2], vertical_alignment="bottom")
    tipo = f1.multiselect("Tipo de sistema", sorted(df.tipo_sistema.unique()), key="busca_tipo")
    impacto = f2.multiselect("Impacto", sorted(df.impacto.unique()), key="busca_impacto")
    f3.button("Limpar filtros", on_click=_limpar_busca, width="stretch",
              icon=":material/filter_alt_off:", type="primary")
    termo = termo.strip()

    if not termo and not tipo and not impacto:
        st.info("Digite um termo ou escolha um filtro para ver os resultados.")
    else:
        sel = df
        if termo:
            mask = pd.Series(False, index=df.index)
            for c in ["departamento", "cargo", "sistema", "perfil"]:
                mask |= df[c].str.contains(termo, case=False, na=False, regex=False)
            sel = df[mask]
        if tipo:
            sel = sel[sel.tipo_sistema.isin(tipo)]
        if impacto:
            sel = sel[sel.impacto.isin(impacto)]

        if sel.empty:
            st.warning("Nenhum resultado para esta busca. Tente outro termo ou remova filtros.")
        else:
            cols = ["departamento", "cargo", "sistema", "perfil", "impacto", "tipo_sistema"]
            m1, m2, m3 = st.columns(3)
            m1.metric("Resultados", len(sel))
            m2.metric("Cargos", sel.groupby(["departamento", "cargo"]).ngroups)
            m3.metric("Sistemas", sel.sistema.nunique())
            _tabela(sel[cols], "impacto", COR_IMPACTO)
            _download(sel[cols], "busca_matriz")
