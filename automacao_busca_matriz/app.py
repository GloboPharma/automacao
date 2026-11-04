"""Consulta da Matriz de Acessos x Cargos (INF-000007).

Executar:  streamlit run app.py
Lê somente a cópia local em dados/. Para atualizar, substitua a cópia.
"""
import io
from datetime import datetime

import pandas as pd
import streamlit as st

from matriz_loader import ARQUIVO_PADRAO, carregar_matriz, listar_cargos

st.set_page_config(page_title="Matriz de Acessos x Cargos", layout="wide")


@st.cache_data(show_spinner="Lendo a planilha...")
def _dados(mtime: float):
    return carregar_matriz(), listar_cargos()


def _excel(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    return buf.getvalue()


def _download(df: pd.DataFrame, nome: str):
    c1, c2, _ = st.columns([1, 1, 6])
    c1.download_button("Baixar CSV", df.to_csv(index=False).encode("utf-8-sig"),
                       f"{nome}.csv", "text/csv")
    c2.download_button("Baixar Excel", _excel(df), f"{nome}.xlsx")


def _agrupar(df: pd.DataFrame) -> pd.DataFrame:
    """Uma linha por sistema, com os perfis juntos."""
    return (
        df.groupby(["sistema", "impacto", "tipo_sistema", "area_dono"], sort=False)
        ["perfil"].agg(lambda s: " | ".join(s)).reset_index()
        .rename(columns={"perfil": "perfis de acesso"})
    )


if not ARQUIVO_PADRAO.exists():
    st.error(f"Planilha não encontrada: {ARQUIVO_PADRAO}")
    st.stop()

df, cargos = _dados(ARQUIVO_PADRAO.stat().st_mtime)

st.title("Matriz de Acessos x Cargos")
st.caption(
    f"Fonte: {ARQUIVO_PADRAO.name} (cópia local, atualizada em "
    f"{datetime.fromtimestamp(ARQUIVO_PADRAO.stat().st_mtime):%d/%m/%Y %H:%M})"
)

modo = st.sidebar.radio("Consultar por", ["Cargo", "Sistema", "Comparar cargos", "Busca livre"])

if modo == "Cargo":
    dep = st.sidebar.selectbox("Departamento", sorted(cargos.departamento.unique()))
    lista = cargos[cargos.departamento == dep].cargo.tolist()
    cargo = st.sidebar.selectbox("Cargo", lista)
    sel = df[(df.departamento == dep) & (df.cargo == cargo)]
    st.subheader(f"{cargo} — {dep}")
    if sel.empty:
        st.info("Este cargo não possui acesso cadastrado na matriz.")
    else:
        res = _agrupar(sel)
        st.metric("Sistemas com acesso", len(res))
        st.dataframe(res, width='stretch', hide_index=True)
        _download(res, f"acessos_{cargo}")

elif modo == "Sistema":
    sistema = st.sidebar.selectbox("Sistema", sorted(df.sistema.unique()))
    sel = df[df.sistema == sistema]
    st.subheader(sistema)
    res = (sel.groupby(["departamento", "cargo"], sort=False)["perfil"]
           .agg(lambda s: " | ".join(s)).reset_index()
           .rename(columns={"perfil": "perfis de acesso"}))
    st.metric("Cargos com acesso", len(res))
    st.dataframe(res, width='stretch', hide_index=True)
    _download(res, f"cargos_{sistema}")

elif modo == "Comparar cargos":
    rotulos = (cargos.departamento + " | " + cargos.cargo).tolist()
    a = st.sidebar.selectbox("Cargo A", rotulos, index=0)
    b = st.sidebar.selectbox("Cargo B", rotulos, index=min(1, len(rotulos) - 1))

    def _perfis(rotulo):
        dep, cargo = rotulo.split(" | ", 1)
        s = df[(df.departamento == dep) & (df.cargo == cargo)]
        return s.groupby("sistema")["perfil"].agg(lambda x: " | ".join(x))

    pa, pb = _perfis(a), _perfis(b)
    comp = pd.concat([pa.rename("Cargo A"), pb.rename("Cargo B")], axis=1).fillna("—")
    comp["situação"] = comp.apply(
        lambda r: "igual" if r["Cargo A"] == r["Cargo B"]
        else "só A" if r["Cargo B"] == "—"
        else "só B" if r["Cargo A"] == "—" else "perfil diferente", axis=1)
    comp = comp.reset_index().rename(columns={"index": "sistema"})
    st.subheader("Comparação")
    st.write(f"**A:** {a}  \n**B:** {b}")
    filtro = st.multiselect("Mostrar", ["só A", "só B", "perfil diferente", "igual"],
                            default=["só A", "só B", "perfil diferente"])
    comp = comp[comp["situação"].isin(filtro)]
    st.dataframe(comp, width='stretch', hide_index=True)
    _download(comp, "comparacao_cargos")

else:
    termo = st.sidebar.text_input("Buscar (cargo, sistema, perfil ou departamento)")
    tipo = st.sidebar.multiselect("Tipo de sistema", sorted(df.tipo_sistema.unique()))
    impacto = st.sidebar.multiselect("Impacto", sorted(df.impacto.unique()))
    sel = df
    if termo.strip():
        t = termo.strip()
        mask = pd.Series(False, index=df.index)
        for c in ["departamento", "cargo", "sistema", "perfil"]:
            mask |= df[c].str.contains(t, case=False, na=False, regex=False)
        sel = df[mask]
    if tipo:
        sel = sel[sel.tipo_sistema.isin(tipo)]
    if impacto:
        sel = sel[sel.impacto.isin(impacto)]
    st.subheader("Busca livre")
    if not termo.strip() and not tipo and not impacto:
        st.info("Digite um termo ou escolha um filtro na barra lateral.")
    else:
        st.metric("Resultados", len(sel))
        cols = ["departamento", "cargo", "sistema", "perfil", "impacto", "tipo_sistema"]
        st.dataframe(sel[cols], width='stretch', hide_index=True)
        _download(sel[cols], "busca_matriz")
