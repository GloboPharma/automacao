import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent.parent))
from matriz_loader import ARQUIVO_PADRAO, carregar_matriz, listar_cargos  # noqa: E402

pytestmark = pytest.mark.skipif(not ARQUIVO_PADRAO.exists(), reason="planilha ausente")


@pytest.fixture(scope="module")
def df():
    return carregar_matriz()


def test_contagens(df):
    cargos = listar_cargos()
    assert len(cargos) == 248
    assert cargos.departamento.nunique() == 32
    assert df.sistema.nunique() == 101


def test_sem_departamento_vazio_nem_na(df):
    assert (df.departamento == "").sum() == 0
    assert not df.perfil.str.upper().isin(["N/A", ""]).any()


def test_cargo_repetido_no_mesmo_departamento_e_rotulado():
    cargos = listar_cargos()
    assert not cargos.duplicated().any()
    assert cargos.cargo.str.contains(r"\[col\. AT\]").any()


def test_cruzamento_conhecido(df):
    sel = df[(df.cargo == "ASSISTENTE ADMINISTRATIVO") & (df.departamento == "PCP")
             & (df.sistema == "ACTIVE DIRECTORY")]
    assert sorted(sel.perfil) == ["G_PCP_RO", "G_PCP_RW"]
