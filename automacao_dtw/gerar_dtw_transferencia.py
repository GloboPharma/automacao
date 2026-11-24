"""
gerar_dtw_transferencia.py
--------------------------------------------------------------------------
Valida saldo de lotes no SAP Business One (HANA) e gera os 3 arquivos DTW
de Transferência de Estoque (OWTR / WTR1 / BTNT) somente com as linhas
aprovadas.

Regra de bloqueio (linha NÃO entra na transferência):
  * lote não existe para o item / item não controlado por lote
  * lote sem saldo no depósito de ORIGEM
  * saldo menor que a quantidade pedida
  * saldo reservado/comprometido (OBTQ.CommitQty) que deixa o disponível
    menor que o pedido   (use --bloquear-qualquer-reserva para bloquear
    o lote se houver QUALQUER reserva)
  * quantidade inválida (zero, negativa, vazia)
  * depósito de destino inexistente ou igual à origem
  * (opcional) lote com status Inacessível / Bloqueado: --bloquear-status

Quando o mesmo lote/depósito de origem aparece em mais de uma linha
(inclusive em abas diferentes), as quantidades são SOMADAS antes da
comparação com o saldo — se a soma não couber, todas as linhas daquele
lote são bloqueadas.

Planilha de entrada: qualquer nome. São lidas todas as abas cujo cabeçalho
tenha "Código do Item", "Lote", uma coluna "Qtd <deposito>" e "Depósito".
O depósito de ORIGEM vem do título da coluna de quantidade
(ex.: "Qtd 01.08" -> origem 01.08) ou de --origem.
O depósito de DESTINO é a coluna "Depósito".

Uso:
  pip install -r requirements.txt     (conexão em .env — ver README.md)
  python gerar_dtw_transferencia.py "Ajuste Estoque PA v2-validacao.xlsx" ^
         --host 10.0.0.10 --porta 30015 --usuario LEITURA --schema SBO_GLOBO

  A senha é pedida no terminal (ou variável de ambiente SAP_HANA_SENHA).

Saídas (pasta --saida, padrão: ao lado da planilha, subpasta DTW_aaaammdd_hhmm):
  OWTR - Transferência de Estoque.xlsx / .txt
  WTR1 - Linhas Transferencia de Estoque.xlsx / .txt
  BTNT - Lotes.xlsx / .txt
  (.txt = "Texto (separado por tabulações)", ANSI, CRLF — o formato que o DTW lê;
   escolha com --formato txt|xlsx|ambos)
  Relatorio_Validacao.xlsx   (todas as linhas, com saldo e motivo)

O script só faz SELECT. Nenhuma escrita no banco.
--------------------------------------------------------------------------
"""
import argparse
import datetime as dt
import getpass
import os
import re
import sys
from collections import OrderedDict, defaultdict

# ---------------------------------------------------------------- .env
def carregar_env():
    """Lê o .env da pasta do script (ou da pasta atual). Não depende de biblioteca.
    Valores vazios são ignorados; variáveis já definidas no Windows têm prioridade."""
    pasta_script = os.path.dirname(os.path.abspath(__file__))
    candidatos = []
    for pasta in (pasta_script, os.getcwd()):
        for nome in (".env", ".env.txt", "env", "env.txt"):
            c = os.path.join(pasta, nome)
            if c not in candidatos:
                candidatos.append(c)
    for caminho in candidatos:
        if not os.path.isfile(caminho):
            continue
        conteudo = None
        for enc in ("utf-8-sig", "utf-16", "cp1252"):
            try:
                with open(caminho, encoding=enc) as f:
                    conteudo = f.read()
                if "=" in conteudo:
                    break
            except (UnicodeError, UnicodeDecodeError):
                conteudo = None
        if not conteudo:
            continue
        lidas = []
        for linha in conteudo.splitlines():
            linha = linha.strip()
            if not linha or linha.startswith("#") or "=" not in linha:
                continue
            if linha.lower().startswith("export "):
                linha = linha[7:]
            chave, valor = linha.split("=", 1)
            chave, valor = chave.strip(), valor.strip()
            if len(valor) >= 2 and valor[0] == valor[-1] and valor[0] in "\"'":
                valor = valor[1:-1]
            if valor and not os.environ.get(chave):
                os.environ[chave] = valor
                lidas.append(chave)
        aviso = "" if os.path.basename(caminho) == ".env" else "  (renomeie para .env)"
        print(f"Configuração lida de: {caminho}{aviso}")
        print(f"  variáveis: {', '.join(lidas) if lidas else 'NENHUMA preenchida'}")
        return caminho
    print(f"AVISO: arquivo .env não encontrado em {pasta_script}")
    return None

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
except ImportError:
    sys.exit("Instale o openpyxl:  pip install openpyxl")

STATUS_LOTE = {"0": "Liberado", "1": "Inacessível", "2": "Bloqueado"}

# Cabeçalhos dos modelos DTW (linha 1 = Service/DI, linha 2 = campo do banco)
OWTR_HDR = [("DocEntry", "DocEntry"), ("Series", "Series"), ("DocDate", "DocDate"),
            ("Comments", "Comments"), ("PriceList", "GroupNum")]
WTR1_HDR = [("ParentKey", "DocNum"), ("LineNum", "LineNum"), ("ItemCode", "ItemCode"),
            ("ItemDescription", "Dscription"), ("Quantity", "Quantity"),
            ("FromWarehouseCode", "FromWhsCod"), ("WarehouseCode", "WhsCode")]
BTNT_HDR = [("ParentKey", "DocNum"), ("LineNum", "LineNum"), ("BatchNumber", "DistNumber"),
            ("Quantity", "Quantity"), ("BaseLineNumber", "DocLineNum"), ("ItemCode", "ItemCode")]


# ---------------------------------------------------------------- leitura
def _norm(s):
    s = str(s or "").strip().lower()
    for a, b in (("ó", "o"), ("ô", "o"), ("ã", "a"), ("á", "a"), ("é", "e"), ("í", "i"), ("ç", "c")):
        s = s.replace(a, b)
    return s


def _qtd(v):
    if v is None or (isinstance(v, str) and not v.strip()):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(" ", "")
    if "," in s:                       # 1.234,5 -> 1234.5
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def ler_planilha(caminho, origem_padrao=None):
    wb = openpyxl.load_workbook(caminho, data_only=True, read_only=True)
    linhas = []
    for ws in wb.worksheets:
        rows = ws.iter_rows(values_only=True)
        try:
            cab = next(rows)
        except StopIteration:
            continue
        idx = {}
        origem_aba = origem_padrao
        for i, c in enumerate(cab):
            n = _norm(c)
            if n in ("codigo do item", "itemcode", "codigo"):
                idx["item"] = i
            elif n in ("nome do item", "descricao", "itemname"):
                idx["nome"] = i
            elif n in ("lote", "batch", "distnumber"):
                idx["lote"] = i
            elif n.startswith("qtd") or n.startswith("quantidade"):
                idx["qtd"] = i
                m = re.search(r"(\d{2}\.\d{2})", str(c))
                if m and not origem_padrao:
                    origem_aba = m.group(1)
            elif n in ("deposito", "deposito destino", "whscode"):
                idx["dest"] = i
            elif n in ("serial", "linha", "seq"):
                idx["serial"] = i
        nomes = {_norm(c) for c in cab}
        if not {"item", "lote", "qtd", "dest"} <= idx.keys():
            continue                    # aba não é de transferência
        if nomes & {"situacao", "saldo", "disponivel", "comprometido"}:
            print(f"  aba '{ws.title}' ignorada (parece resultado de consulta)")
            continue
        if not origem_aba:
            print(f"  aba '{ws.title}' ignorada: sem depósito de origem no título da "
                  f"coluna de quantidade (use --origem)")
            continue
        for nlin, r in enumerate(rows, start=2):
            item = str(r[idx["item"]] or "").strip()
            if not item:
                continue
            linhas.append({
                "aba": ws.title,
                "linha_excel": nlin,
                "serial": r[idx["serial"]] if "serial" in idx else nlin - 1,
                "item": item,
                "nome": str(r[idx["nome"]] or "").strip() if "nome" in idx else "",
                "lote": str(r[idx["lote"]] or "").strip(),
                "qtd": _qtd(r[idx["qtd"]]),
                "origem": origem_aba,
                "destino": str(r[idx["dest"]] or "").strip(),
            })
    wb.close()
    if not linhas:
        sys.exit("Nenhuma aba com as colunas Código do Item / Lote / Qtd / Depósito foi encontrada.")
    return linhas


# ---------------------------------------------------------------- banco
def consultar_hana(args, linhas):
    try:
        from hdbcli import dbapi
    except ImportError:
        sys.exit("Instale o driver do HANA:  pip install hdbcli")
    senha = args.senha or os.environ.get("SAP_HANA_SENHA") or getpass.getpass("Senha HANA: ")
    con = dbapi.connect(address=args.host, port=int(args.porta), user=args.usuario, password=senha)
    try:
        cur = con.cursor()
        cur.execute(f'SET SCHEMA "{args.schema}"')
        itens = sorted({l["item"] for l in linhas})
        whs = sorted({l["origem"] for l in linhas} | {l["destino"] for l in linhas if l["destino"]})
        dados = {"itens": {}, "lotes": {}, "saldos": {}, "depositos": set()}

        for bloco in _blocos(itens, 500):
            ph = ",".join("?" * len(bloco))
            cur.execute(f'SELECT "ItemCode", "ItemName", "ManBtchNum", "frozenFor" '
                        f'FROM "OITM" WHERE "ItemCode" IN ({ph})', bloco)
            for ic, nm, mb, fz in cur.fetchall():
                dados["itens"][ic] = {"nome": nm, "lote": mb == "Y", "inativo": fz == "Y"}

            cur.execute(f'SELECT "ItemCode", "DistNumber", "SysNumber", "Status" '
                        f'FROM "OBTN" WHERE "ItemCode" IN ({ph})', bloco)
            for ic, dn, sn, st in cur.fetchall():
                dados["lotes"][(ic, (dn or "").strip())] = {"sys": sn, "status": str(st)}

            ph_w = ",".join("?" * len(whs))
            cur.execute(f'SELECT B."ItemCode", B."DistNumber", Q."WhsCode", '
                        f'IFNULL(Q."Quantity",0), IFNULL(Q."CommitQty",0) '
                        f'FROM "OBTN" B INNER JOIN "OBTQ" Q '
                        f'  ON Q."ItemCode" = B."ItemCode" AND Q."SysNumber" = B."SysNumber" '
                        f'WHERE B."ItemCode" IN ({ph}) AND Q."WhsCode" IN ({ph_w})',
                        list(bloco) + whs)
            for ic, dn, wc, q, c in cur.fetchall():
                dados["saldos"][(ic, (dn or "").strip(), wc)] = (float(q), float(c))

        ph_w = ",".join("?" * len(whs))
        cur.execute(f'SELECT "WhsCode" FROM "OWHS" WHERE "WhsCode" IN ({ph_w})', whs)
        dados["depositos"] = {r[0] for r in cur.fetchall()}
        return dados
    finally:
        con.close()


def _blocos(lst, n):
    for i in range(0, len(lst), n):
        yield lst[i:i + n]


# ---------------------------------------------------------------- validação
def validar(linhas, dados, args):
    demanda = defaultdict(float)
    for l in linhas:
        if l["qtd"] and l["qtd"] > 0:
            demanda[(l["item"], l["lote"], l["origem"])] += l["qtd"]

    for l in linhas:
        motivos, alertas = [], []
        it = dados["itens"].get(l["item"])
        lt = dados["lotes"].get((l["item"], l["lote"]))
        saldo, comp = dados["saldos"].get((l["item"], l["lote"], l["origem"]), (0.0, 0.0))
        disp = saldo - comp
        dem = demanda.get((l["item"], l["lote"], l["origem"]), 0.0)
        l.update(saldo=saldo, comprometido=comp, disponivel=disp, demanda_total=dem,
                 status_lote=STATUS_LOTE.get(lt["status"], lt["status"]) if lt else "")
        if it and not l["nome"]:
            l["nome"] = it["nome"]

        if l["qtd"] is None or l["qtd"] <= 0:
            motivos.append("QUANTIDADE INVÁLIDA (zero/vazia)")
        if not it:
            motivos.append("ITEM NÃO EXISTE")
        elif not it["lote"]:
            motivos.append("ITEM NÃO É CONTROLADO POR LOTE")
        elif it["inativo"]:
            alertas.append("item inativo")
        if it and not lt:
            motivos.append("LOTE NÃO EXISTE PARA O ITEM")
        if not l["destino"] or l["destino"] not in dados["depositos"]:
            motivos.append(f"DEPÓSITO DESTINO INVÁLIDO ({l['destino']})")
        if l["origem"] not in dados["depositos"]:
            motivos.append(f"DEPÓSITO ORIGEM INVÁLIDO ({l['origem']})")
        if l["destino"] == l["origem"]:
            motivos.append("ORIGEM = DESTINO")

        if lt and l["qtd"] and l["qtd"] > 0:
            if saldo <= 0:
                motivos.append(f"LOTE SEM SALDO NO {l['origem']}")
            elif saldo < dem:
                motivos.append(f"SALDO MENOR QUE O PEDIDO ({saldo:g} < {dem:g})")
            elif disp < dem:
                motivos.append(f"SALDO RESERVADO/COMPROMETIDO (disp. {disp:g} < {dem:g})")
            elif args.bloquear_qualquer_reserva and comp > 0:
                motivos.append(f"LOTE COM RESERVA ({comp:g})")
            if dem > l["qtd"] and not any(m.startswith(("SALDO", "LOTE SEM")) for m in motivos):
                alertas.append(f"lote repetido na planilha, demanda somada {dem:g}")

            if lt["status"] in ("1", "2"):
                if args.bloquear_status:
                    motivos.append(f"LOTE {l['status_lote'].upper()}")
                else:
                    alertas.append(f"lote {l['status_lote']}")

        l["resultado"] = "BLOQUEADO" if motivos else "OK"
        l["motivo"] = "; ".join(motivos)
        l["alerta"] = "; ".join(alertas)
    return linhas


# ---------------------------------------------------------------- saída
def _wb_dtw(hdr, dados_linhas):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append([h[0] for h in hdr])
    ws.append([h[1] for h in hdr])
    for r in dados_linhas:
        ws.append(r)
    return wb


def _fmt_txt(v, sep_dec):
    if v is None:
        return ""
    if isinstance(v, float):
        v = int(v) if v.is_integer() else v
    if isinstance(v, float):
        t = repr(v).replace(".", sep_dec)
    else:
        t = str(v)
    t = t.replace("\r", " ").replace("\n", " ")
    if "\t" in t or '"' in t:          # mesmo critério do Excel
        t = '"' + t.replace('"', '""') + '"'
    return t


def _txt_dtw(hdr, dados_linhas, caminho, args):
    """Igual ao Excel 'Texto (separado por tabulações) (*.txt)':
    tabulação, quebra CRLF, codificação ANSI (Windows-1252)."""
    linhas = [[h[0] for h in hdr], [h[1] for h in hdr]] + list(dados_linhas)
    with open(caminho, "w", encoding=args.codificacao, errors="replace", newline="") as f:
        for r in linhas:
            f.write("\t".join(_fmt_txt(v, args.decimal) for v in r) + "\r\n")


def gerar_dtw(linhas, pasta, args):
    ok = [l for l in linhas if l["resultado"] == "OK"]
    if not ok:
        print("Nenhuma linha aprovada — arquivos DTW não gerados.")
        return 0
    # um documento por par origem -> destino (ou único, se --documento-unico)
    grupos = OrderedDict()
    for l in ok:
        chave = ("UNICO",) if args.documento_unico else (l["origem"], l["destino"])
        grupos.setdefault(chave, []).append(l)

    data = args.data or dt.date.today().strftime("%Y%m%d")
    owtr, wtr1, btnt = [], [], []
    for doc, (chave, ls) in enumerate(grupos.items(), start=1):
        coment = args.comentario
        if not args.documento_unico:
            coment = f"{coment} [{chave[0]} -> {chave[1]}]"
        owtr.append([doc, args.serie, data, coment[:254], -1])
        for n, l in enumerate(ls):
            q = int(l["qtd"]) if float(l["qtd"]).is_integer() else l["qtd"]
            wtr1.append([str(doc), str(n), l["item"], l["nome"], q, l["origem"], l["destino"]])
            btnt.append([str(doc), str(n), l["lote"], q, str(n), l["item"]])
            l["doc_dtw"], l["linha_dtw"] = doc, n

    arquivos = [("OWTR - Transferência de Estoque", OWTR_HDR, owtr),
                ("WTR1 - Linhas Transferencia de Estoque", WTR1_HDR, wtr1),
                ("BTNT - Lotes", BTNT_HDR, btnt)]
    for nome, hdr, rows in arquivos:
        if args.formato in ("xlsx", "ambos"):
            _wb_dtw(hdr, rows).save(os.path.join(pasta, nome + ".xlsx"))
        if args.formato in ("txt", "ambos"):
            _txt_dtw(hdr, rows, os.path.join(pasta, nome + ".txt"), args)
    return len(grupos)


def gerar_relatorio(linhas, pasta):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Validacao"
    cols = [("Aba", "aba"), ("Serial", "serial"), ("Item", "item"), ("Descrição", "nome"),
            ("Lote", "lote"), ("Origem", "origem"), ("Destino", "destino"), ("Qtd pedida", "qtd"),
            ("Demanda total lote", "demanda_total"), ("Saldo origem", "saldo"),
            ("Comprometido", "comprometido"), ("Disponível", "disponivel"),
            ("Status lote", "status_lote"), ("Resultado", "resultado"), ("Motivo", "motivo"),
            ("Alerta", "alerta"), ("Doc DTW", "doc_dtw"), ("Linha DTW", "linha_dtw")]
    ws.append([c[0] for c in cols])
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
    verde = PatternFill("solid", fgColor="E2EFDA")
    verm = PatternFill("solid", fgColor="FCE4D6")
    for l in linhas:
        ws.append([l.get(k) for _, k in cols])
        ws.cell(ws.max_row, 14).fill = verde if l["resultado"] == "OK" else verm
    larg = [22, 7, 10, 45, 15, 8, 8, 11, 12, 12, 12, 11, 12, 11, 50, 35, 8, 9]
    for i, w in enumerate(larg, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    rs = wb.create_sheet("Resumo")
    rs.append(["Aba", "Linhas", "OK", "Bloqueadas", "Qtd OK", "Qtd bloqueada"])
    for aba in OrderedDict.fromkeys(l["aba"] for l in linhas):
        ls = [l for l in linhas if l["aba"] == aba]
        okl = [l for l in ls if l["resultado"] == "OK"]
        rs.append([aba, len(ls), len(okl), len(ls) - len(okl),
                   sum(l["qtd"] or 0 for l in okl),
                   sum(l["qtd"] or 0 for l in ls if l["resultado"] != "OK")])
    for c in rs[1]:
        c.font = Font(bold=True)
    rs.column_dimensions["A"].width = 30
    caminho = os.path.join(pasta, "Relatorio_Validacao.xlsx")
    wb.save(caminho)
    return caminho


# ---------------------------------------------------------------- main
def main(argv=None, dados_mock=None):
    carregar_env()
    p = argparse.ArgumentParser(description="Valida saldo de lotes e gera DTW de Transferência de Estoque")
    p.add_argument("planilha")
    p.add_argument("--host", default=os.environ.get("SAP_HANA_HOST"))
    p.add_argument("--porta", default=os.environ.get("SAP_HANA_PORTA", "30015"))
    p.add_argument("--usuario", default=os.environ.get("SAP_HANA_USUARIO"))
    p.add_argument("--senha", default=None, help="evite; prefira o prompt ou SAP_HANA_SENHA")
    p.add_argument("--schema", default=os.environ.get("SAP_HANA_SCHEMA"), help="banco da empresa (CompanyDB)")
    p.add_argument("--origem", default=None, help="força o depósito de origem (padrão: lido de 'Qtd 01.08')")
    p.add_argument("--serie", type=int, default=int(os.environ.get("DTW_SERIE", "21")), help="Series da OWTR (padrão 21)")
    p.add_argument("--data", default=None, help="DocDate aaaammdd (padrão hoje)")
    p.add_argument("--comentario", default=os.environ.get(
        "DTW_COMENTARIO", "Transferência executada via DTW para ajuste de estoque em massa."))
    p.add_argument("--documento-unico", action="store_true", help="todas as linhas em uma só transferência")
    p.add_argument("--bloquear-qualquer-reserva", action="store_true",
                   help="bloqueia o lote se houver qualquer CommitQty > 0")
    p.add_argument("--bloquear-status", action="store_true",
                   help="bloqueia lotes com status Inacessível/Bloqueado")
    p.add_argument("--saida", default=None, help="pasta de saída")
    p.add_argument("--formato", choices=["txt", "xlsx", "ambos"],
                   default=os.environ.get("DTW_FORMATO", "ambos"),
                   help="arquivos DTW em .txt separado por tabulação, .xlsx ou ambos (padrão)")
    p.add_argument("--decimal", default=".", help="separador decimal no .txt (padrão '.')")
    p.add_argument("--codificacao", default="cp1252",
                   help="codificação do .txt (padrão cp1252 = ANSI do Excel; use utf-16 para 'Texto Unicode')")
    args = p.parse_args(argv)

    linhas = ler_planilha(args.planilha, args.origem)
    print(f"{len(linhas)} linhas lidas de {len({l['aba'] for l in linhas})} aba(s).")

    if dados_mock is not None:
        dados = dados_mock
    else:
        falta = [n for n in ("host", "usuario", "schema") if not getattr(args, n)]
        if falta:
            p.error("informe " + ", ".join("--" + f for f in falta) +
                    "  (ou preencha SAP_HANA_HOST / SAP_HANA_USUARIO / SAP_HANA_SCHEMA no .env)")
        dados = consultar_hana(args, linhas)

    validar(linhas, dados, args)

    pasta = args.saida or os.path.join(os.path.dirname(os.path.abspath(args.planilha)),
                                       "DTW_" + dt.datetime.now().strftime("%Y%m%d_%H%M"))
    os.makedirs(pasta, exist_ok=True)
    ndocs = gerar_dtw(linhas, pasta, args)
    rel = gerar_relatorio(linhas, pasta)

    ok = sum(1 for l in linhas if l["resultado"] == "OK")
    print(f"Aprovadas: {ok}  |  Bloqueadas: {len(linhas) - ok}  |  Documentos DTW: {ndocs}")
    print(f"Arquivos em: {pasta}")
    print(f"Relatório:   {rel}")
    return linhas


if __name__ == "__main__":
    main()
