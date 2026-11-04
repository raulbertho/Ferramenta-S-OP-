"""
Atualiza o arquivo "Order Fill por Base" com os dados do mês corrente.

Order Fill = Faturamento / (Faturamento + Ruptura), em quantidade.

Como usar (rodar toda semana):
    1. Coloque na pasta de entrada:
       - o arquivo base (ex.: Order_Fill_por_Base_2026.xlsx)
       - o de-para de códigos (nome com "de-para")
       - o faturamento de cada marca exportado do Power BI (nome com "Fat" e BK / PLK / SBX)
       - a ruptura de cada marca (nome com "Ruptura" e BK / PLK / SBX)
       Faturamento e ruptura devem cobrir do dia 01 do mês até a data de corte.
    2. Rode:  python atualizar_order_fill.py --pasta "C:\\caminho\\da\\pasta"
       (opcional: --ate 03/11/2026 para escolher a data de corte)
    3. O resultado é salvo na mesma pasta como Order_Fill_por_Base_<ano>_ate_<dd-mm>.xlsx

O script:
    - cria ou atualiza a linha do mês nos blocos mensais (BK, PLK, SBX);
    - nos blocos semanais, mantém o mês anterior e refaz as semanas do mês atual;
    - refaz as semanas do mês atual nas abas Top Cortes e Ofensores CD Z05;
    - refaz a aba Todos Food com o acumulado do mês.
"""
import argparse
import copy
import datetime as dt
import re
import sys
import unicodedata
import warnings
from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl.formula.translate import Translator
from openpyxl.utils import column_index_from_string as CI
from openpyxl.utils import get_column_letter as L

warnings.filterwarnings("ignore")

# ---------------------------------------------------------------- configuração
PASTA_PADRAO = (r"C:\Users\CarolinaChristofolet\BK Brasil\Gerência de Supply Chain - "
                r"00. Planejamento de Demanda\01. Bases Importantes\Order Fill")

# Código do CD -> base. X = Próprio, Z = Compra e Venda.
PROPRIO = {"XK4": "JDI", "XKF": "R3", "XKA": "ARA", "XE4": "JDI", "XE6": "R3"}
COMPRA_VENDA = {"Z14": "JDI", "Z17": "R3", "Z10": "ARA", "Z05": "JP"}

# Colunas de cada aba principal: coluna -> código do CD
ABAS = {
    "BK": dict(marca="bk", rotulo="C", fat="D", rup="E",
               bases={"M": "XK4", "N": "XKF", "O": "XKA", "R": "Z14", "S": "Z17", "T": "Z10", "U": "Z05"},
               total_proprio="L", total_cv="Q",
               media={"G": ("M", "R"), "H": ("N", "S"), "I": ("O", "T"), "J": ("U", None)}),
    "PLK": dict(marca="plk", rotulo="C", fat="D", rup="E",
                bases={"J": "XK4", "K": "XKF", "M": "Z14", "N": "Z17"},
                total_proprio=None, total_cv=None,
                media={"G": ("J", "M"), "H": ("K", "N")}),
    "SBX": dict(marca="sbx", rotulo="B", fat="C", rup="D",
                bases={}, total_proprio=None, total_cv=None, media={}),
}
DETALHES = [  # aba, marca, filtro de CD (None = todos)
    ("Top Cortes BK", "bk", None),
    ("Top Cortes PLK", "plk", None),
    ("Top Cortes SBX", "sbx", None),
    ("Ofensores CD Z05", "bk", ["Z05"]),
]
MESES = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto",
         "Setembro", "Outubro", "Novembro", "Dezembro"]
ND = "n/d"


# ---------------------------------------------------------------- leitura
def sem_acento(s):
    return "".join(c for c in unicodedata.normalize("NFKD", str(s)) if not unicodedata.combining(c)).lower()


def tem_marca(nome, marca):
    return re.search(rf"(^|[^a-z]){marca}([^a-z]|$)", sem_acento(nome)) is not None


def achar(pasta, *palavras, marca=None, excluir=()):
    arqs = [p for p in pasta.glob("*.xls*") if not p.name.startswith("~$")]
    arqs = [p for p in arqs if all(w in sem_acento(p.name) for w in palavras)
            and not any(e in sem_acento(p.name) for e in excluir)]
    if marca:
        arqs = [p for p in arqs if tem_marca(p.stem, marca)]
    if not arqs:
        return None
    return max(arqs, key=lambda p: p.stat().st_mtime)


def ler_faturamento(arq):
    df = pd.read_excel(arq)
    df.columns = [sem_acento(str(c).split("[")[-1].rstrip("]")).strip() for c in df.columns]
    qtd = next(c for c in df.columns if "qtd" in c and "venda" in c)
    data = next(c for c in df.columns if "data emiss" in c)
    df = df[df["empresa"].isin(list(PROPRIO) + list(COMPRA_VENDA))]
    return pd.DataFrame({
        "cod": pd.to_numeric(df["codigo produto"]).astype(int),
        "desc": df["ingredient_master_description"] if "ingredient_master_description" in df else df.get("descricao"),
        "base": df["empresa"], "data": pd.to_datetime(df[data]).dt.normalize(),
        "fat": pd.to_numeric(df[qtd]).fillna(0), "rup": 0})


def ler_ruptura(arq):
    xl = pd.ExcelFile(arq)
    aba = next(s for s in xl.sheet_names if s.strip().upper() == "BASE")
    bruto = pd.read_excel(arq, sheet_name=aba, header=None, nrows=30)
    lin = next(i for i, r in bruto.iterrows() if any(str(v).strip() == "DATA CORTE" for v in r.values))
    df = pd.read_excel(arq, sheet_name=aba, header=lin)
    df.columns = [str(c).strip() for c in df.columns]
    qtd = next(c for c in df.columns if c.upper().startswith("QTD"))
    df = df[pd.to_datetime(df["DATA CORTE"], errors="coerce").notna()
            & pd.to_numeric(df["CODIGO"], errors="coerce").notna()]
    return pd.DataFrame({
        "cod": pd.to_numeric(df["CODIGO"], errors="coerce"), "desc": df["DESCRIÇÃO"],
        "base": df["BASE"].astype(str).str.strip(), "data": pd.to_datetime(df["DATA CORTE"]).dt.normalize(),
        "fat": 0, "rup": pd.to_numeric(df[qtd], errors="coerce").fillna(0)}).astype({"cod": int})


def ler_depara(arq):
    dp = pd.read_excel(arq).dropna(how="all")
    dp = dp[dp["Código"].notna()].copy()
    dp["Código"] = dp["Código"].astype(int)
    dp = dp.drop_duplicates("Código").set_index("Código")
    ativo_linha = set(dp.index[(dp["Item ativo?"] == "Ativo") & (dp["Classificação"] == "Linha")])
    nbo_flag = set(dp.index[dp["NBO"].isin(["NBO", "Sim"])])
    food = set(dp.index[dp["Categoria"] == "Food"])
    return dp, ativo_linha, ativo_linha & nbo_flag, nbo_flag, food


# ---------------------------------------------------------------- semanas
def semanas_do_mes(inicio, fim):
    """Semanas de segunda a domingo, cortadas no mês.
    Se o mês começa num domingo, esse dia entra junto com a semana seguinte."""
    dias = pd.date_range(inicio, fim)
    grupos = {}
    for d in dias:
        grupos.setdefault(d.isocalendar().week, []).append(d)
    lista = [(k, v[0], v[-1]) for k, v in grupos.items()]
    if len(lista) > 1 and lista[0][1] == lista[0][2]:  # primeira semana com 1 dia só
        lista = [(lista[1][0], lista[0][1], lista[1][2])] + lista[2:]
    return [(f"Semana {k}", f"{a:%d/%m}–{b:%d/%m}", a, b) for k, a, b in lista]


def mes_do_rotulo(txt):
    m = re.search(r"\d{2}/(\d{2})", str(txt))
    return int(m.group(1)) if m else None


# ---------------------------------------------------------------- motor de cálculo
class Dados:
    def __init__(self, dados, ativo_linha, nbo, dp):
        self.d, self.al, self.nbo, self.dp = dados, ativo_linha, nbo, dp

    def sel(self, marca, filtro, d0, d1, bases=None):
        d = self.d[marca]
        d = d[(d.data >= d0) & (d.data <= d1) & d.cod.isin(self.al if filtro == "al" else self.nbo)]
        return d[d.base.isin(bases)] if bases else d

    @staticmethod
    def of(d):
        f, r = d.fat.sum(), d.rup.sum()
        return ND if f + r == 0 else round(float(f / (f + r)), 6)

    @staticmethod
    def tot(d):
        return int(d.fat.sum()), int(d.rup.sum())

    def itens(self, d, n=10, so_ruptura=True):
        g = d.groupby("cod").agg(fat=("fat", "sum"), rup=("rup", "sum")).reset_index()
        desc = d.sort_values("fat", ascending=False).drop_duplicates("cod").set_index("cod")["desc"]
        g["desc"] = g.cod.map(lambda c: self.dp["Descricao"].get(c, desc.get(c)))
        if so_ruptura:
            g = g[g.rup > 0]
        g = g.sort_values(["rup", "fat"], ascending=[False, False])
        return (g.head(n) if n else g).to_dict("records")


# ---------------------------------------------------------------- escrita na planilha
def foto(ws):
    linhas = {r: {c: (ws.cell(r, c).value, copy.copy(ws.cell(r, c)._style))
                  for c in range(1, ws.max_column + 1)} for r in range(1, ws.max_row + 1)}
    alturas = {r: ws.row_dimensions[r].height for r in range(1, ws.max_row + 1)}
    mescl = [(m.min_row, m.min_col, m.max_row, m.max_col) for m in ws.merged_cells.ranges]
    return linhas, alturas, mescl, ws.max_column


def reescrever(ws, plano, ft):
    """plano: lista de (linha_origem ou None, {coluna: valor}, linha_modelo)."""
    linhas, alturas, mescl, maxc = ft
    for m in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(m))
    ws.delete_rows(1, ws.max_row + 5)
    nova, modelos = {}, []
    for i, (src, ov, tpl) in enumerate(plano, start=1):
        origem = src or tpl
        base = linhas.get(origem, {})
        for c in range(1, maxc + 1):
            v, st = base.get(c, (None, None))
            if isinstance(v, str) and v.startswith("=") and origem:
                v = Translator(v, origin=f"{L(c)}{origem}").translate_formula(f"{L(c)}{i}")
            if callable(ov.get(c)):
                v = ov[c](i)
            elif c in ov:
                v = ov[c]
            cel = ws.cell(i, c)
            cel.value = v
            if st is not None:
                cel._style = copy.copy(st)
        if alturas.get(origem):
            ws.row_dimensions[i].height = alturas[origem]
        if src:
            nova[src] = i
        elif tpl:
            modelos.append((tpl, i))
    for (r1, c1, r2, c2) in mescl:
        if r1 != r2:
            continue
        destinos = ([nova[r1]] if r1 in nova else []) + [i for t, i in modelos if t == r1]
        for i in destinos:
            ws.merge_cells(start_row=i, start_column=c1, end_row=i, end_column=c2)


def formula_media(a, b, vals):
    if b is None:
        return lambda i: f"={a}{i}"
    if vals.get(a) == ND or vals.get(b) == ND:
        return lambda i: f'=IFERROR(AVERAGE({a}{i},{b}{i}),"n/d")'
    return lambda i: f"=({a}{i}+{b}{i})/2"


def valores_linha(D, cfg, filtro, d0, d1, aba, semanal):
    d = D.sel(cfg["marca"], filtro, d0, d1)
    f, r = D.tot(d)
    ov = {CI(cfg["fat"]): f, CI(cfg["rup"]): r}
    vals = {col: D.of(d[d.base == b]) for col, b in cfg["bases"].items()}
    if cfg["total_proprio"]:
        vals[cfg["total_proprio"]] = D.of(d[d.base.isin(list(PROPRIO))])
    if cfg["total_cv"]:
        vals[cfg["total_cv"]] = D.of(d[d.base.isin(list(COMPRA_VENDA))])
    if aba == "SBX":
        x = D.of(d[d.base == "XE4"])
        vals["F"] = x
        if not semanal:
            vals["I"] = x
    for col, v in vals.items():
        ov[CI(col)] = v
    for col, (a, b) in cfg["media"].items():
        ov[CI(col)] = formula_media(a, b, vals)
    return ov


def atualizar_principal(wb, aba, D, ano, mes, inicio, fim, semanas):
    ws, cfg = wb[aba], ABAS[aba]
    ft = foto(ws)
    linhas = ft[0]
    lab = CI(cfg["rotulo"])
    nome_mes = MESES[mes - 1]
    mes_ant = 12 if mes == 1 else mes - 1
    plano, tipo, filtro = [], None, "al"
    maxr = max(linhas)
    rot = lambda r: linhas.get(r, {}).get(lab, (None,))[0]
    for r in range(1, maxr + 1):
        txt = " ".join(str(linhas[r][c][0]) for c in range(1, 4) if linhas[r][c][0])
        v = rot(r)
        if "Order Fill" in txt:
            filtro = "nbo" if "NBO" in txt else "al"
        if v == "Mês":
            tipo = "m"
        elif v == "Semana":
            tipo = "w"
        dado = tipo and v not in (None, "Mês", "Semana")
        fim_bloco = dado and rot(r + 1) is None

        if tipo == "m" and dado:
            if v == nome_mes:  # mês já existe: atualiza
                plano.append((r, valores_linha(D, cfg, filtro, inicio, fim, aba, False), None))
            else:
                plano.append((r, {}, None))
            if fim_bloco and nome_mes not in [rot(x) for x in range(r - 12, r + 1)]:
                ov = valores_linha(D, cfg, filtro, inicio, fim, aba, False)
                ov[lab] = nome_mes
                plano.append((None, ov, r))
        elif tipo == "w" and dado:
            if mes_do_rotulo(v) == mes_ant:
                plano.append((r, {}, None))
            if fim_bloco:
                for _, rotulo, d0, d1 in semanas:
                    ov = valores_linha(D, cfg, filtro, d0, d1, aba, True)
                    ov[lab] = rotulo
                    plano.append((None, ov, r))
        else:
            plano.append((r, {}, None))
        if fim_bloco:
            tipo = None
    reescrever(ws, plano, ft)


def atualizar_detalhe(wb, aba, marca, bases, D, ano, mes, fim, semanas):
    ws = wb[aba]
    ft = foto(ws)
    linhas = ft[0]
    maxr = max(linhas)
    mes_ant = 12 if mes == 1 else mes - 1
    inicios = [r for r in range(1, maxr + 1) if str(linhas[r][1][0] or "").startswith("Semana")]
    plano = [(r, {}, None) for r in range(1, inicios[0])]
    a2 = str(linhas[2][1][0] or "")
    data_ini_ant = dt.date(ano if mes > 1 else ano - 1, mes_ant, 1)
    a2 = re.sub(r"— [^—]*$", f"— {MESES[mes_ant - 1]} e {MESES[mes - 1]}/{ano} (dados até {fim:%d/%m/%Y})", a2) \
        if "Top Cortes" in aba else \
        re.sub(r"dados de \d{2}/\d{2} a \d{2}/\d{2}/\d{4}", f"dados de {data_ini_ant:%d/%m} a {fim:%d/%m/%Y}", a2)
    plano[1] = (2, {1: a2}, None)
    for s, e in zip(inicios, inicios[1:] + [maxr + 1]):
        if mes_do_rotulo(linhas[s][1][0]) == mes_ant:
            plano += [(r, {}, None) for r in range(s, e)]
    while plano and plano[-1][0] and all(linhas[plano[-1][0]][c][0] is None for c in range(1, 16)):
        plano.pop()
    t = inicios[-1]  # última seção serve de modelo de formatação
    for semana, rotulo, d0, d1 in semanas:
        plano += [(None, {}, None), (None, {}, None)]
        titulo = f"{semana} ({rotulo})"
        plano.append((None, {1: f"{titulo} — Ativos e Linha", 9: f"{titulo} — NBO"}, t))
        plano.append((None, {}, t + 1))
        plano.append((None, {}, t + 2))
        dl, dr = D.sel(marca, "al", d0, d1, bases), D.sel(marca, "nbo", d0, d1, bases)
        lados = [(D.itens(dl), *D.tot(dl), 0), (D.itens(dr), *D.tot(dr), 8)]
        n = max(len(lados[0][0]), len(lados[1][0]), 1)
        for k in range(n):
            ov = {c: None for c in range(1, 16)}
            for itens, F, R, off in lados:
                c0 = 1 + off
                if k < len(itens):
                    x = itens[k]
                    cf, cr = L(c0 + 3), L(c0 + 4)
                    ov.update({c0: k + 1, c0 + 1: int(x["cod"]), c0 + 2: x["desc"],
                               c0 + 3: int(x["fat"]), c0 + 4: int(x["rup"]),
                               c0 + 5: (lambda i, cf=cf, cr=cr: f'=IFERROR({cf}{i}/({cf}{i}+{cr}{i}),"n/d")'),
                               c0 + 6: (lambda i, cr=cr, F=F, R=R:
                                        f'=IFERROR(({F}/({F}+{R}-{cr}{i}))-({F}/({F}+{R})),"n/d")')})
                elif k == 0:
                    ov[c0 + 2] = "Sem rupturas na semana"
            plano.append((None, ov, t + 3))
    reescrever(ws, plano, ft)


def atualizar_food(wb, D, food, nbo_flag, mes, ano, inicio, fim):
    ws = wb["Todos Food"]
    ft = foto(ws)
    d = D.d["sbx"]
    d = d[(d.data >= inicio) & (d.data <= fim) & d.cod.isin(food)]
    itens = [x for x in D.itens(d, n=None, so_ruptura=False) if x["fat"] + x["rup"] > 0]
    ult = 4 + max(len(itens), 1)
    plano = [(r, {}, None) for r in range(1, 5)]
    plano[0] = (1, {1: f"Todos os Itens Food — SBX {inicio:%d} a {fim:%d/%m}"}, None)
    plano[1] = (2, {1: f"Order Fill acumulado de {MESES[mes - 1]} ({inicio:%d/%m} a {fim:%d/%m/%Y}) "
                       "para cada codigo classificado como categoria 'Food' no de-para"}, None)
    plano[2] = (3, {10: f"=SUM(G5:G{ult})/SUM(F5:G{ult})"}, None)
    for k, x in enumerate(itens):
        c = int(x["cod"])
        plano.append((None, {1: k + 1, 2: c, 3: x["desc"], 4: "Sim" if c in D.al else "Não",
                             5: "Sim" if c in nbo_flag else "Não", 6: int(x["fat"]), 7: int(x["rup"]),
                             8: (lambda i: f"=F{i}/(F{i}+G{i})")}, 5))
    reescrever(ws, plano, ft)


# ---------------------------------------------------------------- principal
def main():
    ap = argparse.ArgumentParser(description="Atualiza o Order Fill por Base")
    ap.add_argument("--pasta", help="pasta com os arquivos de entrada (padrão: a pasta do script; "
                                    "se lá não houver planilhas, a pasta Order Fill do SharePoint)")
    ap.add_argument("--base", help="arquivo Order Fill a atualizar (padrão: o mais recente da pasta)")
    ap.add_argument("--ate", help="data de corte dd/mm/aaaa (padrão: último dia com faturamento E ruptura)")
    ap.add_argument("--saida", help="arquivo de saída")
    a = ap.parse_args()
    if a.pasta:
        pasta = Path(a.pasta)
    else:
        pasta = Path(__file__).resolve().parent
        if not achar(pasta, "order", "fill") and Path(PASTA_PADRAO).exists():
            pasta = Path(PASTA_PADRAO)
    print(f"Pasta:   {pasta}")
    if not pasta.exists():
        sys.exit("Essa pasta não existe. Confira o caminho.")

    base = Path(a.base) if a.base else achar(pasta, "order", "fill")
    depara = achar(pasta, "de", "para")
    if not base or not depara:
        planilhas = sorted(p.name for p in pasta.glob("*.xls*"))
        sys.exit("Não encontrei " + ("a planilha Order Fill" if not base else "o de-para") + " nessa pasta.\n"
                 "Coloque nela: planilha Order Fill, de-para, 3 faturamentos e 3 rupturas.\n"
                 "Planilhas encontradas: " + (", ".join(planilhas) if planilhas else "nenhuma"))
    print(f"Base:    {base.name}\nDe-para: {depara.name}")

    dados = {}
    for m in ["bk", "plk", "sbx"]:
        fa = achar(pasta, "fat", marca=m, excluir=("order",))
        ru = achar(pasta, "ruptura", marca=m)
        if not fa or not ru:
            sys.exit(f"Faltou o arquivo de faturamento ou de ruptura da marca {m.upper()}.")
        f, r = ler_faturamento(fa), ler_ruptura(ru)
        if f.empty:
            sys.exit(f"O faturamento {fa.name} veio vazio. Abra no Excel, clique em Atualizar Tudo e salve.")
        print(f"{m.upper():4} fat: {fa.name} (até {f.data.max():%d/%m})  |  ruptura: {ru.name} (até {r.data.max():%d/%m})")
        dados[m] = pd.concat([f, r], ignore_index=True)

    if a.ate:
        fim = pd.Timestamp(dt.datetime.strptime(a.ate, "%d/%m/%Y"))
    else:
        fim = min(min(d[d.fat > 0].data.max(), d[d.rup > 0].data.max()) for d in dados.values())
    inicio = fim.replace(day=1)
    for m in dados:
        dados[m] = dados[m][(dados[m].data >= inicio) & (dados[m].data <= fim)]
    print(f"Período: {inicio:%d/%m/%Y} a {fim:%d/%m/%Y}")

    dp, al, nbo, nbo_flag, food = ler_depara(depara)
    D = Dados(dados, al, nbo, dp)
    semanas = semanas_do_mes(inicio, fim)
    print("Semanas:", ", ".join(f"{s} ({r})" for s, r, *_ in semanas))

    wb = openpyxl.load_workbook(base)
    for aba in ABAS:
        if aba in wb.sheetnames:
            atualizar_principal(wb, aba, D, fim.year, fim.month, inicio, fim, semanas)
    for aba, marca, bases in DETALHES:
        if aba in wb.sheetnames:
            atualizar_detalhe(wb, aba, marca, bases, D, fim.year, fim.month, fim, semanas)
    if "Todos Food" in wb.sheetnames:
        atualizar_food(wb, D, food, nbo_flag, fim.month, fim.year, inicio, fim)

    saida = Path(a.saida) if a.saida else pasta / f"Order_Fill_por_Base_{fim.year}_ate_{fim:%d-%m}.xlsx"
    wb.save(saida)

    print("\nResumo (Ativos e Linha | NBO):")
    for m in ["bk", "plk", "sbx"]:
        x, y = D.sel(m, "al", inicio, fim), D.sel(m, "nbo", inicio, fim)
        print(f"  {m.upper():4} {D.of(x):.1%} | {D.of(y):.1%}")
    fora = []
    for m, d in dados.items():
        g = d[~d.cod.isin(dp.index)].groupby("cod").agg(desc=("desc", "first"), fat=("fat", "sum"),
                                                         rup=("rup", "sum")).reset_index()
        fora += [f"  {m.upper()} {int(r.cod)} {r.desc}: fat {int(r.fat)}, rup {int(r.rup)}" for r in g.itertuples()]
    if fora:
        print("\nItens que não estão no de-para (ficaram fora do cálculo):")
        print("\n".join(fora))
    print(f"\nArquivo salvo: {saida}")


if __name__ == "__main__":
    main()
