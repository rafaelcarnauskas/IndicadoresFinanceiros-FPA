#!/usr/bin/env python3
"""Indicadores financeiros FPA: movimentação detalhada (OngSys) -> planilha resumida + slides.

Fluxo:
  1. A pessoa envia a movimentação detalhada exportada do OngSys (.xls ou .xlsx).
  2. O código processa tudo em memória (nada é gravado em disco).
  3. Ela baixa a planilha resumida (.xlsx) e os slides com gráficos editáveis (.pptx).

Como rodar:
  Interface web:   streamlit run app_indicadores.py
  Linha de comando (teste): python app_indicadores.py movimentacao.xls
                    opções: --incluir-provisoes  --projetado proj.csv  --hist resumo_anterior.xlsx
"""
import io
import re
import sys
import unicodedata
from datetime import datetime

import numpy as np
import pandas as pd
import xlrd

# ----------------------------------------------------------------------------
# Tabelas de apoio (vieram do Anexo 02 da prestação de contas e dos slides)
# ----------------------------------------------------------------------------
AREAS = ["Pessoal", "Materiais", "Administrativas", "Serviços", "Instalações", "Patrimônio"]

# rubrica -> área do Anexo 02. "Pagto provisão" fica fora das 6 áreas, como no Anexo.
RUBRICA_AREA = {
    "2.1.01.01.001": "Pessoal",
    "2.1.01.01.006": "Pessoal",
    "2.1.01.01.007": "Pessoal",
    "2.1.01.01.008": "Pessoal",
    "2.1.01.01.009": "Pessoal",
    "2.1.01.01.010": "Pessoal",
    "2.1.01.01.012": "Pessoal",
    "2.1.01.01.013": "Serviços",
    "2.1.01.01.016": "Pessoal",
    "2.1.01.01.017": "Pessoal",
    "2.1.01.01.018": "Pessoal",
    "2.1.01.01.020": "Pagto provisão",
    "2.1.01.01.021": "Pagto provisão",
    "2.1.01.01.022": "Pagto provisão",
    "2.1.01.01.023": "Pagto provisão",
    "2.1.01.01.024": "Pagto provisão",
    "2.1.01.01.025": "Pagto provisão",
    "2.1.01.01.026": "Pagto provisão",
    "2.1.01.01.027": "Pagto provisão",
    "2.1.01.01.028": "Pagto provisão",
    "2.1.01.01.029": "Pagto provisão",
    "2.1.01.01.030": "Pagto provisão",
    "2.1.01.01.031": "Pagto provisão",
    "2.1.01.01.032": "Pagto provisão",
    "2.1.01.01.033": "Serviços",
    "2.1.01.01.035": "Pagto provisão",
    "2.1.01.01.036": "Pessoal",
    "2.1.01.02.001": "Pessoal",
    "2.1.01.02.002": "Pessoal",
    "2.1.01.02.009": "Administrativas",
    "2.1.01.02.010": "Administrativas",
    "2.1.01.04.003": "Administrativas",
    "2.1.01.04.005": "Materiais",
    "2.1.01.04.006": "Materiais",
    "2.1.01.04.011": "Materiais",
    "2.1.01.04.013": "Materiais",
    "2.1.01.04.014": "Materiais",
    "2.1.01.04.015": "Materiais",
    "2.1.01.04.016": "Materiais",
    "2.1.01.04.017": "Materiais",
    "2.1.01.04.018": "Materiais",
    "2.1.01.04.019": "Materiais",
    "2.1.01.04.023": "Materiais",
    "2.1.01.05.001": "Serviços",
    "2.1.01.05.003": "Serviços",
    "2.1.01.05.008": "Serviços",
    "2.1.01.05.009": "Serviços",
    "2.1.01.05.010": "Serviços",
    "2.1.01.05.011": "Serviços",
    "2.1.01.05.012": "Serviços",
    "2.1.01.05.013": "Serviços",
    "2.1.01.05.014": "Serviços",
    "2.1.01.05.015": "Serviços",
    "2.1.01.05.016": "Serviços",
    "2.1.01.05.024": "Administrativas",
    "2.1.01.05.032": "Serviços",
    "2.1.01.05.033": "Pessoal",
    "2.1.01.05.034": "Pessoal",
    "2.1.01.05.038": "Pessoal",
    "2.1.01.05.039": "Serviços",
    "2.1.01.06.001": "Administrativas",
    "2.1.01.06.002": "Administrativas",
    "2.1.01.06.003": "Administrativas",
    "2.1.01.06.004": "Administrativas",
    "2.1.01.06.005": "Administrativas",
    "2.1.01.06.007": "Administrativas",
    "2.1.01.06.008": "Administrativas",
    "2.1.01.06.009": "Administrativas",
    "2.1.01.06.010": "Administrativas",
    "2.1.01.06.012": "Administrativas",
    "2.1.01.06.013": "Administrativas",
    "2.1.01.06.017": "Administrativas",
    "2.1.01.06.019": "Administrativas",
    "2.1.01.06.025": "Administrativas",
    "2.1.01.06.028": "Administrativas",
    "2.1.01.06.030": "Administrativas",
    "2.1.01.06.031": "Administrativas",
    "2.1.01.06.038": "Administrativas",
    "2.1.01.07.001": "Instalações",
    "2.1.01.07.007": "Instalações",
    "2.1.01.07.010": "Instalações",
    "2.1.01.07.011": "Instalações",
    "2.1.01.07.012": "Instalações",
    "2.1.01.07.013": "Instalações",
    "2.1.01.08.012": "Patrimônio",
    "2.1.01.08.015": "Patrimônio",
    "2.1.01.08.028": "Patrimônio",
    "2.1.01.08.030": "Patrimônio",
    "2.1.01.08.031": "Patrimônio",
    "2.1.01.08.033": "Administrativas",
    "2.1.01.08.034": "Patrimônio",
    "2.1.01.08.048": "Patrimônio",
    "2.1.01.08.049": "Patrimônio",
    "2.1.01.08.050": "Patrimônio",
    "2.1.01.08.051": "Patrimônio",
    "2.1.01.08.052": "Patrimônio",
    "2.1.01.08.053": "Patrimônio",
    "3.1.01.01.001": "Pessoal",
    "3.1.01.01.002": "Pessoal",
    "3.1.01.01.003": "Pessoal",
    "3.1.01.01.004": "Pessoal",
    "3.1.01.01.005": "Pessoal",
    "3.1.01.01.006": "Pessoal",
    "3.1.01.01.007": "Pessoal",
    "3.1.01.01.008": "Pessoal",
    "3.1.01.01.009": "Pessoal",
    "3.1.01.01.010": "Administrativas",
    "3.1.01.01.011": "Administrativas",
}

# rubricas de insumos por curso (ordem do slide)
CURSOS = {
    "2.1.01.04.013": "Gastronomia",
    "2.1.01.04.014": "Manut. e reparo",
    "2.1.01.04.015": "Economia criativa",
    "2.1.01.04.016": "Jardinagem",
    "2.1.01.04.017": "Hig. e limpeza",
    "2.1.01.04.023": "Cabeleireiro",
    "2.1.01.04.018": "Corte e costura",
    "2.1.01.04.019": "Lavanderia",
}

# número do CF no OngSys -> nome no slide
CF_NOMES = {
    "3": "Bom Retiro I", "8": "Bom Retiro II", "4": "Brasilândia", "2": "Campo Belo",
    "6": "Ermelino", "7": "Glicério", "5": "Heliópolis", "9": "Penha", "11": "Sede",
    "1": "Geral", "10": "Geraldo Spínola",
}
ORDEM_CF = ["Bom Retiro I", "Bom Retiro II", "Brasilândia", "Campo Belo", "Ermelino",
            "Glicério", "Heliópolis", "Penha", "Sede"]

FONTES_FPA = {
    "1.1.01.01.002": "Doações PF",
    "1.1.01.01.007": "Arrecadação Captador",
    "1.1.01.01.009": "NFP",
    "1.1.01.03.001": "Aluguéis",
}
COD_RENDIMENTO = "1.1.01.02.001"

QD_CATEGORIAS = {"Pessoal": "RH", "Pagto provisão": "RH", "Administrativas": "Administrativo",
                 "Serviços": "Serviços", "Materiais": "Materiais"}

MESES_ABR = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
MESES_EXT = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto",
             "Setembro", "Outubro", "Novembro", "Dezembro"]

COLUNAS = {  # nome interno -> nome no arquivo do OngSys
    "lanc": "LANÇAMENTO", "tipo": "TIPO", "ident": "IDENTIFICAÇÃO", "data": "DATA PGTO",
    "rub": "Rubrica de Receitas e Despesas", "projeto": "Projeto",
    "cf": "Centro de Formação", "bruto": "VALOR BRUTO", "liq": "VALOR LÍQUIDO",
}


# ----------------------------------------------------------------------------
# Leitura e preparação
# ----------------------------------------------------------------------------
def _norm(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s).strip().lower()


def _to_dt(v, datemode):
    if isinstance(v, datetime):
        return pd.Timestamp(v)
    if isinstance(v, (int, float)) and v > 1:
        return pd.Timestamp(xlrd.xldate_as_datetime(v, datemode))
    s = str(v).strip()
    return pd.to_datetime(s, dayfirst=True, errors="coerce") if s else pd.NaT


def ler_movimentacao(conteudo: bytes, nome_arquivo: str) -> pd.DataFrame:
    """Lê o .xls/.xlsx do OngSys e devolve um DataFrame já normalizado."""
    datemode = 0
    if nome_arquivo.lower().endswith(".xls"):
        wb = xlrd.open_workbook(file_contents=conteudo, ignore_workbook_corruption=True)
        sh = wb.sheet_by_index(0)
        linhas = [sh.row_values(i) for i in range(sh.nrows)]
        datemode = wb.datemode
    else:
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(conteudo), read_only=True, data_only=True)
        linhas = [list(r) for r in wb.worksheets[0].iter_rows(values_only=True)]
    if len(linhas) < 2:
        raise ValueError("A planilha está vazia.")
    cab = [str(c).strip() for c in linhas[0]]
    bruto = pd.DataFrame(linhas[1:], columns=cab)

    achadas, faltando = {}, []
    for interno, nome in COLUNAS.items():
        col = next((c for c in bruto.columns if _norm(c) == _norm(nome)), None)
        if col is None:
            faltando.append(nome)
        else:
            achadas[interno] = col
    if faltando:
        raise ValueError("Não encontrei estas colunas na planilha: " + ", ".join(faltando)
                         + ". Confira se é a movimentação detalhada do OngSys.")

    df = pd.DataFrame({k: bruto[v] for k, v in achadas.items()})
    # a última linha do relatório é o total e vem sem número de lançamento
    df = df[df["lanc"].astype(str).str.strip() != ""].copy()
    for c in ("tipo", "ident", "projeto", "cf", "rub"):
        df[c] = df[c].astype(str).str.strip()
    for c in ("bruto", "liq"):
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    df["data"] = df["data"].map(lambda v: _to_dt(v, datemode))
    partes = df["rub"].str.split(" - ", n=1, expand=True)
    df["cod"] = partes[0].str.strip()
    df["rubrica"] = partes[1].fillna("").str.strip() if partes.shape[1] > 1 else ""
    df["cf_num"] = df["cf"].str.split(" - ", n=1).str[0].str.strip()
    df["area"] = df["cod"].map(RUBRICA_AREA).fillna("Não mapeada")
    return df.reset_index(drop=True)


# ----------------------------------------------------------------------------
# Cálculo dos indicadores
# ----------------------------------------------------------------------------
def _soma(df, col="bruto"):
    return float(df[col].sum())


def processar(df, incluir_provisoes=False, projetado=None, historico_anterior=None):
    """Devolve um dicionário com todas as tabelas dos indicadores e as verificações."""
    res = {}
    periodo = df["data"].dt.to_period("M").dropna()
    mes = periodo.mode().iloc[0] if len(periodo) else pd.Period(datetime.now(), "M")
    res["mes_key"] = str(mes)
    res["mes_label"] = f"{MESES_ABR[mes.month - 1]}/{str(mes.year)[2:]}"
    res["mes_extenso"] = f"{MESES_EXT[mes.month - 1]} de {mes.year}"
    res["incluir_provisoes"] = incluir_provisoes

    up = df["projeto"].str.upper()
    pot = df[up.str.contains("POT")]
    despesa_pot = pot[(pot["tipo"] == "Despesa") & (pot["ident"] != "D")].copy()
    despesa_pot["valor"] = despesa_pot["bruto"]
    despesa_pot["provisao"] = despesa_pot["cod"].str.startswith("3.")
    base = despesa_pot if incluir_provisoes else despesa_pot[~despesa_pot["provisao"]]

    # 1) despesa por área (as 6 áreas do Anexo, sem pagamento de provisão)
    real = base[base["area"].isin(AREAS)].groupby("area")["valor"].sum().reindex(AREAS).fillna(0.0)
    tab_area = pd.DataFrame({"area": AREAS, "realizado": real.values})
    tab_area["projetado"] = np.nan
    if projetado is not None and len(projetado):
        proj = projetado.set_index("area")["projetado"]
        tab_area["projetado"] = tab_area["area"].map(proj)
    tab_area["pct_realizado"] = np.where(tab_area["projetado"] > 0,
                                         tab_area["realizado"] / tab_area["projetado"], np.nan)
    res["area"] = tab_area
    res["realizado_total"] = float(tab_area["realizado"].sum())
    res["projetado_total"] = (float(tab_area["projetado"].sum())
                              if tab_area["projetado"].notna().any() else np.nan)

    # 2) custo por CF
    cf = base.copy()
    cf["cf_nome"] = cf["cf_num"].map(CF_NOMES).fillna(cf["cf"])
    tot_cf = cf.groupby("cf_nome")["valor"].sum()
    res["cf_geral"] = float(tot_cf.get("Geral", 0.0))
    res["cf_tabela"] = tot_cf
    # 3) gasto por curso (rubricas de insumos)
    cur = base[base["cod"].isin(CURSOS)].groupby("cod")["valor"].sum()
    res["curso"] = pd.DataFrame({"curso": list(CURSOS.values()),
                                 "valor": [float(cur.get(c, 0.0)) for c in CURSOS]})

    # 4) rescisões e proporção de gastos com RH
    resc_mask = despesa_pot["rubrica"].str.upper().str.contains("RESCIS")
    rescisoes = _soma(despesa_pot[resc_mask], "valor")
    resc_em_pessoal = _soma(base[(base["area"] == "Pessoal")
                                 & base["rubrica"].str.upper().str.contains("RESCIS")], "valor")
    res["rescisoes"] = rescisoes
    res["demais_rh"] = float(tab_area.loc[tab_area["area"] == "Pessoal", "realizado"].iloc[0]
                             - resc_em_pessoal)

    # 5) FPA próprio
    fpa = df[up.str.contains("FPA PR") & (df["ident"] != "TRANSF") & (df["ident"] != "D")]
    rec = fpa[(fpa["tipo"] == "Receita") & (fpa["cod"] != COD_RENDIMENTO)].copy()
    rec["fonte"] = rec["cod"].map(FONTES_FPA).fillna("Outras receitas")
    fontes = rec.groupby("fonte")["bruto"].sum()
    fontes = fontes[fontes > 0].sort_values(ascending=False)
    res["fpa_fontes"] = fontes
    res["fpa_receita"] = float(fontes.sum())
    res["fpa_rendimentos"] = _soma(fpa[(fpa["tipo"] == "Receita") & (fpa["cod"] == COD_RENDIMENTO)])
    res["fpa_despesa"] = _soma(fpa[(fpa["tipo"] == "Despesa") & fpa["cod"].str.startswith("2")])

    # 6) Emenda parlamentar
    em = df[up.str.contains("EMENDA") & (df["tipo"] == "Despesa") & (df["ident"] != "D")
            & df["cod"].str.startswith("2")]
    res["emenda"] = em.groupby("rubrica")["bruto"].sum().sort_values(ascending=False)
    res["emenda_bolsas"] = int(em["rubrica"].str.upper().str.contains("BOLSA").sum())

    # 7) Quebrada Digital (CF Geraldo Spínola)
    # só o projeto FPA próprio: as bolsas da Emenda (também no CF 10) têm indicador próprio
    qd_base = df[(df["cf_num"] == "10") & up.str.contains("FPA PR")
                 & (df["ident"] != "D") & (df["ident"] != "TRANSF")]
    qd = qd_base[(qd_base["tipo"] == "Despesa") & qd_base["cod"].str.startswith("2")].copy()
    qd["categoria"] = qd["area"].map(QD_CATEGORIAS).fillna("Outros")
    cats = ["RH", "Administrativo", "Serviços", "Materiais", "Outros"]
    qd_tot = qd.groupby("categoria")["bruto"].sum().reindex(cats).fillna(0.0)
    res["qd"] = qd_tot[(qd_tot > 0) | qd_tot.index.isin(cats[:4])]
    res["qd_receita"] = _soma(qd_base[(qd_base["tipo"] == "Receita") & (qd_base["cod"] != COD_RENDIMENTO)])

    # 8) histórico para as séries (acumula com a planilha do mês anterior, se enviada)
    atual = pd.DataFrame([
        (res["mes_key"], "Rescisões (POT II)", rescisoes),
        (res["mes_key"], "Receita FPA próprio", res["fpa_receita"]),
        (res["mes_key"], "Despesa FPA próprio", res["fpa_despesa"]),
    ], columns=["mes", "indicador", "valor"])
    if historico_anterior is not None and len(historico_anterior):
        atual = pd.concat([historico_anterior[["mes", "indicador", "valor"]], atual])
    res["historico"] = (atual.drop_duplicates(["mes", "indicador"], keep="last")
                        .sort_values(["indicador", "mes"]).reset_index(drop=True))

    # 9) verificações (para a pessoa conferir antes de usar)
    prov_val = _soma(despesa_pot[despesa_pot["provisao"]], "valor")
    pagto_prov = _soma(despesa_pot[despesa_pot["area"] == "Pagto provisão"], "valor")
    nao_map = despesa_pot[despesa_pot["area"] == "Não mapeada"]
    ajustes = _soma(pot[(pot["tipo"] == "Despesa") & (pot["ident"] == "D")])
    res["nao_mapeadas"] = (nao_map.groupby(["cod", "rubrica"])["valor"].sum().reset_index()
                           if len(nao_map) else pd.DataFrame(columns=["cod", "rubrica", "valor"]))
    res["verificacoes"] = pd.DataFrame([
        ("Linhas lidas na movimentação", len(df), ""),
        ("Linhas do projeto POT II", len(pot), ""),
        ("Despesa POT II (bruto, sem ajustes bancários)", _soma(despesa_pot, "valor"),
         "Base de todos os gráficos do POT II."),
        ("Ajustes bancários excluídos (IOF e IR s/ rendimento)", ajustes,
         "Não entram, como no Anexo 02 e no arquivo da Nubia."),
        ("Provisões (rubricas 3.x)", prov_val,
         "Incluídas nos gráficos." if incluir_provisoes else "Fora dos gráficos (marque a opção para incluir)."),
        ("Pagamento de provisão (férias, rescisões, custas)", pagto_prov,
         "Fica fora das 6 áreas, como no Anexo 02; entra no custo por CF e nas rescisões."),
        ("Despesa realizada nas 6 áreas", res["realizado_total"], ""),
        ("Rubricas sem área no mapa", float(nao_map["valor"].sum()),
         "Ver aba de pendências; aparecem só no custo por CF."),
        ("Rescisões do mês", rescisoes, "Rubricas com 'rescisão' no nome."),
        ("CF '1 - GERAL' (rateio pendente)", res["cf_geral"], ""),
        ("Rendimentos FPA próprio (fora da receita)", res["fpa_rendimentos"], ""),
    ], columns=["item", "valor", "observação"])
    return res


def tabela_cf(res, somar_geral_em_glicerio=True):
    """Custo por CF na ordem do slide. O 'Geral' pode ser somado em Glicério."""
    t = res["cf_tabela"].copy()
    geral = float(t.get("Geral", 0.0))
    ordem = list(ORDEM_CF)
    if somar_geral_em_glicerio:
        t["Glicério"] = float(t.get("Glicério", 0.0)) + geral
    elif geral:
        ordem.append("Geral")
    return pd.DataFrame({"cf": ordem, "valor": [float(t.get(c, 0.0)) for c in ordem]})


def ler_projetado(conteudo: bytes, nome: str):
    """CSV/XLSX com colunas area e projetado."""
    if nome.lower().endswith((".xls", ".xlsx")):
        d = pd.read_excel(io.BytesIO(conteudo))
    else:
        d = pd.read_csv(io.BytesIO(conteudo), sep=None, engine="python")
    d.columns = [_norm(c) for c in d.columns]
    if not {"area", "projetado"} <= set(d.columns):
        raise ValueError("O arquivo de projetado precisa ter as colunas: area, projetado.")
    d["area"] = d["area"].astype(str).str.strip()
    d["projetado"] = pd.to_numeric(d["projetado"], errors="coerce")
    return d[["area", "projetado"]]


def ler_historico(conteudo: bytes):
    try:
        h = pd.read_excel(io.BytesIO(conteudo), sheet_name="Historico")
        h.columns = [_norm(c) for c in h.columns]
        h["valor"] = pd.to_numeric(h["valor"], errors="coerce")
        return h[["mes", "indicador", "valor"]].dropna()
    except Exception:
        raise ValueError("Não achei a aba 'Historico' nessa planilha. Use a planilha resumida gerada aqui.")


# ----------------------------------------------------------------------------
# Planilha resumida (.xlsx)
# ----------------------------------------------------------------------------
def gerar_xlsx(res, somar_geral_em_glicerio=True) -> bytes:
    cf = tabela_cf(res, somar_geral_em_glicerio)
    resumo = [("POT II", "Despesa realizada (6 áreas)", res["realizado_total"])]
    resumo += [("POT II", f"Despesa realizada: {r.area}", r.realizado) for r in res["area"].itertuples()]
    resumo += [("POT II", f"Custo por CF: {r.cf}", r.valor) for r in cf.itertuples()]
    resumo += [("POT II", f"Gasto por curso: {r.curso}", r.valor) for r in res["curso"].itertuples()]
    resumo += [("POT II", "Rescisões do mês", res["rescisoes"]),
               ("POT II", "Demais gastos com RH", res["demais_rh"]),
               ("FPA próprio", "Receita do mês", res["fpa_receita"]),
               ("FPA próprio", "Despesa do mês", res["fpa_despesa"]),
               ("Quebrada Digital", "Receita destinada ao CF", res["qd_receita"]),
               ("Emenda parlamentar", "Bolsas pagas (lançamentos)", res["emenda_bolsas"])]
    resumo += [("FPA próprio", f"Receita: {k}", v) for k, v in res["fpa_fontes"].items()]
    resumo += [("Quebrada Digital", f"Gasto: {k}", v) for k, v in res["qd"].items()]
    resumo += [("Emenda parlamentar", f"Despesa: {k}", v) for k, v in res["emenda"].items()]
    abas = {
        "Resumo": pd.DataFrame(resumo, columns=["indicador", "item", res["mes_label"]]),
        "POT_por_area": res["area"],
        "POT_por_CF": cf,
        "POT_por_curso": res["curso"],
        "FPA_receitas": res["fpa_fontes"].rename("valor").reset_index().rename(columns={"index": "fonte"}),
        "Emenda": res["emenda"].rename("valor").reset_index(),
        "Quebrada_Digital": res["qd"].rename("valor").reset_index().rename(columns={"index": "categoria"}),
        "Historico": res["historico"],
        "Verificacoes": res["verificacoes"],
        "Pendencias": res["nao_mapeadas"],
    }
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        for nome, tab in abas.items():
            tab.to_excel(xw, sheet_name=nome, index=False)
        from openpyxl.styles import Font
        for ws in xw.book.worksheets:
            for c in ws[1]:
                c.font = Font(bold=True)
            for col in ws.columns:
                letra = col[0].column_letter
                larg = max(len(str(c.value)) if c.value is not None else 0 for c in col[:60])
                ws.column_dimensions[letra].width = min(max(12, larg + 2), 60)
                for c in col[1:]:
                    if isinstance(c.value, float):
                        c.number_format = "0.0%" if "pct" in str(col[0].value) else "#,##0.00"
    return buf.getvalue()


# ----------------------------------------------------------------------------
# Slides (.pptx) com gráficos nativos, editáveis no PowerPoint e importáveis no Canva
# ----------------------------------------------------------------------------
LARANJA, AZUL, ROXO, VERDE, CINZA = "F15A24", "3779E3", "9760CF", "4F7A66", "595959"


def brl(v):
    s = f"{abs(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return ("-" if v < 0 else "") + "R$ " + s


def gerar_pptx(res, somar_geral_em_glicerio=True) -> bytes:
    from pptx import Presentation
    from pptx.chart.data import CategoryChartData
    from pptx.dml.color import RGBColor
    from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
    from pptx.enum.text import PP_ALIGN
    from pptx.util import Inches, Pt

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    vazio = prs.slide_layouts[6]
    rgb = RGBColor.from_string

    def texto(slide, x, y, w, h, txt, tam=14, cor=CINZA, negrito=False, alinhar=None):
        tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = txt
        if alinhar:
            p.alignment = alinhar
        p.font.size, p.font.bold = Pt(tam), negrito
        p.font.color.rgb = rgb(cor)
        return tb

    def novo(titulo, subtitulo=None):
        s = prs.slides.add_slide(vazio)
        texto(s, 0.5, 0.3, 12.3, 0.9, titulo, 30, LARANJA, True)
        if subtitulo:
            texto(s, 0.5, 1.05, 12.3, 0.5, subtitulo, 16, AZUL, True)
        return s

    def barras(slide, x, y, w, h, cats, series, cores, tipo=XL_CHART_TYPE.BAR_CLUSTERED,
               fmt="#,##0", legenda=False, titulo=None):
        cats, series = list(cats), {k: list(v) for k, v in series.items()}
        if tipo == XL_CHART_TYPE.BAR_CLUSTERED:  # o PowerPoint desenha a 1ª categoria embaixo
            cats = cats[::-1]
            series = {k: v[::-1] for k, v in series.items()}
        cd = CategoryChartData()
        cd.categories = cats
        for nome, vals in series.items():
            cd.add_series(nome, [float(v) for v in vals])
        gf = slide.shapes.add_chart(tipo, Inches(x), Inches(y), Inches(w), Inches(h), cd)
        ch = gf.chart
        ch.font.size = Pt(12)
        ch.has_title = bool(titulo)
        if titulo:
            ch.chart_title.text_frame.text = titulo
            ch.chart_title.text_frame.paragraphs[0].font.size = Pt(14)
            ch.chart_title.text_frame.paragraphs[0].font.bold = True
        ch.has_legend = legenda
        if legenda:
            ch.legend.position = XL_LEGEND_POSITION.TOP
            ch.legend.include_in_layout = False
        plot = ch.plots[0]
        plot.gap_width = 50
        plot.has_data_labels = True
        dl = plot.data_labels
        dl.number_format, dl.number_format_is_linked = fmt, False
        dl.font.size = Pt(11)
        dl.position = XL_LABEL_POSITION.OUTSIDE_END
        for s, cor in zip(plot.series, cores):
            s.format.fill.solid()
            s.format.fill.fore_color.rgb = rgb(cor)
        va = ch.value_axis
        va.has_major_gridlines = True
        va.major_gridlines.format.line.color.rgb = rgb("D9D9D9")
        va.tick_labels.number_format, va.tick_labels.number_format_is_linked = fmt, False
        return ch

    def pizza(slide, x, y, w, h, cats, vals, cores):
        cd = CategoryChartData()
        cd.categories = list(cats)
        cd.add_series("Receitas", [float(v) for v in vals])
        ch = slide.shapes.add_chart(XL_CHART_TYPE.PIE, Inches(x), Inches(y), Inches(w), Inches(h), cd).chart
        ch.font.size = Pt(13)
        ch.has_title = False
        ch.has_legend = True
        ch.legend.position = XL_LEGEND_POSITION.RIGHT
        ch.legend.include_in_layout = False
        plot = ch.plots[0]
        plot.has_data_labels = True
        plot.data_labels.show_percentage = True
        plot.data_labels.show_value = False
        plot.data_labels.number_format, plot.data_labels.number_format_is_linked = "0.0%", False
        plot.data_labels.font.size = Pt(13)
        for i, cor in enumerate(cores[:len(cats)]):
            pt = plot.series[0].points[i]
            pt.format.fill.solid()
            pt.format.fill.fore_color.rgb = rgb(cor)

    def serie_mensal(indicadores, cores):
        h = res["historico"]
        meses = sorted(h["mes"].unique())
        rotulos = [f"{MESES_ABR[int(m[5:7]) - 1]}/{m[2:4]}" for m in meses]
        series = {}
        for ind in indicadores:
            sub = h[h["indicador"] == ind].set_index("mes")["valor"]
            series[ind] = [float(sub.get(m, 0.0)) for m in meses]
        return rotulos, series

    prov_nota = ("Provisões (rubricas 3.x) incluídas." if res["incluir_provisoes"]
                 else "Provisões (rubricas 3.x) não incluídas.")

    # capa
    s = prs.slides.add_slide(vazio)
    fundo = s.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
    fundo.fill.solid()
    fundo.fill.fore_color.rgb = rgb(AZUL)
    fundo.line.fill.background()
    texto(s, 0.5, 2.4, 12.3, 1.6, "Financeiro", 80, "FFFFFF", True, PP_ALIGN.CENTER)
    texto(s, 0.5, 4.3, 12.3, 0.8, res["mes_extenso"], 28, "FFFFFF", True, PP_ALIGN.CENTER)

    # receitas FPA próprio
    s = novo(f"FINANCEIRO - TOTALIDADE DAS RECEITAS", f"FPA próprio, {res['mes_extenso']}")
    fontes = res["fpa_fontes"]
    if len(fontes):
        pizza(s, 1.5, 1.6, 10.3, 5.0, fontes.index, fontes.values,
              [AZUL, "F4B76A", "FFDD57", "E879C4", "EE7480", ROXO])
    texto(s, 0.5, 6.75, 12.3, 0.6,
          f"Receita total: {brl(res['fpa_receita'])}. Vitrine Social e a divisão entre doação recorrente e "
          "pontual não constam na movimentação detalhada: complementar manualmente.", 12)

    # POT: despesa realizada e por área
    s = novo("POT Redenção", res["mes_extenso"])
    a = res["area"]
    tem_proj = a["projetado"].notna().any()
    if tem_proj:
        pct = res["realizado_total"] / res["projetado_total"] if res["projetado_total"] else 0
        texto(s, 0.5, 1.55, 5.5, 0.5, f"{pct:.0%} da despesa projetada", 20, AZUL, True, PP_ALIGN.CENTER)
        barras(s, 0.4, 2.1, 5.6, 4.4, ["Despesa projetada", "Despesa realizada"],
               {"Valor": [res["projetado_total"], res["realizado_total"]]}, [AZUL],
               XL_CHART_TYPE.COLUMN_CLUSTERED)
        cats = [f"{r.area} ({r.pct_realizado:.0%})" if pd.notna(r.pct_realizado) else r.area
                for r in a.itertuples()]
        barras(s, 6.3, 1.6, 6.6, 4.9, cats, {"Projetado": a["projetado"].fillna(0),
                                              "Realizado": a["realizado"]}, [AZUL, "F4B76A"],
               legenda=True, titulo="Despesa por rubrica")
    else:
        barras(s, 0.4, 1.6, 5.6, 4.9, ["Despesa realizada"], {"Valor": [res["realizado_total"]]},
               [AZUL], XL_CHART_TYPE.COLUMN_CLUSTERED)
        barras(s, 6.3, 1.6, 6.6, 4.9, a["area"], {"Realizado": a["realizado"]}, [AZUL],
               titulo="Despesa por rubrica")
    texto(s, 0.5, 6.7, 12.3, 0.6,
          f"Despesa realizada: {brl(res['realizado_total'])}. {prov_nota}"
          + ("" if tem_proj else " Projetado não informado."), 12)

    # custo por CF e por curso
    s = novo("POT Redenção", f"Custo por CF e gasto por curso, {res['mes_extenso']}")
    cf = tabela_cf(res, somar_geral_em_glicerio)
    barras(s, 0.4, 1.6, 6.4, 5.0, cf["cf"], {"Custo": cf["valor"]}, [LARANJA], titulo="Custo por CF")
    cu = res["curso"]
    barras(s, 6.9, 1.6, 6.0, 4.2, cu["curso"], {"Gasto": cu["valor"]}, [AZUL],
           titulo="Gasto por curso (insumos)")
    nota_geral = (f"CF '1 - GERAL' ({brl(res['cf_geral'])}) somado em Glicério."
                  if somar_geral_em_glicerio and res["cf_geral"] else "")
    texto(s, 6.9, 5.85, 6.0, 1.2, "Comentário: (editar) " + nota_geral, 12)

    # rescisões
    s = novo("POT Redenção", f"Rescisões, {res['mes_extenso']}")
    rot, ser = serie_mensal(["Rescisões (POT II)"], [AZUL])
    barras(s, 0.4, 1.6, 7.2, 4.9, rot, ser, ["40B6D6"], XL_CHART_TYPE.COLUMN_CLUSTERED,
           titulo="Rescisões, série histórica")
    barras(s, 7.8, 1.6, 5.1, 4.9, ["Demais gastos com RH", "Rescisões"],
           {"Valor": [res["demais_rh"], res["rescisoes"]]}, [ROXO],
           titulo=f"Proporção de gastos, {res['mes_label']}")
    texto(s, 0.5, 6.7, 12.3, 0.6, f"Rescisões do mês: {brl(res['rescisoes'])}. {prov_nota}", 12)

    # FPA próprio: receitas e despesas
    s = novo("FPA PRÓPRIO", res["mes_extenso"])
    rot, ser = serie_mensal(["Receita FPA próprio", "Despesa FPA próprio"], [VERDE, "D9A400"])
    barras(s, 0.4, 1.6, 8.6, 4.9, rot, {"Receita": ser["Receita FPA próprio"],
                                        "Despesas": ser["Despesa FPA próprio"]},
           [VERDE, "D9A400"], XL_CHART_TYPE.COLUMN_CLUSTERED, legenda=True,
           titulo="Série histórica de receitas")
    texto(s, 9.4, 2.2, 3.6, 0.5, "Receita no mês", 18, AZUL, True, PP_ALIGN.CENTER)
    texto(s, 9.4, 2.7, 3.6, 0.6, brl(res["fpa_receita"]), 22, AZUL, True, PP_ALIGN.CENTER)
    texto(s, 9.4, 3.8, 3.6, 0.5, "Despesa no mês", 18, LARANJA, True, PP_ALIGN.CENTER)
    texto(s, 9.4, 4.3, 3.6, 0.6, brl(res["fpa_despesa"]), 22, LARANJA, True, PP_ALIGN.CENTER)
    texto(s, 0.5, 6.7, 12.3, 0.6,
          "Receitas: aluguéis e doações. Despesas: gastos administrativos da fundação que não são "
          "financiados por projetos.", 12)

    # Quebrada Digital
    s = novo("FINANCEIRO - QUEBRADA DIGITAL - CF Geraldo Spínola", res["mes_extenso"])
    qd = res["qd"]
    barras(s, 0.4, 1.6, 8.6, 4.9, qd.index, {"Gastos": qd.values}, [AZUL],
           titulo=f"Gastos {res['mes_label']}")
    texto(s, 9.4, 2.4, 3.6, 0.5, "Receita destinada ao CF", 18, AZUL, True, PP_ALIGN.CENTER)
    texto(s, 9.4, 2.9, 3.6, 0.6, brl(res["qd_receita"]), 22, AZUL, True, PP_ALIGN.CENTER)
    texto(s, 0.5, 6.7, 12.3, 0.6, "Obs.: (editar) explicar gastos fora do padrão do mês.", 12)

    # Emenda parlamentar
    s = novo("FINANCEIRO - EMENDA PARLAMENTAR", res["mes_extenso"])
    em = res["emenda"]
    if len(em):
        barras(s, 0.4, 1.6, 8.6, 4.9, em.index, {"Realizado": em.values}, ["9760CF"],
               titulo="Emenda parlamentar, despesas do mês")
    texto(s, 9.4, 2.4, 3.6, 0.5, "Bolsas pagas no mês", 18, AZUL, True, PP_ALIGN.CENTER)
    texto(s, 9.4, 2.9, 3.6, 0.6, str(res["emenda_bolsas"]), 28, AZUL, True, PP_ALIGN.CENTER)
    texto(s, 0.5, 6.7, 12.3, 0.6,
          "Previsto e aporte do projeto não constam na movimentação detalhada: complementar manualmente.", 12)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


# ----------------------------------------------------------------------------
# Interface (Streamlit) e linha de comando
# ----------------------------------------------------------------------------
def interface():
    import streamlit as st

    st.set_page_config(page_title="Indicadores financeiros FPA", page_icon="📊", layout="wide")

    # senha opcional: defina "senha" nos Secrets do Streamlit para proteger o app
    try:
        senha = st.secrets.get("senha")
    except Exception:
        senha = None
    if senha:
        digitada = st.sidebar.text_input("Senha", type="password")
        if digitada != senha:
            st.info("Digite a senha na barra lateral para continuar.")
            st.stop()

    st.title("Indicadores financeiros")
    st.write("Envie a **movimentação detalhada** exportada do OngSys. O app gera a planilha "
             "resumida e os slides com gráficos. Os dados são processados na memória e não ficam salvos.")

    arq = st.file_uploader("Movimentação detalhada (.xls ou .xlsx)", type=["xls", "xlsx"])
    with st.expander("Opções"):
        incluir = st.checkbox("Incluir provisões (rubricas 3.x) na despesa realizada", value=False)
        geral = st.checkbox("Somar o CF '1 - GERAL' em Glicério (como no slide de agosto)", value=True)
        proj_arq = st.file_uploader("Despesa projetada do mês (opcional, colunas: area, projetado)",
                                    type=["csv", "xlsx", "xls"])
        hist_arq = st.file_uploader("Planilha resumida do mês anterior (opcional, para as séries)",
                                    type=["xlsx"])
    if not arq:
        st.stop()

    try:
        df = ler_movimentacao(arq.getvalue(), arq.name)
        proj = ler_projetado(proj_arq.getvalue(), proj_arq.name) if proj_arq else None
        hist = ler_historico(hist_arq.getvalue()) if hist_arq else None
        res = processar(df, incluir, proj, hist)
        xlsx = gerar_xlsx(res, geral)
        pptx = gerar_pptx(res, geral)
    except Exception as e:
        st.error(f"Não consegui processar o arquivo: {e}")
        st.stop()

    st.success(f"Mês identificado: {res['mes_extenso']}")
    c1, c2 = st.columns(2)
    c1.download_button("Baixar planilha resumida", xlsx,
                       file_name=f"indicadores_resumo_{res['mes_key']}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    c2.download_button("Baixar slides com gráficos", pptx,
                       file_name=f"indicadores_slides_{res['mes_key']}.pptx",
                       mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")

    st.subheader("Verificações")
    v = res["verificacoes"].copy()
    v["valor"] = v["valor"].map(lambda x: f"{x:,.2f}")
    st.dataframe(v, hide_index=True, use_container_width=True)
    if len(res["nao_mapeadas"]):
        st.warning("Há rubricas sem área no mapa. Elas não entram nas 6 áreas.")
        st.dataframe(res["nao_mapeadas"], hide_index=True)

    st.subheader("Pré-visualização")
    a, b, c = st.columns(3)
    a.caption("Despesa por área")
    a.bar_chart(res["area"].set_index("area")["realizado"])
    b.caption("Custo por CF")
    b.bar_chart(tabela_cf(res, geral).set_index("cf")["valor"])
    c.caption("Gasto por curso")
    c.bar_chart(res["curso"].set_index("curso")["valor"])


def cli(args):
    arq = args[0]
    incluir = "--incluir-provisoes" in args
    proj = hist = None
    if "--projetado" in args:
        p = args[args.index("--projetado") + 1]
        proj = ler_projetado(open(p, "rb").read(), p)
    if "--hist" in args:
        p = args[args.index("--hist") + 1]
        hist = ler_historico(open(p, "rb").read())
    df = ler_movimentacao(open(arq, "rb").read(), arq)
    res = processar(df, incluir, proj, hist)
    print(res["verificacoes"].to_string(index=False))
    open(f"indicadores_resumo_{res['mes_key']}.xlsx", "wb").write(gerar_xlsx(res))
    open(f"indicadores_slides_{res['mes_key']}.pptx", "wb").write(gerar_pptx(res))
    print("Gerados:", f"indicadores_resumo_{res['mes_key']}.xlsx", f"indicadores_slides_{res['mes_key']}.pptx")


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1].lower().endswith((".xls", ".xlsx")):
        cli(sys.argv[1:])
    else:
        interface()
