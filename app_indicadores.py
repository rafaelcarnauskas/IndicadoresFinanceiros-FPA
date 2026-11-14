#!/usr/bin/env python3
"""Indicadores financeiros FPA: relatório financeiro (OngSys) -> planilha resumida + slides.

Fluxo:
  1. Envie o relatório financeiro exportado do OngSys (.xls ou .xlsx). Pode ser o arquivo do
     jeito que sai do sistema ou a versão editada pela Nubia (sem mudar o layout).
  2. Envie a planilha de mapeamento de rubricas (rubrica -> grupo).
  3. Opcional: envie a movimentação detalhada para os slides de FPA próprio, Quebrada Digital
     e Emenda parlamentar (o relatório financeiro só traz o POT II).
  4. Baixe a planilha resumida (.xlsx) e os slides (.pptx). Nada é gravado em disco.

Como rodar:
  Interface web:    streamlit run app_indicadores.py
  Linha de comando: python app_indicadores.py relatorio_financeiro.xls --mapa mapa_rubricas.xlsx
                    opções: --mov movimentacao.xls  --sem-provisoes  --projetado proj.csv
                            --hist resumo_anterior.xlsx
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
# Tabelas de apoio
# ----------------------------------------------------------------------------
AREAS = ["Pessoal", "Materiais", "Administrativas", "Serviços", "Instalações", "Patrimônio"]

# Nomes de grupo da planilha de mapeamento -> nomes usados nas 6 áreas e nos slides.
GRUPO_ALIAS = {
    "pessoal": "Pessoal",
    "materiais": "Materiais",
    "administrativo": "Administrativas",
    "administrativas": "Administrativas",
    "servicos": "Serviços",
    "instalacoes": "Instalações",
    "patrimonio": "Patrimônio",
    "pagto provisao": "Pagto provisão",
    "pagamento de provisao": "Pagto provisão",
}

# O grupo de cada rubrica vem da planilha de mapeamento enviada a cada processamento.
# Regra das provisões (é a que reproduz o relatório oficial de set/2026):
#   - provisões (rubricas 3.x) entram na despesa por área: o custo é reconhecido quando
#     o dinheiro vai para a conta de provisão;
#   - pagamentos de provisão (grupo "Pagto provisão": férias, 13º, rescisão, custas...)
#     ficam fora das 6 áreas, senão o mesmo custo seria contado duas vezes;
#   - o custo por CF não usa provisões e usa os pagamentos.

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

# nome interno -> nome da coluna no relatório financeiro do OngSys
COLUNAS_FIN = {
    "lanc": "LANÇAMENTO", "tipo": "TIPO", "ident": "IDENTIFICAÇÃO", "data": "DATA PGTO",
    "cod": "NUM. RUBRICA DE RECEITAS E DESPESAS", "rubrica": "RUBRICA DE RECEITAS E DESPESAS",
    "projeto": "PROJETO", "cf": "Centro de Formação", "curso": "Processo Formativo",
    "valor": "VALOR",
}
# nome interno -> nome da coluna na movimentação detalhada do OngSys
COLUNAS_MOV = {
    "lanc": "LANÇAMENTO", "tipo": "TIPO", "ident": "IDENTIFICAÇÃO", "data": "DATA PGTO",
    "rub": "Rubrica de Receitas e Despesas", "projeto": "Projeto",
    "cf": "Centro de Formação", "bruto": "VALOR BRUTO", "liq": "VALOR LÍQUIDO",
}

# código (id) de rubrica: ex. "2.1.01.01.001" (13 caracteres)
COD_RUBRICA_RE = re.compile(r"^\d\.\d{1,2}\.\d{2}\.\d{2}\.\d{3}")


# ----------------------------------------------------------------------------
# Leitura
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


def _cod(valor) -> str:
    """Id da rubrica (13 primeiros caracteres). Devolve '' se não for um código válido."""
    s = str(valor).strip()
    if COD_RUBRICA_RE.match(s[:13]):
        return s[:13]
    primeiro = s.split(" - ", 1)[0].strip()
    return primeiro if COD_RUBRICA_RE.match(primeiro) else ""


def _ler_linhas(conteudo: bytes, nome_arquivo: str):
    """Linhas da primeira aba do .xls/.xlsx e o datemode (para datas em número do Excel)."""
    if nome_arquivo.lower().endswith(".xls"):
        wb = xlrd.open_workbook(file_contents=conteudo, ignore_workbook_corruption=True)
        sh = wb.sheet_by_index(0)
        return [sh.row_values(i) for i in range(sh.nrows)], wb.datemode
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(conteudo), read_only=True, data_only=True)
    return [list(r) for r in wb.worksheets[0].iter_rows(values_only=True)], 0


def _tabela(linhas, colunas: dict, descricao: str) -> pd.DataFrame:
    """Acha a linha de cabeçalho (entre as 15 primeiras), confere se todas as colunas
    existem e devolve um DataFrame com os nomes internos."""
    alvo = {_norm(n) for n in colunas.values()}
    ini = next((i for i, l in enumerate(linhas[:15])
                if alvo <= {_norm(c) for c in l if c is not None}), None)
    if ini is None:
        achadas = {_norm(c) for l in linhas[:15] for c in l if c is not None}
        faltando = [n for n in colunas.values() if _norm(n) not in achadas]
        raise ValueError(f"Não encontrei estas colunas no {descricao}: {', '.join(faltando)}. "
                         "Confira se o arquivo é o certo e se o cabeçalho não foi alterado.")
    cab = [_norm(c) if c is not None else "" for c in linhas[ini]]
    idx = {k: cab.index(_norm(n)) for k, n in colunas.items()}
    dados = linhas[ini + 1:]
    df = pd.DataFrame({k: [l[i] if i < len(l) else None for l in dados] for k, i in idx.items()})
    df = df.astype(object).where(df.notna(), "")
    # tira linhas vazias e a linha de total (vêm sem número de lançamento)
    return df[df["lanc"].astype(str).str.strip() != ""].copy()


def ler_relatorio_financeiro(conteudo: bytes, nome_arquivo: str, mapa_rubricas: dict):
    """Lê o relatório financeiro do OngSys (original ou editado) e devolve (despesas, info).

    Fica só com as despesas: tira receitas (rendimentos, entradas na conta de provisão,
    descontos) e ajustes bancários (TIPO 'D': IOF, IR sobre rendimento). Os valores de
    despesa vêm negativos no OngSys e são convertidos para positivos.
    """
    if not mapa_rubricas:
        raise ValueError("Envie a planilha de mapeamento de rubricas (colunas: Rubrica, Grupo).")
    linhas, datemode = _ler_linhas(conteudo, nome_arquivo)
    df = _tabela(linhas, COLUNAS_FIN, "relatório financeiro")
    for c in ("lanc", "tipo", "ident", "projeto", "cf", "curso", "rubrica"):
        df[c] = df[c].astype(str).str.strip()
    despesa = df["ident"].map(_norm) == "despesa"
    ajuste = df["tipo"].str.upper() == "D"
    info = {"lidas": len(df), "receitas": int((~despesa).sum()), "ajustes": int((despesa & ajuste).sum())}
    df = df[despesa & ~ajuste].copy()

    df["valor"] = -pd.to_numeric(df["valor"], errors="coerce").fillna(0.0)
    df["data"] = df["data"].map(lambda v: _to_dt(v, datemode))
    cod = df["cod"].map(_cod)
    # linhas de juros vêm do OngSys sem rubrica válida: herdam a do mesmo lançamento e CF
    sem_cod = cod.eq("")
    herdado = cod.where(~sem_cod).groupby([df["lanc"], df["cf"]]).transform("first")
    df["cod"] = cod.where(~sem_cod, herdado.fillna(""))
    info["herdadas"] = int((sem_cod & df["cod"].ne("")).sum())
    df["cf_num"] = df["cf"].str.split(" - ", n=1).str[0].str.strip()
    df["area"] = df["cod"].map(mapa_rubricas).fillna("Não mapeada")
    df["provisao"] = df["cod"].str.startswith("3.")
    if df.empty:
        raise ValueError("Não encontrei nenhuma despesa no relatório financeiro.")
    return df.reset_index(drop=True), info


def ler_movimentacao(conteudo: bytes, nome_arquivo: str, mapa_rubricas: dict) -> pd.DataFrame:
    """Lê a movimentação detalhada do OngSys (usada nos slides de FPA próprio,
    Quebrada Digital e Emenda parlamentar)."""
    if not mapa_rubricas:
        raise ValueError("Envie a planilha de mapeamento de rubricas (colunas: Rubrica, Grupo).")
    linhas, datemode = _ler_linhas(conteudo, nome_arquivo)
    df = _tabela(linhas, COLUNAS_MOV, "arquivo de movimentação detalhada")
    for c in ("tipo", "ident", "projeto", "cf", "rub"):
        df[c] = df[c].astype(str).str.strip()
    for c in ("bruto", "liq"):
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    df["data"] = df["data"].map(lambda v: _to_dt(v, datemode))
    df["cod"] = df["rub"].map(_cod)
    df["rubrica"] = np.where(df["cod"].ne(""), df["rub"].str.slice(13).str.lstrip(" -").str.strip(),
                             df["rub"])
    df["cf_num"] = df["cf"].str.split(" - ", n=1).str[0].str.strip()
    df["area"] = df["cod"].map(mapa_rubricas).fillna("Não mapeada")
    return df.reset_index(drop=True)


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


def ler_mapa_rubricas(conteudo: bytes, nome: str) -> dict:
    """Lê a planilha de mapeamento rubrica -> grupo e devolve {codigo: grupo}.

    Usa as colunas Rubrica e Grupo (Projeto e descrição são ignorados). Linhas sem código
    de rubrica válido (ex.: títulos de seção) ou sem grupo são ignoradas; essas rubricas
    aparecem como "Não mapeada" nas pendências.
    """
    if nome.lower().endswith((".xls", ".xlsx")):
        d = pd.read_excel(io.BytesIO(conteudo))
    else:
        d = pd.read_csv(io.BytesIO(conteudo), sep=None, engine="python")
    norm_cols = {_norm(c): c for c in d.columns}
    col_rub = next((norm_cols[k] for k in norm_cols
                    if k in ("rubrica", "codigo", "conta", "num. rubrica de receitas e despesas")), None)
    col_grp = next((norm_cols[k] for k in norm_cols if k in ("grupo", "area", "grupo da despesa")), None)
    if col_rub is None or col_grp is None:
        raise ValueError("A planilha de mapeamento precisa ter as colunas: Rubrica e Grupo.")
    mapa = {}
    for cod, grp in zip(d[col_rub], d[col_grp]):
        cod = _cod(cod)
        if not cod or pd.isna(grp) or not str(grp).strip():
            continue
        mapa[cod] = GRUPO_ALIAS.get(_norm(grp), str(grp).strip())
    if not mapa:
        raise ValueError("Não encontrei nenhuma rubrica válida na planilha de mapeamento. "
                         "Confira as colunas Rubrica e Grupo.")
    return mapa


def ler_historico(conteudo: bytes):
    try:
        h = pd.read_excel(io.BytesIO(conteudo), sheet_name="Historico")
        h.columns = [_norm(c) for c in h.columns]
        h["valor"] = pd.to_numeric(h["valor"], errors="coerce")
        return h[["mes", "indicador", "valor"]].dropna()
    except Exception:
        raise ValueError("Não achei a aba 'Historico' nessa planilha. Use a planilha resumida gerada aqui.")


# ----------------------------------------------------------------------------
# Cálculo dos indicadores
# ----------------------------------------------------------------------------
def _soma(df, col):
    return float(df[col].sum())


def _mes(datas):
    periodo = datas.dt.to_period("M").dropna()
    return periodo.mode().iloc[0] if len(periodo) else pd.Period(datetime.now(), "M")


def _resc(df):
    return df["rubrica"].str.upper().str.contains("RESCIS")


def processar(fin, mov=None, incluir_provisoes=True, projetado=None, historico_anterior=None,
              info_fin=None):
    """Calcula todos os indicadores. fin = despesas do relatório financeiro (POT);
    mov = movimentação detalhada (opcional, para FPA próprio, Quebrada Digital e Emenda)."""
    res = {"tem_mov": mov is not None, "incluir_provisoes": incluir_provisoes}
    info_fin = info_fin or {}
    pot = fin[fin["projeto"].str.upper().str.contains("POT")]
    if pot.empty:
        raise ValueError("O relatório financeiro não tem linhas do projeto POT.")
    mes = _mes(pot["data"])
    res["mes_key"] = str(mes)
    res["mes_label"] = f"{MESES_ABR[mes.month - 1]}/{str(mes.year)[2:]}"
    res["mes_extenso"] = f"{MESES_EXT[mes.month - 1]} de {mes.year}"

    sem_prov = pot[~pot["provisao"]]
    base_area = pot if incluir_provisoes else sem_prov

    # 1) despesa por área: só as 6 áreas (o grupo "Pagto provisão" fica de fora)
    real = (base_area[base_area["area"].isin(AREAS)].groupby("area")["valor"].sum()
            .reindex(AREAS).fillna(0.0))
    tab_area = pd.DataFrame({"area": AREAS, "realizado": real.values})
    tab_area["projetado"] = np.nan
    if projetado is not None and len(projetado):
        tab_area["projetado"] = tab_area["area"].map(projetado.set_index("area")["projetado"])
    tab_area["pct_realizado"] = np.where(tab_area["projetado"] > 0,
                                         tab_area["realizado"] / tab_area["projetado"], np.nan)
    res["area"] = tab_area
    res["realizado_total"] = float(tab_area["realizado"].sum())
    res["projetado_total"] = (float(tab_area["projetado"].sum())
                              if tab_area["projetado"].notna().any() else np.nan)

    # 2) custo por CF: sem provisões, com pagamentos de provisão
    cf = sem_prov.assign(cf_nome=sem_prov["cf_num"].map(CF_NOMES).fillna(sem_prov["cf"]))
    tot_cf = cf.groupby("cf_nome")["valor"].sum()
    res["cf_geral"] = float(tot_cf.get("Geral", 0.0))
    res["cf_tabela"] = tot_cf

    # 3) gasto por curso (rubricas de insumos)
    cur = sem_prov[sem_prov["cod"].isin(CURSOS)].groupby("cod")["valor"].sum()
    res["curso"] = pd.DataFrame({"curso": list(CURSOS.values()),
                                 "valor": [float(cur.get(c, 0.0)) for c in CURSOS]})

    # 4) rescisões e proporção de gastos com RH
    rescisoes = _soma(sem_prov[_resc(sem_prov)], "valor")
    resc_em_pessoal = _soma(base_area[(base_area["area"] == "Pessoal") & _resc(base_area)], "valor")
    res["rescisoes"] = rescisoes
    res["demais_rh"] = float(tab_area.loc[tab_area["area"] == "Pessoal", "realizado"].iloc[0]
                             - resc_em_pessoal)

    # 5 a 7) projetos que só estão na movimentação detalhada
    res["fpa_fontes"] = pd.Series(dtype=float)
    res["fpa_receita"] = res["fpa_despesa"] = res["fpa_rendimentos"] = res["qd_receita"] = 0.0
    res["emenda"] = pd.Series(dtype=float)
    res["emenda_bolsas"] = 0
    res["qd"] = pd.Series(dtype=float)
    res["aviso_mes"] = ""
    if mov is not None:
        mes_mov = _mes(mov["data"])
        if mes_mov != mes:
            res["aviso_mes"] = (f"A movimentação detalhada é de {MESES_EXT[mes_mov.month - 1]}/{mes_mov.year}, "
                                f"mas o relatório financeiro é de {res['mes_extenso']}.")
        up = mov["projeto"].str.upper()

        # 5) FPA próprio
        fpa = mov[up.str.contains("FPA PR") & (mov["ident"] != "TRANSF") & (mov["ident"] != "D")]
        rec = fpa[(fpa["tipo"] == "Receita") & (fpa["cod"] != COD_RENDIMENTO)]
        fontes = rec.groupby(rec["cod"].map(FONTES_FPA).fillna("Outras receitas"))["bruto"].sum()
        fontes = fontes[fontes > 0].sort_values(ascending=False)
        res["fpa_fontes"] = fontes
        res["fpa_receita"] = float(fontes.sum())
        res["fpa_rendimentos"] = _soma(fpa[(fpa["tipo"] == "Receita") & (fpa["cod"] == COD_RENDIMENTO)], "bruto")
        res["fpa_despesa"] = _soma(fpa[(fpa["tipo"] == "Despesa") & fpa["cod"].str.startswith("2")], "bruto")

        # 6) Emenda parlamentar
        em = mov[up.str.contains("EMENDA") & (mov["tipo"] == "Despesa") & (mov["ident"] != "D")
                 & mov["cod"].str.startswith("2")]
        res["emenda"] = em.groupby("rubrica")["bruto"].sum().sort_values(ascending=False)
        res["emenda_bolsas"] = int(em["rubrica"].str.upper().str.contains("BOLSA").sum())

        # 7) Quebrada Digital (CF Geraldo Spínola, projeto FPA próprio)
        qd_base = mov[(mov["cf_num"] == "10") & up.str.contains("FPA PR")
                      & (mov["ident"] != "D") & (mov["ident"] != "TRANSF")]
        qd = qd_base[(qd_base["tipo"] == "Despesa") & qd_base["cod"].str.startswith("2")]
        cats = ["RH", "Administrativo", "Serviços", "Materiais", "Outros"]
        qd_tot = (qd.groupby(qd["area"].map(QD_CATEGORIAS).fillna("Outros"))["bruto"].sum()
                  .reindex(cats).fillna(0.0))
        res["qd"] = qd_tot[(qd_tot > 0) | qd_tot.index.isin(cats[:4])]
        res["qd_receita"] = _soma(qd_base[(qd_base["tipo"] == "Receita")
                                          & (qd_base["cod"] != COD_RENDIMENTO)], "bruto")

    # 8) histórico para as séries (acumula com a planilha do mês anterior, se enviada)
    linhas_hist = [(res["mes_key"], "Rescisões (POT II)", rescisoes)]
    if mov is not None:
        linhas_hist += [(res["mes_key"], "Receita FPA próprio", res["fpa_receita"]),
                        (res["mes_key"], "Despesa FPA próprio", res["fpa_despesa"])]
    atual = pd.DataFrame(linhas_hist, columns=["mes", "indicador", "valor"])
    if historico_anterior is not None and len(historico_anterior):
        atual = pd.concat([historico_anterior[["mes", "indicador", "valor"]], atual])
    res["historico"] = (atual.drop_duplicates(["mes", "indicador"], keep="last")
                        .sort_values(["indicador", "mes"]).reset_index(drop=True))

    # 9) verificações (para conferir antes de usar)
    prov_val = _soma(pot[pot["provisao"]], "valor")
    pagto_prov = _soma(pot[pot["area"] == "Pagto provisão"], "valor")
    nao_map = pot[pot["area"] == "Não mapeada"]
    res["nao_mapeadas"] = (nao_map.groupby(["cod", "rubrica"])["valor"].sum().reset_index()
                           if len(nao_map) else pd.DataFrame(columns=["cod", "rubrica", "valor"]))
    verif = [
        ("Linhas lidas no relatório financeiro", info_fin.get("lidas", len(fin)), ""),
        ("Linhas de receita retiradas", info_fin.get("receitas", 0),
         "Rendimentos, entradas na conta de provisão e descontos não são despesa."),
        ("Ajustes bancários retirados (linhas)", info_fin.get("ajustes", 0), "IOF e IR sobre rendimento."),
        ("Linhas de juros sem rubrica", info_fin.get("herdadas", 0),
         "Receberam a rubrica do próprio lançamento."),
        ("Linhas de despesa do POT II", len(pot), ""),
        ("Despesa POT II no relatório (tudo)", _soma(pot, "valor"), ""),
        ("Provisões (rubricas 3.x)", prov_val,
         ("Entram na despesa por área." if incluir_provisoes else "Fora da despesa por área.")
         + " Nunca entram no custo por CF."),
        ("Pagamento de provisão (férias, 13º, rescisões, custas)", pagto_prov,
         "Fora das 6 áreas; entra no custo por CF e nas rescisões. Se ficar zerado, confira se a "
         "planilha de mapeamento marca essas rubricas com o grupo 'Pagto provisão'."),
        ("Despesa realizada nas 6 áreas", res["realizado_total"], ""),
        ("Despesas com rubrica sem grupo no mapa", float(nao_map["valor"].sum()),
         "Ver aba Pendencias; entram só no custo por CF."),
        ("Rescisões do mês", rescisoes, "Rubricas com 'rescisão' no nome."),
        ("CF '1 - GERAL' (rateio pendente)", res["cf_geral"], ""),
    ]
    if mov is not None:
        verif.append(("Rendimentos FPA próprio (fora da receita)", res["fpa_rendimentos"], ""))
    res["verificacoes"] = pd.DataFrame(verif, columns=["item", "valor", "observação"])
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
               ("POT II", "Demais gastos com RH", res["demais_rh"])]
    abas = {}
    if res["tem_mov"]:
        resumo += [("FPA próprio", "Receita do mês", res["fpa_receita"]),
                   ("FPA próprio", "Despesa do mês", res["fpa_despesa"]),
                   ("Quebrada Digital", "Receita destinada ao CF", res["qd_receita"]),
                   ("Emenda parlamentar", "Bolsas pagas (lançamentos)", res["emenda_bolsas"])]
        resumo += [("FPA próprio", f"Receita: {k}", v) for k, v in res["fpa_fontes"].items()]
        resumo += [("Quebrada Digital", f"Gasto: {k}", v) for k, v in res["qd"].items()]
        resumo += [("Emenda parlamentar", f"Despesa: {k}", v) for k, v in res["emenda"].items()]
    abas["Resumo"] = pd.DataFrame(resumo, columns=["indicador", "item", res["mes_label"]])
    abas["POT_por_area"] = res["area"]
    abas["POT_por_CF"] = cf
    abas["POT_por_curso"] = res["curso"]
    if res["tem_mov"]:
        abas["FPA_receitas"] = res["fpa_fontes"].rename_axis("fonte").rename("valor").reset_index()
        abas["Emenda"] = res["emenda"].rename_axis("rubrica").rename("valor").reset_index()
        abas["Quebrada_Digital"] = res["qd"].rename_axis("categoria").rename("valor").reset_index()
    abas["Historico"] = res["historico"]
    abas["Verificacoes"] = res["verificacoes"]
    abas["Pendencias"] = res["nao_mapeadas"]

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

    def serie_mensal(indicadores):
        h = res["historico"]
        meses = sorted(h["mes"].unique())
        rotulos = [f"{MESES_ABR[int(m[5:7]) - 1]}/{m[2:4]}" for m in meses]
        series = {}
        for ind in indicadores:
            sub = h[h["indicador"] == ind].set_index("mes")["valor"]
            series[ind] = [float(sub.get(m, 0.0)) for m in meses]
        return rotulos, series

    prov_nota = ("Provisões (rubricas 3.x) incluídas; pagamentos de provisão fora das 6 áreas."
                 if res["incluir_provisoes"] else "Provisões (rubricas 3.x) não incluídas.")

    # capa
    s = prs.slides.add_slide(vazio)
    fundo = s.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
    fundo.fill.solid()
    fundo.fill.fore_color.rgb = rgb(AZUL)
    fundo.line.fill.background()
    texto(s, 0.5, 2.4, 12.3, 1.6, "Financeiro", 80, "FFFFFF", True, PP_ALIGN.CENTER)
    texto(s, 0.5, 4.3, 12.3, 0.8, res["mes_extenso"], 28, "FFFFFF", True, PP_ALIGN.CENTER)

    # receitas FPA próprio (movimentação)
    if res["tem_mov"]:
        s = novo("FINANCEIRO - TOTALIDADE DAS RECEITAS", f"FPA próprio, {res['mes_extenso']}")
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
    nota_geral = (f" CF '1 - GERAL' ({brl(res['cf_geral'])}) somado em Glicério."
                  if somar_geral_em_glicerio and res["cf_geral"] else "")
    texto(s, 6.9, 5.85, 6.0, 1.2,
          "Comentário: (editar) Custo por CF sem provisões (3.x)." + nota_geral, 12)

    # rescisões
    s = novo("POT Redenção", f"Rescisões, {res['mes_extenso']}")
    rot, ser = serie_mensal(["Rescisões (POT II)"])
    barras(s, 0.4, 1.6, 7.2, 4.9, rot, ser, ["40B6D6"], XL_CHART_TYPE.COLUMN_CLUSTERED,
           titulo="Rescisões, série histórica")
    barras(s, 7.8, 1.6, 5.1, 4.9, ["Demais gastos com RH", "Rescisões"],
           {"Valor": [res["demais_rh"], res["rescisoes"]]}, [ROXO],
           titulo=f"Proporção de gastos, {res['mes_label']}")
    texto(s, 0.5, 6.7, 12.3, 0.6, f"Rescisões do mês: {brl(res['rescisoes'])}.", 12)

    if res["tem_mov"]:
        # FPA próprio: receitas e despesas
        s = novo("FPA PRÓPRIO", res["mes_extenso"])
        rot, ser = serie_mensal(["Receita FPA próprio", "Despesa FPA próprio"])
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
            barras(s, 0.4, 1.6, 8.6, 4.9, em.index, {"Realizado": em.values}, [ROXO],
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
    st.write("Envie o **relatório financeiro** exportado do OngSys (o original ou a versão editada) e a "
             "**planilha de mapeamento de rubricas**. A movimentação detalhada é opcional e serve para os "
             "slides de FPA próprio, Quebrada Digital e Emenda parlamentar. Os dados são processados na "
             "memória e não ficam salvos.")

    fin_arq = st.file_uploader("Relatório financeiro do OngSys (.xls ou .xlsx)", type=["xls", "xlsx"])
    mapa_arq = st.file_uploader("Mapeamento de rubricas (colunas: Rubrica, Grupo)",
                                type=["csv", "xlsx", "xls"])
    mov_arq = st.file_uploader("Movimentação detalhada (opcional: FPA próprio, Quebrada Digital e Emenda)",
                               type=["xls", "xlsx"])
    with st.expander("Opções"):
        incluir = st.checkbox("Incluir provisões (rubricas 3.x) na despesa por área (como no relatório oficial)",
                              value=True)
        geral = st.checkbox("Somar o CF '1 - GERAL' em Glicério (como no slide de agosto)", value=True)
        proj_arq = st.file_uploader("Despesa projetada do mês (opcional, colunas: area, projetado)",
                                    type=["csv", "xlsx", "xls"])
        hist_arq = st.file_uploader("Planilha resumida do mês anterior (opcional, para as séries)",
                                    type=["xlsx"])
    if not fin_arq or not mapa_arq:
        st.info("Envie o relatório financeiro e a planilha de mapeamento para continuar.")
        st.stop()

    try:
        mapa = ler_mapa_rubricas(mapa_arq.getvalue(), mapa_arq.name)
        fin, info = ler_relatorio_financeiro(fin_arq.getvalue(), fin_arq.name, mapa)
        mov = ler_movimentacao(mov_arq.getvalue(), mov_arq.name, mapa) if mov_arq else None
        proj = ler_projetado(proj_arq.getvalue(), proj_arq.name) if proj_arq else None
        hist = ler_historico(hist_arq.getvalue()) if hist_arq else None
        res = processar(fin, mov, incluir, proj, hist, info)
        xlsx = gerar_xlsx(res, geral)
        pptx = gerar_pptx(res, geral)
    except Exception as e:
        st.error(f"Não consegui processar os arquivos: {e}")
        st.stop()

    st.success(f"Mês identificado: {res['mes_extenso']}")
    st.caption(f"Mapeamento de rubricas aplicado: {mapa_arq.name} ({len(mapa)} rubricas).")
    if res["aviso_mes"]:
        st.warning(res["aviso_mes"])
    if not res["tem_mov"]:
        st.info("Sem a movimentação detalhada, os slides de FPA próprio, Quebrada Digital e Emenda "
                "não foram gerados.")
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
    st.dataframe(v, hide_index=True, width="stretch")
    if len(res["nao_mapeadas"]):
        st.warning("Há despesas com rubrica sem grupo na planilha de mapeamento. Elas não entram nas 6 áreas.")
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
    def opcao(flag):
        if flag not in args:
            return None
        i = args.index(flag)
        if i + 1 >= len(args):
            sys.exit(f"Erro: faltou o arquivo depois de {flag}.")
        return args[i + 1]

    def abrir(p):
        with open(p, "rb") as f:
            return f.read()

    arq = args[0]
    p = opcao("--mapa")
    if not p:
        sys.exit("Erro: informe a planilha de mapeamento de rubricas com --mapa arquivo.xlsx.")
    mapa = ler_mapa_rubricas(abrir(p), p)
    fin, info = ler_relatorio_financeiro(abrir(arq), arq, mapa)
    p = opcao("--mov")
    mov = ler_movimentacao(abrir(p), p, mapa) if p else None
    p = opcao("--projetado")
    proj = ler_projetado(abrir(p), p) if p else None
    p = opcao("--hist")
    hist = ler_historico(abrir(p)) if p else None
    res = processar(fin, mov, "--sem-provisoes" not in args, proj, hist, info)
    print(res["verificacoes"].to_string(index=False))
    if res["aviso_mes"]:
        print("ATENÇÃO:", res["aviso_mes"])
    nomes = (f"indicadores_resumo_{res['mes_key']}.xlsx", f"indicadores_slides_{res['mes_key']}.pptx")
    with open(nomes[0], "wb") as f:
        f.write(gerar_xlsx(res))
    with open(nomes[1], "wb") as f:
        f.write(gerar_pptx(res))
    print("Gerados:", *nomes)


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1].lower().endswith((".xls", ".xlsx")):
        cli(sys.argv[1:])
    else:
        interface()
