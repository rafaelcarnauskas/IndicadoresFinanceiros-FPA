#!/usr/bin/env python3
"""Indicadores financeiros FPA: relatório financeiro (OngSys) -> planilha resumida + slides.

Arquivos (todos .xls/.xlsx, como saem do OngSys ou das planilhas da FPA):
  - Relatório(s) financeiro(s) do OngSys (obrigatório). Pode enviar vários meses de uma vez,
    ou um relatório de um período longo: o app separa por mês. Pode ser o arquivo original ou
    a versão editada pela Nubia (sem mudar o layout).
  - Mapeamento de rubricas (obrigatório): uma ou mais planilhas no formato das "Estruturas"
    (Projeto, Rubrica, descrição, Grupo). Com a coluna Projeto, cada projeto usa a sua.
  - Orçamento POT (opcional): projetado do mês, projetado x realizado acumulado, acompanhamento
    mensal e insumos.
  - Movimentação detalhada (opcional): FPA próprio, Quebrada Digital e, se o relatório financeiro
    não trouxer o projeto, a Emenda parlamentar.
  - Orçamento Emenda (opcional): previsto x realizado da Emenda.
  - Planilha resumida de um mês anterior (opcional): continua as séries e os acumulados.

Como rodar:
  Interface web:    streamlit run app_indicadores.py
  Linha de comando: python app_indicadores.py relatorio1.xls [relatorio2.xls ...] --mapa estrutura.xlsx
                    opções (podem se repetir): --mapa  --mov  --orcamento  --orc-emenda  --hist
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

# Nomes de grupo das planilhas de mapeamento -> nomes usados nas 6 áreas e nos slides.
# Qualquer outro grupo (ex.: "-", "Pagto provisão") fica fora das 6 áreas.
GRUPO_ALIAS = {
    "pessoal": "Pessoal",
    "material": "Materiais",
    "materiais": "Materiais",
    "administrativo": "Administrativas",
    "administrativas": "Administrativas",
    "servico": "Serviços",
    "servicos": "Serviços",
    "instalacao": "Instalações",
    "instalacoes": "Instalações",
    "patrimonio": "Patrimônio",
    "pagto provisao": "Pagto provisão",
    "pagamento de provisao": "Pagto provisão",
}
FORA = "Não mapeada"

# O grupo de cada rubrica vem da planilha de mapeamento: é ela que decide o que entra
# em cada área (inclusive provisões e pagamentos de provisão). Regras fixas do app:
#   - o custo por CF não usa provisões (rubricas 3.x);
#   - rescisões = rubricas com "rescisão" no nome, sem provisões.

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
# o comparativo de insumos projetado x realizado também mostra o EPI
INSUMOS = {**CURSOS, "2.1.01.04.011": "EPI"}

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
    "1.1.01.01.010": "Vitrine Social",
    "1.1.01.03.001": "Aluguéis",
}
COD_RENDIMENTO = "1.1.01.02.001"

QD_CATEGORIAS = {"Pessoal": "RH", "Pagto provisão": "RH", "Administrativas": "Administrativo",
                 "Serviços": "Serviços", "Materiais": "Materiais"}

MESES_ABR = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
MESES_EXT = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto",
             "Setembro", "Outubro", "Novembro", "Dezembro"]
MES_NUM = {m.lower(): i + 1 for i, m in enumerate(MESES_ABR)}

# nome interno -> nome da coluna no relatório financeiro do OngSys
COLUNAS_FIN = {
    "lanc": "LANÇAMENTO", "tipo": "TIPO", "ident": "IDENTIFICAÇÃO", "data": "DATA PGTO",
    "cod": "NUM. RUBRICA DE RECEITAS E DESPESAS", "rubrica": "RUBRICA DE RECEITAS E DESPESAS",
    "projeto": "PROJETO", "cf": "Centro de Formação", "curso": "Processo Formativo",
    "valor": "VALOR",
}
OPCIONAIS_FIN = {"num_projeto": "NUM. PROJETO"}
# nome interno -> nome da coluna na movimentação detalhada do OngSys
COLUNAS_MOV = {
    "lanc": "LANÇAMENTO", "tipo": "TIPO", "ident": "IDENTIFICAÇÃO", "data": "DATA PGTO",
    "rub": "Rubrica de Receitas e Despesas", "projeto": "Projeto",
    "cf": "Centro de Formação", "bruto": "VALOR BRUTO", "liq": "VALOR LÍQUIDO",
}

# código (id) de rubrica: ex. "2.1.01.01.001" (13 caracteres)
COD_RUBRICA_RE = re.compile(r"^\d\.\d{1,2}\.\d{2}\.\d{2}\.\d{3}")

# nomes antigos de indicadores no histórico (planilhas geradas por versões anteriores)
HIST_ALIAS = {"Rescisões (POT II)": "POT II - rescisões",
              "Receita FPA próprio": "FPA próprio - receita",
              "Despesa FPA próprio": "FPA próprio - despesa"}


# ----------------------------------------------------------------------------
# Leitura
# ----------------------------------------------------------------------------
def _norm(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s).strip().lower()


def _num(v):
    x = pd.to_numeric(v, errors="coerce")
    return 0.0 if pd.isna(x) else float(x)


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


def _proj(valor) -> str:
    """Número do projeto: '1.2.06 - POT2 - PMPS' -> '1.2.06'."""
    s = str(valor).strip()
    return "" if s.lower() in ("", "nan", "none") else s.split(" - ", 1)[0].strip()


def _rotulo(mes_key: str) -> str:
    """'2026-09' -> 'Set/26'."""
    return f"{MESES_ABR[int(mes_key[5:7]) - 1]}/{mes_key[2:4]}"


def _ler_linhas(conteudo: bytes, nome_arquivo: str):
    """Linhas da primeira aba do .xls/.xlsx e o datemode (para datas em número do Excel)."""
    if nome_arquivo.lower().endswith(".xls"):
        wb = xlrd.open_workbook(file_contents=conteudo, ignore_workbook_corruption=True)
        sh = wb.sheet_by_index(0)
        return [sh.row_values(i) for i in range(sh.nrows)], wb.datemode
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(conteudo), read_only=True, data_only=True)
    return [list(r) for r in wb.worksheets[0].iter_rows(values_only=True)], 0


def _tabela(linhas, colunas: dict, descricao: str, opcionais=None) -> pd.DataFrame:
    """Acha a linha de cabeçalho (entre as 15 primeiras), confere se todas as colunas
    existem e devolve um DataFrame com os nomes internos."""
    opcionais = opcionais or {}
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
    idx.update({k: cab.index(_norm(n)) for k, n in opcionais.items() if _norm(n) in cab})
    dados = linhas[ini + 1:]
    df = pd.DataFrame({k: [l[i] if i < len(l) else None for l in dados] for k, i in idx.items()})
    for k in opcionais:
        if k not in df:
            df[k] = ""
    df = df.astype(object).where(df.notna(), "")
    # tira linhas vazias e a linha de total (vêm sem número de lançamento)
    return df[df["lanc"].astype(str).str.strip() != ""].copy()


def _atribuir_mes(df: pd.DataFrame, nome: str) -> pd.DataFrame:
    """Coluna 'mes' (texto AAAA-MM). Se quase todo o arquivo (90%) é de um mês, como num
    relatório filtrado por mês, o arquivo inteiro vai para esse mês. Se não (relatório de um
    período longo), cada linha vai para o mês da data de pagamento."""
    per = df["data"].dt.to_period("M")
    moda = per.mode()
    if not len(moda):
        raise ValueError(f"Não encontrei datas de pagamento em {nome}.")
    moda = moda.iloc[0]
    if (per == moda).mean() >= 0.9:
        df["mes"] = str(moda)
    else:
        df["mes"] = per.astype(str).where(per.notna(), str(moda))
    df["arquivo"] = nome
    return df


def _juntar(partes, descricao):
    """Junta vários arquivos e impede que dois arquivos tragam o mesmo mês (contaria em dobro).
    O mesmo lançamento pode aparecer em meses diferentes (parcelas)."""
    vistos = {}
    for p in partes:
        nome = p["arquivo"].iloc[0]
        for m in p["mes"].unique():
            if m in vistos and vistos[m] != nome:
                raise ValueError(f"Dois arquivos de {descricao} trazem {_rotulo(m)}: {vistos[m]} e {nome}. "
                                 "Envie cada mês uma vez só.")
            vistos[m] = nome
    return pd.concat(partes, ignore_index=True)


def _aplicar_mapa(proj: pd.Series, cod: pd.Series, mapa: dict) -> pd.Series:
    """Grupo de cada linha. Se o projeto da linha tem planilha de mapeamento própria, usa só
    ela; se não, usa o grupo da rubrica nas planilhas enviadas."""
    com_mapa = {p for p, _ in mapa if p}
    geral = {}
    for (p, c), g in mapa.items():
        if p == "" or c not in geral:
            geral[c] = g
    return pd.Series([mapa.get((p, c), FORA) if p in com_mapa else geral.get(c, FORA)
                      for p, c in zip(proj, cod)], index=cod.index)


def ler_mapa_rubricas(conteudo: bytes, nome: str) -> dict:
    """Lê uma planilha de mapeamento (formato das Estruturas) e devolve {(projeto, rubrica): grupo}.

    Usa as colunas Rubrica e Grupo, e Projeto se existir (descrição é ignorada). Linhas sem
    código de rubrica válido (ex.: títulos de seção) ou sem grupo são ignoradas.
    """
    def abrir(header):
        if nome.lower().endswith((".xls", ".xlsx")):
            return pd.read_excel(io.BytesIO(conteudo), dtype=str, header=header)
        return pd.read_csv(io.BytesIO(conteudo), sep=None, engine="python", dtype=str, header=header)

    d = abrir(0)
    norm_cols = {_norm(c): c for c in d.columns}
    col_rub = next((norm_cols[k] for k in norm_cols
                    if k in ("rubrica", "codigo", "conta", "num. rubrica de receitas e despesas")), None)
    col_grp = next((norm_cols[k] for k in norm_cols if k in ("grupo", "area", "grupo da despesa")), None)
    col_proj = next((norm_cols[k] for k in norm_cols if k in ("projeto", "num. projeto")), None)
    if col_rub is None or col_grp is None:
        # planilha sem cabeçalho (ex.: Estrutura FPA): descobre as colunas pelo conteúdo
        d = abrir(None).dropna(how="all")
        taxa = {c: d[c].dropna().map(lambda v: bool(_cod(v))).mean() if d[c].notna().any() else 0
                for c in d.columns}
        col_rub = max(taxa, key=taxa.get)
        if taxa[col_rub] < 0.5:
            raise ValueError(f"A planilha de mapeamento {nome} precisa ter as colunas: Rubrica e Grupo.")
        depois = [c for c in d.columns if c > col_rub and d[c].notna().any()]
        if not depois:
            raise ValueError(f"Não encontrei a coluna Grupo na planilha de mapeamento {nome}.")
        col_grp = depois[-1]
        antes = [c for c in d.columns if c < col_rub
                 and d[c].dropna().astype(str).str.match(r"^\d+\.\d+\.\d+$").mean() > 0.5]
        col_proj = antes[0] if antes else None
    projs = d[col_proj] if col_proj is not None else pd.Series([""] * len(d), index=d.index)
    mapa = {}
    for p, cod, grp in zip(projs, d[col_rub], d[col_grp]):
        cod = _cod(cod)
        if not cod or pd.isna(grp) or not str(grp).strip():
            continue
        mapa[(_proj(p) if pd.notna(p) else "", cod)] = GRUPO_ALIAS.get(_norm(grp), str(grp).strip())
    if not mapa:
        raise ValueError(f"Não encontrei nenhuma rubrica válida na planilha de mapeamento {nome}. "
                         "Confira as colunas Rubrica e Grupo.")
    return mapa


def ler_mapas(arquivos) -> dict:
    """arquivos = lista de (conteúdo, nome). Junta as planilhas de mapeamento."""
    mapa = {}
    for conteudo, nome in arquivos:
        mapa.update(ler_mapa_rubricas(conteudo, nome))
    if not mapa:
        raise ValueError("Envie a planilha de mapeamento de rubricas (colunas: Rubrica, Grupo).")
    return mapa


def ler_relatorio_financeiro(conteudo: bytes, nome_arquivo: str, mapa: dict):
    """Lê um relatório financeiro do OngSys (original ou editado) e devolve (despesas, info).

    Fica só com as despesas: tira receitas (rendimentos, entradas na conta de provisão,
    descontos) e ajustes bancários (TIPO 'D': IOF, IR sobre rendimento). Os valores de
    despesa vêm negativos no OngSys e são convertidos para positivos.
    """
    if not mapa:
        raise ValueError("Envie a planilha de mapeamento de rubricas (colunas: Rubrica, Grupo).")
    linhas, datemode = _ler_linhas(conteudo, nome_arquivo)
    df = _tabela(linhas, COLUNAS_FIN, f"relatório financeiro {nome_arquivo}", OPCIONAIS_FIN)
    for c in ("lanc", "tipo", "ident", "projeto", "cf", "curso", "rubrica"):
        df[c] = df[c].astype(str).str.strip()
    despesa = df["ident"].map(_norm) == "despesa"
    ajuste = df["tipo"].str.upper() == "D"
    info = {"lidas": len(df), "receitas": int((~despesa).sum()), "ajustes": int((despesa & ajuste).sum())}
    df = df[despesa & ~ajuste].copy()
    if df.empty:
        raise ValueError(f"Não encontrei nenhuma despesa no relatório financeiro {nome_arquivo}.")

    df["valor"] = -pd.to_numeric(df["valor"], errors="coerce").fillna(0.0)
    df["data"] = df["data"].map(lambda v: _to_dt(v, datemode))
    cod = df["cod"].map(_cod)
    # linhas de juros vêm do OngSys sem rubrica válida: herdam a do mesmo lançamento e CF
    sem_cod = cod.eq("")
    herdado = cod.where(~sem_cod).groupby([df["lanc"], df["cf"]]).transform("first")
    df["cod"] = cod.where(~sem_cod, herdado.fillna(""))
    info["herdadas"] = int((sem_cod & df["cod"].ne("")).sum())
    df["proj"] = df["num_projeto"].map(_proj)
    df.loc[df["proj"] == "", "proj"] = df.loc[df["proj"] == "", "projeto"].map(_proj)
    df["cf_num"] = df["cf"].str.split(" - ", n=1).str[0].str.strip()
    df["area"] = _aplicar_mapa(df["proj"], df["cod"], mapa)
    df["provisao"] = df["cod"].str.startswith("3.")
    return _atribuir_mes(df, nome_arquivo).reset_index(drop=True), info


def ler_relatorios(arquivos, mapa: dict):
    """arquivos = lista de (conteúdo, nome). Junta os relatórios financeiros."""
    if not arquivos:
        raise ValueError("Envie pelo menos um relatório financeiro do OngSys.")
    partes, info = [], {}
    for conteudo, nome in arquivos:
        df, i = ler_relatorio_financeiro(conteudo, nome, mapa)
        partes.append(df)
        for k, v in i.items():
            info[k] = info.get(k, 0) + v
    return _juntar(partes, "relatório financeiro"), info


def ler_movimentacao(conteudo: bytes, nome_arquivo: str, mapa: dict) -> pd.DataFrame:
    """Lê uma movimentação detalhada do OngSys (usada em FPA próprio, Quebrada Digital e Emenda)."""
    linhas, datemode = _ler_linhas(conteudo, nome_arquivo)
    df = _tabela(linhas, COLUNAS_MOV, f"arquivo de movimentação detalhada {nome_arquivo}")
    for c in ("lanc", "tipo", "ident", "projeto", "cf", "rub"):
        df[c] = df[c].astype(str).str.strip()
    for c in ("bruto", "liq"):
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    df["data"] = df["data"].map(lambda v: _to_dt(v, datemode))
    df["cod"] = df["rub"].map(_cod)
    df["rubrica"] = np.where(df["cod"].ne(""), df["rub"].str.slice(13).str.lstrip(" -").str.strip(),
                             df["rub"])
    df["proj"] = df["projeto"].map(_proj)
    df["cf_num"] = df["cf"].str.split(" - ", n=1).str[0].str.strip()
    df["area"] = _aplicar_mapa(df["proj"], df["cod"], mapa)
    return _atribuir_mes(df, nome_arquivo).reset_index(drop=True)


def ler_movimentacoes(arquivos, mapa: dict):
    if not arquivos:
        return None
    return _juntar([ler_movimentacao(c, n, mapa) for c, n in arquivos], "arquivo de movimentação")


def ler_orcamento(conteudo: bytes, nome: str) -> dict:
    """Lê uma planilha de orçamento (Orçamento POT ou Orçamento Emenda).

    Procura o cabeçalho com "Orçamento TOTAL"; as colunas "Orçamento de out/25" etc. viram meses.
    Cada linha com código de rubrica vira uma linha do orçamento (o "TOTAL DAS DESPESAS" é ignorado).
    """
    linhas, _ = _ler_linhas(conteudo, nome)
    ini = next((i for i, l in enumerate(linhas[:20])
                if any(_norm(c) == "orcamento total" for c in l if c is not None)), None)
    if ini is None:
        raise ValueError(f"Não encontrei a coluna 'Orçamento TOTAL' em {nome}. Confira se é a planilha de orçamento.")
    cab = linhas[ini]
    i_tot = next(j for j, c in enumerate(cab) if c is not None and _norm(c) == "orcamento total")
    meses = {}
    for j, c in enumerate(cab):
        m = re.search(r"([a-z]{3})\s*/\s*(\d{2,4})", _norm(c)) if j > i_tot and c is not None else None
        if m and m.group(1) in MES_NUM:
            ano = int(m.group(2))
            meses[j] = str(pd.Period(year=ano + 2000 if ano < 100 else ano, month=MES_NUM[m.group(1)], freq="M"))
    regs = []
    for l in linhas[ini + 1:]:
        j_cod = next((j for j in range(min(2, len(l))) if _cod(l[j])), None)
        if j_cod is None:
            continue
        desc = next((str(l[j]).strip() for j in range(j_cod + 1, i_tot)
                     if l[j] not in (None, "") and not isinstance(l[j], (int, float))), "")
        reg = {"cod": _cod(l[j_cod]), "desc": desc, "total": _num(l[i_tot]) if i_tot < len(l) else 0.0}
        reg.update({m: (_num(l[j]) if j < len(l) else 0.0) for j, m in meses.items()})
        regs.append(reg)
    if not regs:
        raise ValueError(f"Não encontrei rubricas na planilha de orçamento {nome}.")
    df = pd.DataFrame(regs)
    return {"linhas": df, "meses": sorted(meses.values()), "nome": nome}


def ler_historico(conteudo: bytes):
    try:
        h = pd.read_excel(io.BytesIO(conteudo), sheet_name="Historico", dtype={"mes": str})
        h.columns = [_norm(c) for c in h.columns]
        h["valor"] = pd.to_numeric(h["valor"], errors="coerce")
        h["mes"] = h["mes"].astype(str).str[:7]
        h["indicador"] = h["indicador"].replace(HIST_ALIAS)
        return h[["mes", "indicador", "valor"]].dropna()
    except Exception:
        raise ValueError("Não achei a aba 'Historico' nessa planilha. Use a planilha resumida gerada aqui.")


# ----------------------------------------------------------------------------
# Cálculo dos indicadores
# ----------------------------------------------------------------------------
def _soma(df, col):
    return float(df[col].sum())


def _resc(df):
    return df["rubrica"].str.upper().str.contains("RESCIS")


def _area_tot(df, col="valor"):
    s = df[df["area"].isin(AREAS)].groupby("area")[col].sum()
    return s.reindex(AREAS).fillna(0.0)


def _historico_dos_arquivos(pot, mov, emenda):
    """Valores de cada mês presente nos arquivos enviados (vão para a aba Historico)."""
    regs = []
    for m, g in pot.groupby("mes"):
        a = _area_tot(g)
        regs += [(m, f"POT II - área - {ar}", v) for ar, v in a.items()]
        regs.append((m, "POT II - despesa 6 áreas", float(a.sum())))
        sp = g[~g["provisao"]]
        regs.append((m, "POT II - rescisões", _soma(sp[_resc(sp)], "valor")))
        ins = sp[sp["cod"].isin(INSUMOS)].groupby("cod")["valor"].sum()
        regs += [(m, f"POT II - insumo - {n}", float(ins.get(c, 0.0))) for c, n in INSUMOS.items()]
    if mov is not None:
        up = mov["projeto"].str.upper()
        fpa = mov[up.str.contains("FPA PR") & (mov["ident"] != "TRANSF") & (mov["ident"] != "D")]
        for m, g in fpa.groupby("mes"):
            rec = g[(g["tipo"] == "Receita") & (g["cod"] != COD_RENDIMENTO)]
            regs.append((m, "FPA próprio - receita", _soma(rec[rec["bruto"] > 0], "bruto")))
            regs.append((m, "FPA próprio - despesa",
                         _soma(g[(g["tipo"] == "Despesa") & g["cod"].str.startswith("2")], "bruto")))
    if emenda is not None:
        for (m, c), v in emenda.groupby(["mes", "cod"])["valor"].sum().items():
            regs.append((m, f"Emenda - {c}", float(v)))
    return pd.DataFrame(regs, columns=["mes", "indicador", "valor"])


def _juntar_historico(anterior, novo, meses_pot, meses_mov, meses_emenda):
    """Os meses que vieram nos arquivos substituem o que havia na planilha anterior."""
    if anterior is None or not len(anterior):
        h = novo
    else:
        a = anterior.copy()
        fora = ((a["mes"].isin(meses_pot) & a["indicador"].str.startswith("POT II"))
                | (a["mes"].isin(meses_mov) & a["indicador"].str.startswith("FPA próprio"))
                | (a["mes"].isin(meses_emenda) & a["indicador"].str.startswith("Emenda")))
        h = pd.concat([a[~fora], novo])
    return (h.drop_duplicates(["mes", "indicador"], keep="last")
            .sort_values(["indicador", "mes"]).reset_index(drop=True))


def processar(fin, mapa, mov=None, orc_pot=None, orc_emenda=None, historico_anterior=None, info_fin=None,
              mapa_orc=None):
    """Calcula todos os indicadores.
    fin = despesas dos relatórios financeiros; mov = movimentação detalhada (opcional);
    orc_pot / orc_emenda = planilhas de orçamento (opcionais);
    mapa_orc = mapeamento usado no orçamento (ex.: Estrutura POT); se não vier, usa o mesmo do realizado."""
    info_fin = info_fin or {}
    mapa_orc = mapa_orc or mapa
    avisos = []
    up_fin = fin["projeto"].str.upper()
    pot = fin[up_fin.str.contains("POT")]
    if pot.empty:
        raise ValueError("O relatório financeiro não tem linhas do projeto POT.")
    mes_key = max(pot["mes"])
    mes = pd.Period(mes_key, "M")
    res = {"mes_key": mes_key, "mes_label": _rotulo(mes_key),
           "mes_extenso": f"{MESES_EXT[mes.month - 1]} de {mes.year}"}
    meses_pot = sorted(pot["mes"].unique())
    if len(meses_pot) > 1:
        avisos.append(f"Relatórios financeiros com {len(meses_pot)} meses ({', '.join(_rotulo(m) for m in meses_pot)}). "
                      f"Os slides do mês usam {res['mes_label']}; os outros meses entram nas séries e acumulados.")
    pot_mes = pot[pot["mes"] == mes_key]
    sem_prov = pot_mes[~pot_mes["provisao"]]
    proj_pot = pot_mes["proj"].mode().iloc[0] if pot_mes["proj"].ne("").any() else ""

    # Emenda: do relatório financeiro, se ele trouxer o projeto; se não, da movimentação
    emenda = None
    em_fin = fin[up_fin.str.contains("EMENDA") & fin["cod"].str.startswith("2")]
    if len(em_fin):
        emenda = em_fin[["mes", "cod", "rubrica", "valor"]].copy()
        res["emenda_fonte"] = "relatório financeiro"
    elif mov is not None:
        um = mov["projeto"].str.upper()
        em = mov[um.str.contains("EMENDA") & (mov["tipo"] == "Despesa") & (mov["ident"] != "D")
                 & mov["cod"].str.startswith("2")]
        if len(em):
            emenda = em[["mes", "cod", "rubrica"]].assign(valor=em["bruto"])
            res["emenda_fonte"] = "movimentação detalhada"

    # histórico: planilha anterior + todos os meses dos arquivos
    novo = _historico_dos_arquivos(pot, mov, emenda)
    meses_mov = sorted(mov["mes"].unique()) if mov is not None else []
    meses_em = sorted(emenda["mes"].unique()) if emenda is not None else []
    H = _juntar_historico(historico_anterior, novo, meses_pot, meses_mov, meses_em)
    res["historico"] = H

    def hv(ind, meses=None):
        s = H[H["indicador"] == ind]
        if meses is not None:
            s = s[s["mes"].isin(meses)]
        return float(s["valor"].sum())

    # 1) despesa por área no mês (as 6 áreas da planilha de mapeamento)
    tab_area = pd.DataFrame({"area": AREAS, "realizado": _area_tot(pot_mes).values})
    tab_area["projetado"] = np.nan
    orc_l = None
    if orc_pot is not None:
        orc_l = orc_pot["linhas"].copy()
        orc_l["area"] = _aplicar_mapa(pd.Series([proj_pot] * len(orc_l), index=orc_l.index), orc_l["cod"],
                                      mapa_orc)
        if mes_key in orc_l.columns:
            tab_area["projetado"] = tab_area["area"].map(orc_l.groupby("area")[mes_key].sum()).fillna(0.0)
        else:
            avisos.append(f"O orçamento {orc_pot['nome']} não tem a coluna de {res['mes_label']}: "
                          "sem projetado do mês.")
    tab_area["pct_realizado"] = np.where(tab_area["projetado"] > 0,
                                         tab_area["realizado"] / tab_area["projetado"], np.nan)
    res["area"] = tab_area
    res["realizado_total"] = float(tab_area["realizado"].sum())
    res["projetado_total"] = (float(tab_area["projetado"].sum())
                              if tab_area["projetado"].notna().any() else np.nan)

    # 2) custo por CF: sem provisões
    cf = sem_prov.assign(cf_nome=sem_prov["cf_num"].map(CF_NOMES).fillna(sem_prov["cf"]))
    tot_cf = cf.groupby("cf_nome")["valor"].sum()
    res["cf_geral"] = float(tot_cf.get("Geral", 0.0))
    res["cf_tabela"] = tot_cf

    # 3) gasto por curso no mês (rubricas de insumos)
    cur = sem_prov[sem_prov["cod"].isin(CURSOS)].groupby("cod")["valor"].sum()
    res["curso"] = pd.DataFrame({"curso": list(CURSOS.values()),
                                 "valor": [float(cur.get(c, 0.0)) for c in CURSOS]})

    # 4) rescisões e proporção de gastos com RH
    rescisoes = _soma(sem_prov[_resc(sem_prov)], "valor")
    resc_em_pessoal = _soma(pot_mes[(pot_mes["area"] == "Pessoal") & _resc(pot_mes)], "valor")
    res["rescisoes"] = rescisoes
    res["demais_rh"] = float(tab_area.loc[tab_area["area"] == "Pessoal", "realizado"].iloc[0]
                             - resc_em_pessoal)

    # 5) acumulado do aditivo (precisa do Orçamento POT)
    meses_com_dados = sorted(H.loc[H["indicador"] == "POT II - despesa 6 áreas", "mes"].unique())
    res["tem_orc"] = orc_l is not None
    res["orc_pendencias"] = pd.DataFrame(columns=["cod", "desc", "total"])
    if orc_l is not None:
        janela = [m for m in orc_pot["meses"] if m <= mes_key] or meses_com_dados
        com = [m for m in janela if m in meses_com_dados]
        falta = [m for m in janela if m not in meses_com_dados]
        res["meses_aditivo"] = janela
        res["meses_sem_dados"] = falta
        res["n_mes_aditivo"] = janela.index(mes_key) + 1 if mes_key in janela else None
        if falta:
            avisos.append(f"Faltam dados de {len(falta)} mês(es) do aditivo para o acumulado: "
                          f"{', '.join(_rotulo(m) for m in falta)}. Envie esses relatórios financeiros "
                          "ou a planilha resumida do mês anterior.")
        proj_tot = orc_l.groupby("area")["total"].sum()
        ac = pd.DataFrame({"area": AREAS,
                           "projetado": [float(proj_tot.get(a, 0.0)) for a in AREAS],
                           "realizado": [hv(f"POT II - área - {a}", com) for a in AREAS]})
        ac["pct_realizado"] = np.where(ac["projetado"] > 0, ac["realizado"] / ac["projetado"], np.nan)
        res["acumulado"] = ac
        total_proj = float(ac["projetado"].sum())
        mensal = pd.DataFrame({"mes": com, "despesa": [hv("POT II - despesa 6 áreas", [m]) for m in com]})
        mensal["saldo"] = total_proj - mensal["despesa"].cumsum()
        res["mensal"] = mensal
        ins = pd.DataFrame({"insumo": list(INSUMOS.values()),
                            "projetado": [float(orc_l.loc[orc_l["cod"] == c, "total"].sum()) for c in INSUMOS],
                            "realizado": [hv(f"POT II - insumo - {n}", com) for n in INSUMOS.values()]})
        ins["pct_realizado"] = np.where(ins["projetado"] > 0, ins["realizado"] / ins["projetado"], np.nan)
        res["insumos"] = ins
        res["rescisoes_acum"] = hv("POT II - rescisões", janela)
        res["periodo_acum"] = f"{_rotulo(janela[0])} a {_rotulo(janela[-1])}" if janela else ""
        pend = orc_l[(orc_l["area"] == FORA) & (orc_l["total"] != 0)]
        res["orc_pendencias"] = pend[["cod", "desc", "total"]].reset_index(drop=True)
        res["orc_fora"] = float(orc_l.loc[~orc_l["area"].isin(AREAS) & (orc_l["area"] != FORA), "total"].sum())
        res["orc_total_arquivo"] = float(orc_l["total"].sum())
    else:
        meses_r = sorted(H.loc[H["indicador"] == "POT II - rescisões", "mes"].unique())
        res["rescisoes_acum"] = hv("POT II - rescisões")
        res["periodo_acum"] = f"{_rotulo(meses_r[0])} a {_rotulo(meses_r[-1])}" if meses_r else ""

    # 6 e 7) FPA próprio e Quebrada Digital (movimentação do mês de referência)
    res["fpa_fontes"] = pd.Series(dtype=float)
    res["fpa_receita"] = res["fpa_despesa"] = res["fpa_rendimentos"] = res["qd_receita"] = 0.0
    res["qd"] = pd.Series(dtype=float)
    mov_mes = mov[mov["mes"] == mes_key] if mov is not None else None
    res["tem_mov"] = mov_mes is not None and len(mov_mes) > 0
    if mov is not None and not res["tem_mov"]:
        avisos.append(f"A movimentação detalhada não tem {res['mes_label']}: slides de FPA próprio e "
                      "Quebrada Digital não foram gerados.")
    if res["tem_mov"]:
        up = mov_mes["projeto"].str.upper()
        fpa = mov_mes[up.str.contains("FPA PR") & (mov_mes["ident"] != "TRANSF") & (mov_mes["ident"] != "D")]
        rec = fpa[(fpa["tipo"] == "Receita") & (fpa["cod"] != COD_RENDIMENTO)]
        fontes = rec.groupby(rec["cod"].map(FONTES_FPA).fillna("Outras receitas"))["bruto"].sum()
        fontes = fontes[fontes > 0].sort_values(ascending=False)
        res["fpa_fontes"] = fontes
        res["fpa_receita"] = float(fontes.sum())
        res["fpa_rendimentos"] = _soma(fpa[(fpa["tipo"] == "Receita") & (fpa["cod"] == COD_RENDIMENTO)], "bruto")
        res["fpa_despesa"] = _soma(fpa[(fpa["tipo"] == "Despesa") & fpa["cod"].str.startswith("2")], "bruto")
        qd_base = mov_mes[(mov_mes["cf_num"] == "10") & up.str.contains("FPA PR")
                          & (mov_mes["ident"] != "D") & (mov_mes["ident"] != "TRANSF")]
        qd = qd_base[(qd_base["tipo"] == "Despesa") & qd_base["cod"].str.startswith("2")]
        cats = ["RH", "Administrativo", "Serviços", "Materiais", "Outros"]
        qd_tot = (qd.groupby(qd["area"].map(QD_CATEGORIAS).fillna("Outros"))["bruto"].sum()
                  .reindex(cats).fillna(0.0))
        res["qd"] = qd_tot[(qd_tot > 0) | qd_tot.index.isin(cats[:4])]
        res["qd_receita"] = _soma(qd_base[(qd_base["tipo"] == "Receita")
                                          & (qd_base["cod"] != COD_RENDIMENTO)], "bruto")

    # 8) Emenda parlamentar
    res["tem_emenda"] = emenda is not None or orc_emenda is not None
    res["emenda"] = pd.Series(dtype=float)
    res["emenda_bolsas"] = 0
    res["emenda_tab"] = None
    nomes = {}
    if emenda is not None:
        em_mes = emenda[emenda["mes"] == mes_key]
        res["emenda"] = em_mes.groupby("rubrica")["valor"].sum().sort_values(ascending=False)
        res["emenda_bolsas"] = int(em_mes["rubrica"].str.upper().str.contains("BOLSA").sum())
        nomes = emenda.groupby("cod")["rubrica"].first().to_dict()
    if orc_emenda is not None:
        eo = orc_emenda["linhas"]
        prev = eo.groupby("cod")["total"].sum()
        real = {c[len("Emenda - "):]: v for c, v in
                H[H["indicador"].str.startswith("Emenda - ")].groupby("indicador")["valor"].sum().items()}
        cods = [c for c in prev.index if prev[c] > 0] + [c for c, v in real.items() if v > 0 and prev.get(c, 0) <= 0]
        desc = eo.groupby("cod")["desc"].first().to_dict()
        tab = pd.DataFrame({"cod": cods,
                            "rubrica": [desc.get(c) or nomes.get(c, c) for c in cods],
                            "previsto": [float(prev.get(c, 0.0)) for c in cods],
                            "realizado": [float(real.get(c, 0.0)) for c in cods]})
        tab["pct_realizado"] = np.where(tab["previsto"] > 0, tab["realizado"] / tab["previsto"], np.nan)
        res["emenda_tab"] = tab
        res["emenda_aporte"] = float(prev.sum())
        res["emenda_realizado"] = float(sum(real.values()))
        meses_e = sorted(H.loc[H["indicador"].str.startswith("Emenda - "), "mes"].unique())
        res["emenda_meses"] = meses_e
        if emenda is None and not meses_e:
            avisos.append("Orçamento da Emenda enviado, mas não há despesas da Emenda nos arquivos "
                          "(envie a movimentação detalhada ou o relatório financeiro do projeto).")

    # 9) verificações (para conferir antes de usar)
    pagto_prov = _soma(pot_mes[pot_mes["area"] == "Pagto provisão"], "valor")
    nao_map = pot_mes[pot_mes["area"] == FORA]
    fora_areas = pot_mes[~pot_mes["area"].isin(AREAS) & (pot_mes["area"] != FORA)]
    res["nao_mapeadas"] = (nao_map.groupby(["cod", "rubrica"])["valor"].sum().reset_index()
                           if len(nao_map) else pd.DataFrame(columns=["cod", "rubrica", "valor"]))
    verif = [
        ("Linhas lidas nos relatórios financeiros", info_fin.get("lidas", len(fin)), ""),
        ("Linhas de receita retiradas", info_fin.get("receitas", 0),
         "Rendimentos, entradas na conta de provisão e descontos não são despesa."),
        ("Ajustes bancários retirados (linhas)", info_fin.get("ajustes", 0), "IOF e IR sobre rendimento."),
        ("Linhas de juros sem rubrica", info_fin.get("herdadas", 0), "Receberam a rubrica do próprio lançamento."),
        (f"Linhas de despesa do POT II em {res['mes_label']}", len(pot_mes), ""),
        (f"Despesa POT II em {res['mes_label']} (tudo)", _soma(pot_mes, "valor"), ""),
        ("Provisões (rubricas 3.x)", _soma(pot_mes[pot_mes["provisao"]], "valor"),
         "Entram nas 6 áreas só se a planilha de mapeamento der um grupo a elas. Nunca entram no custo por CF."),
        ("Pagamento de provisão (grupo 'Pagto provisão')", pagto_prov, "Fora das 6 áreas; entra no custo por CF."),
        ("Despesas em grupos fora das 6 áreas", _soma(fora_areas, "valor"),
         "Grupos como '-' ou 'Pagto provisão' na planilha de mapeamento."),
        ("Despesa realizada nas 6 áreas", res["realizado_total"], ""),
        ("Despesas com rubrica sem grupo no mapa", float(nao_map["valor"].sum()),
         "Ver aba Pendencias; entram só no custo por CF."),
        ("Rescisões do mês", rescisoes, "Rubricas com 'rescisão' no nome."),
        ("CF '1 - GERAL' (rateio pendente)", res["cf_geral"], ""),
    ]
    if orc_l is not None:
        verif += [
            ("Projetado do mês (6 áreas)", res["projetado_total"] if pd.notna(res["projetado_total"]) else 0.0,
             f"Coluna {res['mes_label']} do orçamento."),
            ("Orçamento total do arquivo", res["orc_total_arquivo"], orc_pot["nome"]),
            ("Orçamento em grupos fora das 6 áreas", res["orc_fora"], "Grupos como '-' ou 'Pagto provisão'."),
            ("Orçamento com rubrica sem grupo no mapa", float(res["orc_pendencias"]["total"].sum()),
             "Ver aba Pendencias_orcamento; fica fora do projetado."),
            ("Meses do aditivo com dados", len(res["meses_aditivo"]) - len(res["meses_sem_dados"]),
             f"de {len(res['meses_aditivo'])} meses"),
        ]
    if res["tem_mov"]:
        verif.append(("Rendimentos FPA próprio (fora da receita)", res["fpa_rendimentos"], ""))
    res["verificacoes"] = pd.DataFrame(verif, columns=["item", "valor", "observação"])
    res["avisos"] = avisos
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
    if pd.notna(res["projetado_total"]):
        resumo.append(("POT II", "Despesa projetada (6 áreas)", res["projetado_total"]))
    resumo += [("POT II", f"Despesa realizada: {r.area}", r.realizado) for r in res["area"].itertuples()]
    resumo += [("POT II", f"Custo por CF: {r.cf}", r.valor) for r in cf.itertuples()]
    resumo += [("POT II", f"Gasto por curso: {r.curso}", r.valor) for r in res["curso"].itertuples()]
    resumo += [("POT II", "Rescisões do mês", res["rescisoes"]),
               ("POT II", "Demais gastos com RH", res["demais_rh"]),
               ("POT II", f"Rescisões acumuladas ({res['periodo_acum']})", res["rescisoes_acum"])]
    if res["tem_orc"]:
        resumo += [("POT II", f"Acumulado realizado: {r.area}", r.realizado) for r in res["acumulado"].itertuples()]
    if res["tem_mov"]:
        resumo += [("FPA próprio", "Receita do mês", res["fpa_receita"]),
                   ("FPA próprio", "Despesa do mês", res["fpa_despesa"]),
                   ("Quebrada Digital", "Receita destinada ao CF", res["qd_receita"])]
        resumo += [("FPA próprio", f"Receita: {k}", v) for k, v in res["fpa_fontes"].items()]
        resumo += [("Quebrada Digital", f"Gasto: {k}", v) for k, v in res["qd"].items()]
    if res["tem_emenda"]:
        resumo.append(("Emenda parlamentar", "Bolsas pagas no mês (lançamentos)", res["emenda_bolsas"]))
        resumo += [("Emenda parlamentar", f"Despesa do mês: {k}", v) for k, v in res["emenda"].items()]
        if res["emenda_tab"] is not None:
            resumo += [("Emenda parlamentar", "Aporte do projeto", res["emenda_aporte"]),
                       ("Emenda parlamentar", "Realizado acumulado", res["emenda_realizado"])]
    abas = {"Resumo": pd.DataFrame(resumo, columns=["indicador", "item", res["mes_label"]]),
            "POT_por_area": res["area"]}
    if res["tem_orc"]:
        abas["POT_acumulado"] = res["acumulado"]
        abas["POT_mensal"] = res["mensal"].assign(mes=res["mensal"]["mes"].map(_rotulo))
        abas["Insumos_acumulado"] = res["insumos"]
    abas["POT_por_CF"] = cf
    abas["POT_por_curso"] = res["curso"]
    if res["tem_mov"]:
        abas["FPA_receitas"] = res["fpa_fontes"].rename_axis("fonte").rename("valor").reset_index()
        abas["Quebrada_Digital"] = res["qd"].rename_axis("categoria").rename("valor").reset_index()
    if res["tem_emenda"]:
        abas["Emenda_mes"] = res["emenda"].rename_axis("rubrica").rename("valor").reset_index()
        if res["emenda_tab"] is not None:
            abas["Emenda_acumulado"] = res["emenda_tab"]
    abas["Historico"] = res["historico"]
    abas["Verificacoes"] = res["verificacoes"]
    abas["Pendencias"] = res["nao_mapeadas"]
    if res["tem_orc"]:
        abas["Pendencias_orcamento"] = res["orc_pendencias"]

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


def _com_pct(nomes, pcts):
    return [f"{n} ({p:.0%})" if pd.notna(p) else n for n, p in zip(nomes, pcts)]


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

    def capa(titulo, subtitulo=None, tam=60):
        s = prs.slides.add_slide(vazio)
        fundo = s.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
        fundo.fill.solid()
        fundo.fill.fore_color.rgb = rgb(AZUL)
        fundo.line.fill.background()
        texto(s, 0.5, 2.4, 12.3, 1.6, titulo, tam, "FFFFFF", True, PP_ALIGN.CENTER)
        if subtitulo:
            texto(s, 0.5, 4.3, 12.3, 0.8, subtitulo, 28, "FFFFFF", True, PP_ALIGN.CENTER)

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
        meses = sorted(h.loc[h["indicador"].isin(indicadores), "mes"].unique())
        series = {}
        for ind in indicadores:
            sub = h[h["indicador"] == ind].set_index("mes")["valor"]
            series[ind] = [float(sub.get(m, 0.0)) for m in meses]
        return [_rotulo(m) for m in meses], series

    # capa
    capa("Financeiro", res["mes_extenso"], 80)

    # receitas FPA próprio (movimentação)
    if res["tem_mov"]:
        s = novo("FINANCEIRO - TOTALIDADE DAS RECEITAS", f"FPA próprio, {res['mes_extenso']}")
        fontes = res["fpa_fontes"]
        if len(fontes):
            pizza(s, 1.5, 1.6, 10.3, 5.0, fontes.index, fontes.values,
                  [AZUL, "F4B76A", "FFDD57", "E879C4", "EE7480", ROXO])
        texto(s, 0.5, 6.75, 12.3, 0.6,
              f"Receita total: {brl(res['fpa_receita'])}. A divisão entre doação recorrente e pontual "
              "não consta na movimentação detalhada: complementar manualmente.", 12)

    capa("Área Financeira", "Projeto POT II")

    # POT: despesa do mês, projetada x realizada e por área
    s = novo("POT Redenção", res["mes_extenso"])
    a = res["area"]
    tem_proj = a["projetado"].notna().any()
    if tem_proj:
        pct = res["realizado_total"] / res["projetado_total"] if res["projetado_total"] else 0
        texto(s, 0.5, 1.55, 5.5, 0.5, f"{pct:.0%} da despesa projetada", 20, AZUL, True, PP_ALIGN.CENTER)
        barras(s, 0.4, 2.1, 5.6, 4.4, ["Despesa projetada", "Despesa realizada"],
               {"Valor": [res["projetado_total"], res["realizado_total"]]}, [AZUL],
               XL_CHART_TYPE.COLUMN_CLUSTERED)
        barras(s, 6.3, 1.6, 6.6, 4.9, _com_pct(a["area"], a["pct_realizado"]),
               {"Projetado": a["projetado"].fillna(0), "Realizado": a["realizado"]}, [AZUL, "F4B76A"],
               legenda=True, titulo="Despesa por rubrica")
    else:
        barras(s, 0.4, 1.6, 5.6, 4.9, ["Despesa realizada"], {"Valor": [res["realizado_total"]]},
               [AZUL], XL_CHART_TYPE.COLUMN_CLUSTERED)
        barras(s, 6.3, 1.6, 6.6, 4.9, a["area"], {"Realizado": a["realizado"]}, [AZUL],
               titulo="Despesa por rubrica")
    texto(s, 0.5, 6.7, 12.3, 0.6,
          f"Despesa realizada: {brl(res['realizado_total'])}."
          + (f" Despesa projetada: {brl(res['projetado_total'])}." if tem_proj else " Projetado não informado."), 12)

    if res["tem_orc"]:
        n = res["n_mes_aditivo"]
        sub = f"{n}º mês do aditivo" if n else res["mes_extenso"]
        falta = res["meses_sem_dados"]
        nota_falta = (f" Atenção: sem dados de {', '.join(_rotulo(m) for m in falta)}." if falta else "")

        # projetado x realizado acumulado
        s = novo("POT Redenção", f"Projetado x Realizado acumulado - {sub}")
        ac = res["acumulado"]
        barras(s, 0.4, 1.6, 12.5, 5.0, _com_pct(ac["area"], ac["pct_realizado"]),
               {"Projetado": ac["projetado"], "Realizado": ac["realizado"]}, [AZUL, "F4B76A"], legenda=True)
        texto(s, 0.5, 6.7, 12.3, 0.6,
              f"Projetado: orçamento total do aditivo ({brl(float(ac['projetado'].sum()))}). "
              f"Realizado de {res['periodo_acum']}: {brl(float(ac['realizado'].sum()))}.{nota_falta}", 12)

        # acompanhamento mensal
        s = novo("POT Redenção", f"Acompanhamento mensal - aditivo ({res['periodo_acum']})")
        me = res["mensal"]
        if len(me):
            barras(s, 0.4, 1.6, 12.5, 5.0, [_rotulo(m) for m in me["mes"]],
                   {"Despesa do mês": me["despesa"], "Saldo no mês": me["saldo"]}, [AZUL, "F4B76A"],
                   XL_CHART_TYPE.COLUMN_CLUSTERED, fmt='#,##0.00,,"mi"', legenda=True)
        texto(s, 0.5, 6.7, 12.3, 0.6,
              "Saldo = orçamento total do aditivo menos a despesa acumulada até o mês." + nota_falta, 12)

        # insumos projetado x realizado
        s = novo("POT Redenção", f"Comparativo dos insumos - Projetado x Realizado - {sub}")
        ins = res["insumos"]
        barras(s, 0.4, 1.6, 12.5, 5.0, _com_pct(ins["insumo"], ins["pct_realizado"]),
               {"Projetado": ins["projetado"], "Realizado": ins["realizado"]}, [AZUL, "F4B76A"], legenda=True)
        texto(s, 0.5, 6.7, 12.3, 0.6, "OBS: (editar)" + nota_falta, 12)

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
    rot, ser = serie_mensal(["POT II - rescisões"])
    barras(s, 0.4, 1.6, 7.2, 4.9, rot, ser, ["40B6D6"], XL_CHART_TYPE.COLUMN_CLUSTERED,
           titulo="Rescisões, série histórica")
    barras(s, 7.8, 1.6, 5.1, 4.9, ["Demais gastos com RH", "Rescisões"],
           {"Valor": [res["demais_rh"], res["rescisoes"]]}, [ROXO],
           titulo=f"Proporção de gastos, {res['mes_label']}")
    texto(s, 0.5, 6.7, 12.3, 0.6,
          f"Rescisões do mês: {brl(res['rescisoes'])}. Gastos com rescisões ({res['periodo_acum']}): "
          f"{brl(res['rescisoes_acum'])}.", 12)

    if res["tem_mov"]:
        capa("Área Financeira", "Receitas e despesas próprias")
        s = novo("FPA PRÓPRIO", res["mes_extenso"])
        rot, ser = serie_mensal(["FPA próprio - receita", "FPA próprio - despesa"])
        barras(s, 0.4, 1.6, 8.6, 4.9, rot, {"Receita": ser["FPA próprio - receita"],
                                            "Despesas": ser["FPA próprio - despesa"]},
               [VERDE, "D9A400"], XL_CHART_TYPE.COLUMN_CLUSTERED, legenda=True,
               titulo="Série histórica de receitas")
        texto(s, 9.4, 2.2, 3.6, 0.5, "Receita no mês", 18, AZUL, True, PP_ALIGN.CENTER)
        texto(s, 9.4, 2.7, 3.6, 0.6, brl(res["fpa_receita"]), 22, AZUL, True, PP_ALIGN.CENTER)
        texto(s, 9.4, 3.8, 3.6, 0.5, "Despesa no mês", 18, LARANJA, True, PP_ALIGN.CENTER)
        texto(s, 9.4, 4.3, 3.6, 0.6, brl(res["fpa_despesa"]), 22, LARANJA, True, PP_ALIGN.CENTER)
        texto(s, 0.5, 6.7, 12.3, 0.6,
              "Receitas: aluguéis e doações. Despesas: gastos administrativos da fundação que não são "
              "financiados por projetos.", 12)

        capa("Área Financeira", "CF Geraldo Spínola - Quebrada Digital")
        s = novo("FINANCEIRO - QUEBRADA DIGITAL - CF Geraldo Spínola", res["mes_extenso"])
        qd = res["qd"]
        barras(s, 0.4, 1.6, 8.6, 4.9, qd.index, {"Gastos": qd.values}, [AZUL],
               titulo=f"Gastos {res['mes_label']}")
        texto(s, 9.4, 2.4, 3.6, 0.5, "Receita destinada ao CF", 18, AZUL, True, PP_ALIGN.CENTER)
        texto(s, 9.4, 2.9, 3.6, 0.6, brl(res["qd_receita"]), 22, AZUL, True, PP_ALIGN.CENTER)
        texto(s, 9.4, 3.9, 3.6, 0.5, "Despesa total", 18, LARANJA, True, PP_ALIGN.CENTER)
        texto(s, 9.4, 4.4, 3.6, 0.6, brl(float(qd.sum())), 22, LARANJA, True, PP_ALIGN.CENTER)
        texto(s, 0.5, 6.7, 12.3, 0.6, "Obs.: (editar) explicar gastos fora do padrão do mês.", 12)

    if res["tem_emenda"]:
        capa("Área Financeira", "Projeto Emenda Parlamentar")
        s = novo("FINANCEIRO - EMENDA PARLAMENTAR", res["mes_extenso"])
        tab = res["emenda_tab"]
        if tab is not None and len(tab):
            barras(s, 0.4, 1.6, 8.6, 4.9, _com_pct(tab["rubrica"], tab["pct_realizado"]),
                   {"Previsto": tab["previsto"], "Realizado": tab["realizado"]}, [AZUL, ROXO], legenda=True,
                   titulo="Emenda parlamentar - Previsto x Realizado")
            texto(s, 9.4, 1.9, 3.6, 0.5, "Aporte do projeto", 18, AZUL, True, PP_ALIGN.CENTER)
            texto(s, 9.4, 2.4, 3.6, 0.6, brl(res["emenda_aporte"]), 22, AZUL, True, PP_ALIGN.CENTER)
            texto(s, 9.4, 3.3, 3.6, 0.5, "Realizado", 18, ROXO, True, PP_ALIGN.CENTER)
            texto(s, 9.4, 3.8, 3.6, 0.6, brl(res["emenda_realizado"]), 22, ROXO, True, PP_ALIGN.CENTER)
            me = res["emenda_meses"]
            periodo = (f"Realizado acumulado de {len(me)} mês(es) com dados ({_rotulo(me[0])} a {_rotulo(me[-1])})."
                       if me else "Sem despesas da Emenda nos arquivos.")
        else:
            em = res["emenda"]
            if len(em):
                barras(s, 0.4, 1.6, 8.6, 4.9, em.index, {"Realizado": em.values}, [ROXO],
                       titulo="Emenda parlamentar, despesas do mês")
            periodo = "Previsto não informado: envie o Orçamento da Emenda."
        texto(s, 9.4, 4.7, 3.6, 0.5, "Bolsas pagas no mês", 18, AZUL, True, PP_ALIGN.CENTER)
        texto(s, 9.4, 5.2, 3.6, 0.6, str(res["emenda_bolsas"]), 28, AZUL, True, PP_ALIGN.CENTER)
        texto(s, 0.5, 6.7, 12.3, 0.6, periodo, 12)

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
             "**planilha de mapeamento de rubricas**. Os outros arquivos são opcionais e liberam mais slides. "
             "Os dados são processados na memória e não ficam salvos.")

    fin_arqs = st.file_uploader("Relatório financeiro do OngSys (pode enviar vários meses)",
                                type=["xls", "xlsx"], accept_multiple_files=True)
    mapa_arqs = st.file_uploader("Mapeamento de rubricas (uma ou mais Estruturas: Projeto, Rubrica, Grupo)",
                                 type=["csv", "xlsx", "xls"], accept_multiple_files=True)
    orc_arq = st.file_uploader("Orçamento POT (opcional: projetado x realizado, acumulado e insumos)",
                               type=["xls", "xlsx"])
    mapa_orc_arqs = st.file_uploader("Estrutura do orçamento POT (opcional: mapeamento usado no orçamento; "
                                     "se não enviar, usa o mapeamento de rubricas)",
                                     type=["csv", "xlsx", "xls"], accept_multiple_files=True)
    mov_arqs = st.file_uploader("Movimentação detalhada (opcional: FPA próprio, Quebrada Digital e Emenda)",
                                type=["xls", "xlsx"], accept_multiple_files=True)
    orc_em_arq = st.file_uploader("Orçamento Emenda parlamentar (opcional: previsto x realizado)",
                                  type=["xls", "xlsx"])
    with st.expander("Opções"):
        geral = st.checkbox("Somar o CF '1 - GERAL' em Glicério (como no slide de agosto)", value=True)
        hist_arq = st.file_uploader("Planilha resumida de um mês anterior (opcional: séries e acumulados)",
                                    type=["xlsx"])
    if not fin_arqs or not mapa_arqs:
        st.info("Envie o relatório financeiro e a planilha de mapeamento para continuar.")
        st.stop()

    par = lambda arqs: [(a.getvalue(), a.name) for a in arqs]
    try:
        mapa = ler_mapas(par(mapa_arqs))
        fin, info = ler_relatorios(par(fin_arqs), mapa)
        mov = ler_movimentacoes(par(mov_arqs or []), mapa)
        orc = ler_orcamento(orc_arq.getvalue(), orc_arq.name) if orc_arq else None
        orc_em = ler_orcamento(orc_em_arq.getvalue(), orc_em_arq.name) if orc_em_arq else None
        hist = ler_historico(hist_arq.getvalue()) if hist_arq else None
        mapa_orc = ler_mapas(par(mapa_orc_arqs)) if mapa_orc_arqs else None
        res = processar(fin, mapa, mov, orc, orc_em, hist, info, mapa_orc)
        xlsx = gerar_xlsx(res, geral)
        pptx = gerar_pptx(res, geral)
    except Exception as e:
        st.error(f"Não consegui processar os arquivos: {e}")
        st.stop()

    st.success(f"Mês de referência: {res['mes_extenso']}")
    st.caption(f"Mapeamento de rubricas aplicado: {', '.join(a.name for a in mapa_arqs)} ({len(mapa)} rubricas).")
    for aviso in res["avisos"]:
        st.warning(aviso)
    if not res["tem_mov"]:
        st.info("Sem a movimentação detalhada do mês, os slides de FPA próprio e Quebrada Digital não foram gerados.")
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
    if res["tem_orc"] and len(res["orc_pendencias"]):
        st.warning("Há rubricas do orçamento sem grupo na planilha de mapeamento. Elas ficam fora do projetado.")
        st.dataframe(res["orc_pendencias"], hide_index=True)

    st.subheader("Pré-visualização")
    a, b, c = st.columns(3)
    a.caption("Despesa por área")
    a.bar_chart(res["area"].set_index("area")[["realizado"] + (["projetado"] if res["tem_orc"] else [])])
    b.caption("Custo por CF")
    b.bar_chart(tabela_cf(res, geral).set_index("cf")["valor"])
    c.caption("Gasto por curso")
    c.bar_chart(res["curso"].set_index("curso")["valor"])


def cli(args):
    flags_validas = ("--mapa", "--mapa-orc", "--mov", "--orcamento", "--orc-emenda", "--hist")
    pos, flags = [], {}
    i = 0
    while i < len(args):
        a = args[i]
        if a.startswith("--"):
            if a not in flags_validas:
                sys.exit(f"Erro: opção desconhecida {a}. Opções: {', '.join(flags_validas)}.")
            if i + 1 >= len(args):
                sys.exit(f"Erro: faltou o arquivo depois de {a}.")
            flags.setdefault(a, []).append(args[i + 1])
            i += 2
        else:
            pos.append(a)
            i += 1

    def abrir(p):
        with open(p, "rb") as f:
            return f.read(), p

    if not flags.get("--mapa"):
        sys.exit("Erro: informe a planilha de mapeamento de rubricas com --mapa arquivo.xlsx.")
    mapa = ler_mapas([abrir(p) for p in flags["--mapa"]])
    fin, info = ler_relatorios([abrir(p) for p in pos], mapa)
    mov = ler_movimentacoes([abrir(p) for p in flags.get("--mov", [])], mapa)
    orc = ler_orcamento(*abrir(flags["--orcamento"][-1])) if "--orcamento" in flags else None
    orc_em = ler_orcamento(*abrir(flags["--orc-emenda"][-1])) if "--orc-emenda" in flags else None
    hist = ler_historico(abrir(flags["--hist"][-1])[0]) if "--hist" in flags else None
    mapa_orc = ler_mapas([abrir(p) for p in flags["--mapa-orc"]]) if "--mapa-orc" in flags else None
    res = processar(fin, mapa, mov, orc, orc_em, hist, info, mapa_orc)
    v = res["verificacoes"].copy()
    v["valor"] = v["valor"].map(lambda x: f"{x:,.2f}")
    print(v.to_string(index=False))
    for aviso in res["avisos"]:
        print("ATENÇÃO:", aviso)
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
