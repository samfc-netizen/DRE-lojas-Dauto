from __future__ import annotations


import os
import glob
import re
import unicodedata
from typing import Dict, List, Optional, Tuple

import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px

# =========================
# Config
# =========================
st.set_page_config(page_title="DRE + DFC — Lojas", layout="wide")

EXCEL_FIXED_NAME = "DRE apenas lojas.xlsx"

MESES_PT = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"]
MES_NUM_TO_PT = {1: "JAN", 2: "FEV", 3: "MAR", 4: "ABR", 5: "MAI", 6: "JUN",
                 7: "JUL", 8: "AGO", 9: "SET", 10: "OUT", 11: "NOV", 12: "DEZ"}
MES_PT_TO_NUM = {v: k for k, v in MES_NUM_TO_PT.items()}

# DRE: M+1
SHIFT_NEXT_GROUPS = {"DEDUÇÕES (IMPOSTOS SOBRE VENDAS)", "DESPESAS COM PESSOAL"}

# =========================
# Helpers
# =========================
def _norm_txt(s: object) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    t = str(s)
    t = t.replace("\u00a0", " ").replace("–", "-").replace("—", "-")
    t = unicodedata.normalize("NFKD", t)
    t = "".join(ch for ch in t if not unicodedata.combining(ch))
    t = re.sub(r"\s+", " ", t).strip().lower()
    return t

def _auto_find_excel() -> Optional[str]:
    # Prioriza o nome fixo
    if os.path.exists(EXCEL_FIXED_NAME):
        return EXCEL_FIXED_NAME
    # fallback: pega o mais recente
    files: List[str] = []
    for pat in ["*.xlsx", "*.xlsm", "*.xls"]:
        files.extend(glob.glob(pat))
    files = [f for f in files if os.path.isfile(f)]
    if not files:
        return None
    files.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    return files[0]

def excel_signature(path: str) -> Tuple[int, int]:
    stt = os.stat(path)
    return (stt.st_mtime_ns, stt.st_size)

def to_num(v) -> float:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 0.0
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    s = str(v).strip()
    if s == "":
        return 0.0
    s = s.replace("\u00a0", " ").replace("R$", "").strip()
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0

def format_brl(x) -> str:
    try:
        return f"{float(x):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "0,00"

def fmt_brl(x) -> str:
    return f"R$ {format_brl(x)}"

def fmt_pct(x) -> str:
    try:
        return f"{float(x):,.2f}%".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "0,00%"

def parse_mes(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    s = str(v).strip().upper()
    if s.isdigit():
        m = int(s)
        return m if 1 <= m <= 12 else None
    mapa = {
        "JANEIRO": 1, "JAN": 1,
        "FEVEREIRO": 2, "FEV": 2,
        "MARCO": 3, "MARÇO": 3, "MAR": 3,
        "ABRIL": 4, "ABR": 4,
        "MAIO": 5, "MAI": 5,
        "JUNHO": 6, "JUN": 6,
        "JULHO": 7, "JUL": 7,
        "AGOSTO": 8, "AGO": 8,
        "SETEMBRO": 9, "SET": 9,
        "OUTUBRO": 10, "OUT": 10,
        "NOVEMBRO": 11, "NOV": 11,
        "DEZEMBRO": 12, "DEZ": 12,
    }
    return mapa.get(s)

def safe_div(a: float, b: float) -> float:
    return float(a) / float(b) if float(b) != 0 else 0.0

def sintetizar_despesa(nome: str) -> str:
    if nome is None or (isinstance(nome, float) and pd.isna(nome)):
        return "—"
    s = str(nome).strip()
    s = re.sub(r"\s*\(\s*\d+\s*-\s*DESPESAS\s*\)\s*$", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\s*\([^)]*\)\s*$", "", s).strip()
    s = re.sub(r"\s{2,}", " ", s)
    return s if s else "—"

def lojas_effective(lojas_sel: List[str]) -> List[str]:
    """Remove TODAS se houver seleção específica; retorna lista de lojas efetivas."""
    if not lojas_sel:
        return []
    lojas_sel = [str(x).strip() for x in lojas_sel]
    if "TODAS" in lojas_sel and len(lojas_sel) > 1:
        lojas_sel = [x for x in lojas_sel if x != "TODAS"]
    if "TODAS" in lojas_sel:
        return []
    return lojas_sel

# =========================
# Excel IO (cached)
# =========================
@st.cache_data(show_spinner=False)
def get_sheet_names(excel_path: str, sig):
    try:
        return pd.ExcelFile(excel_path).sheet_names
    except Exception:
        return []

@st.cache_data(show_spinner=False)
def read_sheet(excel_path: str, sheet_name: str, sig):
    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name)
    except Exception:
        return None
    df.columns = [str(c).strip() for c in df.columns]
    return df

# =========================
# Prepare bases
# =========================
@st.cache_data(show_spinner=False)
def prep_cmv_fat(excel_path: str, sig) -> Optional[pd.DataFrame]:
    df = read_sheet(excel_path, "CMV E FATURAMENTO", sig)
    if df is None:
        return None
    d = df.copy()
    if "DATA" not in d.columns:
        return None
    d["_dt"] = pd.to_datetime(d["DATA"], errors="coerce", dayfirst=True)
    d["_ano"] = d["_dt"].dt.year
    d["_mes"] = d["_dt"].dt.month
    d["_receita"] = d.get("VR.TOTAL", 0).apply(to_num) if "VR.TOTAL" in d.columns else 0.0
    d["_cmv"] = d.get("CUSTO", 0).apply(to_num) if "CUSTO" in d.columns else 0.0
    d["LOJA"] = d.get("LOJA", "—").astype(str).str.strip()
    d["_loja_norm"] = d["LOJA"].apply(_norm_txt)
    return d

@st.cache_data(show_spinner=False)
def prep_dre_detail(excel_path: str, sig) -> Optional[pd.DataFrame]:
    df = read_sheet(excel_path, "DRE", sig)
    if df is None:
        return None
    d = df.copy()
    if "DTA.PAG" not in d.columns:
        for c in ["DTA.CAD", "DTA.VEN"]:
            if c in d.columns:
                d["DTA.PAG"] = d[c]
                break
    d["_dt"] = pd.to_datetime(d.get("DTA.PAG"), errors="coerce", dayfirst=True)
    d["_ano"] = d["_dt"].dt.year
    d["_mes"] = d["_dt"].dt.month
    d["_v"] = d.get("VAL.PAG", 0).apply(to_num) if "VAL.PAG" in d.columns else d.get("VAL.DUP", 0).apply(to_num)
    d["CONTA DE RESULTADO"] = d.get("CONTA DE RESULTADO", "—").astype(str).str.strip()
    d["DESPESA"] = d.get("DESPESA", "—")
    d["FAVORECIDO"] = d.get("FAVORECIDO", "—")
    d["HISTÓRICO"] = d.get("HISTÓRICO", d.get("HISTORICO", "—"))
    d["LOJA"] = d.get("LOJA", "—").astype(str).str.strip()
    d["_loja_norm"] = d["LOJA"].apply(_norm_txt)
    d["DESPESA_SINT"] = d["DESPESA"].apply(sintetizar_despesa)
    return d

def _sheet_month_value_map(df: pd.DataFrame, ano_ref: int, lojas_sel: List[str]) -> Dict[int, float]:
    """Sheets RECEBIMENTOS/COMPRAS/DEVOLUÇÕES: MESES, ANO, colunas por loja."""
    if df is None:
        return {m: 0.0 for m in range(1, 13)}
    d = df.copy()

    col_ano = "ANO" if "ANO" in d.columns else ("Ano" if "Ano" in d.columns else None)
    col_mes = None
    for c in ["MESES", "MÊS", "MES", "MÊS.", "MES."]:
        if c in d.columns:
            col_mes = c
            break
    if col_ano is None or col_mes is None:
        return {m: 0.0 for m in range(1, 13)}

    d["_ano"] = pd.to_numeric(d[col_ano], errors="coerce").astype("Int64")
    d["_mes"] = d[col_mes].apply(parse_mes)
    d = d[(d["_ano"] == int(ano_ref)) & (d["_mes"].notna())].copy()

    base_cols = {col_ano, col_mes, "_ano", "_mes"}
    lojas_cols = [c for c in d.columns if c not in base_cols and not str(c).startswith("Unnamed")]
    lojas_cols_norm = { _norm_txt(c): c for c in lojas_cols }

    lojas_eff = lojas_effective(lojas_sel)
    if not lojas_eff:
        use_cols = lojas_cols
    else:
        use_cols = []
        for l in lojas_eff:
            key = _norm_txt(l)
            if key in lojas_cols_norm:
                use_cols.append(lojas_cols_norm[key])

    if not use_cols:
        return {m: 0.0 for m in range(1, 13)}

    out = {m: 0.0 for m in range(1, 13)}
    for m in range(1, 13):
        dm = d[d["_mes"] == m]
        out[m] = float(dm[use_cols].applymap(to_num).to_numpy().sum()) if not dm.empty else 0.0
    return out

def _sum_dre_group_by_month(dre_det: pd.DataFrame, ano: int, lojas_sel: List[str], prefix: str) -> Dict[int, float]:
    if dre_det is None or dre_det.empty:
        return {m: 0.0 for m in range(1, 13)}
    d = dre_det[dre_det["_ano"] == int(ano)].copy()

    lojas_eff = lojas_effective(lojas_sel)
    if lojas_eff:
        lojas_eff_norm = set(_norm_txt(x) for x in lojas_eff)
        d = d[d["_loja_norm"].isin(lojas_eff_norm)]

    mask = d["CONTA DE RESULTADO"].astype(str).str.startswith(prefix)
    grp = d[mask].groupby("_mes")["_v"].sum()
    return {m: float(grp.get(m, 0.0)) for m in range(1, 13)}

def _shift_next(month_vals_this_year: Dict[int, float], month_vals_next_year: Dict[int, float]) -> Dict[int, float]:
    out = {}
    for m in range(1, 13):
        out[m] = float(month_vals_this_year.get(m + 1, 0.0)) if m < 12 else float(month_vals_next_year.get(1, 0.0))
    return out

def _shift_prev(month_vals_this_year: Dict[int, float], month_vals_prev_year: Dict[int, float]) -> Dict[int, float]:
    out = {}
    for m in range(1, 13):
        out[m] = float(month_vals_this_year.get(m - 1, 0.0)) if m > 1 else float(month_vals_prev_year.get(12, 0.0))
    return out

# =========================
# Tables
# =========================
def build_matrix_df(rows: List[Tuple[str, Dict[int, float]]],
                    meses_nums: List[int],
                    denom_by_month: Dict[int, float],
                    add_acum: bool = True) -> pd.DataFrame:
    out_rows = []
    for name, bym in rows:
        r = {"Linha": name}
        for m in meses_nums:
            pt = MES_NUM_TO_PT[m]
            val = float(bym.get(m, 0.0))
            denom = float(denom_by_month.get(m, 0.0))
            r[pt] = val
            r[f"{pt}%"] = (val / denom * 100.0) if denom != 0 else 0.0
        out_rows.append(r)
    df = pd.DataFrame(out_rows)

    if add_acum:
        cols_val = [MES_NUM_TO_PT[m] for m in meses_nums]
        df["ACUM"] = df[cols_val].sum(axis=1, skipna=True) if cols_val else 0.0
        denom_acum = float(sum(float(denom_by_month.get(m, 0.0)) for m in meses_nums))
        df["ACUM%"] = (df["ACUM"] / denom_acum * 100.0) if denom_acum != 0 else 0.0

    ordered = ["Linha"]
    for m in meses_nums:
        pt = MES_NUM_TO_PT[m]
        ordered.extend([pt, f"{pt}%"])
    if add_acum:
        ordered.extend(["ACUM", "ACUM%"])
    return df[ordered]

def style_result_rows(df: pd.DataFrame, result_labels: List[str]) -> object:
    def _styler(row):
        styles = [""] * len(row)
        label = str(row.get("Linha", ""))
        if label in result_labels:
            for j, col in enumerate(row.index):
                if col == "Linha":
                    styles[j] = "font-weight: 900;"
                else:
                    try:
                        v = float(row[col])
                    except Exception:
                        continue
                    styles[j] = "font-weight: 800; color: #c00000;" if v < 0 else "font-weight: 800; color: #1f4e79;"
        return styles
    return df.style.apply(_styler, axis=1)

def apply_formats(styler, meses_nums: List[int]) -> object:
    fmt = {}
    for m in meses_nums:
        pt = MES_NUM_TO_PT[m]
        fmt[pt] = lambda x: f"R$ {format_brl(x)}"
        fmt[f"{pt}%"] = lambda x: fmt_pct(x)
    fmt["ACUM"] = lambda x: f"R$ {format_brl(x)}"
    fmt["ACUM%"] = lambda x: fmt_pct(x)
    return styler.format(fmt)

# =========================
# Drill + Despesas + Históricos
# =========================
def drill_kpis(df_matrix: pd.DataFrame, meses_nums: List[int], denom_total: float, label_denom: str, key_prefix: str):
    st.markdown("### Drill — Linha selecionada")
    linha = st.selectbox("Linha", options=df_matrix["Linha"].tolist(), key=f"{key_prefix}_linha")

    cols_val = [MES_NUM_TO_PT[m] for m in meses_nums]
    row = df_matrix[df_matrix["Linha"] == linha].iloc[0]
    vals = pd.Series({c: float(row.get(c, 0.0)) for c in cols_val}, dtype="float64").fillna(0.0)

    total = float(vals.sum())
    media = float(total / max(len(cols_val), 1))
    pct = (total / denom_total * 100.0) if denom_total != 0 else 0.0

    c1, c2, c3 = st.columns(3)
    c1.metric("Total (R$)", fmt_brl(total))
    c2.metric("Média mensal (R$)", fmt_brl(media))
    c3.metric(f"% sobre {label_denom}", fmt_pct(pct))
    return linha

def despesas_e_historicos(dre_det: pd.DataFrame,
                          ano_ref: int,
                          meses_nums_sel: List[int],
                          lojas_sel: List[str],
                          linha_sel: str,
                          denom_periodo: float,
                          is_dre: bool,
                          group_def: Dict[str, Dict],
                          key_prefix: str):
    """
    Segue a Linha selecionada no Drill.
    - denom_periodo: Receita do período (DRE) ou Recebimento do período (DFC)
    """
    if linha_sel not in group_def:
        st.info("Essa linha não tem detalhamento por DESPESA (origem não é a aba DRE).")
        return

    g = group_def[linha_sel]
    prefix = g.get("prefix")
    shift_next = bool(g.get("shift_next", False))

    if dre_det is None or dre_det.empty:
        st.warning("Aba DRE não disponível para detalhamento.")
        return

    lojas_eff = lojas_effective(lojas_sel)
    lojas_eff_norm = set(_norm_txt(x) for x in lojas_eff) if lojas_eff else None

    def _filter_for_months(ano: int, meses: List[int]) -> pd.DataFrame:
        d = dre_det[(dre_det["_ano"] == int(ano)) & (dre_det["_mes"].isin(meses))].copy()
        if lojas_eff_norm is not None:
            d = d[d["_loja_norm"].isin(lojas_eff_norm)]
        if prefix:
            d = d[d["CONTA DE RESULTADO"].astype(str).str.startswith(prefix)]
        return d

    if shift_next:
        meses_this = [m + 1 for m in meses_nums_sel if m < 12]
        need_next_year = 12 in meses_nums_sel
        part1 = _filter_for_months(ano_ref, meses_this) if meses_this else dre_det.iloc[0:0].copy()
        part2 = _filter_for_months(ano_ref + 1, [1]) if need_next_year else dre_det.iloc[0:0].copy()
        base_raw = pd.concat([part1, part2], ignore_index=True)
    else:
        base_raw = _filter_for_months(ano_ref, meses_nums_sel)

    if base_raw.empty:
        st.info("Sem lançamentos para essa linha no período selecionado.")
        return

    denom = float(denom_periodo) if float(denom_periodo) != 0 else 0.0
    pct_label = "% sobre Receita" if is_dre else "% sobre Recebimentos"

    st.markdown("### Despesas (por DESPESA) — seguindo a Linha selecionada")
    agg = (base_raw.groupby("DESPESA_SINT", dropna=False)["_v"].sum()
           .reset_index().rename(columns={"_v": "Valor"}))
    agg[pct_label] = (agg["Valor"] / denom * 100.0) if denom != 0 else 0.0
    agg = agg.sort_values("Valor", ascending=False)

    top_max = min(80, max(5, len(agg)))
    top_def = min(15, top_max)
    topn = st.slider("Top N (despesas)", 5, top_max, top_def, key=f"{key_prefix}_topn_{_norm_txt(linha_sel)}")

    agg_top = agg.head(topn).copy()

    c1, c2 = st.columns([1.2, 1])
    with c1:
        fig = px.bar(agg_top, x="Valor", y="DESPESA_SINT", orientation="h",
                     title=f"Top {topn} despesas — {linha_sel}",
                     hover_data={pct_label: True})
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        st.dataframe(
            agg.style.format({"Valor": lambda x: f"R$ {format_brl(x)}", pct_label: lambda x: fmt_pct(x)}).hide(axis="index"),
            use_container_width=True
        )

    st.markdown("### Histórico da despesa — layout compacto")
    desp_sel = st.selectbox("Despesa (sintetizada)", options=agg["DESPESA_SINT"].tolist(), key=f"{key_prefix}_desp_{_norm_txt(linha_sel)}")
    raw_sel = base_raw[base_raw["DESPESA_SINT"] == desp_sel].copy()

    total_sel = float(raw_sel["_v"].sum())
    pct_sel = (total_sel / denom * 100.0) if denom != 0 else 0.0

    m1, m2, m3 = st.columns(3)
    m1.metric("Total da despesa (R$)", fmt_brl(total_sel))
    m2.metric(pct_label, fmt_pct(pct_sel))
    m3.metric("Qtd. lançamentos", f"{len(raw_sel):,}".replace(",", "."))

    raw_sel["_dt_sort"] = pd.to_datetime(raw_sel.get("DTA.PAG"), errors="coerce", dayfirst=True)
    raw_sel = raw_sel.sort_values(["_dt_sort"], ascending=False).drop(columns=["_dt_sort"])

    tab1, tab2, tab3 = st.tabs(["Sintetizado", "Sintetizado por Favorecido", "Detalhado"])

    with tab1:
        key_hist = "HISTÓRICO" if "HISTÓRICO" in raw_sel.columns else ("HISTORICO" if "HISTORICO" in raw_sel.columns else None)
        if key_hist is None:
            st.info("Sem coluna de HISTÓRICO.")
        else:
            tmp = raw_sel.copy()
            tmp[key_hist] = tmp[key_hist].astype(str).str.strip().replace({"": "—"})
            hist = (tmp.groupby(key_hist, dropna=False)["_v"].sum().reset_index().rename(columns={"_v": "Valor"}))
            hist[pct_label] = (hist["Valor"] / denom * 100.0) if denom != 0 else 0.0
            hist = hist.sort_values("Valor", ascending=False)
            st.dataframe(
                hist.style.format({"Valor": lambda x: f"R$ {format_brl(x)}", pct_label: lambda x: fmt_pct(x)}).hide(axis="index"),
                use_container_width=True
            )

    with tab2:
        if "FAVORECIDO" not in raw_sel.columns:
            st.info("Sem coluna FAVORECIDO.")
        else:
            tmp = raw_sel.copy()
            tmp["FAVORECIDO"] = tmp["FAVORECIDO"].astype(str).str.strip().replace({"": "—"})
            fav = (tmp.groupby("FAVORECIDO", dropna=False)["_v"].sum().reset_index().rename(columns={"_v": "Valor"}))
            fav[pct_label] = (fav["Valor"] / denom * 100.0) if denom != 0 else 0.0
            fav = fav.sort_values("Valor", ascending=False)
            st.dataframe(
                fav.style.format({"Valor": lambda x: f"R$ {format_brl(x)}", pct_label: lambda x: fmt_pct(x)}).hide(axis="index"),
                use_container_width=True
            )

    with tab3:
        cols = [c for c in ["DTA.PAG", "CONTA DE RESULTADO", "DESPESA", "FAVORECIDO", "DUPLICATA", "HISTÓRICO", "VAL.PAG"] if c in raw_sel.columns]
        view = raw_sel[cols].copy() if cols else raw_sel.copy()
        if "VAL.PAG" in view.columns:
            st.dataframe(view.style.format({"VAL.PAG": lambda x: f"R$ {format_brl(to_num(x))}"}).hide(axis="index"), use_container_width=True)
        else:
            st.dataframe(view.hide(axis="index") if hasattr(view, "hide") else view, use_container_width=True)

# =========================
# Sidebar
# =========================
st.sidebar.title("Filtros")

excel_path = _auto_find_excel()
if not excel_path:
    st.sidebar.error("Não encontrei nenhum Excel (.xlsx/.xlsm/.xls) na pasta do app.")
    st.stop()

sig = excel_signature(excel_path)
st.sidebar.caption(f"Excel: **{excel_path}**")

sheet_names = get_sheet_names(excel_path, sig)
need = {"CMV E FATURAMENTO", "DRE", "RECEBIMENTOS", "COMPRAS", "DEVOLUÇÕES"}
missing = [s for s in need if s not in set(sheet_names)]
if missing:
    st.sidebar.error("Faltam abas no Excel: " + ", ".join(missing))
    st.stop()

fat = prep_cmv_fat(excel_path, sig)
dre_det = prep_dre_detail(excel_path, sig)
df_receb = read_sheet(excel_path, "RECEBIMENTOS", sig)
df_compras = read_sheet(excel_path, "COMPRAS", sig)
df_devol = read_sheet(excel_path, "DEVOLUÇÕES", sig)

if fat is None or fat.empty:
    st.sidebar.error("Não consegui preparar a aba CMV E FATURAMENTO (verifique DATA, VR.TOTAL e CUSTO).")
    st.stop()

anos = sorted(set(pd.to_numeric(fat["_ano"], errors="coerce").dropna().astype(int).tolist()))
if dre_det is not None and not dre_det.empty:
    anos += list(set(pd.to_numeric(dre_det["_ano"], errors="coerce").dropna().astype(int).tolist()))
for dfx in [df_receb, df_compras, df_devol]:
    if dfx is not None and "ANO" in dfx.columns:
        anos += list(set(pd.to_numeric(dfx["ANO"], errors="coerce").dropna().astype(int).tolist()))
anos = sorted(set([a for a in anos if a > 1900]))
if not anos:
    st.sidebar.error("Não encontrei anos válidos nas bases.")
    st.stop()

ano_ref = st.sidebar.selectbox("Ano", options=anos, index=len(anos) - 1)

meses_pt_sel = st.sidebar.multiselect("Meses", options=MESES_PT, default=MESES_PT)
meses_nums = [MES_PT_TO_NUM[m] for m in meses_pt_sel] if meses_pt_sel else list(range(1, 13))

# lojas (união CMV + colunas dos agregados)
lojas_from_fat = sorted(set(fat["LOJA"].dropna().astype(str).str.strip().tolist()))

def _lojas_cols(df):
    if df is None:
        return []
    base_cols = {"MESES", "MÊS", "MES", "ANO", "Ano"}
    cols = [c for c in df.columns if str(c) not in base_cols and not str(c).startswith("Unnamed")]
    return [str(c).strip() for c in cols]

lojas_from_aggs = sorted(set(_lojas_cols(df_receb) + _lojas_cols(df_compras) + _lojas_cols(df_devol)))
lojas_all = ["TODAS"] + sorted(set(lojas_from_fat + lojas_from_aggs))
lojas_sel = st.sidebar.multiselect("Lojas", options=lojas_all, default=["TODAS"])
lojas_sel = ["TODAS"] if "TODAS" in lojas_sel and len(lojas_sel) == 1 else lojas_sel  # normaliza

pagina = st.sidebar.radio("Página", ["DRE", "DFC"])

# =========================
# Cálculos DRE
# =========================
def calc_receita_cmv_by_month() -> Tuple[Dict[int, float], Dict[int, float]]:
    d = fat[fat["_ano"] == int(ano_ref)].copy()
    lojas_eff = lojas_effective(lojas_sel)
    if lojas_eff:
        lojas_eff_norm = set(_norm_txt(x) for x in lojas_eff)
        d = d[d["_loja_norm"].isin(lojas_eff_norm)]
    rec = d.groupby("_mes")["_receita"].sum()
    cmv = d.groupby("_mes")["_cmv"].sum()
    return ({m: float(rec.get(m, 0.0)) for m in range(1, 13)},
            {m: float(cmv.get(m, 0.0)) for m in range(1, 13)})

def dre_page():
    st.title("DRE")

    receita_by_m, cmv_by_m = calc_receita_cmv_by_month()

    # Contas da aba DRE (prefixos revisados conforme seu Excel)
    ded_this = _sum_dre_group_by_month(dre_det, ano_ref, lojas_sel, "00004 -")
    ded_next = _sum_dre_group_by_month(dre_det, ano_ref + 1, lojas_sel, "00004 -")
    ded_by_m = _shift_next(ded_this, ded_next)

    pes_this = _sum_dre_group_by_month(dre_det, ano_ref, lojas_sel, "00006 -")
    pes_next = _sum_dre_group_by_month(dre_det, ano_ref + 1, lojas_sel, "00006 -")
    pes_by_m = _shift_next(pes_this, pes_next)

    adm_by_m = _sum_dre_group_by_month(dre_det, ano_ref, lojas_sel, "00007 -")
    com_by_m = _sum_dre_group_by_month(dre_det, ano_ref, lojas_sel, "00009 -")
    fin_by_m = _sum_dre_group_by_month(dre_det, ano_ref, lojas_sel, "00011 -")
    fornec_by_m = _sum_dre_group_by_month(dre_det, ano_ref, lojas_sel, "00012 -")  # (não exibimos na DRE)
    imo_by_m = _sum_dre_group_by_month(dre_det, ano_ref, lojas_sel, "00014 -")
    inv_by_m = _sum_dre_group_by_month(dre_det, ano_ref, lojas_sel, "00015 -")
    ret_by_m = _sum_dre_group_by_month(dre_det, ano_ref, lojas_sel, "00016 -")
    op_by_m  = _sum_dre_group_by_month(dre_det, ano_ref, lojas_sel, "00017 -")

    markup_by_m = {m: safe_div(receita_by_m[m], cmv_by_m[m]) for m in range(1, 13)}
    margem_bruta_by_m = {m: receita_by_m[m] - cmv_by_m[m] for m in range(1, 13)}

    def _sum_costs(m):
        return (ded_by_m[m] + pes_by_m[m] + adm_by_m[m] + com_by_m[m] + fin_by_m[m] +
                imo_by_m[m] + inv_by_m[m] + ret_by_m[m] + op_by_m[m])

    resultado_oper_by_m = {m: margem_bruta_by_m[m] - _sum_costs(m) for m in range(1, 13)}
    res_antes_by_m = {m: resultado_oper_by_m[m] + fin_by_m[m] + ret_by_m[m] + inv_by_m[m] for m in range(1, 13)}

    rows_no_markup = [
        ("RECEITA", receita_by_m),
        ("CMV", cmv_by_m),
        ("MARGEM BRUTA", margem_bruta_by_m),
        ("DEDUÇÕES (IMPOSTOS SOBRE VENDAS)", ded_by_m),
        ("DESPESAS COM PESSOAL", pes_by_m),
        ("DESPESAS ADMINISTRATIVAS", adm_by_m),
        ("DESPESAS COMERCIAIS", com_by_m),
        ("DESPESAS FINANCEIRAS", fin_by_m),
        ("IMOBILIZADO", imo_by_m),
        ("INVESTIMENTOS", inv_by_m),
        ("RETIRADAS SÓCIOS", ret_by_m),
        ("DESPESAS OPERACIONAIS", op_by_m),
        ("RES. ANTES DAS DEP FINANC E RETIRADAS/INVESTIMENTOS", res_antes_by_m),
        ("RESULTADO OPERACIONAL", resultado_oper_by_m),
    ]

    denom = receita_by_m
    df = build_matrix_df(rows_no_markup, meses_nums, denom, add_acum=True)

    # Inserir Markup (sem %)
    mk = {"Linha": "Markup"}
    for m in meses_nums:
        pt = MES_NUM_TO_PT[m]
        mk[pt] = float(markup_by_m.get(m, 0.0))
        mk[f"{pt}%"] = np.nan
    mk["ACUM"] = float(np.nanmean([markup_by_m.get(m, 0.0) for m in meses_nums])) if meses_nums else 0.0
    mk["ACUM%"] = np.nan

    # coloca Markup após CMV
    df = pd.concat([df.iloc[0:2], pd.DataFrame([mk])[df.columns], df.iloc[2:]], ignore_index=True)

    st.subheader("DRE — Valores em R$ e % sobre Receita (Markup sem %)")
    result_labels = ["RESULTADO OPERACIONAL", "RES. ANTES DAS DEP FINANC E RETIRADAS/INVESTIMENTOS"]
    sty = style_result_rows(df, result_labels)
    sty = apply_formats(sty, meses_nums)

    # Formatar Markup como número (2 casas)
    def _fmt_markup(val):
        try:
            return f"{float(val):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except Exception:
            return "0,00"

    for m in meses_nums:
        sty = sty.format({MES_NUM_TO_PT[m]: (lambda x, _f=_fmt_markup: _f(x))}, subset=pd.IndexSlice[df["Linha"] == "Markup", [MES_NUM_TO_PT[m]]])

    st.dataframe(sty.hide(axis="index"), use_container_width=True)

    receita_periodo = float(sum(receita_by_m.get(m, 0.0) for m in meses_nums))
    linha_sel = drill_kpis(df, meses_nums, receita_periodo, "Receita", key_prefix="dre")

    group_def = {
        "DEDUÇÕES (IMPOSTOS SOBRE VENDAS)": {"prefix": "00004 -", "shift_next": True},
        "DESPESAS COM PESSOAL": {"prefix": "00006 -", "shift_next": True},
        "DESPESAS ADMINISTRATIVAS": {"prefix": "00007 -", "shift_next": False},
        "DESPESAS COMERCIAIS": {"prefix": "00009 -", "shift_next": False},
        "DESPESAS FINANCEIRAS": {"prefix": "00011 -", "shift_next": False},
        "IMOBILIZADO": {"prefix": "00014 -", "shift_next": False},
        "INVESTIMENTOS": {"prefix": "00015 -", "shift_next": False},
        "RETIRADAS SÓCIOS": {"prefix": "00016 -", "shift_next": False},
        "DESPESAS OPERACIONAIS": {"prefix": "00017 -", "shift_next": False},
    }
    despesas_e_historicos(dre_det, ano_ref, meses_nums, lojas_sel, linha_sel, receita_periodo, True, group_def, key_prefix="dre")

# =========================
# Cálculos DFC
# =========================
def dfc_page():
    st.title("DFC")

    receb_by_m = _sheet_month_value_map(df_receb, ano_ref, lojas_sel)

    comp_this = _sheet_month_value_map(df_compras, ano_ref, lojas_sel)
    comp_prev = _sheet_month_value_map(df_compras, ano_ref - 1, lojas_sel)
    dev_this  = _sheet_month_value_map(df_devol, ano_ref, lojas_sel)
    dev_prev  = _sheet_month_value_map(df_devol, ano_ref - 1, lojas_sel)

    compras_by_m = _shift_prev(comp_this, comp_prev)   # mês anterior
    devol_by_m   = _shift_prev(dev_this, dev_prev)     # mês anterior

    compras_liq_by_m = {m: compras_by_m[m] - devol_by_m[m] for m in range(1, 13)}
    margem_bruta_by_m = {m: receb_by_m[m] - compras_liq_by_m[m] for m in range(1, 13)}

    # contas da aba DRE (sem shift no DFC)
    ded_by_m    = _sum_dre_group_by_month(dre_det, ano_ref, lojas_sel, "00004 -")
    pes_by_m    = _sum_dre_group_by_month(dre_det, ano_ref, lojas_sel, "00006 -")
    adm_by_m    = _sum_dre_group_by_month(dre_det, ano_ref, lojas_sel, "00007 -")
    com_by_m    = _sum_dre_group_by_month(dre_det, ano_ref, lojas_sel, "00009 -")
    fin_by_m    = _sum_dre_group_by_month(dre_det, ano_ref, lojas_sel, "00011 -")
    imo_by_m    = _sum_dre_group_by_month(dre_det, ano_ref, lojas_sel, "00014 -")
    inv_by_m    = _sum_dre_group_by_month(dre_det, ano_ref, lojas_sel, "00015 -")
    ret_by_m    = _sum_dre_group_by_month(dre_det, ano_ref, lojas_sel, "00016 -")
    op_by_m     = _sum_dre_group_by_month(dre_det, ano_ref, lojas_sel, "00017 -")

    def _sum_costs(m):
        return (ded_by_m[m] + pes_by_m[m] + adm_by_m[m] + com_by_m[m] + fin_by_m[m] +
                imo_by_m[m] + inv_by_m[m] + ret_by_m[m] + op_by_m[m])

    saldo_oper_by_m = {m: margem_bruta_by_m[m] - _sum_costs(m) for m in range(1, 13)}
    saldo_antes_by_m = {m: saldo_oper_by_m[m] + ret_by_m[m] + inv_by_m[m] + fin_by_m[m] for m in range(1, 13)}

    rows = [
        ("RECEBIMENTO", receb_by_m),
        ("COMPRAS (MÊS ANTERIOR)", compras_by_m),
        ("DEVOLUÇÕES (MÊS ANTERIOR)", devol_by_m),
        ("COMPRAS LÍQ", compras_liq_by_m),
        ("MARGEM BRUTA", margem_bruta_by_m),
        ("DEDUÇÕES (IMPOSTOS SOBRE VENDAS)", ded_by_m),
        ("DESPESAS COM PESSOAL", pes_by_m),
        ("DESPESAS ADMINISTRATIVAS", adm_by_m),
        ("DESPESAS COMERCIAIS", com_by_m),
        ("DESPESAS FINANCEIRAS", fin_by_m),
        ("IMOBILIZADO", imo_by_m),
        ("INVESTIMENTOS", inv_by_m),
        ("RETIRADAS SÓCIOS", ret_by_m),
        ("DESPESAS OPERACIONAIS", op_by_m),
        ("SALDO OPERACIONAL ANTES DE RETIRADAS/INVESTIMENTOS/DESP FINANCEIRAS", saldo_antes_by_m),
        ("SALDO OPERACIONAL", saldo_oper_by_m),
    ]

    df = build_matrix_df(rows, meses_nums, receb_by_m, add_acum=True)

    st.subheader("DFC — Valores em R$ e % sobre Recebimento")
    result_labels = ["SALDO OPERACIONAL", "SALDO OPERACIONAL ANTES DE RETIRADAS/INVESTIMENTOS/DESP FINANCEIRAS"]
    sty = style_result_rows(df, result_labels)
    sty = apply_formats(sty, meses_nums)
    st.dataframe(sty.hide(axis="index"), use_container_width=True)

    receb_periodo = float(sum(receb_by_m.get(m, 0.0) for m in meses_nums))
    linha_sel = drill_kpis(df, meses_nums, receb_periodo, "Recebimento", key_prefix="dfc")

    group_def = {
        "DEDUÇÕES (IMPOSTOS SOBRE VENDAS)": {"prefix": "00004 -", "shift_next": False},
        "DESPESAS COM PESSOAL": {"prefix": "00006 -", "shift_next": False},
        "DESPESAS ADMINISTRATIVAS": {"prefix": "00007 -", "shift_next": False},
        "DESPESAS COMERCIAIS": {"prefix": "00009 -", "shift_next": False},
        "DESPESAS FINANCEIRAS": {"prefix": "00011 -", "shift_next": False},
        "IMOBILIZADO": {"prefix": "00014 -", "shift_next": False},
        "INVESTIMENTOS": {"prefix": "00015 -", "shift_next": False},
        "RETIRADAS SÓCIOS": {"prefix": "00016 -", "shift_next": False},
        "DESPESAS OPERACIONAIS": {"prefix": "00017 -", "shift_next": False},
    }
    despesas_e_historicos(dre_det, ano_ref, meses_nums, lojas_sel, linha_sel, receb_periodo, False, group_def, key_prefix="dfc")

# =========================
# Router
# =========================
if pagina == "DRE":
    dre_page()
else:
    dfc_page()
