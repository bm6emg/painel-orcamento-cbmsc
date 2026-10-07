"""Painel de execução orçamentária do CBMSC (FUMCBM) em Streamlit.

Como atualizar: troque os arquivos da pasta `dados/` no GitHub.
  - Análise de despesa do SIGEF (.xlsx). Mantenha a data no nome, ex.: "07-10-2026 - analisedespesa.xlsx".
    Se houver mais de um, o painel usa o de data mais recente.
  - preempenhos.xlsx (opcional): colunas documento, subacao, fonte, natureza, a_empenhar, objeto.
"""
import calendar
import json
import re
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

BASE = Path(__file__).parent
DADOS = BASE / "dados"
TEMPLATE = BASE / "painel_template.html"

COLS = [
    "cdcredor", "nmcredor", "cdsubacao", "nmsubacao", "cdfonterecurso", "nmfonterecurso",
    "cdsubelemento", "nmelemento", "nuano", "numes",
    "vldotacaoinicial", "vldotacaoatualizada", "vlempenhado", "vlliquidado", "vlpago",
]
VALS = ["vldotacaoinicial", "vldotacaoatualizada", "vlempenhado", "vlliquidado", "vlpago"]

st.set_page_config(page_title="Execução orçamentária CBMSC", layout="wide", initial_sidebar_state="collapsed")
st.markdown(
    """<style>
    header, footer, #MainMenu {visibility:hidden;height:0}
    .block-container {padding:0 !important; max-width:100% !important}
    iframe {height:calc(100vh - 4px) !important; width:100% !important; border:0}
    </style>""",
    unsafe_allow_html=True,
)


def data_no_nome(p: Path):
    m = re.search(r"(\d{2})[-_.](\d{2})[-_.](\d{4})", p.name)
    if not m:
        return None
    try:
        return date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    except ValueError:
        return None


def arquivo_sigef():
    cands = [p for p in DADOS.glob("*.xlsx") if "preempenho" not in p.name.lower()]
    if not cands:
        return None
    return max(cands, key=lambda p: (data_no_nome(p) or date.min, p.stat().st_mtime))


def digitos(v) -> str:
    if pd.isna(v):
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return re.sub(r"\D", "", str(v).split(",")[0])


def nome_subacao(s) -> str:
    s = str(s).strip()
    return s.split(": ")[-1].strip() if ": " in s else s


@st.cache_data(show_spinner="Lendo a planilha do SIGEF…")
def montar_dados(caminho: str, mtime: float, pe_caminho: str, pe_mtime: float) -> dict:
    df = pd.read_excel(caminho, usecols=lambda c: c in COLS)
    df[VALS] = df[VALS].fillna(0).astype(float)
    ano = int(df["nuano"].max())
    df = df[df["nuano"] == ano].copy()
    df["f7"] = df["cdfonterecurso"].map(digitos).str[:7]
    df["fonte"] = df["cdfonterecurso"].map(digitos)
    df["el"] = df["cdsubelemento"].map(digitos).str[:6]
    df["sub"] = df["cdsubacao"].astype(int)
    df["subnm"] = df["nmsubacao"].map(nome_subacao)

    fontes = sorted(df["f7"].unique())
    subs = sorted(df["sub"].unique())
    els = sorted(df["el"].unique())
    nome_f = {}
    for f in fontes:
        s = df[df["f7"] == f]
        base = s[s["fonte"].str.endswith("000")]
        nome_f[f] = (base if len(base) else s)["nmfonterecurso"].value_counts().index[0]
    nome_s = df.groupby("sub")["subnm"].first().to_dict()
    nome_e = df.groupby("el")["nmelemento"].agg(lambda x: x.value_counts().index[0]).to_dict()
    fi = {f: i for i, f in enumerate(fontes)}
    si = {s: i for i, s in enumerate(subs)}
    ei = {e: i for i, e in enumerate(els)}

    g = df.groupby(["numes", "f7", "sub", "el"])[VALS].sum().reset_index()
    M = [[int(r.numes), fi[r.f7], si[r.sub], ei[r.el]] + [round(float(getattr(r, c)), 2) for c in VALS]
         for r in g.itertuples(index=False)]

    emp_cred = set(df.loc[df["vlempenhado"] != 0, "cdcredor"])
    c = df[df["cdcredor"].isin(emp_cred)].groupby(["cdcredor", "f7", "sub", "el"])[
        ["vlempenhado", "vlliquidado", "vlpago"]].sum().reset_index()
    c = c[c[["vlempenhado", "vlliquidado", "vlpago"]].abs().sum(axis=1) > 0]
    creds = sorted(c["cdcredor"].unique())
    ci = {k: i for i, k in enumerate(creds)}
    nome_c = df.groupby("cdcredor")["nmcredor"].first().str.strip().to_dict()
    C = [[ci[r.cdcredor], fi[r.f7], si[r.sub], ei[r.el], round(float(r.vlempenhado), 2),
          round(float(r.vlliquidado), 2), round(float(r.vlpago), 2)] for r in c.itertuples(index=False)]

    nm = int(df["numes"].max())
    pos = data_no_nome(Path(caminho))
    hoje = date.today()
    if pos is None:
        pos = hoje if (ano == hoje.year and nm == hoje.month) else date(ano, nm, calendar.monthrange(ano, nm)[1])
    partial = (pos.year == ano and pos.month == nm and pos.day < 28)

    out = {
        "ano": ano,
        "fontes": [[f, nome_f[f]] for f in fontes],
        "subs": [[int(s), nome_s[s]] for s in subs],
        "els": [[e, nome_e[e]] for e in els],
        "cred": [nome_c.get(k, str(k)) for k in creds],
        "M": M, "C": C,
        "pos": pos.isoformat(), "partial": partial,
        "srcLabel": f"Fonte: SIGEF – {Path(caminho).name}",
    }

    if pe_caminho:
        pe = pd.read_excel(pe_caminho, dtype=str)
        pe.columns = [str(x).strip().lower() for x in pe.columns]
        need = ["documento", "subacao", "fonte", "natureza", "a_empenhar"]
        if all(k in pe.columns for k in need):
            def num(v):
                s = str(v).strip()
                if "," in s:
                    s = s.replace(".", "").replace(",", ".")
                try:
                    return float(s)
                except ValueError:
                    return 0.0
            pe = pe.fillna("")
            out["PE"] = [[r["documento"], digitos(r["subacao"]), digitos(r["fonte"]), digitos(r["natureza"]),
                          num(r["a_empenhar"]), r.get("objeto", "")] for _, r in pe.iterrows()]
            out["peSrc"] = f"Fonte: dados/{Path(pe_caminho).name}"
    return out


sig = arquivo_sigef()
if sig is None:
    st.error("Coloque a análise de despesa do SIGEF (.xlsx) na pasta `dados/` do repositório.")
    st.stop()
pe_file = next(iter(DADOS.glob("preempenho*.xlsx")), None)
dados = montar_dados(str(sig), sig.stat().st_mtime,
                     str(pe_file) if pe_file else "", pe_file.stat().st_mtime if pe_file else 0.0)

html = TEMPLATE.read_text(encoding="utf-8").replace(
    "__DATA__", json.dumps(dados, ensure_ascii=False, separators=(",", ":")))
components.html(html, height=900, scrolling=True)
