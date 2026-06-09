from pathlib import Path
from datetime import datetime
import re

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="CETESBRisk Explorer", page_icon="🧪", layout="wide")

DATA_PATH = Path(__file__).parent / "data" / "Analise_detalhada_CETESBRisk_2023_vs_2026_CMA.xlsx"
AUDIT_XLSX_PATH = Path(__file__).parent / "data" / "auditoria_classificacao_SQIs_CETESBRisk_regra_objetiva_v2.xlsx"

GLOBAL_COUNTS = {
    "SQIs comuns comparadas": 819,
    "Comparações individuais": 29484,
    "Novas SQIs incluídas": 59,
    "SQIs removidas": 1,
    "Comparações com alteração": 1009,
    "Comparações sem alteração": 16565,
    "Sem valor/vazios": 11910,
}

DENSITY_MATERIALITY_THRESHOLD = 0.05  # escala decimal: 0,05 = 5%
CHANGED_STATUS = {"aumentou", "diminuiu", "alteração textual/categórica", "alterado"}
RAW_CHANGE_STATUS = CHANGED_STATUS | {"incluído em 2026", "removido em 2026"}

TOX_RELEVANTE = {
    "RfDo",
    "RfCi",
    "SFO",
    "IUR",
    "CARCINOGÊNICO",
    "Classe de Cancer",
    "Mutagenico",
}
REGULATORIO = {"MCL", "Potabilidade"}
FISQUI_RELEVANTE = {
    "H (-)",
    "HLC (atm-m³/mole)",
    "Dia (cm²/s)",
    "Diw (cm²/s)",
    "Koc",
    "S (mg/L)",
    "Kd (L/kg)",
    "log Kow",
    "Pvap (mm Hg)",
    "Csat",
    "B (-)",
    "FA",
    "PC (cm/h)",
    "PF (°C)",
}
DENSITY_PARAM = "Densidade"

CLASS_DESCRIPTIONS = {
    "A": {
        "descricao": "Sem alteração material relevante",
        "recomendacao": "Não indica revisão automática da AR exclusivamente pela atualização da planilha.",
    },
    "B": {
        "descricao": "Alterações menores",
        "recomendacao": "Revisar apenas em cenários muito sensíveis ou resultados próximos aos critérios aceitáveis.",
    },
    "C": {
        "descricao": "Alterações regulatórias/potabilidade",
        "recomendacao": "Revisar enquadramento regulatório/potabilidade; não implica necessariamente recálculo do risco.",
    },
    "D1": {
        "descricao": "Alterações toxicológicas relevantes",
        "recomendacao": "Revisão dirigida recomendada quando a SQI for relevante para o site, especialmente se o resultado anterior estiver próximo ao limite.",
    },
    "D2": {
        "descricao": "Alterações físico-químicas relevantes",
        "recomendacao": "Revisão dirigida em cenários sensíveis à volatilização, particionamento, transporte ou intrusão de vapores.",
    },
}

CLASS_ORDER = ["A", "B", "C", "D1", "D2"]
CLASS_COLORS = {"A": "#2e7d32", "B": "#6c757d", "C": "#f9a825", "D1": "#c62828", "D2": "#ef6c00"}
LEGACY_CLASS_COUNTS = {"A": 296, "B": 471, "C": 18, "D1": 23, "D2": 11}

# Classificador legado, apenas para auditoria e comparação histórica.
LEGACY_CAS_OVERRIDES = {
    "71-43-2": "A",
    "108-88-3": "A",
    "100-41-4": "A",
    "95-47-6": "A",
    "106-42-3": "A",
    "108-38-3": "A",
    "127-18-4": "B",
    "79-01-6": "B",
    "156-59-2": "C",
    "156-60-5": "C",
    "75-35-4": "D1",
    "75-01-4": "D1",
    "64742-89-8": "D2",
}
LEGACY_NAME_OVERRIDES = [
    ("benzene", "A"),
    ("toluene", "A"),
    ("ethylbenzene", "A"),
    ("xylene", "A"),
    ("tetrachloroethylene", "B"),
    ("tetrachloroethene", "B"),
    ("trichloroethylene", "B"),
    ("trichloroethene", "B"),
    ("cis-1,2-dichloroethene", "C"),
    ("trans-1,2-dichloroethene", "C"),
    ("1,1-dichloroethene", "D1"),
    ("vinyl chloride", "D1"),
    ("aliphatic low", "D2"),
    ("c5-c8", "D2"),
    ("aliphatic medium", "A"),
    ("c9-c18", "A"),
    ("aliphatic high", "A"),
    ("c19-c32", "A"),
    ("aromatic medium", "A"),
    ("c9-c10", "A"),
    ("aromatic high", "A"),
    ("c10-c32", "A"),
    ("pfos", "D1"),
    ("pfoa", "D1"),
    ("hfpo", "D1"),
    ("genx", "D1"),
    ("chromium(vi)", "D1"),
    ("chromium vi", "D1"),
    ("formaldehyde", "D1"),
    ("acrylonitrile", "D1"),
    ("chloroprene", "D1"),
]


@st.cache_data
def load_data():
    xls = pd.ExcelFile(DATA_PATH)
    return (
        pd.read_excel(xls, "Comparativo_Completo"),
        pd.read_excel(xls, "SQIs_Incluidas_2026"),
        pd.read_excel(xls, "SQIs_Removidas_2026"),
        pd.read_excel(xls, "Alteracoes_Relevantes"),
        pd.read_excel(xls, "Impacto_CMA_Direto"),
    )


def clean_status(value):
    return str(value).strip().lower()


def changed(row):
    return clean_status(row.get("Status", "")) in CHANGED_STATUS


def raw_changed(row):
    return clean_status(row.get("Status", "")) in RAW_CHANGE_STATUS


def num(value):
    try:
        if pd.isna(value):
            return None
        return float(value)
    except Exception:
        return None


def abs_var(row):
    value = num(row.get("Variação % Parâmetro", 0))
    return 0.0 if value is None else abs(value)


def fmt_value(value):
    if pd.isna(value):
        return "sem valor"
    value_num = num(value)
    if value_num is None:
        return str(value)
    if abs(value_num) >= 1000 or (0 < abs(value_num) < 0.001):
        return f"{value_num:.3E}".replace(".", ",")
    return f"{value_num:g}".replace(".", ",")


def fmt_pct_from_decimal(value, digits=2):
    value_num = num(value)
    if value_num is None:
        return "n.d."
    return f"{value_num * 100:+.{digits}f}%".replace(".", ",")


def is_density_submaterial(row):
    return str(row.get("Parâmetro", "")) == DENSITY_PARAM and abs_var(row) < DENSITY_MATERIALITY_THRESHOLD


def effective_changes(group):
    ch = group[group.apply(changed, axis=1)].copy()
    if ch.empty:
        return ch
    return ch[~ch.apply(is_density_submaterial, axis=1)].copy()


def ignored_by_threshold(group):
    ch = group[group.apply(changed, axis=1)].copy()
    if ch.empty:
        return ch
    return ch[ch.apply(is_density_submaterial, axis=1)].copy()


def classify_objective(group):
    ch_eff = effective_changes(group)
    if ch_eff.empty:
        return "A"

    tox = ch_eff[ch_eff["Parâmetro"].isin(TOX_RELEVANTE)]
    if not tox.empty:
        return "D1"

    phys_rel = ch_eff[ch_eff["Parâmetro"].isin(FISQUI_RELEVANTE)]
    phys_rel_material = phys_rel[phys_rel.apply(abs_var, axis=1) >= DENSITY_MATERIALITY_THRESHOLD]
    if not phys_rel_material.empty:
        return "D2"

    reg = ch_eff[ch_eff["Parâmetro"].isin(REGULATORIO)]
    if not reg.empty:
        return "C"

    return "B"


def classify_legacy(cas, name, group):
    cas = "" if pd.isna(cas) else str(cas).strip()
    name_l = "" if pd.isna(name) else str(name).lower()
    if cas in LEGACY_CAS_OVERRIDES:
        return LEGACY_CAS_OVERRIDES[cas]
    for key, cls in LEGACY_NAME_OVERRIDES:
        if key in name_l:
            return cls
    ch = group[group.apply(changed, axis=1)]
    if ch.empty:
        return "A"
    tox = ch[ch["Parâmetro"].isin(TOX_RELEVANTE)]
    if not tox.empty and (
        not tox[tox["Status"].astype(str).str.contains("textual|categ", case=False, na=False)].empty
        or not tox[tox.apply(lambda r: abs_var(r) >= DENSITY_MATERIALITY_THRESHOLD, axis=1)].empty
    ):
        return "D1"
    reg = ch[ch["Parâmetro"].isin(REGULATORIO)]
    phys = ch[ch["Parâmetro"].isin(FISQUI_RELEVANTE | {DENSITY_PARAM})]
    phys_mat = phys[phys.apply(lambda r: abs_var(r) >= DENSITY_MATERIALITY_THRESHOLD, axis=1)]
    if not reg.empty and phys_mat.empty:
        return "C"
    if not phys_mat.empty:
        return "D2"
    return "B"


def join_params(df, limit=10):
    vals = df["Parâmetro"].dropna().astype(str).unique().tolist() if "Parâmetro" in df.columns else []
    if not vals:
        return ""
    txt = ", ".join(vals[:limit])
    return txt + ("..." if len(vals) > limit else "")


def class_motive(cls, ch, eff, ignored):
    if cls == "A":
        if ch.empty:
            return "Sem alteração real identificada nos parâmetros avaliados."
        if not ignored.empty and eff.empty:
            return "Apenas densidade com |Δ| < 5%; alteração ignorada por limiar de materialidade."
        return "Sem alteração considerada pela regra objetiva."
    if cls == "B":
        if not eff[eff["Parâmetro"].eq(DENSITY_PARAM)].empty:
            return "Alteração residual considerada; densidade com |Δ| ≥ 5% não gera D2 isoladamente."
        return "Alteração residual considerada, sem enquadramento em D1, D2 ou C."
    if cls == "C":
        return "Alteração em MCL/Potabilidade, sem alteração D1 ou D2 para a mesma SQI."
    if cls == "D1":
        return "Alteração em parâmetro toxicológico relevante; qualquer alteração real é material."
    if cls == "D2":
        return "Alteração físico-química relevante, exceto densidade, com |Δ| ≥ 5%."
    return ""


@st.cache_data
def build_master(comp):
    rows = []
    for (cas, a23, a26), group in comp.groupby(["CAS", "Analito 2023", "Analito 2026"], dropna=False):
        name = a26 if pd.notna(a26) and str(a26).strip() else a23
        cls = classify_objective(group)
        ch = group[group.apply(changed, axis=1)].copy()
        eff = effective_changes(group)
        ignored = ignored_by_threshold(group)
        legacy = classify_legacy(cas, name, group)
        rows.append(
            {
                "CAS": cas,
                "Composto": name,
                "Classe": cls,
                "Classe anterior app": legacy,
                "Mudou com nova regra?": "Sim" if cls != legacy else "Não",
                "Descrição da classe": CLASS_DESCRIPTIONS[cls]["descricao"],
                "Nº alterações brutas": int(ch.shape[0]),
                "Parâmetros alterados brutos": join_params(ch),
                "Nº alterações consideradas na classificação": int(eff.shape[0]),
                "Parâmetros considerados na classificação": join_params(eff),
                "Parâmetros ignorados por limiar": join_params(ignored),
                "Motivo da classe": class_motive(cls, ch, eff, ignored),
                "Recomendação": CLASS_DESCRIPTIONS[cls]["recomendacao"],
            }
        )
    return pd.DataFrame(rows)


def badge(cls):
    return (
        f"<span style='background:{CLASS_COLORS.get(cls, '#999')}; color:white; padding:0.35rem 0.6rem; "
        f"border-radius:0.5rem; font-weight:700;'>Classe {cls} · {CLASS_DESCRIPTIONS[cls]['descricao']}</span>"
    )


def report_html(compound, cas, cls, detail, synthesis):
    now = datetime.now().strftime("%d/%m/%Y %H:%M")
    return f"""<html><head><meta charset='utf-8'><title>Relatório CETESBRisk Explorer - {compound}</title><style>body{{font-family:Arial;margin:40px;color:#1F2A44}}h1,h2{{color:#0B2A4A}}.box{{background:#f3f6fa;border-left:5px solid #0B5CAD;padding:15px;margin:20px 0}}table{{border-collapse:collapse;width:100%;font-size:11px}}th,td{{border:1px solid #ddd;padding:6px}}th{{background:#0B5CAD;color:white}}</style></head><body><h1>CETESBRisk Explorer</h1><h2>Relatório técnico por SQI</h2><p><b>Data/hora:</b> {now}</p><p><b>Substância:</b> {compound}</p><p><b>CAS:</b> {cas}</p><p><b>Classe:</b> {cls} - {CLASS_DESCRIPTIONS[cls]['descricao']}</p><div class='box'>{synthesis}</div><h2>Comparativo por parâmetro</h2>{detail.to_html(index=False,border=0)}</body></html>"""


def to_float_series(s):
    return pd.to_numeric(s, errors="coerce")


def add_classification_flags(df):
    out = df.copy()
    out["Alteração bruta?"] = out.apply(raw_changed, axis=1).map({True: "Sim", False: "Não"})
    out["Alteração considerada na classificação?"] = out.apply(lambda r: changed(r) and not is_density_submaterial(r), axis=1).map({True: "Sim", False: "Não"})
    out["Motivo de uso na classificação"] = out.apply(
        lambda r: (
            "Densidade ignorada: |Δ| < 5%" if changed(r) and is_density_submaterial(r)
            else "Parâmetro considerado pela regra objetiva" if changed(r)
            else "Sem alteração considerada"
        ),
        axis=1,
    )
    return out


def top_sqi_by_parameter_changes(comp, fonte, n=5, considered_only=True):
    df = comp.copy()
    df = df[df["Fonte"].astype(str).str.lower() == fonte.lower()]
    if considered_only:
        df = df[df.apply(lambda r: changed(r) and not is_density_submaterial(r), axis=1)].copy()
    else:
        df = df[df.apply(raw_changed, axis=1)].copy()
    if df.empty:
        return pd.DataFrame(columns=["SQI", "CAS", "Nº parâmetros alterados", "Maior variação absoluta (%)", "Principais parâmetros alterados"])

    df["abs_var"] = to_float_series(df.get("Variação % Parâmetro", pd.Series(dtype=float))).abs()
    grouped = (
        df.groupby(["Analito 2026", "CAS"], dropna=False)
        .agg(
            **{
                "Nº parâmetros alterados": ("Parâmetro", "nunique"),
                "Maior variação absoluta (%)": ("abs_var", lambda x: round(x.max() * 100, 2) if pd.notna(x.max()) else None),
                "Principais parâmetros alterados": ("Parâmetro", lambda x: ", ".join(pd.Series(x).dropna().astype(str).unique()[:6])),
            }
        )
        .reset_index()
        .rename(columns={"Analito 2026": "SQI"})
    )
    return grouped.sort_values(["Nº parâmetros alterados", "Maior variação absoluta (%)"], ascending=False).head(n)


def top_cma_impact(comp, direction="reduz", n=5):
    df = comp.copy()
    df = df[df.apply(lambda r: changed(r) and not is_density_submaterial(r), axis=1)].copy()
    df["Estimativa variação CMA % num"] = to_float_series(df.get("Estimativa variação CMA %", pd.Series(dtype=float)))
    df = df.dropna(subset=["Estimativa variação CMA % num"])
    if direction == "reduz":
        df = df[df["Estimativa variação CMA % num"] < 0]
        sort_ascending = True
    else:
        df = df[df["Estimativa variação CMA % num"] > 0]
        sort_ascending = False

    if df.empty:
        return pd.DataFrame(columns=["SQI", "CAS", "Parâmetro crítico", "Variação estimada CMA (%)", "Tendência esperada da CMA"])

    df["abs_cma"] = df["Estimativa variação CMA % num"].abs()
    idx = df.groupby(["Analito 2026", "CAS"], dropna=False)["abs_cma"].idxmax()
    out = df.loc[idx, ["Analito 2026", "CAS", "Parâmetro", "Estimativa variação CMA % num", "Tendência esperada da CMA"]].copy()
    out = out.rename(
        columns={
            "Analito 2026": "SQI",
            "Parâmetro": "Parâmetro crítico",
            "Estimativa variação CMA % num": "Variação estimada CMA (%)",
        }
    )
    out["Variação estimada CMA (%)"] = out["Variação estimada CMA (%)"].round(2)
    return out.sort_values("Variação estimada CMA (%)", ascending=sort_ascending).head(n)


def top_changed_parameters_by_source(comp, fonte, n=5, considered_only=True):
    df = comp.copy()
    df = df[df["Fonte"].astype(str).str.lower() == fonte.lower()]
    if considered_only:
        df = df[df.apply(lambda r: changed(r) and not is_density_submaterial(r), axis=1)].copy()
    else:
        df = df[df.apply(raw_changed, axis=1)].copy()
    if df.empty:
        return pd.DataFrame(columns=["Parâmetro", "Nº de SQIs afetadas"])
    out = (
        df.groupby("Parâmetro", dropna=False)
        .agg(**{"Nº de SQIs afetadas": ("CAS", "nunique")})
        .reset_index()
        .sort_values("Nº de SQIs afetadas", ascending=False)
        .head(n)
    )
    return out


def density_change_summary(comp):
    dens = comp[(comp["Parâmetro"].astype(str) == DENSITY_PARAM) & (comp.apply(changed, axis=1))].copy()
    if dens.empty:
        return pd.DataFrame(columns=["Indicador", "Resultado"])
    dens["abs_var"] = pd.to_numeric(dens["Variação % Parâmetro"], errors="coerce").abs()
    total = dens["CAS"].nunique()
    ignored = dens[dens["abs_var"] < DENSITY_MATERIALITY_THRESHOLD]["CAS"].nunique()
    considered = dens[dens["abs_var"] >= DENSITY_MATERIALITY_THRESHOLD]["CAS"].nunique()
    max_abs = dens["abs_var"].max()
    return pd.DataFrame(
        [
            ["SQIs com alteração bruta de densidade", total],
            ["SQIs com densidade < 5%, ignoradas na classificação", ignored],
            ["SQIs com densidade ≥ 5%, consideradas", considered],
            ["Maior |Δ| de densidade observado", fmt_pct_from_decimal(max_abs, 2).replace("+", "")],
        ],
        columns=["Indicador", "Resultado"],
    )


def density_material_sqi_table(comp, master):
    dens = comp[(comp["Parâmetro"].astype(str) == DENSITY_PARAM) & (comp.apply(changed, axis=1))].copy()
    dens["abs_var"] = pd.to_numeric(dens["Variação % Parâmetro"], errors="coerce").abs()
    dens = dens[dens["abs_var"] >= DENSITY_MATERIALITY_THRESHOLD]
    if dens.empty:
        return pd.DataFrame(columns=["CAS", "SQI", "Classe", "Densidade 2023", "Densidade 2026", "Variação da densidade"])
    class_map = master[["CAS", "Classe"]].drop_duplicates().astype({"CAS": str})
    dens["CAS"] = dens["CAS"].astype(str)
    dens = dens.merge(class_map, on="CAS", how="left")
    out = dens[["CAS", "Analito 2026", "Classe", "Valor 2023", "Valor 2026", "Variação % Parâmetro"]].copy()
    out = out.rename(columns={"Analito 2026": "SQI", "Valor 2023": "Densidade 2023", "Valor 2026": "Densidade 2026"})
    out["Variação da densidade"] = out["Variação % Parâmetro"].apply(lambda x: fmt_pct_from_decimal(x, 2))
    return out.drop(columns=["Variação % Parâmetro"])


def group_rows_by_cas(master, cas_list, interpretation_map=None):
    df = master[master["CAS"].astype(str).isin([str(c) for c in cas_list])].copy()
    if df.empty:
        return pd.DataFrame(columns=["Composto", "Classe", "Interpretação"])
    order = {str(c): i for i, c in enumerate(cas_list)}
    df["_ordem"] = df["CAS"].astype(str).map(order)
    df = df.sort_values("_ordem")
    out = df[["Composto", "Classe", "Motivo da classe", "Recomendação"]].copy()
    out["Interpretação"] = out["Motivo da classe"] + " " + out["Recomendação"]
    if interpretation_map:
        for idx, row in out.iterrows():
            comp = str(row["Composto"])
            for key, value in interpretation_map.items():
                if key.lower() in comp.lower():
                    out.at[idx, "Interpretação"] = value
                    break
    return out[["Composto", "Classe", "Interpretação"]]


def group_rows_by_regex(master, pattern):
    df = master[master["Composto"].astype(str).str.contains(pattern, case=False, na=False, regex=True)].copy()
    if df.empty:
        return pd.DataFrame(columns=["Composto", "Classe", "Interpretação"])
    df = df.sort_values(["Classe", "Composto"])
    out = df[["Composto", "Classe", "Motivo da classe", "Recomendação"]].copy()
    out["Interpretação"] = out["Motivo da classe"] + " " + out["Recomendação"]
    return out[["Composto", "Classe", "Interpretação"]]


def build_interpretive_synthesis(selected, cas, cls, detail):
    ch = detail[detail.apply(changed, axis=1)].copy()
    eff = detail[detail.apply(lambda r: changed(r) and not is_density_submaterial(r), axis=1)].copy()
    ignored = detail[detail.apply(lambda r: changed(r) and is_density_submaterial(r), axis=1)].copy()
    class_desc = CLASS_DESCRIPTIONS[cls]["descricao"]

    if ch.empty:
        return (
            f"A comparação entre as versões CETESBRisk v3.03 e v4.00 indica que a SQI {selected} "
            f"(CAS {cas}) não apresentou alterações nos parâmetros avaliados. O composto foi enquadrado na Classe {cls} "
            f"({class_desc}), indicando que a atualização da planilha, isoladamente, não justifica revisão da avaliação de risco para esta SQI. "
            f"A necessidade de revisão deve ser considerada apenas se houver mudança do modelo conceitual, das vias de exposição, do uso da área ou de outras premissas do estudo."
        )

    ch2 = ch.copy()
    ch2["abs_var"] = pd.to_numeric(ch2.get("Variação % Parâmetro", pd.Series(dtype=float)), errors="coerce").abs()
    ch2 = ch2.sort_values("abs_var", ascending=False)
    top_sentence = ""
    if not ch2.empty:
        top = ch2.iloc[0]
        top_sentence = (
            f"A maior variação percentual observada ocorreu em {top.get('Parâmetro', '')}, "
            f"de {fmt_value(top.get('Valor 2023'))} para {fmt_value(top.get('Valor 2026'))} "
            f"({fmt_pct_from_decimal(top.get('Variação % Parâmetro'), 2)}). "
        )

    def params_txt(df):
        vals = df["Parâmetro"].dropna().astype(str).unique().tolist()
        if not vals:
            return ""
        return ", ".join(vals[:5]) + (" entre outros" if len(vals) > 5 else "")

    tox = eff[eff["Parâmetro"].isin(TOX_RELEVANTE)]
    reg = eff[eff["Parâmetro"].isin(REGULATORIO)]
    phys = eff[eff["Parâmetro"].isin(FISQUI_RELEVANTE | {DENSITY_PARAM})]
    tendencies = eff["Tendência esperada da CMA"].dropna().astype(str).str.lower().tolist() if "Tendência esperada da CMA" in eff else []
    cma_reduce = any(("redu" in t or "dimin" in t or "restritiva" in t) for t in tendencies)
    cma_increase = any(("aument" in t or "menos restritiva" in t) for t in tendencies)

    if cls == "A":
        if not ignored.empty and eff.empty:
            text = (
                f"A SQI {selected} (CAS {cas}) foi enquadrada na Classe A, pois a alteração identificada corresponde à densidade, "
                f"cuja variação em módulo ficou abaixo do limiar de materialidade de 5% adotado para esse parâmetro auxiliar. {top_sentence}"
                f"Como a densidade não é tratada como parâmetro dominante para risco ou CMA e a alteração observada ficou abaixo do limiar definido, "
                f"a expectativa técnica é de ausência de impacto material sobre a avaliação de risco, considerando exclusivamente a atualização da planilha."
            )
        else:
            text = (
                f"A SQI {selected} (CAS {cas}) foi enquadrada na Classe A, pois as alterações identificadas não configuram mudança material "
                f"capaz de alterar de forma relevante o risco calculado ou a CMA. {top_sentence}"
                f"A expectativa técnica é de ausência de impacto material sobre a avaliação de risco, considerando exclusivamente a atualização da planilha."
            )
    elif cls == "B":
        text = (
            f"A SQI {selected} (CAS {cas}) foi enquadrada na Classe B, indicando alteração residual ou auxiliar considerada pela regra objetiva. {top_sentence}"
            f"Essa classe não indica, por si só, revisão automática da avaliação de risco, mas pode justificar verificação pontual em cenários muito sensíveis "
            f"ou quando os resultados anteriores estiverem próximos aos critérios de aceitabilidade."
        )
    elif cls == "C":
        p = params_txt(reg)
        text = (
            f"A SQI {selected} (CAS {cas}) foi enquadrada na Classe C, associada a alterações regulatórias ou de potabilidade. "
            f"{'As alterações observadas envolveram principalmente ' + p + '. ' if p else ''}"
            f"Essa classificação não significa, por si só, alteração material do cálculo de risco ou da CMA. "
            f"O principal efeito esperado é sobre o enquadramento regulatório, atualização de tabelas comparativas ou discussão de potabilidade, "
            f"especialmente quando houver avaliação da via de ingestão de água subterrânea."
        )
    elif cls == "D1":
        p = params_txt(tox)
        if cma_reduce:
            trend = "A tendência esperada é de redução da CMA para a rota controlada pelo parâmetro alterado, tornando a avaliação potencialmente mais restritiva. "
        elif cma_increase:
            trend = "A tendência esperada é de aumento da CMA para a rota controlada pelo parâmetro alterado, tornando a avaliação potencialmente menos restritiva. "
        else:
            trend = "O impacto sobre a CMA deve ser avaliado conforme a via de exposição dominante e o parâmetro efetivamente controlador do risco. "
        text = (
            f"A SQI {selected} (CAS {cas}) foi enquadrada na Classe D1, indicando alteração relevante em parâmetro toxicológico ou em atributo diretamente associado ao cálculo de risco. "
            f"{'As alterações envolveram principalmente ' + p + '. ' if p else ''}{top_sentence}{trend}"
            f"Recomenda-se revisão dirigida quando esta SQI for relevante no modelo conceitual, especialmente se os resultados anteriores estiverem próximos aos limites aceitáveis."
        )
    elif cls == "D2":
        p = params_txt(phys)
        text = (
            f"A SQI {selected} (CAS {cas}) foi enquadrada na Classe D2, indicando alteração relevante em parâmetros físico-químicos. "
            f"{'As alterações envolveram principalmente ' + p + '. ' if p else ''}{top_sentence}"
            f"Essas alterações não representam necessariamente mudança direta de toxicidade, mas podem modificar o comportamento ambiental da substância, "
            f"influenciando volatilização, particionamento, transporte, concentração de saturação ou intrusão de vapores. "
            f"O impacto sobre CMA é modelo-dependente e deve ser avaliado principalmente em cenários sensíveis à inalação, fase vapor ou transporte entre fonte e ponto de exposição."
        )
    else:
        text = f"A SQI {selected} (CAS {cas}) foi enquadrada na Classe {cls} ({class_desc})."

    if cls in ["D1", "D2"]:
        text += " Em termos práticos, a atualização deve ser tratada como gatilho de triagem técnica, não como obrigação automática de refazimento integral da avaliação."
    elif cls == "C":
        text += " Assim, a ação prioritária é verificar o enquadramento legal/regulatório aplicável, sem assumir automaticamente necessidade de recálculo do risco."
    else:
        text += " Assim, a atualização da planilha, isoladamente, não indica necessidade de revisão da avaliação de risco para esta SQI."

    return text


def build_audit_dataframe(master):
    cols = [
        "CAS",
        "Composto",
        "Classe anterior app",
        "Classe",
        "Mudou com nova regra?",
        "Nº alterações brutas",
        "Parâmetros alterados brutos",
        "Nº alterações consideradas na classificação",
        "Parâmetros considerados na classificação",
        "Parâmetros ignorados por limiar",
        "Motivo da classe",
    ]
    return master[cols].copy()


comp, incl, remov, alt, impacto = load_data()
master = build_master(comp)
CLASS_COUNTS = master["Classe"].value_counts().reindex(CLASS_ORDER).fillna(0).astype(int).to_dict()
RAW_CHANGE_COUNT = int(comp[comp.apply(raw_changed, axis=1)].shape[0])
CLASSIFICATION_CHANGE_COUNT = int(comp[comp.apply(lambda r: changed(r) and not is_density_submaterial(r), axis=1)].shape[0])
IGNORED_DENSITY_COUNT = int(comp[comp.apply(lambda r: changed(r) and is_density_submaterial(r), axis=1)].shape[0])


def pct_delta_label(new_value, old_value):
    try:
        old_value = float(old_value)
        new_value = float(new_value)
        if old_value == 0:
            return None
        pct = (new_value - old_value) / old_value * 100
        return f"{pct:+.1f}% vs regra anterior".replace(".", ",")
    except Exception:
        return None

st.sidebar.title("🧪 CETESBRisk Explorer")
st.sidebar.markdown("**Desenvolvido por André Souza**  \nEspecialista em GAC  \n[LinkedIn](https://www.linkedin.com/in/andr%C3%A9-souza-63539517)")
st.sidebar.divider()
page = st.sidebar.radio(
    "Navegação",
    ["Início", "Dashboard geral", "Pesquisa SQI e impacto CMA", "Grupos prioritários", "Alterações em fórmulas", "Highlights Manual CETESB", "Glossário Técnico", "Downloads e notas"],
)
st.sidebar.caption("v2.1 · regra objetiva + auditoria de fórmulas")

if page == "Início":
    st.title("🧪 CETESBRisk Explorer")
    st.subheader("Comparador técnico CETESBRisk v3.03 (2023) × v4.00 (2026)")
    st.markdown(
        """## 1. O que é esta ferramenta?
O **CETESBRisk Explorer** é uma ferramenta de apoio à interpretação técnica das alterações introduzidas entre as versões **CETESBRisk v3.03 (2023)** e **CETESBRisk v4.00 (2026)**.

A ferramenta **não recalcula avaliações de risco** e **não substitui a análise crítica do profissional responsável**. Seu objetivo é identificar alterações que possam justificar revisão dirigida de avaliações de risco previamente elaboradas."""
    )
    st.info("Mensagem-chave: a atualização da v4.00 não implica revisão automática de todas as avaliações de risco. A recomendação é realizar triagem dirigida por SQI, via de exposição e sensibilidade do cenário.")
    st.info("Nota sobre a CETESBRisk v4.01: em 01/06/2026, a CETESB publicou a v4.01 com ajuste de fórmulas na aba EXP para as planilhas Trabalhador Comercial/Industrial e Trabalhador de Obra Civil. A classificação deste aplicativo permanece baseada na comparação v3.03 × v4.00 das bases FisQui e FatTox. Para uso quantitativo oficial em avaliações de risco e cálculo de CMA, recomenda-se sempre utilizar a versão mais recente da planilha CETESB.")

    st.markdown(
        """## 2. O que está sendo avaliado?
Foram comparados os bancos de dados internos das duas versões da planilha CETESBRisk, com foco nas bases **FisQui** e **FatTox**.

A **FatTox** reúne parâmetros toxicológicos, como RfDo, RfCi, SFo e IUR. Alterações nesses parâmetros podem repercutir diretamente no cálculo de risco e nas CMAs.

A **FisQui** reúne parâmetros físico-químicos, como Henry, solubilidade, Koc, Kd, pressão de vapor e Csat. Esses parâmetros não alteram necessariamente a toxicidade, mas podem influenciar volatilização, transporte, particionamento e intrusão de vapores."""
    )
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("""### Base Físico-Química — FisQui
- solubilidade;
- constante de Henry;
- pressão de vapor;
- Koc e Kd;
- difusão;
- Csat;
- particionamento solo-água-ar.""")
    with c2:
        st.markdown("""### Base Toxicológica — FatTox
- RfDo;
- RfCi;
- SFo;
- IUR;
- carcinogenicidade;
- mutagenicidade;
- MCL e potabilidade.""")

    st.markdown("## 3. Escopo da avaliação")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("SQIs comuns comparadas", "819")
    c2.metric("Comparações individuais", "29.484")
    c3.metric("Novas SQIs incluídas", "59")
    c4.metric("SQIs removidas", "1")
    c5.metric("Versões avaliadas", "2023 × 2026")
    st.caption("A comparação detalhada foi realizada para 819 SQIs presentes em ambas as versões. Adicionalmente, foram identificadas 59 novas substâncias incluídas na v4.00 e 1 substância removida em relação à v3.03.")

    st.markdown("## 4. Como os compostos foram classificados?")
    st.markdown(
        """A classificação A/B/C/D1/D2 foi calculada por regra objetiva e hierárquica. A classe atribuída reflete a alteração considerada mais relevante para cada SQI, seguindo a hierarquia **D1 > D2 > C > B > A**. Alterações de densidade são tratadas como parâmetro auxiliar e somente são consideradas quando a variação em módulo for igual ou superior a 5%."""
    )
    classe_df = pd.DataFrame(
        {
            "Classe": ["A", "B", "C", "D1", "D2"],
            "Regra objetiva": [
                "Sem alteração considerada para classificação; inclui densidade isolada com |Δ| < 5%.",
                "Alteração residual/auxiliar considerada; inclui densidade isolada com |Δ| ≥ 5%.",
                "Alteração em MCL ou Potabilidade, se não houver D1 ou D2.",
                "Qualquer alteração real em parâmetro toxicológico relevante.",
                "Alteração físico-química relevante, exceto densidade, com |Δ| ≥ 5%.",
            ],
            "Implicação prática": [
                "Não indica revisão automática da avaliação de risco.",
                "Revisar apenas em cenários muito sensíveis.",
                "Pode exigir revisão regulatória, mas não necessariamente recálculo do risco.",
                "Pode justificar revisão dirigida ou recomendada da avaliação de risco.",
                "Pode justificar revisão dirigida em cenários sensíveis à volatilização, transporte ou intrusão de vapores.",
            ],
        }
    )
    st.dataframe(classe_df, use_container_width=True, hide_index=True)

    st.markdown("## 5. Explicação detalhada das classes")
    with st.expander("Classe A — Sem alteração material relevante"):
        st.markdown(
            """A Classe A reúne substâncias para as quais não foram identificadas alterações materiais nos parâmetros com potencial de modificar de forma relevante o risco calculado, a Concentração Máxima Aceitável (CMA) ou o comportamento ambiental da substância. De forma objetiva, esta classe também inclui casos em que a única alteração identificada foi densidade com variação em módulo inferior a 5%, por se tratar de parâmetro auxiliar com baixa relevância isolada para risco ou CMA.

**Interpretação prática:** a atualização da CETESBRisk não representa gatilho para reavaliação dessas substâncias. Eventual revisão deve decorrer de mudanças no modelo conceitual, nas vias de exposição, no uso da área ou em novos dados ambientais.

**Exemplos identificados nesta comparação:** benzeno, tolueno, etilbenzeno, xilenos e grande parte dos compostos clássicos de petróleo."""
        )
    with st.expander("Classe B — Alterações menores"):
        st.markdown(
            """A Classe B reúne substâncias que apresentaram alterações paramétricas de pequena magnitude, normalmente associadas a refinamentos de valores, ajustes metodológicos ou correções sem potencial relevante de alterar a interpretação do risco. Objetivamente, inclui alterações isoladas de densidade quando a variação em módulo for igual ou superior a 5%, além de outras alterações consideradas pela regra objetiva que não se enquadram como D1, D2 ou C.

Embora não sejam alterações consideradas materiais, podem justificar uma verificação pontual em situações específicas, especialmente quando a substância controla o resultado da avaliação ou quando os riscos calculados anteriormente se encontram muito próximos dos critérios de aceitabilidade.

**Interpretação prática:** revisão geralmente não necessária, mas pode ser recomendável em estudos particularmente sensíveis.

**Exemplos identificados nesta comparação:** hidrazina, cianogênio, cianeto de hidrogênio, amônia, sulfeto de hidrogênio e fosfina."""
        )
    with st.expander("Classe C — Alterações regulatórias e de potabilidade"):
        st.warning("A Classe C é uma das mais sujeitas a interpretação equivocada: alteração regulatória não significa, necessariamente, alteração do risco calculado.")
        st.markdown(
            """A Classe C reúne SQIs com alteração em **MCL** ou **Potabilidade**, desde que não haja alteração toxicológica relevante (D1) ou físico-química relevante (D2) para a mesma SQI.

Esses parâmetros têm importância para enquadramento regulatório e gestão da via de ingestão de água subterrânea, mas não são necessariamente controladores diretos do cálculo de risco na planilha.

**Interpretação prática:** a prioridade é verificar o enquadramento legal/regulatório aplicável, sem assumir automaticamente a necessidade de recálculo do risco."""
        )
    with st.expander("Classe D1 — Alterações toxicológicas relevantes"):
        st.markdown(
            """A Classe D1 reúne substâncias que apresentaram alterações, independentemente da magnitude, em parâmetros diretamente associados à toxicidade e ao cálculo de risco, incluindo dose de referência oral (RfDo), concentração de referência por inalação (RfCi), fator de slope oral (SFo), unidade de risco por inalação (IUR), classificação carcinogênica e mutagenicidade.

Essas alterações possuem potencial direto de repercutir nos resultados da avaliação de risco e, consequentemente, nas CMAs derivadas. Dependendo do parâmetro alterado, a atualização pode tornar os critérios mais restritivos ou menos restritivos, alterando a interpretação de estudos anteriormente desenvolvidos.

**Interpretação prática:** substâncias enquadradas nesta classe devem ser consideradas prioritárias em qualquer triagem para revisão de avaliações de risco, especialmente quando forem relevantes no modelo conceitual ou quando a avaliação anterior estiver próxima dos limites de aceitabilidade.

**Exemplos identificados nesta comparação:** cloreto de vinila, 1,1-DCE, PFOS, PFOA e cromo VI."""
        )
    with st.expander("Classe D2 — Alterações físico-químicas relevantes"):
        st.markdown(
            """A Classe D2 reúne substâncias que apresentaram alterações materiais em parâmetros físico-químicos capazes de influenciar o comportamento ambiental da substância e sua modelagem de exposição, sem alterações em parâmetros toxicológicos.

Entre os parâmetros potencialmente envolvidos estão constante de Henry, pressão de vapor, solubilidade, coeficiente de partição carbono orgânico-água (Koc), coeficiente de distribuição solo-água (Kd), log Kow e concentração de saturação (Csat). A densidade é exceção: mesmo quando sua variação for igual ou superior a 5%, ela não gera D2 isoladamente.

Diferentemente da Classe D1, essas alterações não representam necessariamente mudança de toxicidade. Seu impacto depende do modelo conceitual e da via de exposição considerada, sendo especialmente relevante para avaliações envolvendo volatilização, intrusão de vapores, transporte em fase vapor, particionamento solo-água-ar e modelagem de transporte.

**Interpretação prática:** a necessidade de revisão depende da importância desses processos no cenário avaliado.

**Exemplos identificados nesta comparação:** TPH alifático leve C5-C8 e HFPO-DA/GenX."""
        )

    st.info("Importante: as classes representam uma ferramenta de triagem técnica para priorização de revisões. A classificação não substitui a análise crítica do modelo conceitual da área, das vias de exposição, das substâncias efetivamente presentes e dos resultados históricos disponíveis.")

    st.markdown("## 6. Fluxo decisório: quando revisar uma avaliação anterior?")
    st.markdown(
        """**Perguntas de triagem recomendadas:**
1. O modelo conceitual possui SQIs da Classe D1, como PFAS, cloreto de vinila ou 1,1-DCE?
2. Há SQIs da Classe D2 associadas a volatilização, intrusão de vapores ou transporte em fase vapor?
3. A avaliação anterior estava próxima da margem de aceitabilidade, como HQ ≈ 1 ou risco ≈ 1E-05?
4. A avaliação baseou-se predominantemente em frações genéricas de TPH em vez de compostos individuais de petróleo?

Se a resposta for **sim** para uma ou mais perguntas, recomenda-se **revisão dirigida**. Se todas forem **não**, a atualização da planilha, isoladamente, não indica revisão automática."""
    )


    st.markdown("## 7. Alterações em fórmulas")
    st.markdown(
        """Além da comparação dos parâmetros FisQui e FatTox, foram registradas correções pontuais de fórmulas identificadas entre a CETESBRisk V3.03 e a V4.01. Essas correções não alteram, por si só, a classificação global das SQIs em A/B/C/D1/D2, pois essa classificação está associada às alterações nos parâmetros físico-químicos, toxicológicos ou regulatórios.

As correções de fórmulas devem ser interpretadas como uma camada própria de auditoria da ferramenta. Quando a correção atingir uma célula efetivamente utilizada no cálculo de risco ou CMA, recomenda-se avaliar pontualmente se estudos anteriores dependiam daquela rota, daquele módulo e daquela posição de SQI. A descrição detalhada dessas correções está disponível na aba **Alterações em fórmulas**."""
    )

    st.markdown("## 8. Resultado da classificação")
    r1, r2, r3, r4, r5 = st.columns(5)
    r1.metric("Classe A", str(CLASS_COUNTS.get("A", 0)), delta=pct_delta_label(CLASS_COUNTS.get("A", 0), LEGACY_CLASS_COUNTS["A"]))
    r2.metric("Classe B", str(CLASS_COUNTS.get("B", 0)), delta=pct_delta_label(CLASS_COUNTS.get("B", 0), LEGACY_CLASS_COUNTS["B"]))
    r3.metric("Classe C", str(CLASS_COUNTS.get("C", 0)), delta=pct_delta_label(CLASS_COUNTS.get("C", 0), LEGACY_CLASS_COUNTS["C"]))
    r4.metric("Classe D1", str(CLASS_COUNTS.get("D1", 0)), delta=pct_delta_label(CLASS_COUNTS.get("D1", 0), LEGACY_CLASS_COUNTS["D1"]))
    r5.metric("Classe D2", str(CLASS_COUNTS.get("D2", 0)), delta=pct_delta_label(CLASS_COUNTS.get("D2", 0), LEGACY_CLASS_COUNTS["D2"]))

    st.markdown("## 9. PFAS e compostos correlatos")
    st.markdown("A v4.00 ampliou significativamente a capacidade de avaliação de PFAS e compostos correlatos. O ganho não é apenas numérico: vários compostos passaram a ter maior preenchimento de parâmetros físico-químicos e toxicológicos, reduzindo lacunas que limitavam análises anteriores.")

    st.markdown("## 10. Limitações da classificação")
    st.markdown("A classificação apresentada nesta ferramenta representa uma triagem técnica baseada nas alterações identificadas entre as versões da CETESBRisk. A decisão de revisar ou não uma avaliação de risco deve considerar adicionalmente o modelo conceitual da área, as vias de exposição completas, as SQIs efetivamente presentes, os resultados históricos e a proximidade aos critérios de aceitabilidade.")

    st.markdown("## 11. Conclusão executiva")
    st.markdown(
        f"""A versão **CETESBRisk v4.00** representa um avanço importante na atualização das bases físico-químicas e toxicológicas, bem como na transparência da ferramenta, especialmente pela publicação do **Manual do Usuário**.

Com a regra objetiva adotada, **{CLASS_COUNTS.get('A', 0)} de 819 SQIs** foram classificadas como Classe A. Isso indica que a maior parte dos compostos avaliados não apresentou alteração considerada material para fins de classificação.

A abordagem recomendada é uma **triagem dirigida**, priorizando substâncias classificadas como **D1**, cenários sensíveis à volatilização e intrusão de vapores associados à **D2**, além de grupos de atenção como **PFAS**, **cloreto de vinila**, **1,1-DCE** e **TPH alifático leve C5-C8**.

Alterações isoladas de densidade inferiores a 5% foram tratadas como não materiais para classificação e, por isso, não devem inflar artificialmente a Classe B."""
    )

elif page == "Dashboard geral":
    st.title("Dashboard geral")
    st.write("Exploração quantitativa dos pontos de dados comparados entre as versões v3.03 e v4.00.")

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Pontos de dados analisados", "29.484")
    c2.metric("Alterações brutas", f"{RAW_CHANGE_COUNT:,}".replace(",", "."))
    c3.metric("Alterações consideradas", str(CLASSIFICATION_CHANGE_COUNT), delta=f"{(CLASSIFICATION_CHANGE_COUNT - RAW_CHANGE_COUNT) / RAW_CHANGE_COUNT * 100:+.1f}% vs alterações brutas".replace(".", ","))
    c4.metric("Densidade <5% ignorada", str(IGNORED_DENSITY_COUNT))
    c5.metric("Sem alteração", "16.565")
    st.caption("Alterações brutas incluem diferenças identificadas na base. Alterações consideradas excluem densidade com variação em módulo inferior a 5%, conforme regra objetiva de classificação.")

    st.subheader("Distribuição consolidada dos compostos por classe de alteração")
    chart_df = pd.DataFrame({"Classe": CLASS_ORDER, "N compostos": [CLASS_COUNTS.get(k, 0) for k in CLASS_ORDER]})
    fig = px.bar(chart_df, x="Classe", y="N compostos", color="Classe", text="N compostos", color_discrete_map=CLASS_COLORS, title="Compostos por classe de alteração")
    fig.update_traces(textposition="outside", hovertemplate="<b>Classe %{x}</b><br>Nº de compostos: %{y}<extra></extra>")
    fig.update_layout(yaxis_title="Nº de compostos", xaxis_title="Classe")
    st.plotly_chart(fig, use_container_width=True)

    st.subheader("Composição dos pontos de dados avaliados")
    flow_df = pd.DataFrame({"Categoria": ["Alterações brutas", "Alterações consideradas", "Densidade <5% ignorada", "Sem alteração", "Sem valor/vazios"], "N": [RAW_CHANGE_COUNT, CLASSIFICATION_CHANGE_COUNT, IGNORED_DENSITY_COUNT, 16565, 11910]})
    fig2 = px.bar(flow_df, x="Categoria", y="N", color="Categoria", text="N", title="Status dos pontos de dados avaliados")
    fig2.update_traces(textposition="outside", hovertemplate="<b>%{x}</b><br>Nº de pontos: %{y}<extra></extra>")
    st.plotly_chart(fig2, use_container_width=True)

    st.divider()
    st.subheader("SQIs incluídas e removidas na v4.00")
    st.markdown("A v4.00 incluiu 59 novas Substâncias Químicas de Interesse (SQIs) e removeu 1 registro em relação à v3.03. Essas listas são apresentadas separadamente, não compondo as classes A/B/C/D1/D2 das 819 SQIs comuns.")
    tab_inc, tab_rem = st.tabs(["SQIs incluídas", "SQIs removidas"])
    with tab_inc:
        st.dataframe(incl, use_container_width=True, hide_index=True)
        st.download_button("Baixar SQIs incluídas", data=incl.to_csv(index=False).encode("utf-8-sig"), file_name="sqis_incluidas_v4.csv", mime="text/csv")
    with tab_rem:
        st.dataframe(remov, use_container_width=True, hide_index=True)
        st.download_button("Baixar SQIs removidas", data=remov.to_csv(index=False).encode("utf-8-sig"), file_name="sqis_removidas_v4.csv", mime="text/csv")

    st.divider()
    st.subheader("Rankings de atenção técnica")
    st.markdown("Os rankings usam apenas alterações consideradas pela regra objetiva de classificação. Portanto, densidade com |Δ| < 5% não domina artificialmente a leitura.")
    tab_fisqui, tab_fattox, tab_cma_down, tab_cma_up = st.tabs(["Top 5 FisQui", "Top 5 FatTox", "Top 5 CMA mais restritiva", "Top 5 CMA menos restritiva"])
    with tab_fisqui:
        st.caption("Critério: número de parâmetros FisQui considerados; desempate pela maior variação percentual absoluta.")
        st.dataframe(top_sqi_by_parameter_changes(comp, "FisQui", 5, considered_only=True), use_container_width=True, hide_index=True)
    with tab_fattox:
        st.caption("Critério: número de parâmetros FatTox considerados; desempate pela maior variação percentual absoluta.")
        st.dataframe(top_sqi_by_parameter_changes(comp, "FatTox", 5, considered_only=True), use_container_width=True, hide_index=True)
    with tab_cma_down:
        st.caption("Critério: maior redução percentual estimada de CMA em parâmetros com relação direta/inversa simplificada.")
        st.dataframe(top_cma_impact(comp, "reduz", 5), use_container_width=True, hide_index=True)
    with tab_cma_up:
        st.caption("Critério: maior aumento percentual estimado de CMA em parâmetros com relação direta/inversa simplificada.")
        st.dataframe(top_cma_impact(comp, "aumenta", 5), use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("Parâmetros mais frequentemente alterados")
    st.markdown("A comparação abaixo separa a leitura bruta da base da leitura usada para classificação.")
    tab_raw, tab_cons = st.tabs(["Alterações brutas", "Alterações consideradas para classificação"])
    with tab_raw:
        pcol1, pcol2 = st.columns(2)
        with pcol1:
            st.markdown("**Top 5 parâmetros FisQui mais alterados — bruto**")
            st.dataframe(top_changed_parameters_by_source(comp, "FisQui", 5, considered_only=False), use_container_width=True, hide_index=True)
        with pcol2:
            st.markdown("**Top 5 parâmetros FatTox mais alterados — bruto**")
            st.dataframe(top_changed_parameters_by_source(comp, "FatTox", 5, considered_only=False), use_container_width=True, hide_index=True)
    with tab_cons:
        pcol1, pcol2 = st.columns(2)
        with pcol1:
            st.markdown("**Top 5 parâmetros FisQui considerados**")
            st.dataframe(top_changed_parameters_by_source(comp, "FisQui", 5, considered_only=True), use_container_width=True, hide_index=True)
        with pcol2:
            st.markdown("**Top 5 parâmetros FatTox considerados**")
            st.dataframe(top_changed_parameters_by_source(comp, "FatTox", 5, considered_only=True), use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("Tabela consolidada por composto")
    classes = st.multiselect("Filtrar classes", CLASS_ORDER, default=CLASS_ORDER)
    filtered = master[master["Classe"].isin(classes)].copy()
    st.dataframe(filtered, use_container_width=True, hide_index=True)
    st.download_button("Baixar tabela filtrada", data=filtered.to_csv(index=False).encode("utf-8-sig"), file_name="cetesbrisk_master_filtrado.csv", mime="text/csv")

    st.markdown("### Entendendo como a densidade foi tratada")
    st.markdown("A densidade é tratada como parâmetro físico-químico auxiliar. Ela não classifica uma SQI como D2 isoladamente. Para fins de classificação, alterações de densidade somente são consideradas quando a variação em módulo for igual ou superior a 5%. Alterações de densidade inferiores a 5%, quando isoladas, são tratadas como não materiais e não alteram a classe da SQI. Quando houver alteração de densidade associada a alteração toxicológica, físico-química relevante ou regulatória, prevalece a hierarquia D1 > D2 > C > B > A.")
    st.dataframe(density_change_summary(comp), use_container_width=True, hide_index=True)
    dens_mat = density_material_sqi_table(comp, master)
    if not dens_mat.empty:
        st.markdown("**SQIs com alteração de densidade ≥ 5%**")
        st.dataframe(dens_mat, use_container_width=True, hide_index=True)

elif page == "Pesquisa SQI e impacto CMA":
    st.title("Pesquisa SQI e impacto potencial na CMA")
    search_df = master.sort_values(["Composto", "CAS"]).dropna(subset=["Composto"]).copy()
    search_df["rotulo_busca"] = search_df.apply(lambda r: f"{r['Composto']} | CAS: {r['CAS']}", axis=1)
    options = search_df["rotulo_busca"].tolist()
    default_idx = 0
    benz = search_df[search_df["Composto"].astype(str).str.lower().eq("benzene")]
    if not benz.empty:
        default_label = benz.iloc[0]["rotulo_busca"]
        if default_label in options:
            default_idx = options.index(default_label)
    selected_label = st.selectbox("Selecione uma substância", options, index=default_idx)
    row = search_df[search_df["rotulo_busca"] == selected_label].iloc[0]
    selected = row["Composto"]
    cas = row["CAS"]
    cls = row["Classe"]
    st.header(selected)
    st.markdown(badge(cls), unsafe_allow_html=True)
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("CAS", str(cas))
    m2.metric("Classe", cls)
    m3.metric("Alterações brutas", int(row["Nº alterações brutas"]))
    m4.metric("Alterações consideradas", int(row["Nº alterações consideradas na classificação"]))
    st.subheader("Interpretação")
    st.write(row["Recomendação"])
    detail = comp[(comp["CAS"].astype(str) == str(cas))].copy()
    detail_flags = add_classification_flags(detail)
    show_cols = [
        "Fonte",
        "Parâmetro",
        "Valor 2023",
        "Valor 2026",
        "Status",
        "Variação % Parâmetro",
        "Alteração considerada na classificação?",
        "Motivo de uso na classificação",
        "Tendência esperada da CMA",
        "Estimativa variação CMA %",
        "Observação técnica",
    ]
    st.subheader("Comparativo por parâmetro")
    st.dataframe(detail_flags[show_cols], use_container_width=True, hide_index=True)
    synthesis = build_interpretive_synthesis(selected, cas, cls, detail)
    st.subheader("Síntese técnica estruturada")
    st.text_area("Texto técnico para relatório", synthesis, height=220)
    html = report_html(selected, cas, cls, detail_flags[show_cols], synthesis)
    st.download_button("Baixar relatório HTML desta SQI", data=html.encode("utf-8"), file_name=f"relatorio_cetesbrisk_{str(selected).replace(' ', '_')}.html", mime="text/html")

elif page == "Grupos prioritários":
    st.title("Grupos prioritários")
    st.write("Síntese interpretativa para grupos frequentemente relevantes em avaliações de risco e GAC.")
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["BTEX", "Etenos clorados", "TPH", "PFAS", "Metais"])

    with tab1:
        st.subheader("BTEX")
        st.markdown("Nesta aba são apresentados apenas os compostos clássicos do grupo BTEX: benzeno, tolueno, etilbenzeno e xilenos. Pela nova regra, alterações isoladas de densidade inferiores a 5% não alteram a classificação.")
        btex_cas = ["71-43-2", "108-88-3", "100-41-4", "95-47-6", "108-38-3", "106-42-3"]
        st.dataframe(group_rows_by_cas(master, btex_cas), use_container_width=True, hide_index=True)

    with tab2:
        st.subheader("Etenos clorados")
        chlorinated_cas = ["127-18-4", "79-01-6", "156-59-2", "156-60-5", "75-35-4", "75-01-4"]
        st.dataframe(group_rows_by_cas(master, chlorinated_cas), use_container_width=True, hide_index=True)

    with tab3:
        st.subheader("TPH")
        st.markdown("Para TPH, a principal alteração material identificada está associada à fração alifática leve C5-C8, enquadrada como D2 por alterações físico-químicas relevantes. A tabela é derivada dinamicamente da classificação final.")
        tph_pattern = "total petroleum hydrocarbons|aliphatic hydrocarbon|aromatic.*c9|aromatic.*c10|aliphatic low|aliphatic medium|aliphatic high"
        st.dataframe(group_rows_by_regex(master, tph_pattern), use_container_width=True, hide_index=True)
        st.info("As faixas de TPH possuem caráter complementar e não substituem a avaliação de compostos individuais de petróleo quando identificados.")

    with tab4:
        st.subheader("PFAS")
        st.markdown("A v4.00 ampliou a representação de PFAS e compostos correlatos. O ganho não é apenas a inclusão de novas substâncias, mas também o preenchimento de lacunas físico-químicas e toxicológicas que limitavam avaliações anteriores.")
        pfas_pattern = "perfluoro|fluorotelomer|hfpo|genx|pfos|pfoa|pfbs|pfhxs|pfna|pfda|pfba|pfhpa|pfhxa"
        pfas = group_rows_by_regex(master, pfas_pattern)
        st.dataframe(pfas, use_container_width=True, hide_index=True)

    with tab5:
        st.subheader("Metais")
        st.markdown("Para metais, a interpretação deve considerar forma química, Kd, pH, especiação e particionamento solo-água. Mudanças nesses parâmetros podem afetar mobilidade, lixiviação e transporte, mas o impacto é fortemente dependente do modelo conceitual e das condições hidrogeoquímicas.")
        metais_pattern = "arsenic|barium|beryllium|boron|cadmium|chromium|cobalt|copper|lead|manganese|mercury|molybdenum|nickel|selenium|silver|thallium|vanadium|zinc|antimony|aluminum|iron|lithium|tin|titanium|uranium|lanthanum"
        metais = group_rows_by_regex(master, metais_pattern)
        if metais.empty:
            st.warning("Nenhum metal foi localizado na tabela consolidada com o critério de busca atual.")
        else:
            st.dataframe(metais, use_container_width=True, hide_index=True)
            st.caption("A lista acima é baseada na identificação nominal dos compostos metálicos na tabela consolidada. A interpretação deve considerar a forma química específica de cada SQI.")


elif page == "Alterações em fórmulas":
    st.title("Alterações em fórmulas")
    st.write("Camada de auditoria das correções pontuais de fórmulas identificadas entre versões da CETESBRisk.")

    st.info("Esta aba não classifica SQIs. Ela documenta correções pontuais de fórmulas identificadas entre versões da CETESBRisk e auxilia a triagem de estudos anteriores que possam ter utilizado células, rotas ou posições afetadas.")

    st.markdown(
        """As alterações de fórmula identificadas entre a **CETESBRisk V3.03** e a **V4.01** devem ser interpretadas como correções pontuais da ferramenta, e não como mudança metodológica generalizada na forma de cálculo do risco ou da Concentração Máxima Aceitável (CMA).

Essas correções não alteram, por si só, a classificação global das SQIs em **A/B/C/D1/D2**, pois essa classificação está associada às alterações nos parâmetros físico-químicos, toxicológicos ou regulatórios. No entanto, podem afetar resultados específicos quando a versão anterior da planilha tiver sido utilizada em uma rota, módulo, célula ou posição de SQI diretamente afetada.

Assim, as alterações em fórmulas devem ser tratadas como uma camada própria de auditoria da ferramenta. Quando a correção atingir uma célula efetivamente utilizada no cálculo de risco ou CMA, recomenda-se avaliar pontualmente se estudos anteriores dependiam daquela rota, daquele módulo e daquela posição de SQI."""
    )

    formula_changes = [
        {
            "Correção": "Correção do fator numérico em CMA Solo Cr",
            "Aba": "CMA Solo Cr",
            "Células afetadas": "X13 e X14",
            "Tipo de alteração": "Correção de fator numérico",
            "Alteração principal": "9^3 → 10^3",
            "Módulo afetado": "CMA para solo, risco carcinogênico e não carcinogênico associado à lixiviação para água subterrânea",
            "Impacto potencial": "Mantidos os demais termos constantes, o termo corrigido eleva o fator multiplicativo de 729 para 1000, ou seja, aumento relativo de aproximadamente 37,2% nas células específicas.",
            "Interpretação": "Correção pontual de fórmula, sem mudança metodológica ampla. Pode justificar reavaliação pontual se o usuário tiver utilizado a V3.03 em cenário no qual essa célula, rota e posição da SQI estavam ativas.",
        },
        {
            "Correção": "Correção de referência nomeada em CMA AR Cr-Ad",
            "Aba": "CMA AR Cr-Ad",
            "Células afetadas": "Q16",
            "Tipo de alteração": "Correção de referência nomeada",
            "Alteração principal": "Cf.GasSolo.14 → Cf.GasSolo.4",
            "Módulo afetado": "CMA para ar/vapores, cenário com concentração medida em gás do solo",
            "Impacto potencial": "A V3.03 podia comparar a CMA da SQI posicionada no bloco 4 contra a concentração de gás do solo da SQI 14, gerando pareamento incorreto entre concentração medida e resultado calculado.",
            "Interpretação": "Correção pontual de referência. Pode justificar reavaliação pontual quando o usuário tiver utilizado o módulo de ar/vapores da V3.03 com concentração de gás do solo preenchida na posição afetada.",
        },
        {
            "Correção": "Remoção de EV.c em FI",
            "Aba": "FI",
            "Células afetadas": "I8",
            "Tipo de alteração": "Correção conceitual de fórmula",
            "Alteração principal": "Remoção de EV.c",
            "Módulo afetado": "Fator de ingresso para inalação de partículas, criança",
            "Impacto potencial": "O impacto numérico padrão tende a ser nulo, pois EV.c possui valor padrão igual a 1. Ainda assim, a remoção corrige a estrutura lógica da fórmula.",
            "Interpretação": "Correção conceitual pontual. A variável EV.c corresponde à frequência de eventos para contato dérmico com solo/água, enquanto FI!I8 está associada à rota de inalação de partículas.",
        },
    ]

    summary_df = pd.DataFrame(formula_changes)
    st.subheader("Resumo das correções identificadas")
    st.dataframe(summary_df[["Correção", "Aba", "Células afetadas", "Tipo de alteração", "Alteração principal", "Impacto potencial"]], use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("Detalhamento técnico")

    with st.expander("1. Correção do fator numérico em CMA Solo Cr", expanded=True):
        st.markdown("""**Local da alteração**

| Item | Informação |
|---|---|
| Aba | CMA Solo Cr |
| Células afetadas | X13 e X14 |
| Tipo de alteração | Correção de fator numérico |
| Alteração principal | 9^3 → 10^3 |
| Módulo afetado | CMA para solo, risco carcinogênico e não carcinogênico associado à lixiviação para água subterrânea |

**Fórmula na V3.03, célula X13**
```excel
IF(CENÁRIOS!$I$18=TRUE,IF(PM.2="","-",IF(OR(Sf.ig.2="-",Sf.ig.2="",Sf.ig.2=0,FT!Q9="-",DAF.2="-"),"-",(TR/((FT!Q9*DAF.2)*IF.SoloSub.Lix.AS.c.c*Sf.ig.2))*9^3)),"-")
```

**Fórmula na V4.01, célula X13**
```excel
IF(CENÁRIOS!$I$18=TRUE,IF(PM.2="","-",IF(OR(Sf.ig.2="-",Sf.ig.2="",Sf.ig.2=0,FT!Q9="-",DAF.2="-"),"-",(TR/((FT!Q9*DAF.2)*IF.SoloSub.Lix.AS.c.c*Sf.ig.2))*10^3)),"-")
```

**Fórmula na V3.03, célula X14**
```excel
IF(CENÁRIOS!$I$18=TRUE,IF(PM.2="","-",IF(OR(RfD.ig.cn.2="-",RfD.ig.cn.2="",RfD.ig.cn.2=0,FT!Q9="-",DAF.2="-"),"-",((THI*RfD.ig.cn.2)/((FT!Q9*DAF.2)*IF.SoloSub.Lix.AS.c.nc))*9^3)),"-")
```

**Fórmula na V4.01, célula X14**
```excel
IF(CENÁRIOS!$I$18=TRUE,IF(PM.2="","-",IF(OR(RfD.ig.cn.2="-",RfD.ig.cn.2="",RfD.ig.cn.2=0,FT!Q9="-",DAF.2="-"),"-",((THI*RfD.ig.cn.2)/((FT!Q9*DAF.2)*IF.SoloSub.Lix.AS.c.nc))*10^3)),"-")
```

**Interpretação técnica**

Na V3.03, as células X13 e X14 utilizavam o fator 9^3, equivalente a 729. Na V4.01, esse fator foi corrigido para 10^3, equivalente a 1000. Mantidos todos os demais termos constantes, o termo corrigido eleva o fator multiplicativo em aproximadamente 37,2% nessas células específicas.

Essa alteração tem forte característica de correção pontual de fórmula, e não de mudança metodológica ampla da planilha. Ela não altera a classificação global da SQI, mas pode justificar reavaliação pontual se o usuário tiver utilizado a V3.03 em cenário no qual essa célula, rota e posição da SQI estavam ativas.""")

    with st.expander("2. Correção de referência nomeada em CMA AR Cr-Ad"):
        st.markdown("""**Local da alteração**

| Item | Informação |
|---|---|
| Aba | CMA AR Cr-Ad |
| Célula afetada | Q16 |
| Tipo de alteração | Correção de referência nomeada |
| Alteração principal | Cf.GasSolo.14 → Cf.GasSolo.4 |
| Módulo afetado | CMA para ar/vapores, cenário com concentração medida em gás do solo |

**Fórmula na V3.03, célula Q16**
```excel
IF(OR(P16="ND",P16="-"),0,Cf.GasSolo.14/P16)
```

**Fórmula na V4.01, célula Q16**
```excel
IF(OR(P16="ND",P16="-"),0,Cf.GasSolo.4/P16)
```

**Interpretação técnica**

Na V3.03, a célula Q16 buscava a concentração de gás do solo associada ao nome definido Cf.GasSolo.14. Na V4.01, essa referência foi corrigida para Cf.GasSolo.4.

A correção é relevante porque a célula Q16 está no bloco associado à quarta SQI daquela sequência. Portanto, a referência esperada seria Cf.GasSolo.4, e não Cf.GasSolo.14. Na prática, a V3.03 podia comparar a CMA da SQI posicionada no bloco 4 contra a concentração de gás do solo da SQI 14, gerando pareamento incorreto entre concentração medida e resultado calculado.

Essa alteração também é uma correção pontual de fórmula/referência, não uma mudança metodológica generalizada. Ela não transforma uma SQI em D1 ou D2, mas pode justificar reavaliação pontual quando o usuário tiver utilizado o módulo de ar/vapores da V3.03 com concentração de gás do solo preenchida na posição afetada.""")

    with st.expander("3. Remoção de EV.c em FI"):
        st.markdown("""**Local da alteração**

| Item | Informação |
|---|---|
| Aba | FI |
| Célula afetada | I8 |
| Tipo de alteração | Correção conceitual de fórmula |
| Alteração principal | Remoção de EV.c |
| Módulo afetado | Fator de ingresso para inalação de partículas, criança |

**Fórmula na V3.03, célula I8**
```excel
IF(CENÁRIOS!$H$12=TRUE,((IRaamb.c*EF.c*ETs.c*EV.c*ED.c)/(BW.c*ATc.c)),"-")
```

**Fórmula na V4.01, célula I8**
```excel
IF(CENÁRIOS!$H$12=TRUE,((IRaamb.c*EF.c*ETs.c*ED.c)/(BW.c*ATc.c)),"-")
```

**Interpretação técnica**

Na V3.03, a fórmula de FI!I8 incluía a variável EV.c. Essa variável corresponde à frequência de eventos para contato dérmico com solo/água. No entanto, a célula FI!I8 está associada à rota de inalação de partículas, não à rota de contato dérmico.

Na V4.01, EV.c foi removida da fórmula, tornando a equação mais coerente com a rota de exposição representada. Do ponto de vista numérico, o impacto padrão tende a ser nulo, porque EV.c possui valor padrão igual a 1. Portanto, multiplicar ou não por EV.c não altera o resultado padrão. Ainda assim, a remoção é tecnicamente relevante porque corrige a estrutura lógica da fórmula.

Essa alteração deve ser interpretada como correção conceitual pontual, não como mudança metodológica ampla. Ela não altera a classificação global da SQI, mas deve constar na aba de auditoria de alterações em fórmulas.""")

elif page == "Highlights Manual CETESB":
    st.title("Highlights do Manual do Usuário da CETESBRisk v4.00")
    st.write("Principais ganhos técnicos e pontos de atenção associados à primeira publicação do Manual do Usuário da planilha CETESB.")
    st.markdown("""## De planilha a documento de governança
O maior ganho do Manual do Usuário é transformar a planilha em uma referência documentada de uso, reduzindo ambiguidades e divergências interpretativas entre consultorias, clientes e órgão ambiental.""")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("""### Padronização
O manual consolida orientações de preenchimento, uso das abas e interpretação das saídas, funcionando como referência oficial para treinamento e revisão técnica.""")
    with col2:
        st.markdown("""### Rastreabilidade
Explicita a separação entre abas de entrada, bases internas, cálculos intermediários e resultados de risco/CMA.""")
    with col3:
        st.markdown("""### Transparência de origem
Ajuda a diferenciar parâmetros replicados dos RSLs da USEPA daqueles adaptados pela CETESB, como Kd de metais, meia-vida e fatores de bioacumulação vegetal.""")
    st.markdown("""## Controle de consistência para intrusão de vapores
Um dos pontos de maior relevância técnica é a limitação associada ao parâmetro **Lgw**, profundidade do nível d'água, no modelo de Johnson & Ettinger. O manual explicita que o valor de Lgw não pode ser inferior à soma da espessura da franja capilar e da espessura das fundações.

Essa restrição evita combinações fisicamente inconsistentes em áreas com lençol freático raso e reduz o risco de uso inadequado da ferramenta em cenários de intrusão de vapores.""")
    st.warning("Ponto de atenção: a restrição operacional do parâmetro Lgw na CETESBRisk indica a necessidade de coerência física entre nível d'água, franja capilar e fundação. Em áreas com nível d'água muito raso, recomenda-se não depender exclusivamente do modelo Johnson & Ettinger (J&E), devendo-se complementar a avaliação com linhas de evidência de campo, como gás do solo, subslab ou ar interno, conforme boas práticas de avaliação de intrusão de vapores.")
    st.markdown("""## MCL, potabilidade e cálculo de risco
O manual também ajuda a diferenciar valores regulatórios, como MCL e potabilidade, dos parâmetros efetivamente usados no cálculo de risco. Essa distinção é essencial para interpretar corretamente a Classe C: mudança em potabilidade pode afetar enquadramento regulatório, mas não significa automaticamente alteração do risco calculado.""")

elif page == "Glossário Técnico":
    st.title("Glossário Técnico")
    st.write("Termos e siglas utilizados no CETESBRisk Explorer e em avaliações de risco à saúde humana.")
    glossary = [
        ("Conceitos gerais", "SQI", "Substância Química de Interesse. Substância selecionada para avaliação em função de sua presença, concentração, toxicidade, mobilidade ou relevância no modelo conceitual."),
        ("Conceitos gerais", "CMA", "Concentração Máxima Aceitável. Concentração calculada como aceitável para determinado cenário de exposição e nível de risco adotado."),
        ("Conceitos gerais", "Avaliação de Risco", "Processo técnico utilizado para estimar riscos potenciais à saúde humana associados à exposição a substâncias químicas em diferentes meios ambientais."),
        ("Conceitos gerais", "Modelo Conceitual", "Representação integrada das fontes de contaminação, meios afetados, mecanismos de transporte, vias de exposição e receptores potenciais."),
        ("Parâmetros toxicológicos", "FatTox", "Base de fatores/parâmetros toxicológicos utilizada pela CETESBRisk, incluindo parâmetros de risco carcinogênico e não carcinogênico."),
        ("Parâmetros toxicológicos", "RfDo", "Dose de Referência Oral. Parâmetro usado na avaliação de risco não carcinogênico por ingestão."),
        ("Parâmetros toxicológicos", "RfCi", "Concentração de Referência por Inalação. Parâmetro usado na avaliação de risco não carcinogênico por inalação."),
        ("Parâmetros toxicológicos", "SFo", "Fator de Slope Oral. Parâmetro usado no cálculo de risco carcinogênico por exposição oral."),
        ("Parâmetros toxicológicos", "IUR", "Unidade de Risco por Inalação. Parâmetro usado no cálculo de risco carcinogênico por inalação."),
        ("Parâmetros físico-químicos", "FisQui", "Base de parâmetros físico-químicos utilizada pela CETESBRisk para descrever comportamento ambiental, volatilização, solubilidade, particionamento e transporte."),
        ("Parâmetros físico-químicos", "Densidade", "Parâmetro físico-químico auxiliar. Na regra objetiva do app, não gera D2 isoladamente e só é considerado para classificação quando |Δ| ≥ 5%."),
        ("Parâmetros físico-químicos", "Constante de Henry", "Parâmetro que expressa a tendência de uma substância transferir-se da fase aquosa para a fase gasosa. É relevante para volatilização e intrusão de vapores."),
        ("Parâmetros físico-químicos", "Koc", "Coeficiente de partição carbono orgânico-água. Indica a tendência da substância se associar à matéria orgânica do solo."),
        ("Parâmetros físico-químicos", "Kd", "Coeficiente de distribuição solo-água. Usado para estimar a partição da substância entre fase sólida e fase aquosa."),
        ("Parâmetros físico-químicos", "Kow / log Kow", "Coeficiente de partição octanol-água. Relacionado à hidrofobicidade da substância e ao potencial de partição em fases orgânicas."),
        ("Parâmetros físico-químicos", "Csat", "Concentração de Saturação. Concentração acima da qual pode haver limitação física de solubilidade ou particionamento no meio avaliado."),
        ("Critérios regulatórios", "MCL", "Maximum Contaminant Level. Padrão de potabilidade adotado pela USEPA para água destinada ao consumo humano."),
        ("Critérios regulatórios", "Potabilidade", "Critérios ou padrões aplicáveis à qualidade da água destinada ao consumo humano."),
        ("Modelagem de exposição", "Lgw", "Profundidade do nível d'água utilizada em modelos de intrusão de vapores. Na CETESBRisk v4.00, há restrição para evitar combinações fisicamente inconsistentes com a franja capilar e fundações."),
        ("Modelagem de exposição", "Intrusão de Vapores", "Migração de vapores de substâncias voláteis presentes no solo ou água subterrânea para ambientes internos de edificações."),
        ("Modelagem de exposição", "Johnson & Ettinger", "Modelo utilizado para estimar intrusão de vapores a partir de fontes em solo ou água subterrânea. Possui premissas e limitações, especialmente em áreas com nível d'água raso."),
    ]
    df_gloss = pd.DataFrame(glossary, columns=["Categoria", "Termo", "Definição"])
    termo = st.text_input("Pesquisar termo ou palavra-chave", "")
    if termo:
        mask = df_gloss["Termo"].str.contains(termo, case=False, na=False) | df_gloss["Definição"].str.contains(termo, case=False, na=False) | df_gloss["Categoria"].str.contains(termo, case=False, na=False)
        df_view = df_gloss[mask].copy()
    else:
        df_view = df_gloss.copy()
    for categoria in df_view["Categoria"].drop_duplicates():
        st.subheader(categoria)
        for _, r in df_view[df_view["Categoria"] == categoria].iterrows():
            with st.expander(r["Termo"]):
                st.write(r["Definição"])
    st.download_button("Baixar glossário em CSV", data=df_gloss.to_csv(index=False).encode("utf-8-sig"), file_name="glossario_tecnico_cetesbrisk_explorer.csv", mime="text/csv")

elif page == "Downloads e notas":
    st.title("Downloads e notas metodológicas")
    st.write("Bases utilizadas pelo aplicativo e arquivo derivado de classificação consolidada.")
    with open(DATA_PATH, "rb") as f:
        st.download_button("Baixar base analítica em Excel", data=f, file_name="Analise_detalhada_CETESBRisk_2023_vs_2026_CMA.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.subheader("Arquivos derivados")
    st.download_button("Baixar tabela consolidada por SQI (CSV)", data=master.to_csv(index=False).encode("utf-8-sig"), file_name="cetesbrisk_master_regra_objetiva.csv", mime="text/csv")

    st.subheader("Notas metodológicas")
    st.markdown(
        """- A classificação A/B/C/D1/D2 segue regra objetiva e hierárquica. Alterações toxicológicas reais são classificadas como D1. Alterações físico-químicas relevantes, exceto densidade, são classificadas como D2 quando **|Δ| ≥ 5%**. Alterações em MCL ou Potabilidade são classificadas como C quando não houver D1 ou D2. A densidade é tratada como parâmetro auxiliar: somente é considerada quando **|Δ| ≥ 5%** e, quando isolada, classifica a SQI como B. Alterações de densidade inferiores a 5%, quando isoladas, não alteram a classe da SQI.
- A tendência de CMA é estimativa técnica, não recálculo oficial da planilha CETESB.
- Impactos diretos foram considerados apenas para RfDo, RfCi, SFO e IUR; parâmetros físico-químicos têm impacto modelo-dependente."""
    )
