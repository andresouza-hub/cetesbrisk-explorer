from pathlib import Path
from datetime import datetime
import pandas as pd
import streamlit as st
import plotly.express as px

st.set_page_config(page_title='CETESBRisk Explorer', page_icon='🧪', layout='wide')
DATA_PATH = Path(__file__).parent / 'data' / 'Analise_detalhada_CETESBRisk_2023_vs_2026_CMA.xlsx'

FIXED_CLASS_COUNTS = {'A':296,'B':471,'C':18,'D1':23,'D2':11}
GLOBAL_COUNTS = {'SQIs comuns comparadas':819,'Comparações individuais':29484,'Novas SQIs incluídas':59,'SQIs removidas':1,'Comparações com alteração':1009,'Comparações sem alteração':16565,'Sem valor/vazios':11910}
CLASS_DESCRIPTIONS = {
 'A':{'descricao':'Sem alteração material relevante','recomendacao':'Não indica revisão automática da AR exclusivamente pela atualização da planilha.'},
 'B':{'descricao':'Alterações menores','recomendacao':'Revisar apenas em cenários muito sensíveis ou resultados próximos aos critérios aceitáveis.'},
 'C':{'descricao':'Alterações regulatórias/potabilidade','recomendacao':'Revisar enquadramento regulatório/potabilidade; não implica necessariamente recálculo do risco.'},
 'D1':{'descricao':'Alterações toxicológicas relevantes','recomendacao':'Revisão recomendada quando a SQI for relevante para o site, especialmente se o resultado anterior estiver próximo ao limite.'},
 'D2':{'descricao':'Alterações físico-químicas relevantes','recomendacao':'Revisão dirigida em cenários sensíveis à volatilização, particionamento, transporte ou intrusão de vapores.'}
}
CAS_OVERRIDES = {'71-43-2':'A','108-88-3':'A','100-41-4':'A','95-47-6':'A','106-42-3':'A','108-38-3':'A','127-18-4':'B','79-01-6':'B','156-59-2':'C','156-60-5':'C','75-35-4':'D1','75-01-4':'D1','64742-89-8':'D2'}
NAME_OVERRIDES = [('benzene','A'),('toluene','A'),('ethylbenzene','A'),('xylene','A'),('tetrachloroethylene','B'),('tetrachloroethene','B'),('trichloroethylene','B'),('trichloroethene','B'),('cis-1,2-dichloroethene','C'),('trans-1,2-dichloroethene','C'),('1,1-dichloroethene','D1'),('vinyl chloride','D1'),('aliphatic low','D2'),('c5-c8','D2'),('aliphatic medium','A'),('c9-c18','A'),('aliphatic high','A'),('c19-c32','A'),('aromatic medium','A'),('c9-c10','A'),('aromatic high','A'),('c10-c32','A'),('pfos','D1'),('pfoa','D1'),('hfpo','D1'),('genx','D1'),('chromium(vi)','D1'),('chromium vi','D1'),('formaldehyde','D1'),('acrylonitrile','D1'),('chloroprene','D1')]
DIRECT_TOX = {'RfDo','RfCi','SFO','IUR','CARCINOGÊNICO','Classe de Cancer','Mutagenico'}
REG = {'MCL','Potabilidade'}
PHYS = {'H (-)','HLC (atm-m³/mole)','Dia (cm²/s)','Diw (cm²/s)','Koc','S (mg/L)','Kd (L/kg)','log Kow','Pvap (mm Hg)','Csat','B (-)','FA','PC (cm/h)','PF (°C)','Densidade'}

@st.cache_data
def load_data():
    xls = pd.ExcelFile(DATA_PATH)
    return (pd.read_excel(xls,'Comparativo_Completo'), pd.read_excel(xls,'SQIs_Incluidas_2026'), pd.read_excel(xls,'SQIs_Removidas_2026'), pd.read_excel(xls,'Alteracoes_Relevantes'), pd.read_excel(xls,'Impacto_CMA_Direto'))

def changed(row): return str(row.get('Status','')).lower() in ['aumentou','diminuiu','alteração textual/categórica','alterado']
def mat_var(row,t=0.05):
    try: return abs(float(row.get('Variação % Parâmetro',0))) >= t
    except Exception: return False

def classify(cas,name,g):
    cas = '' if pd.isna(cas) else str(cas).strip(); name_l = '' if pd.isna(name) else str(name).lower()
    if cas in CAS_OVERRIDES: return CAS_OVERRIDES[cas]
    for key,cls in NAME_OVERRIDES:
        if key in name_l: return cls
    ch = g[g.apply(changed,axis=1)]
    if ch.empty: return 'A'
    tox = ch[ch['Parâmetro'].isin(DIRECT_TOX)]
    if not tox.empty and (not tox[tox['Status'].astype(str).str.contains('textual|categ',case=False,na=False)].empty or not tox[tox.apply(mat_var,axis=1)].empty): return 'D1'
    reg = ch[ch['Parâmetro'].isin(REG)]; phys = ch[ch['Parâmetro'].isin(PHYS)]; phys_mat = phys[phys.apply(mat_var,axis=1)]
    if not reg.empty and phys_mat.empty: return 'C'
    if not phys_mat.empty: return 'D2'
    return 'B'

@st.cache_data
def build_master(comp):
    rows=[]
    for (cas,a23,a26),g in comp.groupby(['CAS','Analito 2023','Analito 2026'],dropna=False):
        name = a26 if pd.notna(a26) and str(a26).strip() else a23
        cls=classify(cas,name,g); ch=g[g.apply(changed,axis=1)]
        rows.append({'CAS':cas,'Composto':name,'Classe':cls,'Descrição da classe':CLASS_DESCRIPTIONS[cls]['descricao'],'Nº alterações':int(ch.shape[0]),'Parâmetros alterados':', '.join(ch['Parâmetro'].dropna().astype(str).unique()[:8]),'Recomendação':CLASS_DESCRIPTIONS[cls]['recomendacao']})
    return pd.DataFrame(rows)

def badge(cls):
    colors={'A':'#2e7d32','B':'#6c757d','C':'#f9a825','D1':'#c62828','D2':'#ef6c00'}
    return f"<span style='background:{colors.get(cls,'#999')}; color:white; padding:0.35rem 0.6rem; border-radius:0.5rem; font-weight:700;'>Classe {cls} · {CLASS_DESCRIPTIONS[cls]['descricao']}</span>"

def report_html(compound, cas, cls, detail, synthesis):
    now=datetime.now().strftime('%d/%m/%Y %H:%M')
    return f"""<html><head><meta charset='utf-8'><title>Relatório CETESBRisk Explorer - {compound}</title><style>body{{font-family:Arial;margin:40px;color:#1F2A44}}h1,h2{{color:#0B2A4A}}.box{{background:#f3f6fa;border-left:5px solid #0B5CAD;padding:15px;margin:20px 0}}table{{border-collapse:collapse;width:100%;font-size:11px}}th,td{{border:1px solid #ddd;padding:6px}}th{{background:#0B5CAD;color:white}}</style></head><body><h1>CETESBRisk Explorer</h1><h2>Relatório técnico por SQI</h2><p><b>Data/hora:</b> {now}</p><p><b>Substância:</b> {compound}</p><p><b>CAS:</b> {cas}</p><p><b>Classe:</b> {cls} - {CLASS_DESCRIPTIONS[cls]['descricao']}</p><div class='box'>{synthesis}</div><h2>Comparativo por parâmetro</h2>{detail.to_html(index=False,border=0)}</body></html>"""


def to_float_series(s):
    return pd.to_numeric(s, errors='coerce')

def top_sqi_by_parameter_changes(comp, fonte, n=5):
    df = comp.copy()
    df = df[df['Fonte'].astype(str).str.lower() == fonte.lower()]
    df = df[df.apply(changed, axis=1)].copy()
    if df.empty:
        return pd.DataFrame(columns=['SQI','CAS','Nº parâmetros alterados','Maior variação absoluta (%)','Principais parâmetros alterados'])

    df['abs_var'] = to_float_series(df.get('Variação % Parâmetro', pd.Series(dtype=float))).abs()
    grouped = (
        df.groupby(['Analito 2026','CAS'], dropna=False)
        .agg(**{
            'Nº parâmetros alterados': ('Parâmetro','nunique'),
            'Maior variação absoluta (%)': ('abs_var','max'),
            'Principais parâmetros alterados': ('Parâmetro', lambda x: ', '.join(pd.Series(x).dropna().astype(str).unique()[:6]))
        })
        .reset_index()
        .rename(columns={'Analito 2026':'SQI'})
    )
    grouped['Maior variação absoluta (%)'] = grouped['Maior variação absoluta (%)'].round(2)
    return grouped.sort_values(['Nº parâmetros alterados','Maior variação absoluta (%)'], ascending=False).head(n)

def top_cma_impact(comp, direction='reduz', n=5):
    df = comp.copy()
    df = df[df.apply(changed, axis=1)].copy()
    df['Estimativa variação CMA % num'] = to_float_series(df.get('Estimativa variação CMA %', pd.Series(dtype=float)))
    df = df.dropna(subset=['Estimativa variação CMA % num'])
    if direction == 'reduz':
        df = df[df['Estimativa variação CMA % num'] < 0]
        sort_ascending = True
    else:
        df = df[df['Estimativa variação CMA % num'] > 0]
        sort_ascending = False

    if df.empty:
        return pd.DataFrame(columns=['SQI','CAS','Parâmetro crítico','Variação estimada CMA (%)','Tendência esperada da CMA'])

    df['abs_cma'] = df['Estimativa variação CMA % num'].abs()
    idx = df.groupby(['Analito 2026','CAS'], dropna=False)['abs_cma'].idxmax()
    out = df.loc[idx, ['Analito 2026','CAS','Parâmetro','Estimativa variação CMA % num','Tendência esperada da CMA']].copy()
    out = out.rename(columns={
        'Analito 2026':'SQI',
        'Parâmetro':'Parâmetro crítico',
        'Estimativa variação CMA % num':'Variação estimada CMA (%)'
    })
    out['Variação estimada CMA (%)'] = out['Variação estimada CMA (%)'].round(2)
    return out.sort_values('Variação estimada CMA (%)', ascending=sort_ascending).head(n)

def build_interpretive_synthesis(selected, cas, cls, detail):
    ch = detail[detail.apply(changed, axis=1)].copy()
    class_desc = CLASS_DESCRIPTIONS[cls]['descricao']
    rec = CLASS_DESCRIPTIONS[cls]['recomendacao']

    if ch.empty:
        return (
            f"A comparação entre as versões CETESBRisk v3.03 e v4.00 indica que a SQI {selected} "
            f"(CAS {cas}) não apresentou alterações materiais relevantes nos parâmetros avaliados. "
            f"O composto foi enquadrado na Classe {cls} ({class_desc}), indicando que a atualização da planilha, "
            f"isoladamente, não justifica revisão da avaliação de risco para esta SQI. A necessidade de revisão deve ser "
            f"considerada apenas se houver mudança do modelo conceitual, das vias de exposição, do uso da área ou de outras premissas do estudo."
        )

    tox_params = ['RfDo','RfCi','SFO','IUR','Classe de Cancer','Mutagenico','CARCINOGÊNICO']
    reg_params = ['MCL','Potabilidade']
    phys_params = ['H (-)','HLC (atm-m³/mole)','Koc','Kd (L/kg)','S (mg/L)','Pvap (mm Hg)','Csat','log Kow','Densidade']

    tox = ch[ch['Parâmetro'].isin(tox_params)]
    reg = ch[ch['Parâmetro'].isin(reg_params)]
    phys = ch[ch['Parâmetro'].isin(phys_params)]

    def params_txt(df):
        vals = df['Parâmetro'].dropna().astype(str).unique().tolist()
        if not vals:
            return ""
        return ", ".join(vals[:5]) + (" entre outros" if len(vals) > 5 else "")

    ch2 = ch.copy()
    ch2['abs_var'] = pd.to_numeric(ch2.get('Variação % Parâmetro', pd.Series(dtype=float)), errors='coerce').abs()
    ch2 = ch2.dropna(subset=['abs_var']).sort_values('abs_var', ascending=False)
    if not ch2.empty:
        top = ch2.iloc[0]
        try:
            top_var = f"{float(top.get('Variação % Parâmetro')):.1f}%"
        except Exception:
            top_var = str(top.get('Variação % Parâmetro',''))
        top_sentence = f"A maior variação percentual observada ocorreu em {top.get('Parâmetro','')}, de {top.get('Valor 2023','')} para {top.get('Valor 2026','')} ({top_var}). "
    else:
        top_sentence = ""

    tendencies = ch['Tendência esperada da CMA'].dropna().astype(str).str.lower().tolist()
    cma_reduce = any(('redu' in t or 'dimin' in t or 'restritiva' in t) for t in tendencies)
    cma_increase = any(('aument' in t or 'menos restritiva' in t) for t in tendencies)

    if cls == 'A':
        text = (
            f"A SQI {selected} (CAS {cas}) foi enquadrada na Classe A, pois as alterações identificadas não configuram mudança material "
            f"capaz de alterar de forma relevante o risco calculado ou a CMA. {top_sentence}"
            f"A expectativa técnica é de ausência de impacto material sobre a avaliação de risco, considerando exclusivamente a atualização da planilha."
        )
    elif cls == 'B':
        text = (
            f"A SQI {selected} (CAS {cas}) foi enquadrada na Classe B, indicando alterações menores ou refinamentos paramétricos. {top_sentence}"
            f"Não há indicação automática de revisão da avaliação de risco, embora a revisão possa ser considerada em cenários muito sensíveis "
            f"ou quando a avaliação anterior estiver próxima dos critérios de aceitabilidade."
        )
    elif cls == 'C':
        p = params_txt(reg)
        text = (
            f"A SQI {selected} (CAS {cas}) foi enquadrada na Classe C, associada a alterações regulatórias ou de potabilidade. "
            f"{'As alterações observadas envolveram principalmente ' + p + '. ' if p else ''}"
            f"Essa classificação não significa, por si só, alteração material do cálculo de risco ou da CMA. "
            f"O principal efeito esperado é sobre o enquadramento regulatório, atualização de tabelas comparativas ou discussão de potabilidade, "
            f"especialmente quando houver avaliação da via de ingestão de água subterrânea."
        )
    elif cls == 'D1':
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
    elif cls == 'D2':
        p = params_txt(phys)
        text = (
            f"A SQI {selected} (CAS {cas}) foi enquadrada na Classe D2, indicando alteração relevante em parâmetros físico-químicos. "
            f"{'As alterações envolveram principalmente ' + p + '. ' if p else ''}{top_sentence}"
            f"Essas alterações não representam necessariamente mudança direta de toxicidade, mas podem modificar o comportamento ambiental da substância, "
            f"influenciando volatilização, particionamento, transporte, concentração de saturação ou intrusão de vapores. "
            f"O impacto sobre CMA é modelo-dependente e deve ser avaliado principalmente em cenários sensíveis à inalação, fase vapor ou transporte entre fonte e ponto de exposição."
        )
    else:
        text = f"A SQI {selected} (CAS {cas}) foi enquadrada na Classe {cls} ({class_desc}). {rec}"

    if cls in ['D1','D2']:
        text += " Em termos práticos, a atualização deve ser tratada como gatilho de triagem técnica, não como obrigação automática de refazimento integral da avaliação."
    elif cls == 'C':
        text += " Assim, a ação prioritária é verificar o enquadramento legal/regulatório aplicável, sem assumir automaticamente necessidade de recálculo do risco."
    else:
        text += " Assim, a atualização da planilha, isoladamente, não indica necessidade de revisão da avaliação de risco para esta SQI."

    return text


comp,incl,remov,alt,impacto=load_data(); master=build_master(comp)
CLASS_COUNTS = master['Classe'].value_counts().reindex(['A','B','C','D1','D2']).fillna(0).astype(int).to_dict()

st.sidebar.title('🧪 CETESBRisk Explorer')
st.sidebar.markdown('**Desenvolvido por André Souza**  \nEspecialista em GAC  \n[LinkedIn](https://www.linkedin.com/in/andr%C3%A9-souza-63539517)')
st.sidebar.divider()
page=st.sidebar.radio('Navegação',['Início','Dashboard geral','Pesquisa SQI e impacto CMA','Grupos prioritários','Highlights Manual CETESB','Glossário Técnico','Downloads e notas'])
st.sidebar.caption('v1.1 · contagens corrigidas')

if page=='Início':
    st.title('🧪 CETESBRisk Explorer')
    st.subheader('Comparador técnico CETESBRisk v3.03 (2023) × v4.00 (2026)')
    st.markdown('''## 1. O que é esta ferramenta?
O **CETESBRisk Explorer** é uma ferramenta de apoio à interpretação técnica das alterações introduzidas entre as versões **CETESBRisk v3.03 (2023)** e **CETESBRisk v4.00 (2026)**.

A ferramenta **não recalcula avaliações de risco** e **não substitui a análise crítica do profissional responsável**. Seu objetivo é identificar alterações que possam justificar revisão dirigida de avaliações de risco previamente elaboradas.''')
    st.info('Mensagem-chave: a atualização da v4.00 não implica revisão automática de todas as avaliações de risco. A recomendação é realizar triagem dirigida por SQI, via de exposição e sensibilidade do cenário.')
    st.markdown('''## 2. O que está sendo avaliado?
Foram comparados os bancos de dados internos das duas versões da planilha CETESBRisk, com foco nas bases **FisQui** e **FatTox**.

A **FatTox** reúne parâmetros toxicológicos, como RfDo, RfCi, SFo e IUR. Alterações nesses parâmetros podem repercutir diretamente no cálculo de risco e nas CMAs.

A **FisQui** reúne parâmetros físico-químicos, como Henry, solubilidade, Koc, Kd, pressão de vapor e Csat. Esses parâmetros não alteram necessariamente a toxicidade, mas podem influenciar volatilização, transporte, particionamento e intrusão de vapores. Em muitos casos, as alterações foram apenas refinamentos numéricos decimais; em compostos específicos, entretanto, as variações podem alterar cenários modelados, com impacto modelo-dependente.''')
    c1,c2=st.columns(2)
    with c1: st.markdown('''### Base Físico-Química — FisQui
- solubilidade;\n- constante de Henry;\n- pressão de vapor;\n- Koc e Kd;\n- difusão;\n- Csat;\n- particionamento solo-água-ar.''')
    with c2: st.markdown('''### Base Toxicológica — FatTox
- RfDo;\n- RfCi;\n- SFo;\n- IUR;\n- carcinogenicidade;\n- mutagenicidade;\n- MCL e potabilidade.''')
    st.markdown('## 3. Escopo da avaliação')
    c1,c2,c3,c4,c5=st.columns(5)
    c1.metric('SQIs comuns comparadas','819'); c2.metric('Comparações individuais','29.484'); c3.metric('Novas SQIs incluídas','59'); c4.metric('SQIs removidas','1'); c5.metric('Versões avaliadas','2023 × 2026')
    st.caption('A comparação detalhada foi realizada para 819 SQIs presentes em ambas as versões. Adicionalmente, foram identificadas 59 novas substâncias incluídas na v4.00 e 1 substância removida em relação à v3.03.')
    st.markdown('''## 4. Como os compostos foram classificados?
As alterações identificadas foram agrupadas em cinco classes de criticidade, considerando o tipo de parâmetro alterado e o potencial de repercussão sobre risco, CMA, transporte, exposição ou enquadramento regulatório.''')
    classe_df=pd.DataFrame({'Classe':['A','B','C','D1','D2'],'Significado':['Sem alteração material relevante','Alteração menor','Alteração regulatória/potabilidade','Alteração toxicológica relevante','Alteração físico-química relevante'],'Implicação prática':['Não indica revisão automática da avaliação de risco.','Revisar apenas em cenários muito sensíveis.','Pode exigir revisão regulatória, mas não necessariamente recálculo do risco.','Pode justificar revisão dirigida ou recomendada da avaliação de risco.','Pode justificar revisão dirigida em cenários sensíveis à volatilização, transporte ou intrusão de vapores.']})
    st.dataframe(classe_df,use_container_width=True,hide_index=True)
    st.markdown('## 5. Explicação detalhada das classes')

    with st.expander('Classe A — Sem alteração material relevante'):
        st.markdown("""
        A Classe A reúne substâncias para as quais não foram identificadas alterações materiais nos parâmetros com potencial de modificar de forma relevante o risco calculado, a Concentração Máxima Aceitável (CMA) ou o comportamento ambiental da substância.

        Embora possam existir pequenas diferenças de nomenclatura, arredondamentos, ajustes de precisão numérica ou alterações sem repercussão prática esperada, não foram observadas modificações significativas nos parâmetros toxicológicos ou físico-químicos capazes de justificar, por si só, a revisão de avaliações de risco previamente elaboradas.

        **Interpretação prática:** a atualização da CETESBRisk não representa gatilho para reavaliação dessas substâncias. Eventual revisão deve decorrer de mudanças no modelo conceitual, nas vias de exposição, no uso da área ou em novos dados ambientais.

        **Exemplos identificados nesta comparação:** benzeno, tolueno, etilbenzeno, xilenos e grande parte dos compostos clássicos de petróleo.
        """)

    with st.expander('Classe B — Alterações menores'):
        st.markdown("""
        A Classe B reúne substâncias que apresentaram alterações paramétricas de pequena magnitude, normalmente associadas a refinamentos de valores, ajustes metodológicos ou correções sem potencial relevante de alterar a interpretação do risco.

        Embora não sejam alterações consideradas materiais, podem justificar uma verificação pontual em situações específicas, especialmente quando a substância controla o resultado da avaliação ou quando os riscos calculados anteriormente se encontram muito próximos dos critérios de aceitabilidade.

        **Interpretação prática:** revisão geralmente não necessária, mas pode ser recomendável em estudos particularmente sensíveis.

        **Exemplos identificados nesta comparação:** PCE/tetracloroeteno e TCE/tricloroeteno.
        """)

    with st.expander('Classe C — Alterações regulatórias e de potabilidade'):
        st.warning('A Classe C é uma das mais sujeitas a interpretação equivocada: alteração regulatória não significa, necessariamente, alteração do risco calculado.')
        st.markdown("""
        A Classe C reúne substâncias que apresentaram alterações principalmente em critérios regulatórios, valores de potabilidade ou parâmetros utilizados para comparação legal, sem alteração material dos parâmetros que controlam diretamente o cálculo de risco.

        Uma alteração de critério de potabilidade, como o Maximum Contaminant Level (MCL), pode modificar o enquadramento regulatório da substância, alterar tabelas comparativas ou afetar discussões relacionadas ao uso da água subterrânea. Isso, entretanto, não significa necessariamente que a CMA ou o risco calculado foram alterados.

        **Interpretação prática:** o foco deve estar na revisão dos critérios de comparação e do enquadramento regulatório, e não no recálculo automático da avaliação de risco.

        **Exemplos identificados nesta comparação:** cis-1,2-DCE e trans-1,2-DCE.
        """)

    with st.expander('Classe D1 — Alterações toxicológicas relevantes'):
        st.markdown("""
        A Classe D1 reúne substâncias que apresentaram alterações relevantes em parâmetros diretamente associados à toxicidade e ao cálculo de risco, incluindo dose de referência oral (RfDo), concentração de referência por inalação (RfCi), fator de slope oral (SFo), unidade de risco por inalação (IUR), classificação carcinogênica e mutagenicidade.

        Essas alterações possuem potencial direto de repercutir nos resultados da avaliação de risco e, consequentemente, nas CMAs derivadas. Dependendo do parâmetro alterado, a atualização pode tornar os critérios mais restritivos ou menos restritivos, alterando a interpretação de estudos anteriormente desenvolvidos.

        **Interpretação prática:** substâncias enquadradas nesta classe devem ser consideradas prioritárias em qualquer triagem para revisão de avaliações de risco, especialmente quando forem relevantes no modelo conceitual ou quando a avaliação anterior estiver próxima dos limites de aceitabilidade.

        **Exemplos identificados nesta comparação:** cloreto de vinila, 1,1-DCE, PFOS, PFOA, HFPO-DA/GenX e cromo VI.
        """)

    with st.expander('Classe D2 — Alterações físico-químicas relevantes'):
        st.markdown("""
        A Classe D2 reúne substâncias que apresentaram alterações relevantes em parâmetros físico-químicos capazes de influenciar o comportamento ambiental da substância e sua modelagem de exposição.

        Entre os parâmetros potencialmente envolvidos estão constante de Henry, pressão de vapor, solubilidade, coeficiente de partição carbono orgânico-água (Koc), coeficiente de distribuição solo-água (Kd), log Kow e concentração de saturação (Csat).

        Diferentemente da Classe D1, essas alterações não representam necessariamente mudança de toxicidade. Seu impacto depende do modelo conceitual e da via de exposição considerada, sendo especialmente relevante para avaliações envolvendo volatilização, intrusão de vapores, transporte em fase vapor, particionamento solo-água-ar e modelagem de transporte.

        **Interpretação prática:** a necessidade de revisão depende da importância desses processos no cenário avaliado.

        **Exemplo identificado nesta comparação:** TPH alifático leve C5-C8.
        """)

    st.info(
        "Importante: as classes representam uma ferramenta de triagem técnica para priorização de revisões. "
        "A classificação não substitui a análise crítica do modelo conceitual da área, das vias de exposição, "
        "das substâncias efetivamente presentes e dos resultados históricos disponíveis."
    )

    st.markdown('''## 6. Fluxo decisório: quando revisar uma avaliação anterior?
**Perguntas de triagem recomendadas:**
1. O modelo conceitual possui SQIs da Classe D1, como PFAS, cloreto de vinila ou 1,1-DCE?
2. Há SQIs da Classe D2 associadas a volatilização, intrusão de vapores ou transporte em fase vapor?
3. A avaliação anterior estava próxima da margem de aceitabilidade, como HQ ≈ 1 ou risco ≈ 1E-05?
4. A avaliação baseou-se predominantemente em frações genéricas de TPH em vez de compostos individuais de petróleo?

Se a resposta for **sim** para uma ou mais perguntas, recomenda-se **revisão dirigida**. Se todas forem **não**, a atualização da planilha, isoladamente, não indica revisão automática.''')
    st.markdown('## 7. Resultado da classificação')
    r1,r2,r3,r4,r5=st.columns(5); r1.metric('Classe A',str(CLASS_COUNTS.get('A',0))); r2.metric('Classe B',str(CLASS_COUNTS.get('B',0))); r3.metric('Classe C',str(CLASS_COUNTS.get('C',0))); r4.metric('Classe D1',str(CLASS_COUNTS.get('D1',0))); r5.metric('Classe D2',str(CLASS_COUNTS.get('D2',0)))
    st.markdown('''## 8. PFAS e compostos correlatos
A v4.00 ampliou significativamente a capacidade de avaliação de PFAS e compostos correlatos. O ganho não é apenas numérico: vários compostos passaram a ter maior preenchimento de parâmetros físico-químicos e toxicológicos, reduzindo lacunas que limitavam análises anteriores.

PFOS, PFOA, HFPO-DA/GenX e sais relacionados devem ser tratados como grupos de atenção, especialmente pela combinação de persistência, mobilidade, evolução regulatória e alterações relevantes em parâmetros toxicológicos.''')
    st.markdown('''## 9. Limitações da classificação
A classificação apresentada nesta ferramenta representa uma **triagem técnica** baseada nas alterações identificadas entre as versões da CETESBRisk. A decisão de revisar ou não uma avaliação de risco deve considerar adicionalmente o modelo conceitual da área, as vias de exposição completas, as SQIs efetivamente presentes, os resultados históricos e a proximidade aos critérios de aceitabilidade.''')
    st.markdown('''## 10. Conclusão executiva
A versão **CETESBRisk v4.00** representa um avanço importante na atualização das bases físico-químicas e toxicológicas, bem como na transparência da ferramenta, especialmente pela publicação do **Manual do Usuário**.

Entretanto, a maior parte dos compostos avaliados não apresentou alteração material relevante. Dessa forma, a atualização **não implica revisão automática das avaliações de risco existentes**.

A abordagem recomendada é uma **triagem dirigida**, priorizando substâncias classificadas como **D1**, cenários sensíveis à volatilização e intrusão de vapores associados à **D2**, além de grupos de atenção como **PFAS**, **cloreto de vinila**, **1,1-DCE** e **TPH alifático leve C5-C8**.''')

elif page=='Dashboard geral':
    st.title('Dashboard geral')
    st.write('Exploração quantitativa dos pontos de dados comparados entre as versões v3.03 e v4.00.')

    c1,c2,c3,c4=st.columns(4)
    c1.metric('Pontos de dados analisados','29.484')
    c2.metric('Com alteração','1.009','3,4% do total')
    c3.metric('Sem alteração','16.565')
    c4.metric('Sem valor/vazios','11.910')

    st.subheader('Distribuição consolidada dos compostos por classe de alteração')
    chart_df=pd.DataFrame({'Classe':list(CLASS_COUNTS.keys()),'N compostos':list(CLASS_COUNTS.values())})
    fig=px.bar(chart_df,x='Classe',y='N compostos',color='Classe',text='N compostos',hover_data={'Classe':True,'N compostos':True},title='Compostos por classe de alteração')
    fig.update_traces(textposition='outside', hovertemplate='<b>Classe %{x}</b><br>Nº de compostos: %{y}<extra></extra>')
    fig.update_layout(yaxis_title='Nº de compostos',xaxis_title='Classe')
    st.plotly_chart(fig,use_container_width=True)

    st.subheader('Composição dos pontos de dados avaliados')
    flow_df=pd.DataFrame({'Categoria':['Com alteração','Sem alteração','Sem valor/vazios'],'N':[1009,16565,11910]})
    fig2=px.bar(flow_df,x='Categoria',y='N',color='Categoria',text='N',title='Status dos 29.484 pontos de dados avaliados')
    fig2.update_traces(textposition='outside', hovertemplate='<b>%{x}</b><br>Nº de pontos: %{y}<extra></extra>')
    st.plotly_chart(fig2,use_container_width=True)

    st.divider()
    st.subheader('Rankings de atenção técnica')
    st.markdown('Os rankings abaixo ajudam a identificar SQIs com maior número de alterações em parâmetros físico-químicos, toxicológicos e maior impacto potencial estimado sobre CMA. A interpretação deve considerar o modelo conceitual, as vias de exposição e a relevância da SQI no estudo.')

    tab_fisqui,tab_fattox,tab_cma_down,tab_cma_up=st.tabs(['Top 5 FisQui','Top 5 FatTox','Top 5 CMA mais restritiva','Top 5 CMA menos restritiva'])

    with tab_fisqui:
        st.markdown('**SQIs com maior número de alterações em parâmetros físico-químicos (FisQui).**')
        st.caption('Critério: número de parâmetros FisQui alterados; desempate pela maior variação percentual absoluta.')
        st.dataframe(top_sqi_by_parameter_changes(comp,'FisQui',5),use_container_width=True,hide_index=True)

    with tab_fattox:
        st.markdown('**SQIs com maior número de alterações em parâmetros toxicológicos/regulatórios (FatTox).**')
        st.caption('Critério: número de parâmetros FatTox alterados; desempate pela maior variação percentual absoluta.')
        st.dataframe(top_sqi_by_parameter_changes(comp,'FatTox',5),use_container_width=True,hide_index=True)

    with tab_cma_down:
        st.markdown('**SQIs cujas alterações tendem a reduzir a CMA, tornando o critério potencialmente mais restritivo.**')
        st.caption('Critério: maior redução percentual estimada de CMA em parâmetros com relação direta/inversa simplificada.')
        st.dataframe(top_cma_impact(comp,'reduz',5),use_container_width=True,hide_index=True)

    with tab_cma_up:
        st.markdown('**SQIs cujas alterações tendem a aumentar a CMA, tornando o critério potencialmente menos restritivo.**')
        st.caption('Critério: maior aumento percentual estimado de CMA em parâmetros com relação direta/inversa simplificada.')
        st.dataframe(top_cma_impact(comp,'aumenta',5),use_container_width=True,hide_index=True)

    st.divider()
    st.subheader('Tabela consolidada por composto')
    classes=st.multiselect('Filtrar classes',['A','B','C','D1','D2'],default=['A','B','C','D1','D2'])
    filtered=master[master['Classe'].isin(classes)].copy()
    st.dataframe(filtered,use_container_width=True,hide_index=True)
    st.download_button('Baixar tabela filtrada',data=filtered.to_csv(index=False).encode('utf-8-sig'),file_name='cetesbrisk_master_filtrado.csv',mime='text/csv')

elif page=='Pesquisa SQI e impacto CMA':
    st.title('Pesquisa SQI e impacto potencial na CMA')
    search_df = master.sort_values(['Composto','CAS']).dropna(subset=['Composto']).copy()
    search_df['rotulo_busca'] = search_df.apply(lambda r: f"{r['Composto']} | CAS: {r['CAS']}", axis=1)
    options = search_df['rotulo_busca'].tolist()
    default_idx = 0
    benz = search_df[search_df['Composto'].astype(str).str.lower().eq('benzene')]
    if not benz.empty:
        default_label = benz.iloc[0]['rotulo_busca']
        if default_label in options:
            default_idx = options.index(default_label)
    selected_label = st.selectbox('Selecione uma substância', options, index=default_idx)
    row = search_df[search_df['rotulo_busca'] == selected_label].iloc[0]
    selected = row['Composto']; cas = row['CAS']; cls = row['Classe']
    st.header(selected); st.markdown(badge(cls),unsafe_allow_html=True)
    m1,m2,m3=st.columns(3); m1.metric('CAS',str(cas)); m2.metric('Classe',cls); m3.metric('Nº alterações',int(row['Nº alterações']))
    st.subheader('Interpretação'); st.write(row['Recomendação'])
    detail=comp[(comp['CAS'].astype(str)==str(cas))].copy(); show_cols=['Fonte','Parâmetro','Valor 2023','Valor 2026','Status','Variação % Parâmetro','Tendência esperada da CMA','Estimativa variação CMA %','Observação técnica']
    st.subheader('Comparativo por parâmetro'); st.dataframe(detail[show_cols],use_container_width=True,hide_index=True)
    synthesis=build_interpretive_synthesis(selected,cas,cls,detail)
    st.subheader('Síntese técnica estruturada'); st.text_area('Texto técnico para relatório',synthesis,height=220)
    html=report_html(selected,cas,cls,detail[show_cols],synthesis)
    st.download_button('Baixar relatório HTML desta SQI',data=html.encode('utf-8'),file_name=f"relatorio_cetesbrisk_{str(selected).replace(' ','_')}.html",mime='text/html')

elif page=='Grupos prioritários':
    st.title('Grupos prioritários'); st.write('Síntese interpretativa para grupos frequentemente relevantes em avaliações de risco e GAC.')
    tab1,tab2,tab3,tab4,tab5=st.tabs(['BTEX','Etenos clorados','TPH','PFAS','Metais'])
    with tab1:
        st.subheader('BTEX'); st.markdown('Os BTEX são compostos frequentes em áreas com hidrocarbonetos de petróleo. Na comparação realizada, benzeno e principais BTEX permaneceram sem alteração material relevante, não indicando revisão automática da avaliação de risco apenas pela atualização da planilha.')
        btex=master[master['Composto'].str.lower().str.contains('benzene|toluene|ethylbenzene|xylene',na=False)]; st.dataframe(btex[['CAS','Composto','Classe','Recomendação']],use_container_width=True,hide_index=True)
    with tab2:
        st.subheader('Etenos clorados'); st.markdown('Os etenos clorados são relevantes em cenários de intrusão de vapores e exposição por inalação. PCE e TCE apresentaram alterações menores; cis-1,2-DCE e trans-1,2-DCE merecem atenção regulatória/potabilidade; 1,1-DCE e cloreto de vinila foram enquadrados como D1.')
        data=pd.DataFrame([['PCE / Tetracloroeteno','B','Alteração menor; não indica revisão automática.'],['TCE / Tricloroeteno','B','Alteração menor; não indica revisão automática.'],['cis-1,2-DCE','C','Atenção regulatória/potabilidade.'],['trans-1,2-DCE','C','Atenção regulatória/potabilidade.'],['1,1-DCE','D1','Revisão dirigida/recomendada se for SQI relevante.'],['Cloreto de vinila','D1','RfCi reduzido de 1,0E-01 para 2,0E-02 mg/m³; atenção à inalação/intrusão de vapores.']],columns=['Composto','Classe','Interpretação']); st.dataframe(data,use_container_width=True,hide_index=True)
    with tab3:
        st.subheader('TPH'); st.markdown('Para TPH, a principal alteração material identificada está associada à fração alifática leve C5-C8, enquadrada como D2 por alterações físico-químicas relevantes. As demais faixas avaliadas permaneceram sem alteração material relevante.')
        data=pd.DataFrame([['TPH alifático baixo C5-C8','D2','Revisão dirigida; alteração físico-química relevante.'],['TPH alifático médio C9-C18','A','Sem alteração material relevante.'],['TPH alifático alto C19-C32','A','Sem alteração material relevante.'],['TPH aromático médio C9-C10','A','Sem alteração material relevante.'],['TPH aromático alto C10-C32','A','Sem alteração material relevante.']],columns=['Faixa','Classe','Interpretação']); st.dataframe(data,use_container_width=True,hide_index=True); st.info('As faixas de TPH possuem caráter complementar e não substituem a avaliação de compostos individuais de petróleo quando identificados.')
    with tab4:
        st.subheader('PFAS'); st.markdown('A v4.00 ampliou a representação de PFAS e compostos correlatos. O ganho não é apenas a inclusão de novas substâncias, mas também o preenchimento de lacunas físico-químicas e toxicológicas que limitavam avaliações anteriores.')
        pfas=master[master['Composto'].str.lower().str.contains('pfos|pfoa|hfpo|perfluoro|genx',na=False)]; st.dataframe(pfas[['CAS','Composto','Classe','Recomendação']],use_container_width=True,hide_index=True)
    with tab5:
        st.subheader('Metais'); st.markdown('Para metais, a interpretação deve considerar forma química, Kd, pH e particionamento solo-água. Mudanças nesses parâmetros podem afetar mobilidade, lixiviação e transporte, mas seu impacto é fortemente dependente do modelo conceitual e das condições hidrogeoquímicas.')

elif page=='Highlights Manual CETESB':
    st.title('Highlights do Manual do Usuário da CETESBRisk v4.00')
    st.write('Principais ganhos técnicos e pontos de atenção associados à primeira publicação do Manual do Usuário da planilha CETESB.')
    st.markdown('''## De planilha a documento de governança
O maior ganho do Manual do Usuário é transformar a planilha em uma referência documentada de uso, reduzindo ambiguidades e divergências interpretativas entre consultorias, clientes e órgão ambiental.''')
    col1,col2,col3=st.columns(3)
    with col1: st.markdown('''### Padronização
O manual consolida orientações de preenchimento, uso das abas e interpretação das saídas, funcionando como referência oficial para treinamento e revisão técnica.''')
    with col2: st.markdown('''### Rastreabilidade
Explicita a separação entre abas de entrada, bases internas, cálculos intermediários e resultados de risco/CMA.''')
    with col3: st.markdown('''### Transparência de origem
Ajuda a diferenciar parâmetros replicados dos RSLs da USEPA daqueles adaptados pela CETESB, como Kd de metais, meia-vida e fatores de bioacumulação vegetal.''')
    st.markdown('''## Controle de consistência para intrusão de vapores
Um dos pontos de maior relevância técnica é a limitação associada ao parâmetro **Lgw**, profundidade do nível d'água, no modelo de Johnson & Ettinger. O manual explicita que o valor de Lgw não pode ser inferior à soma da espessura da franja capilar e da espessura das fundações.

Essa restrição evita combinações fisicamente inconsistentes em áreas com lençol freático raso e reduz o risco de uso inadequado da ferramenta em cenários de intrusão de vapores.''')
    st.warning('Ponto de atenção: em áreas rasas, com franja capilar próxima à fundação, o modelo J&E pode não ser aplicável de forma direta. A planilha v4.00 torna essa restrição mais explícita e operacional.')
    st.markdown('''## MCL, potabilidade e cálculo de risco
O manual também ajuda a diferenciar valores regulatórios, como MCL e potabilidade, dos parâmetros efetivamente usados no cálculo de risco. Essa distinção é essencial para interpretar corretamente a Classe C: mudança em potabilidade pode afetar enquadramento regulatório, mas não significa automaticamente alteração do risco calculado.''')


elif page=='Glossário Técnico':
    st.title('Glossário Técnico')
    st.write('Termos e siglas utilizados no CETESBRisk Explorer e em avaliações de risco à saúde humana.')

    glossary = [
        ('Conceitos gerais', 'SQI', 'Substância Química de Interesse. Substância selecionada para avaliação em função de sua presença, concentração, toxicidade, mobilidade ou relevância no modelo conceitual.'),
        ('Conceitos gerais', 'CMA', 'Concentração Máxima Aceitável. Concentração calculada como aceitável para determinado cenário de exposição e nível de risco adotado.'),
        ('Conceitos gerais', 'Avaliação de Risco', 'Processo técnico utilizado para estimar riscos potenciais à saúde humana associados à exposição a substâncias químicas em diferentes meios ambientais.'),
        ('Conceitos gerais', 'Modelo Conceitual', 'Representação integrada das fontes de contaminação, meios afetados, mecanismos de transporte, vias de exposição e receptores potenciais.'),
        ('Parâmetros toxicológicos', 'FatTox', 'Base de fatores/parâmetros toxicológicos utilizada pela CETESBRisk, incluindo parâmetros de risco carcinogênico e não carcinogênico.'),
        ('Parâmetros toxicológicos', 'RfDo', 'Dose de Referência Oral. Parâmetro usado na avaliação de risco não carcinogênico por ingestão.'),
        ('Parâmetros toxicológicos', 'RfCi', 'Concentração de Referência por Inalação. Parâmetro usado na avaliação de risco não carcinogênico por inalação.'),
        ('Parâmetros toxicológicos', 'SFo', 'Fator de Slope Oral. Parâmetro usado no cálculo de risco carcinogênico por exposição oral.'),
        ('Parâmetros toxicológicos', 'IUR', 'Unidade de Risco por Inalação. Parâmetro usado no cálculo de risco carcinogênico por inalação.'),
        ('Parâmetros físico-químicos', 'FisQui', 'Base de parâmetros físico-químicos utilizada pela CETESBRisk para descrever comportamento ambiental, volatilização, solubilidade, particionamento e transporte.'),
        ('Parâmetros físico-químicos', 'Constante de Henry', 'Parâmetro que expressa a tendência de uma substância transferir-se da fase aquosa para a fase gasosa. É relevante para volatilização e intrusão de vapores.'),
        ('Parâmetros físico-químicos', 'Koc', 'Coeficiente de partição carbono orgânico-água. Indica a tendência da substância se associar à matéria orgânica do solo.'),
        ('Parâmetros físico-químicos', 'Kd', 'Coeficiente de distribuição solo-água. Usado para estimar a partição da substância entre fase sólida e fase aquosa.'),
        ('Parâmetros físico-químicos', 'Kow / log Kow', 'Coeficiente de partição octanol-água. Relacionado à hidrofobicidade da substância e ao potencial de partição em fases orgânicas.'),
        ('Parâmetros físico-químicos', 'Csat', 'Concentração de Saturação. Concentração acima da qual pode haver limitação física de solubilidade ou particionamento no meio avaliado.'),
        ('Critérios regulatórios', 'MCL', 'Maximum Contaminant Level. Padrão de potabilidade adotado pela USEPA para água destinada ao consumo humano.'),
        ('Critérios regulatórios', 'Potabilidade', 'Critérios ou padrões aplicáveis à qualidade da água destinada ao consumo humano.'),
        ('Modelagem de exposição', 'Lgw', 'Profundidade do nível d’água utilizada em modelos de intrusão de vapores. Na CETESBRisk v4.00, há restrição para evitar combinações fisicamente inconsistentes com a franja capilar e fundações.'),
        ('Modelagem de exposição', 'Intrusão de Vapores', 'Migração de vapores de substâncias voláteis presentes no solo ou água subterrânea para ambientes internos de edificações.'),
        ('Modelagem de exposição', 'Johnson & Ettinger', 'Modelo utilizado para estimar intrusão de vapores a partir de fontes em solo ou água subterrânea. Possui premissas e limitações, especialmente em áreas com nível d’água raso.')
    ]

    df_gloss = pd.DataFrame(glossary, columns=['Categoria','Termo','Definição'])
    termo = st.text_input('Pesquisar termo ou palavra-chave', '')
    if termo:
        mask = (
            df_gloss['Termo'].str.contains(termo, case=False, na=False) |
            df_gloss['Definição'].str.contains(termo, case=False, na=False) |
            df_gloss['Categoria'].str.contains(termo, case=False, na=False)
        )
        df_view = df_gloss[mask].copy()
    else:
        df_view = df_gloss.copy()

    for categoria in df_view['Categoria'].drop_duplicates():
        st.subheader(categoria)
        for _, r in df_view[df_view['Categoria'] == categoria].iterrows():
            with st.expander(r['Termo']):
                st.write(r['Definição'])

    st.download_button(
        'Baixar glossário em CSV',
        data=df_gloss.to_csv(index=False).encode('utf-8-sig'),
        file_name='glossario_tecnico_cetesbrisk_explorer.csv',
        mime='text/csv'
    )

elif page=='Downloads e notas':
    st.title('Downloads e notas metodológicas'); st.write('Bases utilizadas pelo aplicativo.')
    with open(DATA_PATH,'rb') as f: st.download_button('Baixar base analítica em Excel',data=f,file_name='Analise_detalhada_CETESBRisk_2023_vs_2026_CMA.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    st.subheader('Notas metodológicas')
    st.markdown('''- A tendência de CMA é estimativa técnica, não recálculo oficial da planilha CETESB.\n- Impactos diretos foram considerados apenas para RfDo, RfCi, SFO e IUR.\n- Parâmetros físico-químicos têm impacto modelo-dependente.\n- A classificação mestre foi ajustada para refletir o parecer técnico consolidado.''')
