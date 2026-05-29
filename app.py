from pathlib import Path
from datetime import datetime
import pandas as pd
import streamlit as st
import plotly.express as px

st.set_page_config(page_title='CETESBRisk Explorer', page_icon='🧪', layout='wide')
DATA_PATH = Path(__file__).parent / 'data' / 'Analise_detalhada_CETESBRisk_2023_vs_2026_CMA.xlsx'

FIXED_CLASS_COUNTS = {'A':670,'B':3,'C':25,'D1':31,'D2':90}
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

comp,incl,remov,alt,impacto=load_data(); master=build_master(comp)

st.sidebar.title('🧪 CETESBRisk Explorer')
st.sidebar.markdown('**Desenvolvido por André Souza**  \nEspecialista em GAC')
st.sidebar.divider()
page=st.sidebar.radio('Navegação',['Início','Dashboard geral','Pesquisa SQI e impacto CMA','Grupos prioritários','Highlights Manual CETESB','Downloads e notas'])
st.sidebar.caption('v0.4 · revisão executiva')

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
    with st.expander('Classe A — Sem alteração material relevante'): st.markdown('Compostos sem alterações significativas nos parâmetros capazes de repercutir materialmente sobre risco ou CMA. Exemplo: benzeno e grande parte dos BTEX clássicos.')
    with st.expander('Classe B — Alterações menores'): st.markdown('Compostos com pequenas alterações paramétricas, normalmente sem repercussão prática relevante. Exemplos: PCE e TCE.')
    with st.expander('Classe C — Alterações regulatórias/potabilidade'):
        st.warning('A Classe C não significa automaticamente que o risco calculado mudou. Ela indica alterações em MCL, potabilidade ou critérios regulatórios, sem alteração material dos parâmetros principais de cálculo de risco/CMA.')
        st.markdown('''**Interpretação prática:**\n- pode exigir revisão de enquadramento regulatório;\n- pode exigir atualização de tabelas comparativas;\n- pode ser relevante para potabilidade ou ingestão de água subterrânea;\n- **não implica automaticamente recálculo da avaliação de risco**.''')
    with st.expander('Classe D1 — Alterações toxicológicas relevantes'): st.markdown('Alterações em parâmetros como RfDo, RfCi, SFo, IUR, carcinogenicidade e mutagenicidade. Exemplos: cloreto de vinila, 1,1-DCE, PFOS, PFOA e cromo VI.')
    with st.expander('Classe D2 — Alterações físico-químicas relevantes'): st.markdown('Alterações em parâmetros capazes de influenciar volatilização, transporte, particionamento, exposição, intrusão de vapores e Csat. Exemplo: TPH alifático leve C5-C8.')
    st.markdown('''## 6. Fluxo decisório: quando revisar uma avaliação anterior?
**Perguntas de triagem recomendadas:**
1. O modelo conceitual possui SQIs da Classe D1, como PFAS, cloreto de vinila ou 1,1-DCE?
2. Há SQIs da Classe D2 associadas a volatilização, intrusão de vapores ou transporte em fase vapor?
3. A avaliação anterior estava próxima da margem de aceitabilidade, como HQ ≈ 1 ou risco ≈ 1E-05?
4. A avaliação baseou-se predominantemente em frações genéricas de TPH em vez de compostos individuais de petróleo?

Se a resposta for **sim** para uma ou mais perguntas, recomenda-se **revisão dirigida**. Se todas forem **não**, a atualização da planilha, isoladamente, não indica revisão automática.''')
    st.markdown('## 7. Resultado da classificação')
    r1,r2,r3,r4,r5=st.columns(5); r1.metric('Classe A','670'); r2.metric('Classe B','3'); r3.metric('Classe C','25'); r4.metric('Classe D1','31'); r5.metric('Classe D2','90')
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
    c1,c2,c3,c4=st.columns(4); c1.metric('Pontos de dados analisados','29.484'); c2.metric('Com alteração','1.009','3,4% do total'); c3.metric('Sem alteração','16.565'); c4.metric('Sem valor/vazios','11.910')
    st.subheader('Distribuição consolidada dos compostos por classe de alteração')
    chart_df=pd.DataFrame({'Classe':list(FIXED_CLASS_COUNTS.keys()),'N compostos':list(FIXED_CLASS_COUNTS.values())})
    fig=px.bar(chart_df,x='Classe',y='N compostos',color='Classe',text='N compostos',hover_data={'Classe':True,'N compostos':True},title='Compostos por classe de alteração'); fig.update_traces(textposition='outside'); fig.update_layout(yaxis_title='Nº de compostos',xaxis_title='Classe')
    st.plotly_chart(fig,use_container_width=True)
    st.subheader('Composição dos pontos de dados avaliados')
    flow_df=pd.DataFrame({'Categoria':['Com alteração','Sem alteração','Sem valor/vazios'],'N':[1009,16565,11910]})
    fig2=px.bar(flow_df,x='Categoria',y='N',color='Categoria',text='N',title='Status dos 29.484 pontos de dados avaliados'); fig2.update_traces(textposition='outside')
    st.plotly_chart(fig2,use_container_width=True)
    st.subheader('Tabela consolidada por composto')
    classes=st.multiselect('Filtrar classes',['A','B','C','D1','D2'],default=['A','B','C','D1','D2']); filtered=master[master['Classe'].isin(classes)].copy()
    st.dataframe(filtered,use_container_width=True,hide_index=True)
    st.download_button('Baixar tabela filtrada',data=filtered.to_csv(index=False).encode('utf-8-sig'),file_name='cetesbrisk_master_filtrado.csv',mime='text/csv')

elif page=='Pesquisa SQI e impacto CMA':
    st.title('Pesquisa SQI e impacto potencial na CMA')
    options=master.sort_values('Composto')['Composto'].dropna().unique().tolist(); selected=st.selectbox('Selecione uma substância',options,index=options.index('Benzene') if 'Benzene' in options else 0)
    row=master[master['Composto']==selected].iloc[0]; cas=row['CAS']; cls=row['Classe']
    st.header(selected); st.markdown(badge(cls),unsafe_allow_html=True)
    m1,m2,m3=st.columns(3); m1.metric('CAS',str(cas)); m2.metric('Classe',cls); m3.metric('Nº alterações',int(row['Nº alterações']))
    st.subheader('Interpretação'); st.write(row['Recomendação'])
    detail=comp[(comp['CAS'].astype(str)==str(cas))].copy(); show_cols=['Fonte','Parâmetro','Valor 2023','Valor 2026','Status','Variação % Parâmetro','Tendência esperada da CMA','Estimativa variação CMA %','Observação técnica']
    st.subheader('Comparativo por parâmetro'); st.dataframe(detail[show_cols],use_container_width=True,hide_index=True)
    ch=detail[detail.apply(changed,axis=1)]; lines=[]
    for _,r in ch.iterrows(): lines.append(f"- {r.get('Parâmetro','')}: tendência CMA = {r.get('Tendência esperada da CMA','')}; estimativa = {r.get('Estimativa variação CMA %','')}; observação: {r.get('Observação técnica','')}")
    cma_text='Não foram identificadas alterações materiais com tendência estimável de impacto sobre CMA para esta SQI.' if not lines else '\n'.join(lines)
    synthesis=f"Com base na comparação entre as versões CETESBRisk v3.03 e v4.00, a SQI {selected} foi enquadrada na Classe {cls} ({CLASS_DESCRIPTIONS[cls]['descricao']}). {CLASS_DESCRIPTIONS[cls]['recomendacao']} Síntese do impacto potencial sobre CMA: {cma_text}"
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

elif page=='Downloads e notas':
    st.title('Downloads e notas metodológicas'); st.write('Bases utilizadas pelo aplicativo.')
    with open(DATA_PATH,'rb') as f: st.download_button('Baixar base analítica em Excel',data=f,file_name='Analise_detalhada_CETESBRisk_2023_vs_2026_CMA.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    st.subheader('Notas metodológicas')
    st.markdown('''- A tendência de CMA é estimativa técnica, não recálculo oficial da planilha CETESB.\n- Impactos diretos foram considerados apenas para RfDo, RfCi, SFO e IUR.\n- Parâmetros físico-químicos têm impacto modelo-dependente.\n- A classificação mestre foi ajustada para refletir o parecer técnico consolidado.''')
