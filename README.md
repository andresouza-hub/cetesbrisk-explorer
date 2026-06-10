# CETESBRisk Explorer

Aplicativo Streamlit para apoiar a interpretação técnica das alterações entre as planilhas CETESBRisk V3.03 e V4.01, com foco nas bases FisQui e FatTox, além de registrar correções pontuais de fórmulas identificadas entre as versões.

## Versão atual

**v2.2 - regra objetiva com densidade >= 5%, auditoria de fórmulas e triagem por vias/matrizes/receptores**

Esta versão implementa a classificação objetiva e hierárquica das SQIs nas classes A/B/C/D1/D2, conforme decisão metodológica consolidada, documenta quatro correções de fórmulas e amplia a síntese técnica por SQI com vias, matrizes e receptores potencialmente afetados.

## Regra de classificação implementada

A classificação A/B/C/D1/D2 segue regra objetiva e hierárquica. A classificação reflete a alteração considerada mais relevante para cada SQI, seguindo a hierarquia:

```text
D1 > D2 > C > B > A
```

Regras aplicadas:

- **Classe A**: sem alteração considerada material para classificação; inclui casos em que a única alteração identificada foi densidade com variação em módulo inferior a 5%.
- **Classe B**: alterações menores; inclui alteração isolada de densidade quando a variação em módulo for igual ou superior a 5%, além de outras alterações consideradas pela regra objetiva que não se enquadram como D1, D2 ou C.
- **Classe C**: alteração em MCL ou Potabilidade, desde que não exista alteração D1 ou D2 para a mesma SQI.
- **Classe D1**: qualquer alteração real em parâmetro diretamente associado à toxicidade e ao cálculo de risco, incluindo RfDo, RfCi, SFO, IUR, classificação carcinogênica e mutagenicidade.
- **Classe D2**: alteração físico-química relevante, exceto densidade, com variação em módulo igual ou superior a 5%.

A densidade é tratada como parâmetro auxiliar. Ela não classifica a SQI como D2 isoladamente. Alterações de densidade inferiores a 5%, quando isoladas, não alteram a classe da SQI.

## Contagem final por classe

| Classe | Nº de SQIs |
|---|---:|
| A | 725 |
| B | 14 |
| C | 15 |
| D1 | 52 |
| D2 | 13 |
| **Total** | **819** |

## Principais alterações desta versão

- Remoção da classificação operacional baseada em overrides por CAS ou nome para a classificação final.
- Correção do enquadramento de cis-1,2-DCE e trans-1,2-DCE, que deixam de ser Classe C por não apresentarem alteração em MCL ou Potabilidade.
- Implementação do limiar de materialidade para densidade em módulo: `|Δ| >= 5%`.
- Atualização dos textos explicativos da aba Início.
- Atualização da síntese técnica estruturada da aba Pesquisa SQI e impacto CMA, preservando o formato detalhado já existente e adicionando vias, matrizes e receptores potencialmente afetados.
- Atualização do Dashboard geral, rankings, tabela consolidada, quadro de densidade e notas metodológicas.
- Inclusão de arquivo de auditoria da classificação na pasta `data/`.
- Inclusão da aba **Alterações em fórmulas**, documentando quatro correções: 9^3 → 10^3 em CMA Solo Cr, Cf.GasSolo.14 → Cf.GasSolo.4 em CMA AR Cr-Ad, remoção de EV.c em FI e correção estrutural material em CMA AS Cr!J56.
- Manutenção da melhoria textual MF-017 em Highlights Manual CETESB, separando a limitação operacional do parâmetro Lgw da recomendação técnica de linhas de evidência para intrusão de vapores.

## Arquivos do projeto

```text
app.py
README.md
requirements.txt
.streamlit/config.toml
data/Analise_detalhada_CETESBRisk_2023_vs_2026_CMA.xlsx
data/auditoria_classificacao_SQIs_CETESBRisk_regra_objetiva_v2.xlsx
```

## Observação importante

A ferramenta não recalcula avaliações de risco e não substitui a análise crítica do profissional responsável. A classificação A/B/C/D1/D2 é uma triagem técnica para priorizar revisões dirigidas, considerando exclusivamente a atualização da planilha e das bases internas comparadas.


## Alterações em fórmulas

A aba **Alterações em fórmulas** não classifica SQIs. Ela documenta correções pontuais de fórmulas identificadas entre versões da CETESBRisk e auxilia a triagem de estudos anteriores que possam ter utilizado células, rotas ou posições afetadas. Essas correções não alteram, por si só, a classificação A/B/C/D1/D2, mas podem justificar reavaliação pontual quando a célula corrigida tiver sido efetivamente utilizada em cálculo de risco ou CMA.


## v2.2 - atualização consolidada

Principais ajustes da v2.2:

- Padronização da comparação para **CETESBRisk V3.03 × V4.01**.
- Remoção dos indicadores percentuais com setas no item de resultado da classificação da aba Início.
- Reorganização da aba Início, com o bloco **Alterações em fórmulas** antes do fluxo decisório.
- Ampliação do fluxo decisório para incluir células, rotas, módulos e posições de SQI afetadas por correções de fórmula.
- Inclusão da correção estrutural material em **CMA AS Cr!J56**.
- Atualização do quadro-resumo da aba Alterações em fórmulas para quatro correções identificadas.
- Ajuste da redação da correção **FI!I8**, explicitando que se refere à inalação de partículas de solo por crianças.
- Inclusão, na aba Pesquisa SQI e impacto CMA, de bloco interpretativo sobre vias, matrizes e receptores potencialmente afetados.
