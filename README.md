# CETESBRisk Explorer - correção robustez preservando classificação

Correções desta versão:
- Preserva a regra de classificação aprovada anteriormente, mantendo as contagens A=296, B=471, C=18, D1=23, D2=11.
- Mantém os ajustes de robustez nos rankings de parâmetros, usando apenas alterações reais via `changed()`.
- Mantém grupos prioritários filtrados separadamente, evitando clorobenzenos/nitrobenzenos em BTEX sem alterar a classificação global.
- Corrige erro de string no texto do ponto de atenção sobre Johnson & Ettinger.

Observação metodológica:
A tentativa anterior de tornar `NAME_OVERRIDES` mais restritivo alterou a classificação global. Essa alteração foi revertida, pois o usuário decidiu manter a regra vigente.
