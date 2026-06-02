# CETESBRisk Explorer v1.0

Versão pública consolidada.

Atualizações principais:
- Aba Início com textos técnicos ampliados das Classes A, B, C, D1 e D2.
- Ressalva metodológica sobre aplicabilidade das bases FisQui/FatTox nas quatro planilhas CETESBRisk v4.00.
- Nova aba Glossário Técnico com pesquisa e download em CSV.
- Link do LinkedIn do autor na barra lateral.
- Dashboard com rankings técnicos e síntese interpretativa por SQI.

## Rodar localmente
```bash
streamlit run app.py --server.port 8510
```

## v1.1
- Aba "Pesquisa SQI e impacto CMA" com busca única por nome da SQI, nome parcial ou CAS Number.
- Resultado exibe nome, CAS Number e classe no rótulo de seleção.

## v1.2
- Corrigida a exibição da etiqueta de classe na aba de pesquisa.
- Busca única por nome da SQI, nome parcial ou CAS Number.
- Quando houver múltiplos resultados, o app mostra apenas um seletor de resultados encontrados.

## v1.3
- A aba "Pesquisa SQI e impacto CMA" agora abre apenas com uma barra de pesquisa.
- A lista de resultados só aparece depois que o usuário digita nome, nome parcial ou CAS Number.
- Quando há apenas um resultado, ele é selecionado automaticamente.
