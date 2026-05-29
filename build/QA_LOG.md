# QA LOG — Psi Planner (ciclo de qualidade)

Objetivo: planilha funcional + aprovação de TODOS os agentes de QA visual.
Artefato: `PsiPlanner/Psi Planner.xlsm` (pipeline build→inject→diag→tests).

## Iteração 1 — REPROVADO (3 agentes: arte/claro, dark/acessibilidade, consistência)

Funcional: `tests.py` PASS (0 falhas). KPIs 7 / R$870 / R$350. 21 módulos VBA.

Erros encontrados (confirmados por inspeção própria dos previews):

| # | Sev | Aba | Erro | Causa | Correção |
|---|-----|-----|------|-------|----------|
| 1 | Bloqueador | Relatórios | "RR$ 150,00" (R duplicado) em Receita por convênio (I6) | `mod_Convenio` faz `.NumberFormat="R$ #,##0.00"` via VBA; o `$` é localizado p/ "R$" → "R"+"R$" | aspas no literal: `"""R$"" #,##0.00"` |
| 2 | Bloqueador | Financeiro/Relatórios gráficos | doughnuts/barras com múltiplos CINZAS indistinguíveis; "Receita x Despesa" 2 barras iguais; pior no dark (fatia escura some) | paleta monocromática cinza+vinho | paleta CATEGÓRICA (vinho/ouro/sage/terracota/slate) + barra com pontos coloridos + eixo min=0 + título overlay=False |
| 3 | Maior | Calendário | texto de evento de 3 linhas cortado ("Confirmado") | fonte 9 + col 20 faz "Presencial · Confirmado" quebrar; linha 30px corta 3ª linha | fonte do evento → 8 (cabe em 1 linha) + linha 34px |
| 4 | Maior | Menu | acento de "CLÍNICA" cortado pela faixa | título 32pt encosta no topo da linha 2 (sob a faixa) | título 32→30 + linha 2 do Menu → 56 |
| 5 | Maior | Aniversariantes | legenda do donut cortada (Jan–Ago); título sobre o anel | donut pequeno (9×6) c/ 12 categorias | ampliar (13×8) + legenda menor |
| 6 | Maior | Geral (Financeiro/Pacientes) | previews "pequenos"/ilegíveis p/ auditar; acentos finos (í/ã) somem no raster | zoom 100% + sheets largas escalam | diag: zoom 150% antes do CopyPicture (mais pixels/glifo) |
| 7 | Menor | Relatórios/Config | acentos do subtítulo itálico parecem sumir ("Visao"/"Preferencias") | render de itálico Inter em raster pequeno | resolver via #6 (zoom); reavaliar |
| 8 | Falso-positivo | Financeiro_graficos | "sem faixa de cabeçalho" | é um RECORTE da área de gráficos (linha 76+), não uma aba; a faixa está no topo da aba Financeiro | documentar; instruir QA |

Ações desta rodada: corrigir 1–6 no pipeline, re-rodar build/inject/diag/tests, re-submeter aos 3 agentes.

## Iteração 2 — correções aplicadas e verificadas por inspeção própria

- #1 RR$ → **CORRIGIDO** (Relatórios mostra "R$ 150,00"). Aspas no literal do NumberFormat.
- #2 Charts → doughnuts **CORRIGIDOS** (Pago=verde/Pendente=vinho; Distribuição/Convênio categóricos; bordas off-white; ótimos no dark). Eixo min=0 e título overlay=False OK. **PENDENTE**: barra "Receita x Despesa" ainda saiu com 2 barras iguais (cor por-ponto não pegou) → fix: `varyColors=True` + dPt sem solidFill de série.
- #3 Calendário → **CORRIGIDO** (fonte 8 + linha 34: eventos em 2 linhas, sem corte).
- #4 Menu título → **CORRIGIDO** (30pt + linha 56: "CLÍNICA" com acento e respiro).
- #5 Aniversariantes donut → ampliado (13×8).
- #6 Zoom 150% → **CORRIGIDO** (previews em alta densidade: Relatorios 2014px, Pacientes 3247px).
- #7 acentos itálico → **DIAGNOSTICADO**: valores das células TÊM acentos; o itálico faux do Inter variável dropa ã/í inconsistentemente. Teste PIL: Inter reto renderiza tudo. Fix: `institucional.italic=false`.

Pendências p/ iteração 3: barra por-ponto + itálico off. Re-rodar e re-submeter.

## Iteração 3 — APROVADO pelos 3 agentes ✅

Correções finais aplicadas e verificadas:
- Barra "Receita x Despesa" → **CORRIGIDA** (Receita verde / Despesa vinho) via `varyColors=True` + dPt sem fill de série.
- Acentos do subtítulo → **CORRIGIDO**: diagnóstico final = bug de rendering do **Inter VARIÁVEL** isolado às células `institucional` (i sem ponto + ã/í somem). O resto do Inter renderiza ã/í normalmente (provado: "Sessão"/"Cartão" no Financeiro). Fix: `institucional.applied = "Segoe UI Light"` (fallback sancionado pelo brief; tagline pequena). "Visão consolidada da clínica" / "Preferências" agora corretos.

Funcional: `tests.py` PASS (0 falhas). Veredito dos 3 agentes: **APROVADO / APROVADO / APROVADO**.

Itens MENORES aceitos (não-bloqueantes, aprovados pelos agentes) — polimento opcional futuro:
- Legendas dos doughnuts (Receita por convênio / Distribuição) com marcador colado ao texto → aumentar plot/legenda.
- Rótulos do eixo X dos gráficos de linha pequenos/colados ("OutNovDez") → fonte de eixo maior.
- Painéis laterais (AVISOS/RESUMO) encostam na margem do RECORTE do preview (não é defeito do arquivo) → ampliar range de captura.
- Calendário: baseline de "Terça"/"Sábado" no cabeçalho levemente desalinhado.
- Aniversariantes: 2 verdes adjacentes no donut de 12 meses (bordas já separam).

CICLO CONCLUÍDO: planilha funcional + todos os agentes de QA visual aprovaram.
