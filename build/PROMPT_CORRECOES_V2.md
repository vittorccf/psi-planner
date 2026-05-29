# Prompt — Pacote de correções v2 (Psi Planner / Clínica Vita)

> Cole este bloco em uma nova conversa do Claude Code com a planilha aberta no
> mesmo workspace. Ele é auto-suficiente: descreve persona, constraints,
> entregáveis aba por aba, pipeline de execução e validação por sub-agente.

---

```
PERSONA: Senior UX/UI Designer + Engineer Excel/openpyxl/VBA.
ARTEFATO ALVO: PsiPlanner/Psi Planner.xlsm
                (saída de build_xlsx.py → inject_vba.py).

CONSTRAINTS NÃO NEGOCIÁVEIS
- Fontes instaladas nesta máquina: Cambria (títulos) + Segoe UI
  Light/Semilight (corpo) + Segoe MDL2 Assets (ícones).
  Tan Pearl e Dream Avenue NÃO existem — usar sempre os fallbacks.
- Paleta atual: preto 000000, mute 3B3B3B, line E2E2E2,
  cinza_claro F4F4F4, vinho 5A1E2A, vinho_escuro 3A1018.
- KPIs / avisos / MEI são fórmulas nativas. Macros só para interação.
- AccessVBOM já está habilitado (HKCU). Sempre matar EXCEL órfão
  entre runs COM (Stop-Process EXCEL antes de cada `python ...`).
- UserForms criados via COM Designer não recompilam — usar fluxo
  InputBox quando precisar de input do usuário.

ENTREGÁVEL — aplicar e validar:

1. PASTÉIS NOVOS NA PALETA (visual_spec.json/meta.palette)
     pastel_green:  D9EAD3   (sucesso / confirmado / ativo / SIM)
     pastel_red:    F4CCCC   (cancelado / pendente / inativo / NÃO)
     pastel_yellow: FFF2CC   (lembrete / remarcado)
   Os tints atuais (done / wait / alert) passam a apontar para
   esses pastéis para manter o resto do sistema coerente.

2. CALENDÁRIO (range C6:H21)
   - "Confirmado" OU "Atendido" OU "Realizado" → fill pastel_green,
     font color 1F3B1F.
   - "Cancelado" OU "Faltou"                    → fill pastel_red,
     font color 3A1018.
   - "Remarcado"                                → fill pastel_yellow,
     font color 5A4A1E.
   - Expandir o dropdown de status para incluir
     "Confirmado / Atendido / Realizado / Remarcado / Cancelado / Faltou".

3. ANIVERSARIANTES
   - Status (G7:G56): "Ativo" → pastel_green;
                       "Inativo" → pastel_red;
                       "Pausado" → pastel_yellow;
                       "Alta" → surface_deep neutro.
   - Mensagem enviada (H7:H56):
       "Sim" → pastel_green; "Não" → pastel_red.
   - Gráfico doughnut "Aniversariantes por mês":
       12 cores distintas (rampa terracota → vinho → oliva escuro),
       com o mês corrente forçado a vinho 5A1E2A para destacar.

4. FINANCEIRO (range G7:G56)
   - "Pago"      → pastel_green (font 1F3B1F)
   - "Pendente"  → pastel_red   (font 5A1E2A)
   - "Cancelado" → pastel_red   (font 3A1018)
   - "Lembrete"  → pastel_yellow (font 5A4A1E)

5. TIPOGRAFIA
   - Títulos das abas e da clínica em CAIXA ALTA + bold + maiores:
       titulo_clinica  26 → 32  (bold true)
       titulo_aba      18 → 26  (bold true)
       Conteúdo (cells): "Clínica Vita" → "CLÍNICA VITA";
         "Financeiro" → "FINANCEIRO"; "Pacientes" → "PACIENTES";
         "Calendário" → "CALENDÁRIO";
         "Aniversariantes" → "ANIVERSARIANTES";
         "Relatórios" → "RELATÓRIOS";
         "Configurações" → "CONFIGURAÇÕES".
   - Subtítulos em NEGRITO (fonte "subtitulo" → bold true):
       LANÇAMENTOS / INDICADORES / DESPESAS / GRÁFICOS /
       ACESSO RÁPIDO / AVISOS / RESUMO DO DIA / LISTA DE ESPERA /
       DO MÊS / DADOS DA CLÍNICA / PARÂMETROS / PRIVACIDADE.
   - Aumentar row_height da linha 2 para 50–56 para acomodar
     os títulos maiores sem cortar descendentes.

6. MENU INICIAL — botões de Acesso Rápido
   No inject_vba.py, estilo "nav":
       fill (90,30,42)        ← vinho_elegante
       line None              ← sem borda
       txt  (255,255,255)
       font "Segoe UI Semibold"
       bold True
   - O glifo (Segoe MDL2 Assets) também precisa ser branco
     (não vinho) para legibilidade sobre o fundo escuro.
   - Estilos "primary" e "back" permanecem inalterados.

7. CONFIGURAÇÕES — toggle de privacidade do convênio
   - Inserir bloco "PRIVACIDADE" (B17, subtitulo bold).
   - Linha "Exibir informações de convênio" em B18.
   - Célula C18: dropdown {Sim, Não}, default "Sim",
     named_range CFG_ExibirConvenio.
   - Novo módulo VBA mod_Convenio com `AplicarExibicaoConvenio()`:
       quando CFG_ExibirConvenio = "Não"
         · Pacientes:   coluna J (Convênio) → Hidden = True
         · Relatórios:  H6 valor → "" ; I6 NumberFormat = ";;;"
         · Relatórios:  ChartObjects(3) (Receita por convênio)
                        → Visible = False
       quando "Sim" → restaura tudo.
   - Hooks em ThisWorkbook:
       Workbook_Open         → AplicarExibicaoConvenio
       Workbook_SheetChange  → reagir quando Target ∩
                               Configurações!C18 não for vazio.

PIPELINE DE EXECUÇÃO (na ordem)
   a) Editar PsiPlanner/build/visual_spec.json
       (paleta + fontes + valores em CAIXA ALTA + bloco PRIVACIDADE)
   b) Editar PsiPlanner/build/build_xlsx.py
       (CF com pastéis + dropdown C18 + named range
        CFG_ExibirConvenio + gráfico aniversariantes colorido)
   c) Editar PsiPlanner/build/inject_vba.py
       (STYLES["nav"] vinho/bold + glifo branco + mod_Convenio
        + hooks Workbook_Open / Workbook_SheetChange)
   d) Stop-Process EXCEL ; python build_xlsx.py
   e) Stop-Process EXCEL ; python inject_vba.py
   f) Stop-Process EXCEL ; python diag.py
       (regenera os 8 PNGs em PsiPlanner/*.png)
   g) Se o preview de Configurações cortar o bloco PRIVACIDADE,
      ampliar o range em diag.py de "A1:D16" para "A1:D20".

VALIDAÇÃO POR SUB-AGENTE
Spawnar um sub-agente que abre via Read:
  - Menu_Inicial_preview.png
  - Calendario_preview.png
  - Aniversariantes_preview.png
  - Financeiro_preview.png
  - Configuracoes_preview.png
E confere o checklist:
  1. Pastéis aplicados nos status corretos (3 abas).
  2. Gráfico de aniversariantes colorido (não monocromático).
  3. Títulos em caixa alta e visivelmente maiores.
  4. Subtítulos em negrito.
  5. Botões do Menu Inicial em vinho com texto branco.
  6. Bloco PRIVACIDADE presente em Configurações com
     "Exibir informações de convênio: Sim".
Resposta em < 200 palavras, com OK/FALHOU + 1 linha de evidência
para cada item.

ANTI-PADRÕES (rejeitar)
- Trocar pastéis por cores saturadas (vermelho HTML, verde flúor).
- Usar "Tan Pearl" / "Dream Avenue" (não instaladas).
- Mexer em UserForm via COM (quebra com erro fm20).
- Apagar fórmulas SUMIF/SUMIFS dos KPIs em vez de usar NumberFormat
  ";;;" para esconder valor sob o toggle de convênio.
- Definir áreas de impressão (_xlnm.Print_Area) no .xlsm — trava
  abertura nesta máquina.
```

---

## Como rodar

```powershell
# 1) Aplicar edits nos 3 arquivos conforme o prompt acima
# 2) Pipeline:
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
python D:\Vittor\Projetos\Saluti\PsiPlanner\build\build_xlsx.py

Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
python D:\Vittor\Projetos\Saluti\PsiPlanner\build\inject_vba.py

Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
python D:\Vittor\Projetos\Saluti\PsiPlanner\build\diag.py
```

Saída final: `PsiPlanner/Psi Planner.xlsm` + 8 PNGs em `PsiPlanner/`.
