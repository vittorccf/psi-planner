# -*- coding: utf-8 -*-
"""Gera o HTML de especificacao visual da planilha Psi Planner com previews embutidos em base64."""
import base64, os

DELIV = r"D:\Vittor\Projetos\Saluti\PsiPlanner"
OUT_HTML = os.path.join(DELIV, "build", "psi_planner_spec.html")

def b64(name):
    with open(os.path.join(DELIV, name), "rb") as f:
        return "data:image/png;base64," + base64.b64encode(f.read()).decode()

IMG = {
    "menu": b64("Menu_Inicial_preview.png"),
    "fin": b64("Financeiro_preview.png"),
    "fin_g": b64("Financeiro_graficos_preview.png"),
    "pac": b64("Pacientes_preview.png"),
    "cal": b64("Calendario_preview.png"),
    "aniv": b64("Aniversariantes_preview.png"),
    "rel": b64("Relatorios_preview.png"),
    "cfg": b64("Configuracoes_preview.png"),
}

# ---- paleta ----
PALETTE = [
    ("Preto Principal", "#000000", "Títulos, textos principais, ícones minimalistas", "dark"),
    ("Cinza Escuro", "#3B3B3B", "Menus laterais, indicadores, botões discretos, gráficos", "dark"),
    ("Cinza Médio", "#A6A6A6", "Bordas finas, subtítulos, linhas da agenda", "dark"),
    ("Cinza Claro", "#F4F4F4", "Fundo de cards, áreas de respiro, campos editáveis", "light"),
    ("Vinho Elegante", "#5A1E2A", "Botões ativos, destaques do dashboard, indicadores", "dark"),
    ("Vinho Escuro", "#3A1018", "Hover dos botões, alertas sofisticados, premium", "dark"),
]

def swatches():
    out = []
    for nome, hexv, uso, kind in PALETTE:
        txt = "#FFFFFF" if kind == "dark" else "#3B3B3B"
        bord = "border:1px solid #E2E2E2;" if kind == "light" else ""
        out.append(f"""
        <div class="sw">
          <div class="sw-chip" style="background:{hexv};color:{txt};{bord}">{hexv}</div>
          <div class="sw-name">{nome}</div>
          <div class="sw-uso">{uso}</div>
        </div>""")
    return "\n".join(out)

# ---- abas: (id imagem, titulo, subtitulo, lista de bullets em grupos) ----
def li(items):
    return "".join(f"<li>{i}</li>" for i in items)

def grupos(g):
    out = []
    for titulo, items in g:
        out.append(f'<div class="grp"><h4>{titulo}</h4><ul>{li(items)}</ul></div>')
    return '<div class="grupos">' + "".join(out) + "</div>"

ABAS = [
    {
        "img": "menu", "num": "01", "titulo": "Menu Inicial",
        "sub": "Dashboard principal — porta de entrada da plataforma",
        "grupos": grupos([
            ("Cabeçalho", ["Nome da clínica em fonte editorial fina", "Frase institucional discreta", "Data automática do dia"]),
            ("Cards de resumo rápido", ["Pacientes ativos", "Receita do mês", "Valor pendente", "Próxima consulta", "Faltas do mês", "Aniversariantes do mês"]),
            ("Navegação", ["6 botões minimalistas com ícones monocromáticos", "Efeito hover discreto via VBA", "Fundo branco e borda fina cinza"]),
            ("Avisos automáticos", ["Pacientes sem sessão há 30+ dias", "Faturamento próximo do teto MEI", "Pagamentos pendentes", "Aniversariantes da semana"]),
        ]),
    },
    {
        "img": "fin", "num": "02", "titulo": "Financeiro",
        "sub": "Lançamentos, recebimentos e indicadores automáticos",
        "grupos": grupos([
            ("Tabela de lançamentos", ["Data, paciente, tipo de atendimento", "Forma de pagamento: Pix, Dinheiro, Cartão, Convênio", "Valor e Status (Pago, Pendente, Lembrete, Cancelado)", "Nota fiscal, número da NF e observações"]),
            ("Indicadores", ["Total recebido e total pendente", "Lucro líquido automático", "Percentual de inadimplência", "Faturamento anual"]),
            ("Limite MEI", ["Indicador do teto de R$ 81.000/ano", "Alerta discreto a 70% (cinza escuro)", "Alerta preto/branco acima de 90%"]),
            ("Extras", ["Geração manual de recibo", "Controle de despesas por categoria", "Filtro por mês, ano e paciente"]),
        ]),
    },
    {
        "img": "fin_g", "num": "02", "titulo": "Financeiro — Dashboard",
        "sub": "Gráficos minimalistas monocromáticos em tons de cinza",
        "grupos": grupos([
            ("Visualizações", ["Receita mensal", "Receita x despesas", "Pago x pendente", "Evolução financeira", "Distribuição por forma de pagamento"]),
            ("Estilo", ["Gráficos monocromáticos em cinza", "Destaques estratégicos em vinho", "Sem cores vibrantes ou degradês"]),
        ]),
    },
    {
        "img": "pac", "num": "03", "titulo": "Pacientes",
        "sub": "Cadastro clínico organizado e visualmente limpo",
        "grupos": grupos([
            ("Cadastro completo", ["Dados pessoais, contato e endereço", "Idade automática pela data de nascimento", "Convênio, profissão e início do acompanhamento", "Frequência, valor da sessão e status"]),
            ("Controle clínico", ["Histórico e quantidade de sessões", "Total faturado por paciente", "Controle de pacotes e saldo restante"]),
            ("Alertas automáticos", ["Sem sessão há 30+ dias", "Inadimplente", "Aniversariante do mês", "Pacote próximo do fim"]),
            ("Filtros inteligentes", ["Status, convênio, frequência", "Mês de aniversário", "Faixa de faturamento"]),
        ]),
    },
    {
        "img": "cal", "num": "04", "titulo": "Calendário",
        "sub": "Agenda semanal minimalista e funcional",
        "grupos": grupos([
            ("Grade visual", ["Segunda a sábado, 07h às 22h", "Fundo branco, linhas cinza claro", "Cabeçalhos em cinza escuro", "Horários ocupados em vinho suave"]),
            ("Registro de consultas", ["Paciente e horário", "Tipo: Presencial ou Online", "Status: Confirmado, Remarcado, Faltou, Cancelado, Realizado"]),
            ("Recursos automáticos", ["Contagem de consultas do dia", "Taxa de cancelamento e de faltas", "Horários vagos e lista de espera"]),
            ("Ações", ["Adicionar consulta", "Remarcar", "Finalizar atendimento", "Busca rápida por paciente"]),
        ]),
    },
    {
        "img": "aniv", "num": "05", "titulo": "Aniversariantes",
        "sub": "Tela exclusiva e elegante para controle de aniversários",
        "grupos": grupos([
            ("Tabela automática", ["Nome, data de nascimento e idade atual", "Dia do aniversário e telefone", "Status do paciente"]),
            ("Recursos", ["Separação automática por mês", "Destaque discreto para a semana", "Contador de aniversariantes do mês", "Campo Mensagem enviada? (Sim/Não)"]),
            ("Dashboard", ["Aniversariantes por mês", "Pacientes ativos aniversariantes", "Lista dos próximos aniversários"]),
        ]),
    },
    {
        "img": "rel", "num": "06", "titulo": "Relatórios",
        "sub": "Painel analítico minimalista",
        "grupos": grupos([
            ("Relatórios automáticos", ["Receita mensal e anual", "Quantidade de atendimentos", "Pacientes ativos", "Taxa de cancelamento e inadimplência", "Pacientes por convênio", "Comparativo entre meses"]),
            ("Recursos", ["Gráficos monocromáticos", "Filtros por período", "Exportação para PDF", "Indicadores automáticos"]),
        ]),
    },
    {
        "img": "cfg", "num": "07", "titulo": "Configurações",
        "sub": "Parâmetros gerais da plataforma",
        "grupos": grupos([
            ("Identidade", ["Nome da clínica e frase institucional", "Dados do profissional"]),
            ("Parâmetros", ["Valor padrão da sessão", "Teto MEI de referência", "Categorias de despesa"]),
        ]),
    },
]

def aba_pages():
    out = []
    for a in ABAS:
        out.append(f"""
  <section class="page aba">
    <div class="aba-head">
      <span class="aba-num">{a['num']}</span>
      <div>
        <h2>{a['titulo']}</h2>
        <p class="aba-sub">{a['sub']}</p>
      </div>
    </div>
    <div class="mock"><img src="{IMG[a['img']]}" alt="{a['titulo']}"></div>
    {a['grupos']}
  </section>""")
    return "\n".join(out)

HTML = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<style>
  @page {{ size: A4; margin: 16mm 15mm 16mm 15mm; }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  html {{ -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
  body {{
    font-family: "Segoe UI", "Corbel", sans-serif; font-weight: 300;
    color: #3B3B3B; font-size: 11px; line-height: 1.6;
  }}
  h1, h2, h3, h4, .serif {{ font-family: "Constantia", "Cambria", Georgia, serif; }}
  .page {{ page-break-after: always; position: relative; }}
  .page:last-child {{ page-break-after: auto; }}

  /* ---------- CAPA ---------- */
  .cover {{ height: 252mm; display: flex; flex-direction: column; }}
  .cover .kicker {{ letter-spacing: .42em; font-size: 11px; color: #A6A6A6; text-transform: uppercase; }}
  .cover .rule {{ height: 1px; background: #E2E2E2; margin: 20px 0 0; }}
  .cover .center {{ flex: 1; display: flex; flex-direction: column; justify-content: center; }}
  .cover h1 {{ font-size: 64px; font-weight: 400; color: #000; letter-spacing: .01em; line-height: 1.05; }}
  .cover h1 .wine {{ color: #5A1E2A; }}
  .cover .tagline {{ font-style: italic; color: #A6A6A6; font-size: 16px; margin-top: 14px; font-family: "Constantia", Georgia, serif; }}
  .cover .desc {{ margin-top: 34px; max-width: 78%; font-size: 12.5px; color: #3B3B3B; line-height: 1.8; }}
  .cover .meta {{ display: flex; gap: 48px; margin-top: 40px; }}
  .cover .meta b {{ display:block; font-family:"Constantia",Georgia,serif; color:#000; font-size:13px; font-weight:400;}}
  .cover .meta span {{ font-size: 9.5px; letter-spacing: .24em; text-transform: uppercase; color:#A6A6A6;}}
  .cover .foot {{ font-size: 9.5px; letter-spacing: .22em; text-transform: uppercase; color:#A6A6A6; }}
  .cover .wbar {{ height: 4px; width: 64px; background: #5A1E2A; margin-bottom: 28px; }}

  /* ---------- SECTION TITLE ---------- */
  .sec-title {{ display:flex; align-items:baseline; gap:14px; border-bottom:1px solid #E2E2E2; padding-bottom:12px; margin-bottom:26px; }}
  .sec-title .n {{ font-family:"Constantia",Georgia,serif; color:#5A1E2A; font-size:13px; letter-spacing:.1em;}}
  .sec-title h2 {{ font-size:26px; font-weight:400; color:#000; }}

  /* ---------- VISAO GERAL ---------- */
  .overview {{ display:grid; grid-template-columns:1fr 1fr; gap:14px; }}
  .ov-card {{ border:1px solid #EAEAEA; padding:16px 18px; background:#FBFBFB; }}
  .ov-card .t {{ font-family:"Constantia",Georgia,serif; font-size:14px; color:#000; }}
  .ov-card .t .n {{ color:#A6A6A6; font-size:11px; margin-right:8px; }}
  .ov-card p {{ margin-top:6px; color:#6b6b6b; font-size:10.5px; line-height:1.55; }}

  /* ---------- PALETA ---------- */
  .swatches {{ display:grid; grid-template-columns:1fr 1fr 1fr; gap:16px; margin-bottom:34px; }}
  .sw-chip {{ height:78px; border-radius:3px; display:flex; align-items:flex-end; padding:8px 10px; font-size:11px; letter-spacing:.06em; }}
  .sw-name {{ font-family:"Constantia",Georgia,serif; font-size:13px; color:#000; margin-top:9px; }}
  .sw-uso {{ font-size:9.5px; color:#9a9a9a; margin-top:3px; line-height:1.45; }}

  /* ---------- TIPOGRAFIA ---------- */
  .type-grid {{ display:grid; grid-template-columns:1fr 1fr; gap:22px; }}
  .type-card {{ border:1px solid #EAEAEA; padding:22px 24px; }}
  .type-card .lbl {{ font-size:9.5px; letter-spacing:.24em; text-transform:uppercase; color:#A6A6A6; }}
  .type-card .name {{ font-family:"Constantia",Georgia,serif; font-size:30px; color:#000; margin:6px 0 4px; font-weight:400;}}
  .type-card .name.sans {{ font-family:"Segoe UI","Corbel",sans-serif; font-weight:300; }}
  .type-card .role {{ font-size:10.5px; color:#6b6b6b; line-height:1.6; margin-top:10px; }}
  .type-card .chars {{ margin-top:12px; font-size:10px; color:#9a9a9a; }}
  .type-card .chars span {{ display:inline-block; border:1px solid #ECECEC; padding:2px 9px; margin:0 5px 5px 0; border-radius:12px; }}

  /* ---------- ABAS ---------- */
  .aba-head {{ display:flex; gap:18px; align-items:flex-start; border-bottom:1px solid #E2E2E2; padding-bottom:14px; margin-bottom:20px; }}
  .aba-num {{ font-family:"Constantia",Georgia,serif; font-size:34px; color:#EBD7DC; line-height:1; }}
  .aba-head h2 {{ font-size:24px; font-weight:400; color:#000; }}
  .aba-sub {{ font-style:italic; color:#A6A6A6; font-size:12px; font-family:"Constantia",Georgia,serif; margin-top:2px; }}
  .mock {{ border:1px solid #E6E6E6; padding:8px; background:#fff; box-shadow:0 1px 6px rgba(0,0,0,.05); margin-bottom:20px; }}
  .mock img {{ width:100%; height:auto; display:block; }}
  .grupos {{ display:grid; grid-template-columns:1fr 1fr; gap:10px 26px; }}
  .grp h4 {{ font-size:11px; font-weight:400; color:#5A1E2A; letter-spacing:.04em; margin-bottom:5px; padding-bottom:4px; border-bottom:1px solid #F0E6E8; }}
  .grp ul {{ list-style:none; }}
  .grp li {{ position:relative; padding-left:13px; font-size:10px; color:#555; line-height:1.65; }}
  .grp li:before {{ content:""; position:absolute; left:0; top:7px; width:4px; height:4px; background:#C9A6AE; border-radius:50%; }}

  /* ---------- NOTAS ---------- */
  .notes {{ display:grid; grid-template-columns:1fr 1fr; gap:20px; }}
  .note {{ border:1px solid #EAEAEA; padding:18px 20px; }}
  .note h4 {{ font-family:"Constantia",Georgia,serif; font-size:14px; color:#000; font-weight:400; margin-bottom:9px; }}
  .note ul {{ list-style:none; }}
  .note li {{ position:relative; padding-left:14px; font-size:10px; color:#555; line-height:1.7; }}
  .note li:before {{ content:"—"; position:absolute; left:0; color:#5A1E2A; }}
  .cond {{ display:flex; gap:8px; flex-wrap:wrap; margin-top:6px; }}
  .tag {{ font-size:9px; padding:4px 12px; border-radius:3px; letter-spacing:.04em; }}

  .foot-note {{ position:absolute; bottom:-8mm; left:0; right:0; text-align:center; font-size:8px; letter-spacing:.22em; text-transform:uppercase; color:#C8C8C8; }}
</style>
</head>
<body>

  <!-- ===================== CAPA ===================== -->
  <section class="page cover">
    <div class="kicker">Especificação Visual da Planilha</div>
    <div class="rule"></div>
    <div class="center">
      <div class="wbar"></div>
      <h1>Psi <span class="wine">Planner</span></h1>
      <div class="tagline">Gestão clínica premium para psicólogos</div>
      <p class="desc">Plataforma em Excel com navegação por macros VBA, dashboards minimalistas e identidade visual editorial. Este documento apresenta a estrutura completa, a paleta oficial, a tipografia e o mockup de cada aba da planilha.</p>
      <div class="meta">
        <div><span>Cliente exemplo</span><b>Clínica Vita</b></div>
        <div><span>Formato</span><b>Excel .xlsm — VBA</b></div>
        <div><span>Data</span><b>27 de maio de 2026</b></div>
      </div>
    </div>
    <div class="foot">Salutti · Projeto Psi Planner</div>
  </section>

  <!-- ===================== VISAO GERAL ===================== -->
  <section class="page">
    <div class="sec-title"><span class="n">I</span><h2>Estrutura geral</h2></div>
    <p style="margin-bottom:24px;color:#6b6b6b;max-width:88%;">A planilha é organizada em sete abas interligadas. Toda a navegação acontece por botões com macros VBA, sem rolagem entre planilhas, reproduzindo a experiência de um software clínico boutique.</p>
    <div class="overview">
      <div class="ov-card"><div class="t"><span class="n">01</span>Menu Inicial</div><p>Dashboard com cards de resumo, navegação e avisos automáticos.</p></div>
      <div class="ov-card"><div class="t"><span class="n">02</span>Financeiro</div><p>Lançamentos, indicadores, limite MEI, despesas e gráficos.</p></div>
      <div class="ov-card"><div class="t"><span class="n">03</span>Pacientes</div><p>Cadastro clínico, pacotes, histórico e alertas automáticos.</p></div>
      <div class="ov-card"><div class="t"><span class="n">04</span>Calendário</div><p>Agenda semanal 07h-22h, status e métricas de atendimento.</p></div>
      <div class="ov-card"><div class="t"><span class="n">05</span>Aniversariantes</div><p>Controle por mês, destaque da semana e mensagens.</p></div>
      <div class="ov-card"><div class="t"><span class="n">06</span>Relatórios</div><p>Painel analítico com indicadores e exportação em PDF.</p></div>
      <div class="ov-card"><div class="t"><span class="n">07</span>Configurações</div><p>Identidade da clínica e parâmetros gerais da plataforma.</p></div>
      <div class="ov-card" style="background:#5A1E2A;border-color:#5A1E2A;"><div class="t" style="color:#fff;">Navegação VBA</div><p style="color:#E6D2D7;">Botões com hover discreto conectam todas as abas em um clique.</p></div>
    </div>
    <div class="foot-note">Psi Planner — Especificação Visual</div>
  </section>

  <!-- ===================== IDENTIDADE ===================== -->
  <section class="page">
    <div class="sec-title"><span class="n">II</span><h2>Identidade visual</h2></div>
    <h3 style="font-size:12px;font-weight:400;letter-spacing:.18em;text-transform:uppercase;color:#A6A6A6;margin-bottom:18px;">Paleta oficial</h3>
    <div class="swatches">
      {swatches()}
    </div>
    <h3 style="font-size:12px;font-weight:400;letter-spacing:.18em;text-transform:uppercase;color:#A6A6A6;margin-bottom:16px;">Tipografia</h3>
    <div class="type-grid">
      <div class="type-card">
        <div class="lbl">Títulos · Editorial</div>
        <div class="name">Tan Pearl</div>
        <div class="role">Nome da clínica, títulos principais, cabeçalhos, frases institucionais e destaques do dashboard. Tamanhos 24-32pt no nome, 16-20pt nos títulos.</div>
        <div class="chars"><span>Sofisticada</span><span>Editorial</span><span>Luxuosa</span><span>Delicada</span></div>
      </div>
      <div class="type-card">
        <div class="lbl">Conteúdo · Interface</div>
        <div class="name sans">Dream Avenue</div>
        <div class="role">Botões, tabelas, dados financeiros, cadastros, menus e campos. Botões 10-11pt, tabelas 10pt, cards 12-14pt.</div>
        <div class="chars"><span>Minimalista</span><span>Limpa</span><span>Moderna</span><span>Confortável</span></div>
      </div>
    </div>
    <p style="margin-top:20px;font-size:9.5px;color:#B0B0B0;font-style:italic;">Regras: respiro visual amplo, hierarquia elegante, uso minimalista de negrito, peso fino/regular e caixa alta apenas em títulos específicos.</p>
    <div class="foot-note">Psi Planner — Especificação Visual</div>
  </section>

  <!-- ===================== ABAS ===================== -->
  {aba_pages()}

  <!-- ===================== NOTAS TECNICAS ===================== -->
  <section class="page">
    <div class="sec-title"><span class="n">III</span><h2>Notas técnicas</h2></div>
    <div class="notes">
      <div class="note">
        <h4>Automação VBA</h4>
        <ul>
          <li>Navegação entre abas por botões com macros</li>
          <li>Efeito hover discreto (cinza claro) nos botões</li>
          <li>Geração manual de recibo</li>
          <li>Ações: adicionar, remarcar e finalizar consulta</li>
        </ul>
      </div>
      <div class="note">
        <h4>Limite MEI</h4>
        <ul>
          <li>Referência de R$ 81.000 por ano</li>
          <li>Alerta a 70% — cinza escuro discreto</li>
          <li>Acima de 90% — preto com texto branco</li>
          <li>Indicador permanente no painel financeiro</li>
        </ul>
      </div>
      <div class="note">
        <h4>Formatação condicional</h4>
        <div class="cond">
          <span class="tag" style="background:#3B3B3B;color:#fff;">Pago — cinza escuro</span>
          <span class="tag" style="background:#A6A6A6;color:#fff;">Pendente — cinza médio</span>
          <span class="tag" style="background:#000;color:#fff;">Cancelado/Faltou — preto</span>
          <span class="tag" style="background:#5A1E2A;color:#fff;">Destaque — vinho</span>
        </div>
      </div>
      <div class="note">
        <h4>Direção estética</h4>
        <ul>
          <li>Luxo minimalista e sofisticação discreta</li>
          <li>Bordas finas e sombras suaves</li>
          <li>Sem cores vibrantes ou poluição visual</li>
          <li>Inspiração em dashboards boutique premium</li>
        </ul>
      </div>
    </div>
    <div class="foot-note">Psi Planner — Especificação Visual</div>
  </section>

</body>
</html>"""

with open(OUT_HTML, "w", encoding="utf-8") as f:
    f.write(HTML)
print("HTML gerado:", OUT_HTML, "-", os.path.getsize(OUT_HTML), "bytes")
