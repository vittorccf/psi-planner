# -*- coding: utf-8 -*-
"""Reabre o .xlsm salvo e valida: componentes VBA, controles do form, navegação, KPIs."""
import win32com.client as win32
FINAL = r"D:\Vittor\Projetos\Saluti\PsiPlanner\Psi Planner.xlsm"

excel = win32.gencache.EnsureDispatch("Excel.Application")
# Visível: CopyPicture(xlScreen) captura em branco quando a janela não está
# renderizada (Excel invisível) — manter visível garante o screenshot.
excel.Visible = True
excel.DisplayAlerts = False
excel.AutomationSecurity = 1  # habilita macros
try:
    excel.WindowState = -4143  # xlNormal (evita janela minimizada)
except Exception:
    pass
try:
    wb = excel.Workbooks.Open(FINAL)
    vbp = wb.VBProject
    print("Componentes VBA:")
    for c in vbp.VBComponents:
        print("  -", c.Name, "type", c.Type, "linhas", c.CodeModule.CountOfLines)
    # controles do form
    try:
        frm = vbp.VBComponents("frmConsulta")
        ctrls = [ctl.Name for ctl in frm.Designer.Controls]
        print("Controles frmConsulta:", ctrls)
    except Exception as e:
        print("Erro lendo form:", e)
    # navegação
    excel.Run("IrFinanceiro"); print("IrFinanceiro -> ativa:", excel.ActiveSheet.Name)
    excel.Run("IrCalendario"); print("IrCalendario -> ativa:", excel.ActiveSheet.Name)
    excel.Run("IrMenu"); print("IrMenu -> ativa:", excel.ActiveSheet.Name)
    # KPIs (valores em cache)
    m = wb.Worksheets("Menu Inicial")
    print("KPI ativos B7 =", m.Range("B7").Value)
    print("KPI receita C7 =", m.Range("C7").Value)
    print("KPI pendente D7 =", m.Range("D7").Value)
    print("KPI prox consulta F7 =", m.Range("F7").Value)
    print("KPI faltas G7 =", m.Range("G7").Value)
    print("KPI aniver H7 =", m.Range("H7").Value)
    print("AVISO SemSessao J6 =", m.Range("J6").Value)
    print("AVISO MEI J9 =", m.Range("J9").Value)
    print("AVISO Pend J12 =", m.Range("J12").Value)
    print("AVISO Aniv J15 =", m.Range("J15").Value)
    f = wb.Worksheets("Financeiro")
    print("FIN M6 recebido =", f.Range("M6").Value, "| M11 fatAnual =", f.Range("M11").Value, "| M10 =", f.Range("M10").Value)
    p = wb.Worksheets("Pacientes")
    print("PAC idade F7 =", p.Range("F7").Value, "| saldo R7 =", p.Range("R7").Value)
    a = wb.Worksheets("Aniversariantes")
    print("ANIV idade D7 =", a.Range("D7").Value, "| total mes K5 =", a.Range("K5").Value)
    print("KPI prox consulta F7 =", m.Range("F7").Value)
    print("AVISO MEI J9 =", m.Range("J9").Value)
    print("AVISO Pend J12 =", m.Range("J12").Value)
    print("Shapes por aba:", {wb.Worksheets(i+1).Name: wb.Worksheets(i+1).Shapes.Count for i in range(wb.Worksheets.Count)})
    # screenshots: CopyPicture -> Chart -> Export PNG
    import time as _t
    def shot(ws, rng_addr, png):
        # CopyPicture pode falhar por contencao de clipboard -> retry.
        last = None
        for attempt in range(4):
            try:
                ws.Activate()
                try:
                    excel.ActiveWindow.ScrollColumn = 1
                    excel.ActiveWindow.Zoom = 150  # captura em alta densidade (acentos/abas largas legíveis)
                except Exception: pass
                _t.sleep(0.35)
                rng = ws.Range(rng_addr)
                rng.CopyPicture(1, 2)  # xlScreen, xlBitmap
                co = ws.ChartObjects().Add(10, 10, rng.Width, rng.Height)
                co.Chart.ChartArea.Format.Line.Visible = False
                co.Chart.Paste()
                _t.sleep(0.6)
                co.Chart.Export(png)
                co.Delete()
                print("Screenshot ->", png)
                return
            except Exception as e:
                last = e
                _t.sleep(0.6)
        print("Screenshot falhou (%s): %r" % (png, last))
    base = r"D:\Vittor\Projetos\Saluti\PsiPlanner"
    SHOTS = [
        ("Menu Inicial",    "A1:L24",   "Menu_Inicial"),
        ("Financeiro",      "A1:M30",   "Financeiro"),
        ("Financeiro",      "A76:N116", "Financeiro_graficos"),
        ("Calendário",      "A1:K25",   "Calendario"),
        ("Aniversariantes", "A1:Q24",   "Aniversariantes"),
        ("Pacientes",       "A1:T16",   "Pacientes"),
        ("Relatórios",      "A1:O32",   "Relatorios"),
        ("Configurações",   "A1:D26",   "Configuracoes"),
    ]
    def dump(suffix):
        for sheet, rng_addr, nome in SHOTS:
            shot(wb.Worksheets(sheet), rng_addr, base + ("\\%s%s.png" % (nome, suffix)))

    # 1) tema CLARO (default) -> previews sem sufixo
    dump("")
    # 2) tema ESCURO -> repinta via AplicarTema e gera *_dark.png
    print("== Alternando para tema Escuro ==")
    wb.Worksheets("Configurações").Range("C21").Value = "Escuro"
    excel.Run("AplicarTema")
    _t.sleep(0.4)
    print("CFG_Tema =", wb.Worksheets("Configurações").Range("C21").Value)
    dump("_dark")
    # checagem rapida: fundo do Menu deve ter mudado (sem tocar formulas)
    m2 = wb.Worksheets("Menu Inicial")
    print("Menu B2 Interior =", m2.Range("B2").Interior.Color, "| B7 (KPI) =", m2.Range("B7").Value)

    wb.Close(SaveChanges=False)
    print("DIAG OK")
finally:
    excel.Quit()
