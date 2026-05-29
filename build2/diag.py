# -*- coding: utf-8 -*-
"""Reabre o .xlsm salvo e valida: componentes VBA, controles do form, navegação, KPIs."""
import win32com.client as win32
FINAL = r"D:\Vittor\Projetos\Saluti\PsiPlanner\modelo2\Psi Planner 2.xlsm"

excel = win32.gencache.EnsureDispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
excel.AutomationSecurity = 1  # habilita macros
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
    import time as _t, os as _os
    def shot(ws, rng_addr, png):
        # robusto: retry se a captura sair vazia/pequena (timing de render do COM)
        for _att in range(3):
            try:
                ws.Activate()
                try:
                    excel.ActiveWindow.ScrollRow = 1
                    excel.ActiveWindow.ScrollColumn = 1
                except Exception: pass
                excel.CalculateFull()
                _t.sleep(0.8)
                rng = ws.Range(rng_addr)
                rng.CopyPicture(1, 2)  # xlScreen, xlBitmap
                co = ws.ChartObjects().Add(10, 10, rng.Width, rng.Height)
                co.Chart.ChartArea.Format.Line.Visible = False
                co.Chart.Paste()
                _t.sleep(1.0)
                co.Chart.Export(png)
                co.Delete()
                if _os.path.exists(png) and _os.path.getsize(png) > 6000:
                    print("Screenshot ->", png); return
                print("Screenshot vazio/pequeno (tent %d): %s" % (_att + 1, png))
                _t.sleep(0.6)
            except Exception as e:
                print("Screenshot falhou (%s) tent %d: %r" % (png, _att + 1, e))
                _t.sleep(0.6)
        print("Screenshot FALHOU apos retries:", png)
    base = r"D:\Vittor\Projetos\Saluti\PsiPlanner\modelo2"
    shot(m, "A1:L21", base + r"\Menu_Inicial_preview.png")
    shot(f, "A1:N30", base + r"\Financeiro_preview.png")
    shot(f, "A76:N116", base + r"\Financeiro_graficos_preview.png")
    cal = wb.Worksheets("Calendário")
    shot(cal, "A1:L25", base + r"\Calendario_preview.png")
    a2 = wb.Worksheets("Aniversariantes")
    shot(a2, "A1:Q22", base + r"\Aniversariantes_preview.png")
    shot(p, "A1:U16", base + r"\Pacientes_preview.png")
    shot(wb.Worksheets("Relatórios"), "A1:O32", base + r"\Relatorios_preview.png")
    shot(wb.Worksheets("Configurações"), "A1:E16", base + r"\Configuracoes_preview.png")
    wb.Close(SaveChanges=False)
    print("DIAG OK")
finally:
    excel.Quit()
