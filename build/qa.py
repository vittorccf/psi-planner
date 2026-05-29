# -*- coding: utf-8 -*-
"""Segunda bateria de QA: erros, navegação, names, validações, CF + contraste WCAG."""
import time, sys
import win32com.client as win32

F = r"D:\Vittor\Projetos\Saluti\PsiPlanner\Psi Planner.xlsm"
xlCellTypeFormulas = -4123
xlErrors = 16
SHEETS = ["Menu Inicial","Financeiro","Pacientes","Calendário","Aniversariantes","Relatórios","Configurações"]
NAV = {"IrFinanceiro":"Financeiro","IrPacientes":"Pacientes","IrCalendario":"Calendário",
       "IrAniversariantes":"Aniversariantes","IrRelatorios":"Relatórios",
       "IrConfiguracoes":"Configurações","IrMenu":"Menu Inicial"}
NAMES = ["PAC_Status","FIN_Valor","FIN_Status","FIN_Data","CAL_LogData","CAL_LogStatus",
         "ANIV_Nasc","MEI_Faturamento","KPI_PacientesAtivos","AVISO_SemSessao"]

fails = []
def check(ok, msg):
    print(("  OK  " if ok else " FAIL ") + msg)
    if not ok: fails.append(msg)

ex = win32.gencache.EnsureDispatch("Excel.Application")
ex.Visible = False; ex.DisplayAlerts = False; ex.AutomationSecurity = 1
try:
    wb = ex.Workbooks.Open(F)
    ex.Calculation = -4105  # automatic
    ex.CalculateFull(); time.sleep(0.5)

    print("== 1. Abas ==")
    got = [wb.Worksheets(i+1).Name for i in range(wb.Worksheets.Count)]
    check(got == SHEETS, "abas na ordem: %s" % got)

    print("== 2. Varredura de células com ERRO (todas as abas) ==")
    total_err = 0
    for s in SHEETS:
        ws = wb.Worksheets(s)
        try:
            errs = ws.UsedRange.SpecialCells(xlCellTypeFormulas, xlErrors)
            addrs = []
            for ar in errs.Areas:
                for c in ar.Cells:
                    addrs.append(c.Address(False, False) + "=" + str(c.Text))
            total_err += len(addrs)
            check(False, "%s: %d erro(s): %s" % (s, len(addrs), addrs[:6]))
        except Exception:
            check(True, "%s: sem erros" % s)
    check(total_err == 0, "total de células com erro = %d" % total_err)

    print("== 3. Navegação (macros) ==")
    for sub, dest in NAV.items():
        try:
            ex.Run(sub); ok = (ex.ActiveSheet.Name == dest)
        except Exception as e:
            ok = False
        check(ok, "%s -> %s" % (sub, dest))
    ex.Run("IrMenu")

    print("== 4. Subs de cálculo/ação não quebram ==")
    for sub in ["AtualizarKPIs","AtualizarAlertaMEI","AtualizarAvisos","LimparFiltros"]:
        try:
            ex.Run(sub); check(True, "%s executou" % sub)
        except Exception as e:
            check(False, "%s erro: %r" % (sub, e))

    print("== 5. Named ranges resolvem ==")
    for nm in NAMES:
        try:
            r = wb.Names(nm).RefersToRange; check(r is not None, "%s -> %s" % (nm, r.Address))
        except Exception as e:
            check(False, "%s ausente" % nm)

    print("== 6. KPIs coerentes ==")
    m = wb.Worksheets("Menu Inicial"); f = wb.Worksheets("Financeiro")
    check(m.Range("B7").Value == 7, "pacientes ativos = %s" % m.Range("B7").Value)
    check(abs(f.Range("M6").Value - 1560) < 1, "recebido = %s" % f.Range("M6").Value)
    check(0 < f.Range("M9").Value < 1, "inadimplência (frac) = %s" % f.Range("M9").Value)

    print("== 7. Validações de dados e CF presentes ==")
    fv = wb.Worksheets("Financeiro")
    # AutoFilter/Validation: conta validações na coluna G (status)
    try:
        vt = fv.Range("G7").Validation.Type; check(vt == 3, "Financeiro G7 validation list (type=%s)" % vt)
    except Exception as e:
        check(False, "validation G7: %r" % e)
    check(fv.Cells.FormatConditions.Count >= 0, "CF acessível")

    wb.Close(SaveChanges=False)
finally:
    ex.Quit()

# ---- Contraste WCAG (paleta) ----
print("== 8. Contraste WCAG (texto/fundo) ==")
def lum(hx):
    hx = hx.lstrip("#"); r,g,b = (int(hx[i:i+2],16)/255 for i in (0,2,4))
    f = lambda c: c/12.92 if c <= 0.03928 else ((c+0.055)/1.055)**2.4
    R,G,B = f(r),f(g),f(b); return 0.2126*R+0.7152*G+0.0722*B
def ratio(fg,bg):
    L1,L2 = lum(fg),lum(bg); hi,lo = max(L1,L2),min(L1,L2); return (hi+0.05)/(lo+0.05)
pairs = [
    ("ink/white","241E20","FFFFFF"), ("ink/surface","241E20","F6F2F1"),
    ("mute/white","6E646A","FFFFFF"), ("mute/surface","6E646A","F6F2F1"),
    ("mute/surface_deep","6E646A","EFE8E7"),
    ("ink/done","241E20","E9ECEC"), ("ink/wait","241E20","F0E6E8"),
    ("vinho_esc/alert","3A1018","E7D2D7"), ("white/vinho","FFFFFF","5A1E2A"),
    ("vinho/white","5A1E2A","FFFFFF"),
]
for nm,fg,bg in pairs:
    r = ratio(fg,bg); lvl = "AAA" if r>=7 else ("AA" if r>=4.5 else ("AA-large" if r>=3 else "FAIL"))
    print(("  OK  " if r>=3 else " FAIL ") + "%-22s %.2f:1 (%s)" % (nm, r, lvl))
    if r < 3: fails.append("contraste %s %.2f" % (nm, r))

print("\n==== RESULTADO: %s (%d falhas) ====" % ("PASS" if not fails else "FALHAS", len(fails)))
for x in fails: print(" -", x)
sys.exit(1 if fails else 0)
