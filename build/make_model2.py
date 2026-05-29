# -*- coding: utf-8 -*-
"""Gera o pipeline do Modelo 2 (build2/) a partir do build/ (Modelo 1).
Reproduzível: re-rodar regenera build2 do zero. Ordem: copia -> replaces
semânticos (paths/branding/fontes/cores-RGB) -> passe de hex -> escreve.
Edições complexas (gerar_pdf paleta/tipografia/capa; tests PART D) são feitas
depois por Edit direto nos arquivos de build2 (não re-rodar depois disso)."""
import json, os, re, shutil

ROOT = r"D:\Vittor\Projetos\Saluti\PsiPlanner"
SRC  = os.path.join(ROOT, "build")
DST  = os.path.join(ROOT, "build2")
OUT  = os.path.join(ROOT, "modelo2")
os.makedirs(DST, exist_ok=True); os.makedirs(OUT, exist_ok=True)

for f in ["build_xlsx.py","inject_vba.py","diag.py","gerar_pdf.py","tests.py","vba_modules.txt"]:
    shutil.copy(os.path.join(SRC,f), os.path.join(DST,f))

# ---- visual_spec.json: paleta quente + fontes ----
spec = json.load(open(os.path.join(SRC,"visual_spec.json"), encoding="utf-8"))
PAL2 = {"preto":"211711","ink":"2E2019","mute":"6E5640","line":"E6DBC9",
        "cinza_claro":"F3ECDF","surface_deep":"ECE2D2","cinza_medio":"B6A488",
        "cinza_escuro":"4A3B2E","vinho_elegante":"7B4A2D","vinho_escuro":"4A2E1A",
        "branco":"FFFFFF","done":"E8E2D5","wait":"F2E7D4","alert":"E9D2BE"}
for k,v in PAL2.items(): spec["meta"]["palette"][k]=v
spec["meta"]["title"]="Psi Planner 2 - Especificacao Visual Premium"
PLAY="Playfair Display"; POP="Poppins Light"
ROLE={"titulo_clinica":(PLAY,"Golsuka"),"titulo_aba":(PLAY,"Golsuka"),
      "kpi_valor":(PLAY,"Golsuka"),"kpi_valor_pequeno":(PLAY,"Golsuka"),
      "aviso_titulo":(PLAY,"Golsuka"),"institucional":(PLAY,"Playfair Display Italic"),
      "subtitulo":(POP,"Poppins Light"),"cabecalho_tabela":(POP,"Poppins Light"),
      "kpi_label":(POP,"Poppins Light"),"conteudo":(POP,"Poppins Light"),
      "conteudo_pequeno":(POP,"Poppins Light"),"botao":(POP,"Poppins Light"),
      "card_titulo":(POP,"Poppins Light"),"data_auto":(POP,"Poppins Light")}
for k,fd in spec["fonts"].items():
    if k in ROLE: fd["applied"],fd["intended"]=ROLE[k]
json.dump(spec, open(os.path.join(DST,"visual_spec.json"),"w",encoding="utf-8"),
          ensure_ascii=False, indent=2)

# ---- replaces semânticos por arquivo (antes do hex) ----
P1=r"D:\Vittor\Projetos\Saluti\PsiPlanner"
SEM = {
 "build_xlsx.py":[
   (P1+r"\build"+'"', P1+r"\build2"+'"'),
   (P1+r"\build\PsiPlanner_base.xlsx", P1+r"\build2\PsiPlanner2_base.xlsx"),
   ('PRODUTO = "Psi Planner"  # nome do produto (planner financeiro-clínico)',
    'PRODUTO = "Psi Planner 2"  # modelo 2 — identidade quente (creme/espresso/cognac)'),
   ('"Assinatura Psi Planner Pro"', '"Assinatura Psi Planner 2 Pro"'),
   ('menu["B1"].font = Font(name="Segoe UI Semilight", size=9,',
    'menu["B1"].font = Font(name="Poppins Light", size=9,'),
   ('mk.font = Font(name="Segoe UI Semilight", size=11,',
    'mk.font = Font(name="Poppins Light", size=11,'),
   ('_ws.oddFooter.left.text = "Psi Planner"', '_ws.oddFooter.left.text = "Psi Planner 2"'),
   ('"Psi Planner — controle financeiro', '"Psi Planner 2 — controle financeiro'),
 ],
 "inject_vba.py":[
   (P1+r"\build"+'"', P1+r"\build2"+'"'),
   (r'SRC_XLSX = BUILD + r"\PsiPlanner_base.xlsx"', r'SRC_XLSX = BUILD + r"\PsiPlanner2_base.xlsx"'),
   (P1+r"\Psi Planner.xlsm", P1+r"\modelo2\Psi Planner 2.xlsm"),
   ('"nav":     dict(fill=(255,255,255), line=(59,59,59),   lw=0.75, txt=(0,0,0),     font="Segoe UI Semilight", size=11),',
    '"nav":     dict(fill=(255,255,255), line=(110,86,64),  lw=0.75, txt=(46,32,25),  font="Poppins Light", size=11),'),
   ('"primary": dict(fill=(58,16,24),    line=None,         lw=0,    txt=(255,255,255), font="Segoe UI Semilight", size=11),',
    '"primary": dict(fill=(123,74,45),   line=None,         lw=0,    txt=(255,255,255), font="Poppins Light", size=11),'),
   ('"back":    dict(fill=(244,244,244), line=(166,166,166),lw=0.75, txt=(59,59,59),  font="Segoe UI Light",     size=10),',
    '"back":    dict(fill=(243,236,223), line=(182,164,136),lw=0.75, txt=(46,32,25),  font="Poppins Light",      size=10),'),
   ('VINHO_RGB = rgb((90, 30, 42))', 'VINHO_RGB = rgb((123, 74, 45))'),
   ('_setprop("Title", "Psi Planner")', '_setprop("Title", "Psi Planner 2")'),
   ('_setprop("Author", "Psi Planner")', '_setprop("Author", "Psi Planner 2")'),
   ('_setprop("Company", "Psi Planner")', '_setprop("Company", "Psi Planner 2")'),
 ],
 "diag.py":[
   (P1+r"\Psi Planner.xlsm", P1+r"\modelo2\Psi Planner 2.xlsm"),
   ('base = r"'+P1+'"', 'base = r"'+P1+r"\modelo2"+'"'),
 ],
 "gerar_pdf.py":[
   ('DELIV = r"'+P1+'"', 'DELIV = r"'+P1+r"\modelo2"+'"'),
   ('OUT_HTML = os.path.join(DELIV, "build", "psi_planner_spec.html")',
    'OUT_HTML = os.path.join(DELIV, "psi_planner2_spec.html")'),
 ],
 "tests.py":[
   (P1+r"\Psi Planner.xlsm", P1+r"\modelo2\Psi Planner 2.xlsm"),
   ('DELIV = r"'+P1+'"', 'DELIV = r"'+P1+r"\modelo2"+'"'),
   ('SPEC_HTML = os.path.join(DELIV, "build", "psi_planner_spec.html")',
    'SPEC_HTML = os.path.join(DELIV, "psi_planner2_spec.html")'),
 ],
}
HEX = {"5A1E2A":"7B4A2D","3A1018":"4A2E1A","241E20":"2E2019","6E646A":"6E5640",
 "9A8F92":"B79B7B","6B6065":"8A6E50","EFE8E7":"ECE2D2","F0E6E8":"F2E7D4",
 "F4F4F4":"F3ECDF","E7E1E0":"E6DBC9","E9ECEC":"E8E2D5","E7D2D7":"E9D2BE",
 "E2E2E2":"E6DBC9","EAEAEA":"E8DFD0","ECECEC":"E8DFD0","6B6B6B":"6E5640",
 "9A9A9A":"9C8A70","FBFBFB":"FBF7F0","EBD7DC":"E8D6BE","E8E8E8":"E6DBC9",
 "E6E6E6":"E2D6C4","E6D2D7":"E9D2BE","C9A6AE":"C2A07A","C8C8C8":"BCAE98",
 "B0B0B0":"A8957D","F6F2F1":"F3ECDF"}
pat = re.compile("("+"|".join(HEX.keys())+")", re.IGNORECASE)
def repl(m): return HEX[m.group(0).upper()]

for f,reps in SEM.items():
    p=os.path.join(DST,f); t=open(p,encoding="utf-8").read()
    for old,new in reps:
        if old not in t: print("  !! NO-OP em %s: %r" % (f, old[:60]))
        t=t.replace(old,new)
    t=pat.sub(repl,t)            # hex por último
    t=t.replace("PSI PLANNER","PSI PLANNER 2")
    open(p,"w",encoding="utf-8").write(t)
    print("ok", f)
print("DONE make_model2 v2")
