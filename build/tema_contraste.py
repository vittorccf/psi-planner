# -*- coding: utf-8 -*-
"""Calcula ratios WCAG da paleta (claro refinado + escuro proposto)."""

def lum(hx):
    hx = hx.lstrip("#")
    r, g, b = (int(hx[i:i+2], 16) / 255 for i in (0, 2, 4))
    f = lambda c: c/12.92 if c <= 0.03928 else ((c+0.055)/1.055)**2.4
    R, G, B = f(r), f(g), f(b)
    return 0.2126*R + 0.7152*G + 0.0722*B

def ratio(fg, bg):
    L1, L2 = lum(fg), lum(bg)
    hi, lo = max(L1, L2), min(L1, L2)
    return (hi+0.05)/(lo+0.05)

def show(title, pairs):
    print("\n== %s ==" % title)
    for nm, fg, bg in pairs:
        r = ratio(fg, bg)
        lvl = "AAA" if r >= 7 else ("AA" if r >= 4.5 else ("AA-large" if r >= 3 else "FAIL"))
        flag = "OK " if r >= 4.5 else ("~lg" if r >= 3 else "XXX")
        print("  [%s] %-26s %5.2f:1 (%s)  %s on %s" % (flag, nm, r, lvl, fg, bg))

# ---------------- TEMA CLARO (refinado) ----------------
L_SURF   = "FAFAF8"  # off-white base (era FFFFFF)
L_ELEV   = "F4F4F4"  # cinza_claro (cards/cabecalhos)
L_LINE   = "EDEDED"  # divisoria suave (era E2E2E2)
L_INK    = "000000"
L_MUTE   = "3B3B3B"
L_VINHO  = "5A1E2A"
# pasteis (com texto escuro)
show("CLARO - texto x fundo", [
    ("ink/surface",      L_INK,  L_SURF),
    ("mute/surface",     L_MUTE, L_SURF),
    ("ink/elev",         L_INK,  L_ELEV),
    ("mute/elev",        L_MUTE, L_ELEV),
    ("vinho/surface",    L_VINHO, L_SURF),
    ("vinho/elev",       L_VINHO, L_ELEV),
    ("branco/vinho",     "FFFFFF", L_VINHO),
])
show("CLARO - pasteis de status (texto escuro x pastel)", [
    ("verde txt/done",   "1F3B1F", "D9EAD3"),
    ("verm txt/red",     "5A1E2A", "F4CCCC"),
    ("amar txt/yellow",  "5A4A1E", "FFF2CC"),
])

# ---------------- TEMA ESCURO (proposto: "Vinho Noturno") ----------------
D_BASE   = "1A1719"  # surface base - carvao quente (NAO preto)
D_ELEV   = "2A2528"  # surface elevada (cards / linhas alternadas)
D_ELEV2  = "322B2E"  # elevada 2 (chips neutros / "Alta")
D_LINE   = "3C353A"  # divisoria sutil
D_TXT    = "E8E4E6"  # texto primario off-white (quente)
D_MUTE   = "A89FA4"  # texto secundario/mute
D_VINHO_FILL = "8E3B4C"  # vinho-acento para FILL (chips/botoes) sobre base
D_VINHO_TXT  = "D08395"  # vinho-acento para TEXTO sobre base (mais claro)
# pasteis escuros (luminosidade reduzida + texto claro)
D_GREEN_BG = "2E4A34"; D_GREEN_TX = "CDE9D2"
D_RED_BG   = "5C2A31"; D_RED_TX   = "F3CDD2"
D_YEL_BG   = "4D421F"; D_YEL_TX   = "F6E6B0"

show("ESCURO - texto x surface", [
    ("txt/base",         D_TXT,  D_BASE),
    ("txt/elev",         D_TXT,  D_ELEV),
    ("txt/elev2",        D_TXT,  D_ELEV2),
    ("mute/base",        D_MUTE, D_BASE),
    ("mute/elev",        D_MUTE, D_ELEV),
    ("vinho-txt/base",   D_VINHO_TXT, D_BASE),
    ("vinho-txt/elev",   D_VINHO_TXT, D_ELEV),
    ("txt/vinho-fill",   D_TXT,  D_VINHO_FILL),
    ("offwhite/vinho-fl", "F3E9EC", D_VINHO_FILL),
])
show("ESCURO - linha/divisoria (UI 3:1)", [
    ("line/base",        D_LINE, D_BASE),
    ("elev/base",        D_ELEV, D_BASE),
])
show("ESCURO - pasteis de status (texto claro x chip)", [
    ("verde",            D_GREEN_TX, D_GREEN_BG),
    ("vermelho",         D_RED_TX,   D_RED_BG),
    ("amarelo",          D_YEL_TX,   D_YEL_BG),
    # texto que remapeia para o chip vermelho (pendente=vinho-txt, cancelado=red-txt)
    ("vinho-txt/redchip", D_VINHO_TXT, D_RED_BG),
    ("red-txt/redchip",   D_RED_TX,    D_RED_BG),
    ("txt/elev (Alta)",   D_TXT,       D_ELEV2),
    # chips sobre a base (3:1 p/ distinguir o chip do fundo)
    ("greenbg/base",     D_GREEN_BG, D_BASE),
    ("redbg/base",       D_RED_BG,   D_BASE),
    ("yelbg/base",       D_YEL_BG,   D_BASE),
])
