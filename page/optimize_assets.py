# -*- coding: utf-8 -*-
"""Pipeline de otimizacao de assets da landing Psi Planner."""
import os
from PIL import Image

SRC = r"D:/Vittor/Projetos/Saluti/PsiPlanner"
LOGO_SRC = os.path.join(SRC, "build", "assets", "psi_monograma_ouro.png")
OUT = r"D:/Vittor/Projetos/Saluti/PsiPlanner/page/assets"
MOCK = os.path.join(OUT, "mockups")

WINE = (90, 30, 42)        # #5A1E2A
WINE_DARK = (26, 23, 25)   # #1A1719
GOLD = (201, 168, 106)     # #C9A86A

# mapeamento origem -> destino, largura maxima
JOBS = [
    ("Menu_Inicial_preview.png",      "menu-light.webp",     1100),
    ("Menu_Inicial_dark.png",         "menu-dark.webp",      1100),
    ("Financeiro_preview.png",        "financeiro.webp",     1100),
    ("Financeiro_dark.png",           "financeiro-dark.webp",1100),
    ("Calendario_preview.png",        "calendario.webp",     1100),
    ("Financeiro_graficos_dark.png",  "graficos-dark.webp",  1100),
    ("Pacientes_preview.png",         "pacientes.webp",      1100),
    ("Relatorios_preview.png",        "relatorios.webp",     1100),
    ("Aniversariantes_preview.png",   "aniversariantes.webp",1100),
    ("Configuracoes_preview.png",     "configuracoes.webp",  1100),
]

def conv(src, dst, maxw, q=82):
    im = Image.open(src).convert("RGB")
    if im.width > maxw:
        h = int(im.height * maxw / im.width)
        im = im.resize((maxw, h), Image.LANCZOS)
    im.save(dst, "WEBP", quality=q, method=6)
    return os.path.getsize(dst)

total = 0
for s, d, w in JOBS:
    sp = os.path.join(SRC, s)
    dp = os.path.join(MOCK, d)
    if os.path.exists(sp):
        sz = conv(sp, dp, w)
        total += sz
        print(f"{d:24s} {sz/1024:7.1f} KB")
    else:
        print(f"FALTA: {s}")

# Logo: copia o monograma e versao webp
logo = Image.open(LOGO_SRC).convert("RGBA")
logo.save(os.path.join(OUT, "logo-monograma.png"))
logo.save(os.path.join(OUT, "logo-monograma.webp"), "WEBP", quality=90, method=6)
print("logo-monograma.png/.webp OK", logo.size)

# Favicon 64x64 (monograma em fundo vinho)
fav = Image.new("RGBA", (64, 64), WINE + (255,))
mono = logo.resize((48, 48), Image.LANCZOS)
fav.paste(mono, (8, 8), mono)
fav.save(os.path.join(OUT, "favicon.png"))
# ICO multi-size
fav.save(os.path.join(OUT, "favicon.ico"), sizes=[(16,16),(32,32),(48,48)])
print("favicon.png/.ico OK")

# Apple touch icon 180
at = Image.new("RGBA", (180, 180), WINE + (255,))
m2 = logo.resize((130, 130), Image.LANCZOS)
at.paste(m2, (25, 25), m2)
at.convert("RGB").save(os.path.join(OUT, "apple-touch-icon.png"))
print("apple-touch-icon.png OK")

# OG image 1200x630 fundo vinho + monograma dourado centralizado + barras finas
og = Image.new("RGB", (1200, 630), WINE)
mlogo = logo.resize((220, 220), Image.LANCZOS)
og.paste(mlogo, (490, 150), mlogo)
from PIL import ImageDraw
d = ImageDraw.Draw(og)
d.line([(420, 420), (780, 420)], fill=GOLD, width=3)
d.line([(420, 130), (780, 130)], fill=GOLD, width=3)
og.save(os.path.join(OUT, "og-image.jpg"), "JPEG", quality=85)
print("og-image.jpg OK")

print(f"\nTotal mockups: {total/1024:.1f} KB")
