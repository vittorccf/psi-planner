# Psi Planner — Landing Page

Página de vendas estática (HTML + Tailwind CSS + JS vanilla) que direciona para o checkout do Kiwify.

> **Conceito:** *Seu consultório organizado. Sua mente tranquila.*

---

## 📁 Estrutura

```
page/
├── index.html                  # Landing principal (12 seções)
├── politica-de-privacidade.html
├── termos-de-uso.html
├── politica-de-reembolso.html
├── css/
│   ├── input.css               # Fonte do Tailwind (tokens + componentes)
│   └── styles.css              # CSS COMPILADO e minificado (é este que a página usa)
├── js/
│   └── main.js                 # Cookie banner LGPD, pixels, tracking de CTAs
├── assets/
│   ├── logo-monograma.webp/.png
│   ├── favicon.png / favicon.ico / apple-touch-icon.png
│   ├── og-image.jpg            # Imagem de compartilhamento (placeholder gerada)
│   └── mockups/*.webp          # Telas reais do produto (otimizadas)
├── tailwind.config.js          # Tokens da marca
├── optimize_assets.py          # Pipeline que gera os assets a partir dos previews
└── README.md
```

---

## 🚀 Deploy

A página é 100% estática — basta hospedar a pasta `page/` em qualquer serviço de arquivos estáticos.

### Testar localmente
```bash
cd page
python -m http.server 8080
# abra http://localhost:8080
```

### Opções de hospedagem (grátis)
- **Netlify / Vercel / Cloudflare Pages:** arraste a pasta `page/` ou conecte o repositório. Sem build necessário (o CSS já vem compilado).
- **GitHub Pages:** suba o conteúdo de `page/` na branch de publicação.
- **Hospedagem própria (cPanel/FTP):** envie a pasta `page/` para a raiz do domínio.

### Recompilar o CSS (só se editar classes no HTML)
O `styles.css` já está pronto. Se alterar/adicionar classes Tailwind no HTML, recompile:
```bash
cd page
npx tailwindcss@3.4.17 -c tailwind.config.js -i css/input.css -o css/styles.css --minify
```

---

## ⚠️ PLACEHOLDERS — você PRECISA preencher antes de publicar

Procure por `{{...}}` em todos os arquivos `.html` e `.js` e substitua.

### 1. Checkout e contato (obrigatórios)
| Placeholder | Onde aparece | O que é |
|---|---|---|
| `{{LINK_KIWIFY}}` | todos os botões/CTAs, `index.html` (JSON-LD) | URL do checkout Kiwify |
| `{{EMAIL_SUPORTE}}` | rodapé, políticas | E-mail de atendimento |
| `{{WHATSAPP}}` | rodapé, política de reembolso | Número no formato `5562999999999` (DDI+DDD+número) |

### 2. Identificação do vendedor (rodapé + políticas — exigido por LGPD/CDC)
| Placeholder | O que é |
|---|---|
| `{{NOME_VENDEDOR}}` | Nome completo ou razão social |
| `{{CPF_CNPJ}}` | CPF ou CNPJ |
| `{{CIDADE_UF}}` | Cidade/UF (ex.: Goiânia/GO) |
| `{{DOMINIO}}` | Domínio do site (ex.: psiplanner.com.br) — usado em canonical e Open Graph |
| `{{DATA_ATUALIZACAO}}` | Data das políticas (ex.: 28/05/2026) |

### 3. Rastreamento / pixels (opcionais — funciona sem eles)
| Placeholder | Onde | O que é |
|---|---|---|
| `{{GA_ID}}` | `index.html` (`window.GA_ID`) | ID do Google Analytics 4 (`G-XXXXXXXXXX`) |
| `{{META_PIXEL_ID}}` | `index.html` (`window.META_PIXEL_ID`) | ID do Meta (Facebook) Pixel |

> Os pixels só carregam **após o aceite no banner de cookies** (LGPD). Se deixar os placeholders, simplesmente não carregam — a página funciona normalmente. O **pixel da Kiwify** é configurado no painel da própria Kiwify (não precisa de código aqui).

### 4. Depoimentos (prova social — substituir pelos reais)
| Placeholder | O que é |
|---|---|
| `{{DEPOIMENTO_1/2/3}}` | Texto do depoimento real (com permissão de uso) |
| `{{NOME_DEPOIMENTO_1/2/3}}` | Nome de quem deu o depoimento |
| `{{CRP_OU_CIDADE_1/2/3}}` | CRP, profissão ou cidade |

### 5. Vídeo (opcional)
- `{{VIDEO}}` — comentário em `index.html` (seção "Por dentro do planner"). Se tiver um vídeo de demonstração, incorpore o embed nesse ponto.

---

## ✍️ Sobre a copy

- **Toda a copy de venda (dor, solução, FAQ) é uma PROPOSTA** baseada no nicho (psicólogo MEI) e na estrutura de uma LP de referência. Revise e ajuste o tom à sua voz.
- Trechos que dependem de decisão/dado seu estão marcados na página com a etiqueta amarela **"revisar"** (depoimentos e política de atualização).
- **Nenhum dado factual foi inventado** (contato, CPF, depoimentos, link) — tudo é placeholder.
- As **políticas** (privacidade/termos/reembolso) são modelos Brasil (LGPD + CDC). **Recomendamos revisão jurídica** antes de publicar.

---

## 🎨 Identidade aplicada
- **Cores:** vinho `#5A1E2A`, vinho escuro `#3A1018`, vinho claro `#8E3B4C`, dourado `#C9A86A`, off-white `#FAFAF8`, carvão `#1A1719`. Dourado usado **apenas como acento**.
- **Tipografia:** títulos em **Prata** (serif display), corpo em **Inter** — via Google Fonts.
- **Logo:** monograma "P" art déco dourado + wordmark "PSI PLANNER".

---

## ✅ Acessibilidade & performance
- Mobile-first, responsiva (390 / tablet / 1440), sem overflow.
- Contraste WCAG AA; navegação por teclado com `:focus-visible`; skip-link; `prefers-reduced-motion`.
- FAQ em `<details>` (acessível, sem depender de JS).
- Imagens em WebP otimizadas, com `width/height`, `alt` e `loading="lazy"`.
