/* ============================================================
   Psi Planner — landing page (JS vanilla)
   - Ano dinâmico no rodapé
   - Banner de consentimento LGPD (carrega pixels só após aceite)
   - Rastreamento de cliques nos CTAs (Kiwify / GA / Meta)
   ============================================================ */
(function () {
  "use strict";

  // ---- Ano dinâmico ----
  var anoEl = document.getElementById("ano");
  if (anoEl) anoEl.textContent = new Date().getFullYear();

  // ---- Consentimento de cookies (LGPD) ----
  var KEY = "psiplanner_consent"; // "accept" | "reject"
  var banner = document.getElementById("cookie-banner");

  function getConsent() {
    try { return localStorage.getItem(KEY); } catch (e) { return null; }
  }
  function setConsent(v) {
    try { localStorage.setItem(KEY, v); } catch (e) {}
  }

  function showBanner() { if (banner) banner.classList.remove("hidden"); }
  function hideBanner() { if (banner) banner.classList.add("hidden"); }

  // ---- Carregamento dos pixels (somente após consentimento) ----
  var pixelsLoaded = false;
  function loadPixels() {
    if (pixelsLoaded) return;
    pixelsLoaded = true;

    // Google Analytics 4
    var GA_ID = window.GA_ID;
    if (GA_ID && GA_ID.indexOf("{{") === -1) {
      var s = document.createElement("script");
      s.async = true;
      s.src = "https://www.googletagmanager.com/gtag/js?id=" + GA_ID;
      document.head.appendChild(s);
      window.dataLayer = window.dataLayer || [];
      window.gtag = function () { window.dataLayer.push(arguments); };
      window.gtag("js", new Date());
      window.gtag("config", GA_ID);
    }

    // Meta Pixel
    var PIXEL = window.META_PIXEL_ID;
    if (PIXEL && PIXEL.indexOf("{{") === -1) {
      !(function (f, b, e, v, n, t, s) {
        if (f.fbq) return; n = f.fbq = function () { n.callMethod ? n.callMethod.apply(n, arguments) : n.queue.push(arguments); };
        if (!f._fbq) f._fbq = n; n.push = n; n.loaded = !0; n.version = "2.0"; n.queue = [];
        t = b.createElement(e); t.async = !0; t.src = v; s = b.getElementsByTagName(e)[0]; s.parentNode.insertBefore(t, s);
      })(window, document, "script", "https://connect.facebook.net/en_US/fbevents.js");
      window.fbq("init", PIXEL);
      window.fbq("track", "PageView");
    }
  }

  // Estado inicial do consentimento
  var consent = getConsent();
  if (consent === "accept") {
    loadPixels();
  } else if (consent !== "reject") {
    showBanner();
  }

  var btnAccept = document.getElementById("cookie-accept");
  var btnReject = document.getElementById("cookie-reject");
  if (btnAccept) btnAccept.addEventListener("click", function () { setConsent("accept"); hideBanner(); loadPixels(); });
  if (btnReject) btnReject.addEventListener("click", function () { setConsent("reject"); hideBanner(); });

  // ---- Rastreamento de CTAs ----
  function trackCTA(label) {
    if (window.gtag) window.gtag("event", "click_cta", { cta_local: label, value: 97, currency: "BRL" });
    if (window.fbq) window.fbq("track", "InitiateCheckout", { content_name: "Psi Planner", value: 97, currency: "BRL" });
  }
  document.querySelectorAll("[data-cta]").forEach(function (el) {
    el.addEventListener("click", function () { trackCTA(el.getAttribute("data-cta")); });
  });
})();
