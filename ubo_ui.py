"""Dizajni i përbashkët i platformës (i njëjtë me aitools.uboconsulting.com).

- TOOLS: lista e veglave; prej saj ndërtohen menyja (Home.py), kartat e kryefaqes dhe kreu i çdo vegle,
  prandaj emri ose përshkrimi i një vegle ndryshohet vetëm këtu.
- base_css() / page_script(): ngjyrat sipas temës (e çelët / e errët), butoni i temës dhe dritarja
  "Ju lutem prisni" – për çdo faqe.
- tool_css() / tool_header() / card(): stili, kreu dhe kutitë e faqeve të veglave.
Fontet, ngjyra kryesore dhe rrumbullakimet e Streamlit-it vijnë nga tema te .streamlit/config.toml.
"""

import json
import os
import re
import time
from dataclasses import dataclass
from urllib.parse import quote

LOADED_AT = time.time()     # Home.py e ringarkon modulin kur skedari është më i ri se kjo

SECTIONS = {
    # grupi: (ngjyra, ngjyra e errët për tekst mbi sfond të çelët)
    "Pyetësorët": ("#22C6A0", "#0B8F76"),
    "Përkthimet": ("#3E7FF5", "#2560E6"),
    "Analiza": ("#EC1E8C", "#C4157A"),
}


@dataclass(frozen=True)
class Tool:
    section: str
    title: str
    description: str
    url_path: str
    file: str
    icon: str
    tags: str
    badge: str = ""

    @property
    def colors(self):
        return section_colors(self.section)


TOOLS = [
    Tool("Pyetësorët", "Gjenero XLS për KoboToolbox",
         "Kthe pyetësorin në formular XLS për KoboToolbox, me gjuhët dhe filtrat e tij.",
         "Gjenero_XLS", "veglat/1_Gjenero_XLS.py", "icons/survey-xmark.svg", "DOCX · XLSX · PDF", "AI"),
    Tool("Përkthimet", "Përkthim i dokumenteve Excel",
         "Përkthe dokumentet Excel automatikisht me inteligjencë artificiale.",
         "Perkthim_Excel_Files_AI", "veglat/2_Perkthim_Excel_Files_AI.py", "icons/file-excel.svg", "XLSX", "AI"),
    Tool("Përkthimet", "Përkthim i dokumenteve Word",
         "Përkthe dokumentet Word shpejt dhe saktë me inteligjencë artificiale.",
         "Perkthim_Word_Documents_AI", "veglat/4_Perkthim_Word_Documents_AI.py", "icons/file-word.svg", "DOCX", "AI"),
    Tool("Përkthimet", "Përkthim zyrtar",
         "Përkthime të verifikuara për dokumentet që kërkojnë saktësi të plotë.",
         "Perkthe_Zyrtarisht", "veglat/3_Perkthe_Zyrtarisht.py", "icons/language-exchange.svg", "DOCX · XLSX"),
    Tool("Analiza", "Analiza MaxDiff",
         "Analizo të dhënat me metodën MaxDiff për rezultate të thelluara.",
         "MaxDiff_Analysis", "veglat/MaxDiff_Analysis.py", "icons/analyse.svg", "CSV"),
    Tool("Analiza", "Grupimi i pyetjeve të hapura",
         "Grupo përgjigjet e pyetjeve të hapura sipas kategorive.",
         "Grupimi_i_pyetjeve_të_hapura", "veglat/Grupimi_i_pyetjeve_të_hapura.py", "icons/grouping.svg", "XLSX", "AI"),
]

# Logo si skedar (brenda HTML-së, 58 KB, e vononte faqen me ~2.5 s), në PNG (720x360, e krijuar nga SVG-ja):
# serveri i Streamlit Cloud i jep skedarët statikë që nuk janë PNG/JPG/GIF si tekst, prandaj SVG-ja nuk shfaqej.
# Në temën e errët teksti "UBO CONSULTING" është i bardhë; të dyja versionet vendosen në faqe dhe CSS-i
# shfaq atë të temës aktive.
# Adresa relative (pa "/" në fillim): në Streamlit Cloud aplikacioni është nën /~/+/, dhe "/app/static/…"
# do të kërkohej te faqja e jashtme, jo te aplikacioni.
LOGO_URL = "app/static/UBO-Logo.png"
LOGO_URL_DARK = "app/static/UBO-Logo-dark.png"


def logo_imgs():
    return (f'<img class="ubo-logo-light" src="{LOGO_URL}" alt="UBO Consulting">'
            f'<img class="ubo-logo-dark" src="{LOGO_URL_DARK}" alt="UBO Consulting">')


# ---------------------------------------------------------------------------
# Tema: e çelët / e errët
# Streamlit e ndërron temën në shfletues pa e rifreskuar faqen (butoni "Theme" përdor menunë e tij),
# prandaj ngjyrat tona janë në CSS për të dyja temat: skripti i faqes vendos html[data-ubo-theme]
# sipas temës që Streamlit po shfaq realisht. Hamendja e parë vjen nga serveri (st.context.theme).
# ---------------------------------------------------------------------------

def is_dark():
    """True kur shfletuesi e shfaq aplikacionin me temën e errët (sipas hapjes së fundit të faqes)."""
    try:
        import streamlit as st
        return st.context.theme.type == "dark"
    except Exception:
        return False


def section_colors(section):
    """(ngjyra, ngjyra për tekst mbi sfond të çelët) e grupit; për temën e errët CSS-i e zbut vetë."""
    return SECTIONS[section]


PALETTE = {
    "light": {
        "bg": "#FBFBFC", "surface": "#FFFFFF", "surface-2": "#F5F6F8", "ink": "#26262B", "muted": "#6A6F79",
        "line": "#ECECF1", "line-2": "#E4E8F2", "sel-bg": "#EEF4FE", "sel-ink": "#1C3F99", "sel-border": "#2560E6",
        "hover-bg": "#F7F9FF", "hover-border": "#B9CCF7", "sidebar": "#F6F7FA", "wash": "1",
        "shadow": "0 1px 2px #1e1e2d0d, 0 10px 30px #1e1e2d0d",
        "shadow-hover": "0 2px 8px #1e1e2d0f, 0 26px 60px #1e1e2d24",
        "logo-light": "block", "logo-dark": "none",
    },
    "dark": {
        "bg": "#111217", "surface": "#1A1B22", "surface-2": "#23252E", "ink": "#ECEDF2", "muted": "#A3A7B3",
        "line": "#2A2C36", "line-2": "#343744", "sel-bg": "rgba(91, 141, 246, .16)", "sel-ink": "#B7CDFF",
        "sel-border": "#5B8DF6", "hover-bg": "#20232D", "hover-border": "#4A5A85", "sidebar": "#15161C", "wash": "1.3",
        "shadow": "0 1px 2px #0006, 0 10px 30px #0005",
        "shadow-hover": "0 2px 8px #0007, 0 26px 60px #0008",
        "logo-light": "none", "logo-dark": "block",
    },
}


def _vars(name):
    return "".join(f"--ubo-{k}: {v};" for k, v in PALETTE[name].items())


def base_css():
    """Për çdo faqe: ngjyrat e të dy temave, butoni "Theme" dhe dritarja "Ju lutemi prisni pak"."""
    guess = "dark" if is_dark() else "light"
    return ("<style>:root {" + _vars(guess) + "}"
            + 'html[data-ubo-theme="light"] {' + _vars("light") + "}"
            + 'html[data-ubo-theme="dark"] {' + _vars("dark") + "}"
            + """
/* skripti i faqes (iframe pa pamje) nuk zë vend */
.st-key-ubo_js { display: none; }

/* logo e temës aktive */
img.ubo-logo-light { display: var(--ubo-logo-light) !important; }
img.ubo-logo-dark { display: var(--ubo-logo-dark) !important; }

/* ---------- Tema e errët: ngjyrat e UBO më të buta, pa hije rreth ikonave ---------- */
html[data-ubo-theme="dark"] .lp-icon, html[data-ubo-theme="dark"] .ubo-head-icon, html[data-ubo-theme="dark"] .ubo-step {
    background: color-mix(in srgb, var(--c, var(--ubo-c)) 72%, #2a2c36) !important; box-shadow: none !important;
}
html[data-ubo-theme="dark"] .lp-card, html[data-ubo-theme="dark"] .lp-barlabel, html[data-ubo-theme="dark"] .ubo-head {
    --cd: color-mix(in srgb, var(--c) 62%, #ffffff) !important;
}
html[data-ubo-theme="dark"] [class*="st-key-ubo_card"]::before { opacity: .55; }

/* ---------- Butoni "Theme" (në shiritin lart djathtas) dhe lista e tij ---------- */
#ubo-theme-wrap { position: relative; display: inline-flex; align-items: center; margin-right: 8px; flex-shrink: 0; }
#ubo-theme-btn {
    display: inline-flex; align-items: center; gap: 7px; height: 32px; padding: 0 10px 0 12px; white-space: nowrap;
    border-radius: 999px; border: 1px solid var(--ubo-line-2); background: var(--ubo-surface); color: var(--ubo-ink);
    font: 600 13px Inter, system-ui, sans-serif; cursor: pointer; transition: border-color .15s, background-color .15s;
}
#ubo-theme-btn:hover, #ubo-theme-wrap.open #ubo-theme-btn { border-color: var(--ubo-sel-border); background: var(--ubo-hover-bg); }
#ubo-theme-btn svg { width: 16px; height: 16px; }
#ubo-theme-btn .chev { width: 14px; height: 14px; opacity: .7; transition: transform .15s; }
#ubo-theme-wrap.open #ubo-theme-btn .chev { transform: rotate(180deg); }
#ubo-theme-menu {
    display: none; position: absolute; top: calc(100% + 6px); right: 0; min-width: 150px; padding: 5px; z-index: 999992;
    background: var(--ubo-surface); border: 1px solid var(--ubo-line); border-radius: 12px; box-shadow: var(--ubo-shadow-hover);
}
#ubo-theme-wrap.open #ubo-theme-menu { display: block; }
#ubo-theme-menu button {
    display: flex; align-items: center; gap: 9px; width: 100%; padding: 8px 10px; border: 0; border-radius: 8px;
    background: transparent; color: var(--ubo-ink); font: 500 13.5px Inter, system-ui, sans-serif; cursor: pointer; text-align: left;
}
#ubo-theme-menu button:hover { background: var(--ubo-hover-bg); }
#ubo-theme-menu button svg { width: 16px; height: 16px; }
#ubo-theme-menu button .tick { margin-left: auto; opacity: 0; color: var(--ubo-sel-border); }
#ubo-theme-menu button.on { font-weight: 700; }
#ubo-theme-menu button.on .tick { opacity: 1; }
/* menyja e Streamlit-it nuk duket për çastin kur butoni "Theme" e përdor për të ndërruar temën */
html.ubo-switching [data-testid="stMainMenuPopover"] { opacity: 0 !important; }

/* ---------- Dritarja "Ju lutemi prisni pak" ---------- */
#ubo-wait {
    position: fixed; inset: 0; z-index: 999990; display: none; align-items: center; justify-content: center;
    background: color-mix(in srgb, var(--ubo-bg) 55%, transparent); backdrop-filter: blur(2px);
}
#ubo-wait.show { display: flex; animation: ubo-fade .2s ease; }
@keyframes ubo-fade { from { opacity: 0; } to { opacity: 1; } }
#ubo-wait .box {
    width: min(390px, calc(100vw - 40px)); padding: 28px 26px 22px; text-align: center;
    background: var(--ubo-surface); border: 1px solid var(--ubo-line); border-radius: 20px;
    box-shadow: var(--ubo-shadow-hover); color: var(--ubo-ink); font-family: Inter, system-ui, sans-serif;
}
#ubo-wait .ring {
    width: 46px; height: 46px; margin: 0 auto 16px; border-radius: 50%;
    background: conic-gradient(#22C6A0, #3E7FF5, #F4C81E, #EC1E8C, #22C6A0);
    -webkit-mask: radial-gradient(farthest-side, transparent calc(100% - 5px), #000 calc(100% - 4px));
    mask: radial-gradient(farthest-side, transparent calc(100% - 5px), #000 calc(100% - 4px));
    animation: ubo-spin 1s linear infinite;
}
@keyframes ubo-spin { to { transform: rotate(360deg); } }
#ubo-wait .t { font: 700 19px "Space Grotesk", system-ui, sans-serif; letter-spacing: -.01em; }
#ubo-wait .s { color: var(--ubo-muted); font-size: 14px; margin: 6px 0 18px; line-height: 1.55; }
#ubo-wait button {
    border: 1px solid var(--ubo-line-2); background: var(--ubo-surface); color: #D63A3A; border-radius: 10px;
    padding: 9px 16px; font: 600 14px Inter, system-ui, sans-serif; cursor: pointer;
}
#ubo-wait button:hover { border-color: #D63A3A; background: rgba(214, 58, 58, .08); }
#ubo-wait button:disabled { opacity: .6; cursor: default; }
html.ubo-idle [data-testid="stSpinner"] { display: none !important; }
#ubo-toast {
    position: fixed; left: 50%; bottom: 28px; z-index: 999993; transform: translate(-50%, 20px); opacity: 0;
    pointer-events: none; transition: opacity .25s, transform .25s;
    padding: 12px 18px; border-radius: 12px; background: var(--ubo-surface); color: var(--ubo-ink);
    border: 1px solid var(--ubo-line); box-shadow: var(--ubo-shadow-hover); font: 500 14px Inter, system-ui, sans-serif;
}
#ubo-toast.show { opacity: 1; transform: translate(-50%, 0); }
</style>""")


def page_script():
    """JavaScript për çdo faqe (në një iframe pa pamje, punon mbi faqen kryesore):
    - html[data-ubo-theme]: tema që Streamlit po shfaq (sipas ngjyrës së tekstit), për ngjyrat tona;
    - butoni "Theme" me listën Light / Dark: e ndërron temën përmes menysë së Streamlit-it, pa rifreskim,
      që dokumentet e ngarkuara të mos humbasin (nëse menyja nuk gjendet: localStorage + rifreskim);
    - dritarja "Ju lutemi prisni pak" kur faqja punon më shumë se 1.2 s, me butonin "Anulo kërkesën",
      që klikon butonin Stop të Streamlit-it."""
    paths = ["/"] + ["/" + quote(t.url_path) for t in TOOLS]
    return """<script>
(function () {
  const P = window.parent, D = P.document;
  const PATHS = %s;
  const ICON = {
    theme: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="9"/><path d="M12 3a9 9 0 0 1 0 18z" fill="currentColor"/></svg>',
    light: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/></svg>',
    dark: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12.8A9 9 0 1 1 11.2 3a7 7 0 0 0 9.8 9.8z"/></svg>',
    chev: '<svg class="chev" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 9l6 6 6-6"/></svg>',
    tick: '<svg class="tick" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12l5 5L20 7"/></svg>'
  };

  // The theme Streamlit is showing right now: dark when its text colour is light
  function currentTheme() {
    const app = D.querySelector('[data-testid="stApp"]');
    if (!app) return null;
    const m = P.getComputedStyle(app).color.match(/\\d+(\\.\\d+)?/g);
    if (!m) return null;
    const lum = (0.299 * m[0] + 0.587 * m[1] + 0.114 * m[2]) / 255;
    return lum > 0.5 ? "dark" : "light";
  }

  function syncTheme() {
    const t = currentTheme();
    if (t && D.documentElement.dataset.uboTheme !== t) D.documentElement.dataset.uboTheme = t;
    D.querySelectorAll("#ubo-theme-menu button").forEach(b => b.classList.toggle("on", b.dataset.theme === t));
  }

  // Switch through Streamlit's own menu (no reload: the uploaded files stay)
  // The user's choice (our own key, the same for every page); without one the platform opens in Light
  function wantedTheme() {
    try { return P.localStorage.getItem("ubo-theme") || "light"; } catch (e) { return "light"; }
  }

  // auto = the default applied on page load: then never the reload fallback (it could loop)
  function setTheme(name, auto) {
    if (!auto) { try { P.localStorage.setItem("ubo-theme", name); } catch (e) {} }
    const menuBtn = D.querySelector('[data-testid="stMainMenuButton"]');
    if (!menuBtn) return auto ? null : fallback(name);
    D.documentElement.classList.add("ubo-switching");
    menuBtn.click();
    let tries = 0;
    const pick = () => {
      const pop = D.querySelector('[data-testid="stMainMenuPopover"]');
      const opt = pop && Array.from(pop.querySelectorAll('[role="menuitemradio"]'))
        .find(e => new RegExp(name + "$", "i").test((e.innerText || "").trim()));
      if (opt) {
        opt.click();
        setTimeout(() => {
          if (D.querySelector('[data-testid="stMainMenuPopover"]')) menuBtn.click();   // close the menu
          D.documentElement.classList.remove("ubo-switching");
          syncTheme();
        }, 120);
      } else if (++tries < 15) {
        setTimeout(pick, 60);
      } else {
        D.documentElement.classList.remove("ubo-switching");
        if (D.querySelector('[data-testid="stMainMenuPopover"]')) menuBtn.click();
        if (!auto) fallback(name);
      }
    };
    setTimeout(pick, 60);
  }

  // Fallback: store the choice where Streamlit reads it (one key per address the app was opened with) and reload
  function fallback(name) {
    const value = JSON.stringify(name === "dark" ? "Dark" : "Light");
    new Set(PATHS.concat([P.location.pathname])).forEach(p => {
      try { P.localStorage.setItem("stActiveTheme-" + p + "-v2", value); } catch (e) {}
    });
    P.location.reload();
  }

  function ensureThemeButton() {
    const anchor = D.querySelector('[data-testid="stToolbarActions"]') || D.querySelector('[data-testid="stMainMenu"]');
    const bar = anchor ? anchor.parentElement : null;
    let wrap = D.getElementById("ubo-theme-wrap");
    if (wrap && (!bar || bar.contains(wrap))) return;
    if (wrap) wrap.remove();
    wrap = D.createElement("div");
    wrap.id = "ubo-theme-wrap";
    wrap.innerHTML = '<button id="ubo-theme-btn" type="button" aria-haspopup="true">' + ICON.theme + '<span>Theme</span>' + ICON.chev + '</button>'
      + '<div id="ubo-theme-menu" role="menu">'
      + '<button type="button" data-theme="light" role="menuitem">' + ICON.light + '<span>Light</span>' + ICON.tick + '</button>'
      + '<button type="button" data-theme="dark" role="menuitem">' + ICON.dark + '<span>Dark</span>' + ICON.tick + '</button></div>';
    wrap.querySelector("#ubo-theme-btn").onclick = (e) => { e.stopPropagation(); wrap.classList.toggle("open"); syncTheme(); };
    wrap.querySelectorAll("#ubo-theme-menu button").forEach(b => b.onclick = (e) => {
      e.stopPropagation(); wrap.classList.remove("open");
      // after this click has finished: otherwise Streamlit's menu, opened during it, takes it for a click
      // outside and closes again at once
      if (b.dataset.theme !== currentTheme()) setTimeout(() => setTheme(b.dataset.theme), 30);
    });
    if (bar) { bar.insertBefore(wrap, bar.firstChild); }
    else { wrap.style.cssText = "position:fixed;top:12px;right:64px;z-index:999991"; D.body.appendChild(wrap); }
    if (!P.__uboThemeClose) {
      P.__uboThemeClose = true;
      D.addEventListener("click", () => { const w = D.getElementById("ubo-theme-wrap"); if (w) w.classList.remove("open"); });
    }
  }

  function ensureWaitBox() {
    let box = D.getElementById("ubo-wait");
    if (box) return box;
    box = D.createElement("div");
    box.id = "ubo-wait";
    box.innerHTML = '<div class="box"><div class="ring"></div><div class="t">Ju lutemi prisni pak</div>'
      + '<div class="s">Po punojmë me kërkesën tuaj.<br>Për dokumente më të mëdha, kjo zgjat pak më shumë.</div>'
      + '<button type="button">Anulo kërkesën</button></div>';
    box.querySelector("button").onclick = function () {
      const stop = Array.from(D.querySelectorAll('[data-testid="stStatusWidget"] button'))
        .find(b => /stop/i.test(b.textContent));
      if (stop) { stop.click(); this.disabled = true; this.textContent = "Po anulohet…"; P.__uboCancelled = true; }
    };
    D.body.appendChild(box);
    return box;
  }

  let since = null;
  function tick() {
    syncTheme();
    // once per page load: show the chosen theme (Light by default), also when the computer is in dark mode
    if (!P.__uboThemeApplied) {
      const t = currentTheme();
      if (t && D.querySelector('[data-testid="stMainMenuButton"]')) {
        P.__uboThemeApplied = true;
        if (t !== wantedTheme()) setTheme(wantedTheme(), true);
      }
    }
    ensureThemeButton();
    const box = ensureWaitBox();
    const running = !!D.querySelector('[data-testid="stStatusWidgetRunningIcon"], [data-testid="stStatusWidgetRunningManIcon"]');
    if (running) {
      since = since || Date.now();
      if (Date.now() - since > 1200 && !box.classList.contains("show")) {
        const b = box.querySelector("button"); b.disabled = false; b.textContent = "Anulo kërkesën";
        box.classList.add("show");
      }
    } else {
      since = null;
      box.classList.remove("show");
      if (P.__uboCancelled) { P.__uboCancelled = false; showToast("Kërkesa u anulua. Mund ta nisni përsëri kur të doni."); }
    }
    // a stopped page keeps its last spinner ("Claude po lexon…") on screen; hide it while nothing runs
    D.documentElement.classList.toggle("ubo-idle", !running);
  }

  function showToast(text) {
    let t = D.getElementById("ubo-toast");
    if (!t) { t = D.createElement("div"); t.id = "ubo-toast"; D.body.appendChild(t); }
    t.textContent = text;
    t.classList.add("show");
    clearTimeout(P.__uboToastTimer);
    P.__uboToastTimer = setTimeout(() => t.classList.remove("show"), 4500);
  }
  if (P.__uboTimer) P.clearInterval(P.__uboTimer);
  P.__uboTimer = P.setInterval(tick, 200);
  tick();
})();
</script>""" % json.dumps(paths)


def interruptible(fn, *args, **kwargs):
    """Runs a long AI call (Claude / Gemini) so that "Anulo kërkesën" (Streamlit's Stop) stops it at once.

    The call runs in a thread; meanwhile the page sends a tiny update every 0.25 s, which is where Streamlit
    notices a stop request. When the page is stopped, the call gets the stop signal too
    (questionnaire_ai.CANCEL) and ends its request instead of running on and costing."""
    import threading
    import streamlit as st
    from questionnaire_ai import CANCEL

    stop = threading.Event()
    result = {}

    def work():
        CANCEL.set(stop)
        try:
            result["value"] = fn(*args, **kwargs)
        except BaseException as e:      # re-raised in the page's thread
            result["error"] = e

    worker = threading.Thread(target=work, daemon=True)
    worker.start()
    heartbeat = st.empty()
    try:
        while worker.is_alive():
            worker.join(0.25)
            heartbeat.empty()
    finally:
        if worker.is_alive():
            stop.set()
    if "error" in result:
        raise result["error"]
    return result["value"]


# ---------------------------------------------------------------------------
# Veglat
# ---------------------------------------------------------------------------

def icon_svg(tool, size=24):
    """Ikona e veglës (SVG i vogël, 1–2 KB, futet drejt në HTML). Madhësia vendoset edhe në vetë SVG-në,
    që ikona të mos dalë e madhe për një çast para se të vijë CSS-i."""
    if not os.path.exists(tool.icon):
        return ""
    with open(tool.icon, "r", encoding="utf-8") as f:
        svg = f.read()
    return re.sub(r"<svg\b", f'<svg width="{size}" height="{size}"', svg, count=1)


# Sfondi me njollat e lehta me katër ngjyrat e UBO, si te kryefaqja dhe faqja e hyrjes së platformës
WASH_CSS = """
[data-testid="stApp"] {
    background:
        radial-gradient(38vw 38vw at 12% -4%, rgba(34, 198, 160, calc(.10 * var(--ubo-wash))), transparent 60%),
        radial-gradient(34vw 34vw at 88% 4%, rgba(62, 127, 245, calc(.09 * var(--ubo-wash))), transparent 60%),
        radial-gradient(40vw 40vw at 4% 96%, rgba(244, 200, 30, calc(.08 * var(--ubo-wash))), transparent 62%),
        radial-gradient(40vw 40vw at 96% 92%, rgba(236, 30, 140, calc(.07 * var(--ubo-wash))), transparent 60%),
        var(--ubo-bg);
    background-attachment: fixed;
}
[data-testid="stHeader"] { background: transparent; }
"""


def tool_css(tool):
    """Stili i faqeve të veglave: sfondi, kreu, kutitë (kartat), opsionet, kutitë e ngarkimit, menyja anësore."""
    c, cd = tool.colors
    return "<style>" + WASH_CSS + f":root {{ --ubo-c: {c}; --ubo-cd: {cd}; }}" + """
/* ---------- Kartat: çdo pjesë e veglës në një kuti, si kartat e kryefaqes ---------- */
[class*="st-key-ubo_card"] {
    position: relative; overflow: hidden; gap: 1.1rem;
    background: var(--ubo-surface); border: 1px solid var(--ubo-line); border-radius: 18px;
    padding: 22px 24px 22px; margin-bottom: .35rem;
    box-shadow: var(--ubo-shadow);
}
/* vija me ngjyrën e grupit në krye të kartës */
[class*="st-key-ubo_card"]::before {
    content: ""; position: absolute; left: 0; right: 0; top: 0; height: 3px; background: var(--ubo-c); opacity: .85;
}
.ubo-card-head {
    display: flex; align-items: center; gap: 10px; margin-bottom: 4px;
    font-family: "Space Grotesk", system-ui, sans-serif; font-weight: 700; font-size: 17.5px;
    letter-spacing: -.01em; color: var(--ubo-ink);
}
.ubo-step {
    display: inline-grid; place-items: center; flex-shrink: 0; width: 26px; height: 26px; border-radius: 8px;
    background: var(--ubo-c); color: #fff; font-size: 13.5px; font-weight: 700;
    box-shadow: 0 4px 12px color-mix(in srgb, var(--ubo-c) 40%, transparent);
}
.ubo-card-sub { color: var(--ubo-muted); font-size: 13.5px; margin: 4px 0 2px 36px; }

/* ---------- Opsionet (radio, checkbox) si kuti të zgjedhshme ---------- */
[data-testid="stRadioGroup"] { flex-direction: row !important; flex-wrap: wrap; gap: 10px !important; margin-top: 4px; }
/* opsionet e përkthimit njëri nën tjetrin */
[class*="st-key-tr_mode"] [data-testid="stRadioGroup"] { flex-direction: column !important; align-items: flex-start; }
label[data-testid="stRadioOption"], [data-testid="stCheckbox"] > label {
    background: var(--ubo-surface); border: 1px solid var(--ubo-line-2); border-radius: 11px;
    padding: 9px 15px 9px 12px; margin: 0 !important; cursor: pointer;
    transition: border-color .14s, background-color .14s, box-shadow .14s;
}
label[data-testid="stRadioOption"]:hover, [data-testid="stCheckbox"] > label:hover {
    border-color: var(--ubo-hover-border); background: var(--ubo-hover-bg);
}
label[data-testid="stRadioOption"]:has(input:checked), [data-testid="stCheckbox"] > label:has(input:checked) {
    border-color: var(--ubo-sel-border); background: var(--ubo-sel-bg);
    box-shadow: 0 0 0 3px color-mix(in srgb, var(--ubo-sel-border) 14%, transparent);
}
label[data-testid="stRadioOption"]:has(input:checked) p, [data-testid="stCheckbox"] > label:has(input:checked) p {
    color: var(--ubo-sel-ink); font-weight: 600;
}
/* emrat e fushave (p.sh. "Përkthimi:") pak më të theksuar, me pak hapësirë nga opsionet */
label[data-testid="stWidgetLabel"] p { font-weight: 600; color: var(--ubo-ink); }
label[data-testid="stWidgetLabel"] { margin-bottom: 4px; }
[data-testid="stCheckbox"] [data-testid="stWidgetLabel"] p { font-weight: 500; }
[data-testid="stCheckbox"] [data-testid="stWidgetLabel"] { margin-bottom: 0; }

/* ---------- Titujt brenda veglave më të vegjël ---------- */
[data-testid="stMain"] h2 { font-size: 1.3rem !important; padding: .6rem 0 .3rem !important; }
[data-testid="stMain"] h3 { font-size: 1.08rem !important; padding: .4rem 0 .2rem !important; }

/* ---------- Kutitë e ngarkimit ---------- */
[data-testid="stFileUploaderDropzone"] {
    background: var(--ubo-surface); border: 2px dashed var(--ubo-line-2); border-radius: 13px;
    transition: border-color .14s, background-color .14s;
}
[data-testid="stFileUploaderDropzone"]:hover { border-color: var(--ubo-sel-border); background: var(--ubo-sel-bg); }
/* në shqip ("Upload" / "200MB per file" i shton vetë Streamlit) */
[data-testid="stFileUploaderDropzone"] button [data-testid="stMarkdownContainer"] p { font-size: 0 !important; }
[data-testid="stFileUploaderDropzone"] button [data-testid="stMarkdownContainer"] p::after {
    content: "Ngarko"; font-size: 14px;
}
[data-testid="stFileUploaderDropzoneInstructions"] span { font-size: 0 !important; }
[data-testid="stFileUploaderDropzoneInstructions"] span::after {
    content: "Tërhiqeni dokumentin këtu"; font-size: 13px; color: var(--ubo-muted);
}

/* ---------- Kreu i veglës ---------- */
.ubo-head { display: flex; align-items: center; gap: 16px; margin: 4px 0 0; }
.ubo-head-icon {
    width: 52px; height: 52px; flex-shrink: 0; border-radius: 15px; display: grid; place-items: center;
    background: var(--c); box-shadow: 0 8px 22px color-mix(in srgb, var(--c) 40%, transparent);
}
.ubo-head-icon svg { width: 26px; height: 26px; fill: #fff !important; }
.ubo-eyebrow {
    font-weight: 700; font-size: 12px; letter-spacing: .14em; text-transform: uppercase; color: var(--cd);
}
.ubo-title {
    font-family: "Space Grotesk", system-ui, sans-serif; font-weight: 700;
    font-size: clamp(26px, 3.4vw, 34px); line-height: 1.15; letter-spacing: -.02em; color: var(--ubo-ink); margin-top: 2px;
}
.ubo-lede { color: var(--ubo-muted); font-size: 16px; line-height: 1.55; margin: 14px 0 0 !important; max-width: 640px; }
.ubo-rule {
    height: 2px; border-radius: 2px; opacity: .5; margin: 22px 0 10px;
    background: linear-gradient(90deg, #22C6A0, #3E7FF5 34%, #F4C81E 66%, #EC1E8C);
}

/* ---------- Menyja anësore ---------- */
[data-testid="stSidebar"] { background: var(--ubo-sidebar); }
.ubo-nav-sec {
    font-size: 11.5px; font-weight: 700; letter-spacing: .12em; text-transform: uppercase;
    color: var(--ubo-muted); margin: 1.1rem 0 0.2rem 0.3rem;
}
.ubo-nav-sec i { display: inline-block; width: 7px; height: 7px; border-radius: 50%; margin-right: 7px; vertical-align: 1px; }
/* logo lart në menu, e vogël, me një vijë poshtë saj */
a.ubo-home-logo {
    display: block; margin: 0 0 12px; padding: 0 0 16px 8px; transition: opacity .15s;
    border-bottom: 1px solid var(--ubo-line);
}
a.ubo-home-logo:hover { opacity: .8; }
/* Streamlit u jep figurave object-fit: scale-down, që nuk e zmadhon SVG-në 120x60 */
a.ubo-home-logo img {
    width: 125px !important; max-width: none !important; height: auto !important;
    aspect-ratio: 2 / 1; object-fit: contain !important; display: block;
}
/* lidhjet e menysë afër njëra-tjetrës, si te menyja e Streamlit-it */
[data-testid="stSidebarUserContent"] [data-testid="stVerticalBlock"] { gap: 0.15rem; }
[data-testid="stSidebarUserContent"] [data-testid="stMarkdownContainer"] { margin-bottom: 0 !important; }
</style>"""


def card(title, step=None, subtitle="", key=None):
    """Një pjesë e veglës në kuti (me vijën e ngjyrës së grupit, numrin e hapit dhe titullin).
    Përdorim: `with ubo_ui.card("Ngarko pyetësorin", step=1): ...`"""
    import streamlit as st
    key = key or re.sub(r"\W+", "_", title.lower()).strip("_")
    box = st.container(key=f"ubo_card_{key}")
    step_html = f'<span class="ubo-step">{step}</span>' if step is not None else ""
    sub_html = f'<div class="ubo-card-sub">{subtitle}</div>' if subtitle else ""
    box.markdown(f'<div class="ubo-card-head">{step_html}{title}</div>{sub_html}', unsafe_allow_html=True)
    return box


# Ikonat e veglave në menunë anësore (Material Symbols të Streamlit-it)
NAV_ICONS = {
    "": ":material/home:",
    "Gjenero_XLS": ":material/fact_check:",
    "Perkthim_Excel_Files_AI": ":material/table_view:",
    "Perkthim_Word_Documents_AI": ":material/description:",
    "Perkthe_Zyrtarisht": ":material/translate:",
    "MaxDiff_Analysis": ":material/leaderboard:",
    "Grupimi_i_pyetjeve_të_hapura": ":material/category:",
}


def sidebar_css(active):
    """Menyja anësore me ngjyra: çdo lidhje me ngjyrën e grupit të saj (ikona, sfondi kur kalon miu),
    vegla aktive me sfond, vijë anësore dhe tekst me ngjyrën e grupit; sfondi i menysë merr lehtë
    ngjyrën e grupit të veglës aktive."""
    links = ""
    for t in TOOLS:
        c, cd = SECTIONS[t.section]
        for href in {t.url_path, quote(t.url_path)}:
            links += f'[data-testid="stSidebar"] a[href="{href}"] {{ --c: {c}; --cd: {cd}; }} '
    hrefs = sorted({active.url_path, quote(active.url_path)})
    active_sel = ", ".join(f'[data-testid="stSidebar"] a[href="{h}"]' for h in hrefs)
    active_text = ", ".join(f'[data-testid="stSidebar"] a[href="{h}"] p' for h in hrefs)
    return "<style>" + links + """
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, color-mix(in srgb, var(--ubo-c) 9%, var(--ubo-sidebar)) 0%,
                var(--ubo-sidebar) 320px) !important;
}
[data-testid="stSidebar"] a[data-testid="stPageLink-NavLink"] {
    border-radius: 10px; padding: 6px 10px; margin: 1px 0;
    transition: background-color .15s, box-shadow .15s;
}
[data-testid="stSidebar"] a[data-testid="stPageLink-NavLink"] [data-testid="stIconMaterial"] {
    color: var(--c, var(--ubo-muted)); font-size: 1.15rem;
}
[data-testid="stSidebar"] a[data-testid="stPageLink-NavLink"]:hover {
    background: color-mix(in srgb, var(--c, #8a8fa0) 11%, transparent) !important;
}
""" + active_sel + """ {
    background: color-mix(in srgb, var(--c) 15%, var(--ubo-surface)) !important;
    box-shadow: inset 3px 0 0 var(--c);
}
""" + active_text + """ {
    color: var(--cd) !important; font-weight: 700 !important;
}
/* emrat e grupeve me ngjyrën e grupit */
.ubo-nav-sec { color: var(--cd) !important; }
html[data-ubo-theme="dark"] [data-testid="stSidebar"] a, html[data-ubo-theme="dark"] .ubo-nav-sec {
    --cd: color-mix(in srgb, var(--c) 62%, #ffffff) !important;
}
</style>"""


def tool_header(tool):
    """Kreu i përbashkët i çdo vegle, bashkë me stilin e faqes (në një element, që të shfaqen njëherësh):
    ikona me ngjyrën e grupit, grupi, titulli dhe përshkrimi."""
    c, cd = tool.colors
    return tool_css(tool) + f"""<div class="ubo-head" style="--c:{c};--cd:{cd}">
<div class="ubo-head-icon">{icon_svg(tool, 26)}</div>
<div><div class="ubo-eyebrow">{tool.section}</div><div class="ubo-title">{tool.title}</div></div>
</div>
<p class="ubo-lede">{tool.description}</p>
<div class="ubo-rule"></div>"""
