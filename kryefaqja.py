import streamlit as st

import ubo_ui


st.set_page_config(
    page_title="UBO AI Tools",
    layout="wide",
)

# ---------------------------------------------------------------
# STILI – i njëjtë me platformën aitools.uboconsulting.com
# (fontet Space Grotesk / Inter, katër ngjyrat e UBO, kartat me ikonë me ngjyrë).
# Ngjyrat e sfondit, kartave dhe tekstit vijnë nga tema aktive (ubo_ui.base_css: --ubo-*).
# ---------------------------------------------------------------
CSS = """
<style>
@import url("https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600;700&display=swap");

:root {
    --teal: #22C6A0; --blue: #3E7FF5; --yellow: #F4C81E; --magenta: #EC1E8C;
    --rainbow: linear-gradient(90deg, var(--teal), var(--blue) 34%, var(--yellow) 66%, var(--magenta));
    --disp: "Space Grotesk", system-ui, sans-serif;
    --body: "Inter", system-ui, sans-serif;
}

/* Sfondi: njollat e lehta me ngjyrat e UBO në qoshet e faqes */
[data-testid="stApp"] {
    background:
        radial-gradient(38vw 38vw at 12% -4%, rgba(34, 198, 160, calc(.16 * var(--ubo-wash))), transparent 60%),
        radial-gradient(34vw 34vw at 88% 4%, rgba(62, 127, 245, calc(.13 * var(--ubo-wash))), transparent 60%),
        radial-gradient(40vw 40vw at 4% 96%, rgba(244, 200, 30, calc(.13 * var(--ubo-wash))), transparent 62%),
        radial-gradient(40vw 40vw at 96% 92%, rgba(236, 30, 140, calc(.12 * var(--ubo-wash))), transparent 60%),
        var(--ubo-bg);
    background-attachment: fixed;
}
[data-testid="stHeader"] { background: transparent; }

/* Kryefaqja pa menu anësore: kur kthehesh nga një vegël, menyja e saj (ose hija e saj) nuk mbetet */
section[data-testid="stSidebar"],
[data-testid="stExpandSidebarButton"],
[data-testid="stSidebarCollapseButton"] { display: none !important; }

.block-container { max-width: 1140px; padding: 1.5rem 30px 0; }

.lp { font-family: var(--body); color: var(--ubo-ink); }

/* ---------- Kreu ---------- */
a.lp-logo { display: inline-block; height: 64px; transition: opacity .15s; }
a.lp-logo:hover { opacity: .8; }
a.lp-logo img { height: 100%; width: auto; object-fit: contain !important; display: block; }

.lp-hero { padding: 48px 0 10px; }
.lp-h1 {
    font-family: var(--disp); font-weight: 700;
    font-size: clamp(28px, 4vw, 44px); line-height: 1.15; letter-spacing: -.025em;
    margin: 0; color: var(--ubo-ink); white-space: nowrap;
}
/* fjala e theksuar, me shiritin katërngjyrësh poshtë */
.lp-hl { position: relative; white-space: nowrap; z-index: 0; }
.lp-hl::after {
    content: ""; position: absolute; left: -2px; right: -2px; bottom: .2em; height: .3em;
    z-index: -1; border-radius: 4px; opacity: .9; background: var(--rainbow);
}
.lp-lede {
    font-size: clamp(16px, 2vw, 19px); line-height: 1.55; color: var(--ubo-muted);
    margin: 18px 0 0 !important; max-width: 600px;
}

/* ---------- Rreshtat e grupeve ----------
   Rreshti 1: Pyetësorët (1 kartë) + Analiza (2 karta); rreshti 2: Përkthimet (3 karta).
   Në kompjuter secili element merr vendin e tij në rrjetin me 3 kolona (--col / --row);
   në telefon gjithçka shkon njëra pas tjetrës, grup pas grupi. */
.lp-row { display: grid; grid-template-columns: repeat(3, 1fr); column-gap: 22px; row-gap: 18px; margin-top: 40px; }
.lp-row > * { grid-column: var(--col); grid-row: var(--row); }

.lp-barlabel {
    display: flex; align-items: center; gap: 12px; align-self: end;
    font-weight: 700; font-size: 12.5px; letter-spacing: .14em; text-transform: uppercase; color: var(--cd);
}
.lp-barlabel i { width: 9px; height: 9px; border-radius: 50%; background: var(--c); flex-shrink: 0; }
/* vija e grupit me ngjyrën e tij, që dy grupet në të njëjtin rresht të dallohen */
.lp-rule {
    height: 2px; flex: 1; border-radius: 2px; opacity: .7;
    background: linear-gradient(90deg, var(--c), color-mix(in srgb, var(--c) 10%, transparent));
}

/* ---------- Kartat ---------- */
a.lp-card {
    position: relative; display: flex; flex-direction: column;
    min-height: 236px; padding: 26px 26px 22px;
    background: var(--ubo-surface); border: 1px solid var(--ubo-line); border-radius: 20px;
    color: inherit; text-decoration: none !important; overflow: hidden;
    box-shadow: var(--ubo-shadow);
    transition: transform .24s cubic-bezier(.2, .7, .2, 1), box-shadow .24s, border-color .24s;
}
/* vija me ngjyrë që shfaqet sipër kur kalon me miun */
a.lp-card::before {
    content: ""; position: absolute; left: 0; right: 0; top: 0; height: 4px;
    background: var(--c); transform: scaleX(0); transform-origin: left;
    transition: transform .3s cubic-bezier(.2, .7, .2, 1);
}
/* ndriçimi i lehtë me ngjyrën e kartës në qoshen lart djathtas */
a.lp-card::after {
    content: ""; position: absolute; inset: 0; pointer-events: none; opacity: 0; transition: opacity .28s;
    background: radial-gradient(120% 90% at 100% 0%, color-mix(in srgb, var(--c) 15%, transparent), transparent 55%);
}
a.lp-card:hover {
    transform: translateY(-6px);
    box-shadow: var(--ubo-shadow-hover);
    border-color: color-mix(in srgb, var(--c) 55%, var(--ubo-line));
}
a.lp-card:hover::before { transform: scaleX(1); }
a.lp-card:hover::after { opacity: 1; }
a.lp-card:focus-visible { outline: 2px solid var(--c); outline-offset: 3px; }

.lp-ctop { display: flex; align-items: center; justify-content: space-between; margin-bottom: 20px; }
.lp-icon {
    width: 52px; height: 52px; border-radius: 15px; display: grid; place-items: center;
    background: var(--c); box-shadow: 0 8px 22px color-mix(in srgb, var(--c) 45%, transparent);
}
.lp-icon svg { width: 26px; height: 26px; fill: #fff !important; }
.lp-badge {
    font-weight: 700; font-size: 11px; letter-spacing: .09em; text-transform: uppercase;
    border-radius: 999px; padding: 5px 11px; color: var(--ubo-muted);
    border: 1px solid var(--ubo-line); background: var(--ubo-surface-2);
}
.lp-ctitle {
    font-family: var(--disp); font-weight: 700; font-size: 22px; line-height: 1.2;
    letter-spacing: -.02em; margin: 0 0 9px; color: var(--ubo-ink);
}
.lp-cdesc { margin: 0; color: var(--ubo-muted); font-size: 14.5px; line-height: 1.55; }
.lp-cfoot {
    margin-top: auto; padding-top: 20px;
    display: flex; align-items: center; justify-content: space-between; gap: 12px;
}
.lp-tags { font-weight: 600; font-size: 12px; color: var(--ubo-muted); }
.lp-go {
    display: inline-flex; align-items: center; gap: 7px;
    font-weight: 700; font-size: 14.5px; color: var(--cd); white-space: nowrap;
}
.lp-go span { display: inline-block; transition: transform .24s; }
a.lp-card:hover .lp-go span { transform: translateX(4px); }

.lp-foot { color: var(--ubo-muted); font-size: 14px; padding: 42px 0 36px; }
.lp-foot b { color: var(--ubo-ink); font-weight: 700; }

/* Shfaqja e butë e elementeve kur hapet faqja */
@keyframes lp-rise { from { opacity: 0; transform: translateY(14px); } to { opacity: 1; transform: none; } }
.lp-anim { opacity: 0; animation: lp-rise .45s cubic-bezier(.2, .7, .2, 1) forwards; }
@media (prefers-reduced-motion: reduce) {
    .lp-anim { animation: none; opacity: 1; }
    a.lp-card, .lp-go span { transition: none; }
}

@media (max-width: 820px) {
    .lp-row { grid-template-columns: 1fr; margin-top: 30px; row-gap: 14px; }
    .lp-row > * { grid-column: auto; grid-row: auto; }
    .lp-row .lp-barlabel { margin-top: 12px; }
    .lp-h1 { white-space: normal; }
}
@media (max-width: 600px) {
    .block-container { padding: 3.5rem 20px 0; }
    a.lp-logo { height: 50px; }
    .lp-hero { padding: 36px 0 6px; }
    a.lp-card { min-height: 0; }
}
</style>
"""

# Rreshtat e kryefaqes: grupet në secilin rresht (bashkë zënë 3 kolona)
ROWS = [["Pyetësorët", "Analiza"], ["Përkthimet"]]


def tool_card(tool, col, delay):
    c, cd = tool.colors
    badge_html = f'<span class="lp-badge">{tool.badge}</span>' if tool.badge else ""
    return f"""<a href="/{tool.url_path}" target="_self" class="lp-card lp-anim" style="--c:{c};--cd:{cd};--col:{col};--row:2;animation-delay:{delay:.2f}s">
<div class="lp-ctop"><div class="lp-icon">{ubo_ui.icon_svg(tool, 26)}</div>{badge_html}</div>
<div class="lp-ctitle">{tool.title}</div>
<p class="lp-cdesc">{tool.description}</p>
<div class="lp-cfoot"><span class="lp-tags">{tool.tags}</span><span class="lp-go">Hap <span>→</span></span></div>
</a>"""


# ---------------------------------------------------------------
# FAQJA (stili dhe përmbajtja në një element, që të shfaqen njëherësh)
# ---------------------------------------------------------------
html = [CSS, f"""<div class="lp">
<a class="lp-logo lp-anim" href="/" target="_self" title="Kryefaqja">{ubo_ui.logo_imgs()}</a>
<div class="lp-hero">
<div class="lp-h1 lp-anim" style="animation-delay:.04s">Platforma e AI dhe <span class="lp-hl">Automatizimit</span></div>
<p class="lp-lede lp-anim" style="animation-delay:.08s">Veglat e UBO për pyetësorë, përkthime dhe analiza të të dhënave. Zgjidhni një vegël për të vazhduar.</p>
</div>"""]

delay = 0.12
for row in ROWS:
    items, col = [], 1
    for section in row:
        tools = [t for t in ubo_ui.TOOLS if t.section == section]
        c, cd = ubo_ui.section_colors(section)
        items.append(f'<div class="lp-barlabel lp-anim" style="--c:{c};--cd:{cd};--col:{col} / span {len(tools)};'
                     f'--row:1;animation-delay:{delay:.2f}s"><i></i>{section}<span class="lp-rule"></span></div>')
        for tool in tools:
            items.append(tool_card(tool, col, delay))
            col += 1
            delay += 0.03
    html.append(f'<div class="lp-row">{"".join(items)}</div>')

html.append('<div class="lp-foot"><b>UBO Consulting</b> · Platforma e AI dhe Automatizimit</div></div>')
st.markdown("\n".join(html), unsafe_allow_html=True)
