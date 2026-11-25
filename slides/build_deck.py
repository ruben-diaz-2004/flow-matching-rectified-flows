"""Genera FlowMatching_RectifiedFlows.pptx con python-pptx.

Requisitos: ejecutar antes los notebooks (figuras en ../notebooks/figuras) y extra_figs.py (assets/).
Uso (desde slides/):  ..\\.venv\\Scripts\\python build_deck.py
"""
import os
from PIL import Image
from lxml import etree
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_AXIS_CROSSES, XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION, XL_MARKER_STYLE
from pptx.oxml.ns import qn

FIG = "../notebooks/figuras"
AS = "assets"
os.makedirs(f"{AS}/crops", exist_ok=True)

# ------------------------------------------------------------------ paleta
def rgb(h): return RGBColor.from_string(h)
NAVY, TEAL, CRIMSON, GOLD, BLUE = "14213D", "00798C", "D1495B", "EDAE49", "30638E"
INK, MUTED, TINT, TINT2, WHITE = "1F2933", "5F6B7A", "EAF3F4", "FBEDEF", "FFFFFF"
HEAD, BODY, CODE = "Cambria", "Calibri", "Consolas"

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
BLANK = prs.slide_layouts[6]
SW, SH = 13.333, 7.5
_slide_no = [0]
DECK_SHORT = "Flow Matching & Rectified Flows"


# ------------------------------------------------------------------ helpers
import re
_SUB = re.compile(r"_(?:\{([^}]*)\}|([^\W_]))")


def _runs(p, text, size, color, bold=False, italic=False, font=BODY, hl=TEAL):
    """Texto con **negrita resaltada** y subíndices x_t / x_{min} (escapar con \\_)."""
    parts = text.split("**")
    for i, part in enumerate(parts):
        if not part:
            continue
        part = part.replace("\\_", "\x00")
        pieces, pos = [], 0
        if font != CODE:
            for m in _SUB.finditer(part):
                pieces.append((part[pos:m.start()], False)); pieces.append((m.group(1) or m.group(2), True)); pos = m.end()
        pieces.append((part[pos:], False))
        for seg, is_sub in pieces:
            if not seg:
                continue
            r = p.add_run(); r.text = seg.replace("\x00", "_")
            f = r.font; f.size = Pt(size); f.name = font; f.italic = italic
            f.bold = bold or (i % 2 == 1)
            f.color.rgb = rgb(hl if (i % 2 == 1 and hl) else color)
            if is_sub:
                f._rPr.set("baseline", "-25000")


def text(slide, x, y, w, h, s, size=16, color=INK, bold=False, italic=False, font=BODY,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, hl=TEAL, spacing=None, name=None):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if name: tb.name = name
    tf = tb.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    lines = s if isinstance(s, list) else [s]
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if spacing: p.space_after = Pt(spacing)
        _runs(p, line, size, color, bold, italic, font, hl)
    return tb


def bullets(slide, x, y, w, h, items, size=16, color=INK, gap=8, hl=TEAL, bullet_color=TEAL):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, it in enumerate(items):
        lvl = 0
        if isinstance(it, tuple): it, lvl = it
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(gap)
        pPr = p._p.get_or_add_pPr()
        ind = Emu(Inches(0.28)); marl = Emu(Inches(0.28 + 0.32 * lvl))
        pPr.set("marL", str(marl)); pPr.set("indent", str(-ind))
        bc = etree.SubElement(pPr, qn("a:buClr")); sc = etree.SubElement(bc, qn("a:srgbClr")); sc.set("val", bullet_color)
        bf = etree.SubElement(pPr, qn("a:buFont")); bf.set("typeface", "Arial")
        bu = etree.SubElement(pPr, qn("a:buChar")); bu.set("char", "•" if lvl == 0 else "–")
        _runs(p, it, size - 2 * lvl, color if lvl == 0 else MUTED, hl=hl)
    return tb


def rect(slide, x, y, w, h, fill=TINT, radius=0.06, line=None, shape=MSO_SHAPE.ROUNDED_RECTANGLE, name=None):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if name: s.name = name
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = radius
    if fill is None: s.fill.background()
    else:
        s.fill.solid(); s.fill.fore_color.rgb = rgb(fill)
    if line: s.line.color.rgb = rgb(line); s.line.width = Pt(1.25)
    else: s.line.fill.background()
    s.shadow.inherit = False
    s.text_frame.margin_left = s.text_frame.margin_right = Inches(0.15)
    return s


def circle_num(slide, x, y, d, label, fill=TEAL, color=WHITE, size=16):
    c = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(d), Inches(d))
    c.fill.solid(); c.fill.fore_color.rgb = rgb(fill); c.line.fill.background(); c.shadow.inherit = False
    tf = c.text_frame; tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    _runs(p, str(label), size, color, bold=True, font=HEAD)
    return c


def img_size(path):
    with Image.open(path) as im:
        return im.size


def picture(slide, path, x, y, w=None, h=None, box=None, align="center"):
    """Inserta imagen; con box=(w,h) la encaja manteniendo proporción."""
    pw, ph = img_size(path)
    if box:
        bw, bh = box
        s = min(bw / pw, bh / ph); w, h = pw * s, ph * s
        if align == "center":
            x, y = x + (bw - w) / 2, y + (bh - h) / 2
        elif align == "top":
            x = x + (bw - w) / 2
    elif w and not h: h = w * ph / pw
    elif h and not w: w = h * pw / ph
    return slide.shapes.add_picture(path, Inches(x), Inches(y), Inches(w), Inches(h))


def eq(slide, name, x, y, scale=1.0, white=False, center_w=None):
    """Ecuación renderizada (300 dpi, 28 pt): tamaño natural * scale."""
    path = f"{AS}/eq/{name}{'_w' if white else ''}.png"
    pw, ph = img_size(path)
    w, h = pw / 300 * scale, ph / 300 * scale
    if center_w: x = x + (center_w - w) / 2
    slide.shapes.add_picture(path, Inches(x), Inches(y), Inches(w), Inches(h))
    return w, h


def crop(src, box_frac, name):
    """Recorta una figura por fracciones (l, t, r, b)."""
    out = f"{AS}/crops/{name}.png"
    with Image.open(src) as im:
        W, H = im.size
        l, t, r, b = box_frac
        im.crop((int(l * W), int(t * H), int(r * W), int(b * H))).save(out)
    return out


def notes(slide, s):
    slide.notes_slide.notes_text_frame.text = s.strip()


def footer(slide, dark=False):
    _slide_no[0] += 1
    c = "8C9BAB" if dark else MUTED
    text(slide, 0.6, 7.0, 6, 0.3, DECK_SHORT, size=10, color=c)
    text(slide, SW - 1.6, 7.0, 1.0, 0.3, str(_slide_no[0]), size=10, color=c, align=PP_ALIGN.RIGHT)


def bg(slide, color):
    f = slide.background.fill; f.solid(); f.fore_color.rgb = rgb(color)


def content(kicker, title, sub=None):
    s = prs.slides.add_slide(BLANK); bg(s, WHITE)
    text(s, 0.6, 0.38, 12, 0.3, kicker.upper(), size=12, color=TEAL, bold=True)
    text(s, 0.6, 0.68, 12.1, 0.75, title, size=32, color=INK, bold=True, font=HEAD, name="Title")
    if sub: text(s, 0.6, 1.38, 12.1, 0.4, sub, size=16, color=MUTED, italic=True)
    footer(s)
    return s


def divider(num, title, sub, kicker="Sección"):
    s = prs.slides.add_slide(BLANK); bg(s, NAVY)
    picture(s, f"{AS}/cover.png", 7.4, 0, h=SH)
    rect(s, 7.4, 0, 0.01, 0.01, fill=NAVY)  # (ancla para orden)
    text(s, 0.8, 2.2, 6.3, 1.2, num, size=72, color=CRIMSON, bold=True, font=HEAD)
    text(s, 0.8, 3.45, 6.4, 1.6, title, size=40, color=WHITE, bold=True, font=HEAD)
    text(s, 0.8, 4.95, 6.2, 1.0, sub, size=18, color="C9D3DE", italic=True)
    footer(s, dark=True)
    return s


def code_box(slide, x, y, w, h, title, lines, size=13):
    rect(slide, x, y, w, h, fill="1B2638", radius=0.04)
    text(slide, x + 0.25, y + 0.15, w - 0.5, 0.35, title, size=13, color=GOLD, bold=True)
    tb = slide.shapes.add_textbox(Inches(x + 0.25), Inches(y + 0.55), Inches(w - 0.4), Inches(h - 0.7))
    tf = tb.text_frame; tf.word_wrap = False
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(2)
        if "#" in ln:
            c_idx = ln.index("#"); code, com = ln[:c_idx], ln[c_idx:]
        else:
            code, com = ln, ""
        r = p.add_run(); r.text = code if code else " "
        r.font.name = CODE; r.font.size = Pt(size); r.font.color.rgb = rgb("E6EDF3")
        if com:
            r2 = p.add_run(); r2.text = com
            r2.font.name = CODE; r2.font.size = Pt(size); r2.font.color.rgb = rgb("7FB8C2")


def stat(slide, x, y, w, h, big, label, color=TEAL, fill=TINT, big_size=40, label_size=14):
    rect(slide, x, y, w, h, fill=fill)
    text(slide, x + 0.25, y + 0.2, w - 0.5, 0.8, big, size=big_size, color=color, bold=True, font=HEAD)
    text(slide, x + 0.25, y + 0.2 + big_size / 72 * 1.25, w - 0.5, h - 0.4 - big_size / 72 * 1.25, label,
         size=label_size, color=INK)


def source(slide, s, y=6.62):
    text(slide, 0.6, y, 12.1, 0.3, s, size=10, color=MUTED, italic=True)


def style_chart(chart, size=12):
    chart.font.size = Pt(size); chart.font.name = BODY; chart.font.color.rgb = rgb(MUTED)


def log_axis(axis):
    scaling = axis._element.find(qn("c:scaling"))
    lb = etree.SubElement(scaling, qn("c:logBase")); lb.set("val", "10")
    scaling.remove(lb); scaling.insert(0, lb)


def gridlines(axis, color="E1E6EB"):
    axis.has_major_gridlines = True
    gl = axis.major_gridlines.format.line; gl.color.rgb = rgb(color); gl.width = Pt(0.75)


def arrow(slide, x, y, w=0.45, h=0.4, color="B8C2CC"):
    a = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(x), Inches(y), Inches(w), Inches(h))
    a.fill.solid(); a.fill.fore_color.rgb = rgb(color); a.line.fill.background(); a.shadow.inherit = False


# ================================================================== DIAPOSITIVAS
# 1 — Portada ---------------------------------------------------------------
s = prs.slides.add_slide(BLANK); bg(s, NAVY)
picture(s, f"{AS}/cover.png", 6.2, 0, h=SH)
text(s, 0.8, 1.0, 6, 0.4, "APRENDIZAJE PROFUNDO · CURSO 2026–27", size=13, color=GOLD, bold=True)
text(s, 0.8, 1.9, 6.6, 2.2, ["Flow Matching &", "Rectified Flows"], size=48, color=WHITE, bold=True, font=HEAD)
text(s, 0.8, 4.0, 5.9, 1.2, "Aprender a transportar ruido en datos: de los flujos continuos a Stable Diffusion 3",
     size=20, color="C9D3DE", italic=True)
text(s, 0.8, 5.6, 6, 0.4, "[Nombre Apellido]  ·  [Nombre Apellido]", size=16, color=WHITE)
_slide_no[0] += 1
notes(s, """
[P1 · 0:00–0:45]
Presentación. "Hoy vamos a ver cómo se entrenan los generadores de imágenes más recientes, como Stable Diffusion 3 o Flux, y por qué ya no son 'modelos de difusión' en sentido estricto."
La imagen de fondo la genera nuestro propio notebook: cada línea es la trayectoria de un punto de ruido que el modelo transporta hasta la distribución de datos (dos lunas).
""")

# 2 — Hoja de ruta ------------------------------------------------------------
s = content("Hoja de ruta", "Una sola idea, cinco pasos")
blocks = [("Generar = transportar", "Flujos continuos (CNF) y por qué son caros de entrenar", "P1 · 8 min"),
          ("Flow Matching", "Entrenar el flujo con regresión, sin simular la ODE", "P1 · 12 min"),
          ("Conexión con difusión", "Mismo modelo, otra parametrización", "P2 · 6 min"),
          ("Rectified Flow", "Enderezar trayectorias → 1–2 pasos", "P2 · 10 min"),
          ("Demo y estado del arte", "Notebooks · SD3 · modelos de un paso", "P2 · 9 min")]
x0, cw, gap = 0.6, 2.3, 0.15
for i, (t_, d_, m_) in enumerate(blocks):
    x = x0 + i * (cw + gap)
    rect(s, x, 2.1, cw, 3.6, fill=TINT if i < 2 else TINT2)
    circle_num(s, x + 0.25, 2.35, 0.65, i + 1, fill=TEAL if i < 2 else CRIMSON, size=20)
    text(s, x + 0.25, 3.25, cw - 0.45, 0.9, t_, size=19, color=INK, bold=True, font=HEAD)
    text(s, x + 0.25, 4.15, cw - 0.45, 1.1, d_, size=14, color=INK)
    text(s, x + 0.25, 5.25, cw - 0.45, 0.3, m_, size=12, color=MUTED, bold=True)
text(s, 0.6, 6.0, 12.1, 0.5, "Material: diapositivas + 2 notebooks (Colab, CPU) con todo el código y las figuras de esta presentación.",
     size=14, color=MUTED, italic=True)
notes(s, """
[P1 · 0:45–1:30]
La presentación sigue un único hilo: generar datos es transportar una distribución sencilla (ruido) hasta la de los datos.
1-2 los presento yo (P1): flujos continuos y Flow Matching. 3-5 mi compañero/a (P2): relación con difusión, Rectified Flow, demo y estado del arte.
Todas las figuras con datos 2D salen de los notebooks que os dejamos: se ejecutan en CPU en pocos minutos.
""")

# 3 — Generar es transportar --------------------------------------------------
s = content("01 · Motivación", "Generar es transportar una distribución en otra")
text(s, 0.6, 1.6, 12.1, 0.9,
     "Tenemos muestras de una distribución desconocida **q** (los datos) y sabemos muestrear una fácil **p₀ = 𝒩(0, I)**. "
     "Buscamos un mapa **T** tal que si x₀ ~ p₀, entonces T(x₀) ~ q.", size=18)
picture(s, f"{FIG}/nb1_datasets.png", 0.6, 2.75, box=(12.1, 3.2))
text(s, 0.6, 6.05, 12.1, 0.5, "GANs, VAEs y normalizing flows aprenden T en un solo salto. Difusión y flujos continuos lo construyen **poco a poco, en el tiempo**.",
     size=15, color=MUTED)
notes(s, """
[P1 · 1:30–3:00]
Planteamiento general del modelado generativo. Izquierda: ruido gaussiano, fácil de muestrear. Derecha: tres distribuciones 2D que usaremos como 'datos' (en imágenes, cada punto sería una imagen).
Queremos un transporte T que lleve una nube a la otra. Distintas familias lo hacen de forma distinta: GAN/VAE/NF en un paso; difusión y flujos continuos de forma gradual, siguiendo una dinámica en el tiempo t ∈ [0,1].
Convención que usaremos todo el rato: t = 0 es ruido, t = 1 son datos (como Lipman et al. y Liu et al.; ojo, SD3 usa la contraria).
""")

# 4 — Qué falla en difusión ----------------------------------------------------
s = content("01 · Motivación", "La difusión funciona, pero es lenta y compleja")
stat(s, 0.6, 1.75, 3.85, 2.6, "~1000", "evaluaciones de la red para generar **una** muestra con DDPM (T = 1000 pasos)", color=CRIMSON, fill=TINT2)
stat(s, 4.75, 1.75, 3.85, 2.6, "FID 451", "si integramos la ODE de difusión (VP) en **un solo paso** en CIFAR-10: ruido puro", color=CRIMSON, fill=TINT2)
stat(s, 8.9, 1.75, 3.85, 2.6, "SDE + …", "schedule de ruido, parametrización (ε, x, v), ponderación por t, muestreador: muchas decisiones", color=CRIMSON, fill=TINT2, big_size=36)
rect(s, 0.6, 4.75, 12.15, 1.45, fill=TINT)
text(s, 0.9, 4.95, 11.6, 1.1,
     ["**Pregunta de la sesión:** ¿podemos aprender un transporte continuo que sea",
      "simple de entrenar (una regresión), determinista y rápido de muestrear (pocos pasos)?"], size=19, color=INK)
source(s, "Ho et al. (2020) DDPM; FID de la VP ODE con 1 paso: Liu et al. (2023), Tabla 1.")
notes(s, """
[P1 · 3:00–4:30]
Supongo que conocéis DDPM / score matching. Tres problemas prácticos:
1) Coste de muestreo: cientos o miles de pasos, cada uno una pasada por una U-Net/Transformer enorme.
2) Si intentamos ir rápido con pocos pasos, la calidad se desploma: la ODE de difusión con un paso da FID 451 en CIFAR-10 (es ruido). La razón (lo veremos) es que sus trayectorias son curvas.
3) Formulación con muchas piezas: SDE directa e inversa, schedules, parametrizaciones…
La pregunta que guía la charla: ¿se puede hacer algo más simple y más rápido? Spoiler: Flow Matching + Rectified Flow.
""")

# 5 — CNF -----------------------------------------------------------------------
s = content("01 · Flujos continuos", "Flujos continuos: una ODE aprendida")
eq(s, "ode", 0.8, 1.85, scale=0.85)
eq(s, "pushforward", 0.8, 2.95, scale=0.85)
bullets(s, 0.6, 3.95, 6.6, 2.6, [
    "**v_θ(x, t)**: campo de velocidades = red neuronal (MLP, U-Net, Transformer)",
    "**φ_t**: mapa de flujo, x₀ ↦ x_t. Es invertible y suave",
    "**p_t**: densidad de las partículas en el instante t (push-forward de p₀)",
    "**Muestrear = resolver la ODE** desde t = 0 hasta t = 1",
], size=16)
picture(s, crop(f"{FIG}/nb1_samples_traj.png", (0.675, 0.0, 1.0, 1.0), "traj_only"), 7.6, 1.6, box=(5.2, 5.0))
notes(s, """
[P1 · 4:30–6:00]
Continuous Normalizing Flows (Chen et al., 2018, Neural ODE). En vez de una cadena discreta de capas invertibles, definimos un campo de velocidades v_θ(x,t) y dejamos que los puntos 'fluyan'.
Cada punto de ruido x₀ sigue una trayectoria; al llegar a t = 1 debería ser una muestra de los datos. La figura (de nuestro notebook) muestra trayectorias reales: puntos grises de ruido, puntos rojos generados.
Ventaja: cualquier arquitectura vale para v_θ (no hace falta que sea invertible: la ODE garantiza invertibilidad).
""")

# 6 — Continuidad / MLE caro ---------------------------------------------------
s = content("01 · Flujos continuos", "Entrenar una CNF por máxima verosimilitud es caro")
rect(s, 0.6, 1.7, 6.0, 2.35, fill=TINT)
text(s, 0.85, 1.85, 5.5, 0.35, "Ecuación de continuidad (conservación de masa)", size=14, color=TEAL, bold=True)
eq(s, "continuity", 0.6, 2.35, scale=0.72, center_w=6.0)
text(s, 0.85, 3.3, 5.5, 0.6, "El campo v_t determina cómo evoluciona la densidad p_t.", size=14, color=INK)
rect(s, 6.85, 1.7, 5.9, 2.35, fill=TINT)
text(s, 7.1, 1.85, 5.4, 0.35, "Cambio de variables instantáneo", size=14, color=TEAL, bold=True)
eq(s, "logp", 6.85, 2.3, scale=0.72, center_w=5.9)
text(s, 7.1, 3.3, 5.4, 0.6, "Permite calcular log p₁(x) exacto… integrando la ODE.", size=14, color=INK)
items = [("1", "Simular la ODE en **cada** iteración de entrenamiento (decenas o cientos de pasos)"),
         ("2", "Estimar la traza del Jacobiano ∇·v_θ (Hutchinson): ruidoso y caro en alta dimensión"),
         ("3", "Dinámicas aprendidas cada vez más complejas → la ODE se vuelve rígida y lenta")]
for i, (n_, t_) in enumerate(items):
    circle_num(s, 0.6, 4.4 + i * 0.62, 0.45, n_, fill=CRIMSON, size=14)
    text(s, 1.25, 4.45 + i * 0.62, 11.3, 0.45, t_, size=16)
text(s, 0.6, 6.3, 12.1, 0.35, "Idea: ¿y si **no** simulamos y hacemos directamente **regresión** del campo de velocidades?", size=16, color=INK)
notes(s, """
[P1 · 6:00–8:30]
La ecuación de continuidad dice que la masa de probabilidad se conserva: si sé v_t, sé cómo evoluciona p_t. De ahí sale la fórmula del log-likelihood instantáneo (derecha), que permite entrenar por máxima verosimilitud (FFJORD, Grathwohl et al. 2019).
El problema es computacional: para cada batch hay que integrar la ODE (y la divergencia) con un solver adaptativo; en imágenes esto es carísimo y escala mal. Por eso las CNF no despegaron en imágenes de alta resolución.
Transición: Flow Matching cambia el planteamiento por completo: en lugar de maximizar verosimilitud simulando, hacemos regresión de un campo conocido. 'Simulation-free training'.
""")

# 7 — Divider Flow Matching ------------------------------------------------------
s = divider("02", "Flow Matching", "Entrenar flujos continuos sin simularlos")
notes(s, "[P1 · 8:30] Transición a la idea central de la presentación.")

# 8 — Pérdida FM ----------------------------------------------------------------
s = content("02 · Flow Matching", "La idea: hacer regresión del campo de velocidades")
text(s, 0.6, 1.65, 12.1, 0.8, "Elegimos un **camino de probabilidad** p_t que vaya de p₀ (ruido) a q (datos), generado por un campo u_t, y ajustamos v_θ a u_t:",
     size=18)
rect(s, 1.6, 2.55, 10.1, 1.25, fill=TINT)
eq(s, "lfm", 1.6, 2.75, scale=0.95, center_w=10.1)
rect(s, 0.6, 4.25, 5.95, 1.95, fill=TINT)
text(s, 0.85, 4.4, 5.5, 0.4, "Lo bueno", size=16, color=TEAL, bold=True)
text(s, 0.85, 4.85, 5.5, 1.3, "Es una simple regresión por mínimos cuadrados: sin ODE, sin trazas, sin log-verosimilitud. Si la pérdida es 0, el flujo genera exactamente p_t.", size=15)
rect(s, 6.8, 4.25, 5.95, 1.95, fill=TINT2)
text(s, 7.05, 4.4, 5.5, 0.4, "El problema", size=16, color=CRIMSON, bold=True, hl=CRIMSON)
text(s, 7.05, 4.85, 5.5, 1.3, "No conocemos ni p_t ni u_t: para eso necesitaríamos conocer q. ¿Cómo construimos un objetivo calculable?", size=15)
source(s, "Lipman, Chen, Ben-Hamu, Nickel & Le (2023). Flow Matching for Generative Modeling. ICLR.")
notes(s, """
[P1 · 8:45–10:15]
Flow Matching (Lipman et al., ICLR 2023). Si alguien nos diera un campo u_t que transporta el ruido hasta los datos, entrenar sería trivial: regresión con MSE en puntos x muestreados de p_t y tiempos t uniformes.
Pero ese campo depende de la distribución de datos, que no conocemos analíticamente. La pérdida así escrita es intratable. La clave del paper es el siguiente truco.
""")

# 9 — Caminos condicionales ------------------------------------------------------
s = content("02 · Flow Matching", "El truco: construir p_t a partir de caminos condicionales")
rect(s, 0.6, 1.65, 5.95, 2.4, fill=TINT)
text(s, 0.85, 1.8, 5.5, 0.35, "Para cada dato x₁: un camino trivial", size=15, color=TEAL, bold=True)
eq(s, "condpath", 0.6, 2.3, scale=0.8, center_w=5.95)
eq(s, "condvel", 0.6, 3.05, scale=0.8, center_w=5.95)
rect(s, 6.8, 1.65, 5.95, 2.4, fill=TINT)
text(s, 7.05, 1.8, 5.5, 0.35, "Marginalizando sobre los datos", size=15, color=TEAL, bold=True)
eq(s, "marginal_p", 6.8, 2.3, scale=0.8, center_w=5.95)
eq(s, "marginal_u", 6.8, 3.1, scale=0.8, center_w=5.95)
picture(s, f"{FIG}/nb1_prob_path.png", 0.6, 4.3, box=(12.15, 2.3))
notes(s, """
[P1 · 10:15–12:30]
Truco: definimos el camino 'dato a dato'. Fijado un dato x₁, el camino condicional es la interpolación lineal entre ruido x₀ y x₁; su velocidad es constante: x₁ − x₀.
(Con x₀ gaussiano, p_t(x|x₁) = N(t x₁, (1−t)² I): una gaussiana que se va estrechando hacia x₁.)
El camino marginal p_t es la mezcla de todos estos caminos condicionales, y el campo marginal u_t es la media de las velocidades condicionales de todos los pares que pasan por x en el instante t.
Figura de abajo: snapshots de p_t (de ruido a lunas) y, a la derecha, pares condicionales: las rectas se cruzan. Recordad esto, porque es la clave de Rectified Flow.
""")

# 10 — Teorema --------------------------------------------------------------------
s = content("02 · Flow Matching", "Teorema: basta con la pérdida condicional")
rect(s, 0.6, 1.65, 12.15, 1.95, fill=NAVY)
eq(s, "lcfm", 0.6, 1.85, scale=0.9, white=True, center_w=12.15)
eq(s, "theorem", 0.6, 2.8, scale=0.85, white=True, center_w=12.15)
text(s, 0.6, 3.9, 12, 0.4, "Idea de la demostración", size=16, color=TEAL, bold=True)
circle_num(s, 0.6, 4.4, 0.42, 1, size=13)
eq(s, "proof1", 1.2, 4.35, scale=0.68)
text(s, 7.3, 4.42, 5.4, 0.5, "el último término no depende de θ", size=14, color=MUTED, italic=True)
circle_num(s, 0.6, 5.15, 0.42, 2, size=13)
eq(s, "proof2", 1.2, 5.08, scale=0.62)
circle_num(s, 0.6, 5.9, 0.42, 3, size=13)
text(s, 1.2, 5.92, 11.5, 0.5, "El término cruzado coincide (definición de u_t como esperanza condicional) ⇒ **𝓛_{CFM} = 𝓛_{FM} + constante**",
     size=16)
source(s, "Lipman et al. (2023), Teorema 2. Mismo resultado con caminos gaussianos generales y σ_{min} > 0.")
notes(s, """
[P1 · 12:30–15:00]
Este es el resultado central. En vez de regresar el campo marginal (desconocido), regresamos la velocidad condicional x₁ − x₀, que conocemos para cada par muestreado.
Demostración en tres líneas: desarrollamos el cuadrado; el término ‖u‖² no depende de θ; el término ‖v_θ‖² es igual en ambas pérdidas; y el término cruzado coincide porque u_t(x) es justo la esperanza de la velocidad condicional dado x_t = x (intercambiar esperanzas).
Por tanto las dos pérdidas difieren en una constante: mismos gradientes, mismo minimizador. Es el mismo argumento que denoising score matching vs. score matching (Vincent, 2011).
Consecuencia práctica: la pérdida nunca llega a cero (esa constante es la varianza irreducible del objetivo condicional).
""")

# 11 — Algoritmo ------------------------------------------------------------------
s = content("02 · Flow Matching", "El algoritmo completo cabe en una diapositiva")
code_box(s, 0.6, 1.65, 6.9, 2.9, "Entrenamiento (Conditional Flow Matching)", [
    "for x1 in dataloader:                 # datos",
    "    x0 = torch.randn_like(x1)         # ruido",
    "    t  = torch.rand(len(x1), 1)       # t ~ U[0,1]",
    "    xt = (1 - t) * x0 + t * x1        # interpolar",
    "    loss = ((v(xt, t) - (x1 - x0))**2).mean()",
    "    loss.backward(); opt.step(); opt.zero_grad()",
], size=13)
code_box(s, 7.75, 1.65, 5.0, 2.9, "Muestreo (Euler)", [
    "x = torch.randn(n, d)       # x0 ~ N(0,I)",
    "h = 1 / steps",
    "for i in range(steps):",
    "    x = x + h * v(x, i * h)",
    "return x                    # ~ q",
], size=14)
for i, (b_, d_) in enumerate([("Sin simulación", "la ODE sólo aparece al muestrear"),
                              ("Una regresión MSE", "estable y fácil de escalar"),
                              ("Cualquier arquitectura", "MLP, U-Net, Transformer (DiT)")]):
    x = 0.6 + i * 4.1
    rect(s, x, 4.9, 3.9, 1.25, fill=TINT)
    text(s, x + 0.25, 5.05, 3.4, 0.4, b_, size=17, color=TEAL, bold=True)
    text(s, x + 0.25, 5.5, 3.4, 0.5, d_, size=15, color=INK)
notes(s, """
[P1 · 15:00–16:30]
Este es literalmente el bucle de entrenamiento del notebook 1. Seis líneas. Comparad con DDPM: aquí no hay schedule de betas, ni alphas acumuladas, ni pesos por t.
El muestreo: integrar la ODE con Euler (o cualquier solver: punto medio, RK45, DPM-solver…).
Coste de entrenamiento por iteración: una sola pasada forward/backward, igual que difusión, pero muchísimo menor que una CNF por máxima verosimilitud.
""")

# 12 — Resultado 2D -----------------------------------------------------------------
s = content("02 · Flow Matching", "Funciona: trayectorias que no se cruzan")
picture(s, f"{FIG}/nb1_samples_traj.png", 0.6, 1.6, box=(12.15, 4.1))
bullets(s, 0.6, 5.85, 12.1, 0.9, [
    "Las trayectorias **no se cruzan** (unicidad de la ODE), aunque las rectas condicionales sí lo hacían ⇒ las trayectorias son **curvas**",
    "MLP de 3 capas · 6 000 iteraciones · ~20 s en CPU (Notebook 1)",
], size=15, gap=4)
notes(s, """
[P1 · 16:30–18:00]
Resultado del notebook 1 en las dos lunas: izquierda datos reales, centro generados con 100 pasos de Euler, derecha trayectorias.
Fijaos en la observación clave: la ODE aprendida no puede cruzar trayectorias (por unicidad de soluciones). Como el campo marginal promedia las rectas que se cruzan, las trayectorias salen curvas. Esto tendrá consecuencias en el número de pasos: lo retomará mi compañero/a en Rectified Flow.
""")

# 13 — Resultados en imágenes --------------------------------------------------------
s = content("02 · Flow Matching", "En imágenes: mejor calidad con la mitad de evaluaciones")
cd = CategoryChartData(); cd.categories = ["DDPM", "FM (camino OT)"]
cd.add_series("FID", (7.48, 6.35))
g = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.6), Inches(1.7), Inches(5.9), Inches(4.4), cd)
ch = g.chart; style_chart(ch); ch.has_legend = False
ch.has_title = True; ch.chart_title.text_frame.text = "FID en CIFAR-10 (↓ mejor)"
ch.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(15); ch.chart_title.text_frame.paragraphs[0].runs[0].font.color.rgb = rgb(INK)
pl = ch.plots[0]; pl.gap_width = 80; pl.has_data_labels = True
pl.data_labels.number_format = "0.00"; pl.data_labels.number_format_is_linked = False
pl.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END; pl.data_labels.font.size = Pt(14); pl.data_labels.font.bold = True
for i, c in enumerate([MUTED, TEAL]):
    pt = pl.series[0].points[i]; pt.format.fill.solid(); pt.format.fill.fore_color.rgb = rgb(c)
ch.value_axis.visible = False; ch.value_axis.has_major_gridlines = False
ch.category_axis.tick_labels.font.size = Pt(14); ch.category_axis.format.line.color.rgb = rgb("C9D1D9")
cd = CategoryChartData(); cd.categories = ["DDPM", "FM (camino OT)"]
cd.add_series("NFE", (274, 142))
g = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(6.85), Inches(1.7), Inches(5.9), Inches(4.4), cd)
ch = g.chart; style_chart(ch); ch.has_legend = False
ch.has_title = True; ch.chart_title.text_frame.text = "Evaluaciones de la red al muestrear (↓ mejor)"
ch.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(15); ch.chart_title.text_frame.paragraphs[0].runs[0].font.color.rgb = rgb(INK)
pl = ch.plots[0]; pl.gap_width = 80; pl.has_data_labels = True
pl.data_labels.number_format = "0"; pl.data_labels.number_format_is_linked = False
pl.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END; pl.data_labels.font.size = Pt(14); pl.data_labels.font.bold = True
for i, c in enumerate([MUTED, TEAL]):
    pt = pl.series[0].points[i]; pt.format.fill.solid(); pt.format.fill.fore_color.rgb = rgb(c)
ch.value_axis.visible = False; ch.value_axis.has_major_gridlines = False
ch.category_axis.tick_labels.font.size = Pt(14); ch.category_axis.format.line.color.rgb = rgb("C9D1D9")
text(s, 0.6, 6.2, 12.1, 0.4, "Misma arquitectura U-Net, sólo cambia el objetivo de entrenamiento. NFE con solver adaptativo (dopri5).", size=14, color=INK)
source(s, "Lipman et al. (2023), Tabla 1 (CIFAR-10). El paper reporta mejoras análogas en ImageNet 32×32, 64×64 y 128×128.")
notes(s, """
[P1 · 18:00–20:00]
Resultados del paper original en CIFAR-10. Mismo modelo y presupuesto; sólo cambia el objetivo: DDPM vs Flow Matching con el camino lineal ('OT path').
FM consigue mejor FID (6.35 vs 7.48) y necesita aproximadamente la mitad de evaluaciones de la red para muestrear con un solver adaptativo (142 vs 274).
Interpretación: caminos más rectos → ODE más fácil de integrar. Con esto termino mi parte. Paso la palabra a [P2], que os va a contar qué relación tiene esto con difusión y cómo llegar a uno o dos pasos.
""")

# 14 — Divider difusión -------------------------------------------------------------
s = divider("03", "Conexión con difusión", "Mismo modelo, otra parametrización")
notes(s, "[P2 · 20:00] Cambio de presentador. 'Seguramente os estáis preguntando: ¿esto no es lo mismo que difusión?'")

# 15 — Difusión caso particular ------------------------------------------------------
s = content("03 · Conexión con difusión", "La difusión gaussiana es un caso particular")
rect(s, 0.6, 1.65, 12.15, 1.05, fill=TINT)
eq(s, "gausspath", 0.6, 1.85, scale=0.85, center_w=12.15)
picture(s, f"{AS}/schedules.png", 0.6, 2.95, box=(7.6, 3.6), align="top")
bullets(s, 8.5, 3.05, 4.25, 3.6, [
    "**FM-OT / Rectified Flow**: α_t = t, σ_t = 1 − t → caminos rectos",
    "**Difusión VP** (DDPM, coseno): α_t² + σ_t² = 1 → caminos curvos",
    "**Difusión VE** (SMLD): α_t = 1, σ_t decreciente",
    "Todos se entrenan con la **misma pérdida CFM**, cambiando sólo (α_t, σ_t)",
], size=15, gap=8)
source(s, "Lipman et al. (2023) §4.1; Albergo & Vanden-Eijnden (2023) Stochastic Interpolants; Gao et al. (2024) Diffusion meets Flow Matching.")
notes(s, """
[P2 · 20:15–22:00]
Generalizamos el camino condicional: x_t = α_t x₁ + σ_t x₀, con velocidad condicional α̇ x₁ + σ̇ x₀. Cualquier difusión gaussiana es de esta forma: VP (DDPM) tiene α² + σ² = 1, VE (SMLD/NCSN) tiene α = 1.
A la izquierda: los schedules. A la derecha: el camino de un único par. Con el camino lineal el par se une en línea recta; con VP describe un arco.
Mensaje: Flow Matching no es 'un modelo distinto' a la difusión: es un marco más general donde la difusión es un caso concreto, y en el que podemos elegir caminos más simples. Este marco unificado también se conoce como 'stochastic interpolants'.
""")

# 16 — Parametrizaciones ---------------------------------------------------------------
s = content("03 · Conexión con difusión", "Velocidad, ruido, dato o score: lo mismo")
cards = [("ε-predicción (DDPM)", "x0hat"), ("x-predicción (denoiser)", "x1hat"), ("score (SMLD / SDE)", "score")]
for i, (t_, e_) in enumerate(cards):
    x = 0.6 + i * 4.1
    rect(s, x, 1.65, 3.9, 1.5, fill=TINT)
    text(s, x + 0.2, 1.75, 3.5, 0.35, t_, size=14, color=TEAL, bold=True)
    eq(s, e_, x, 2.2, scale=0.62, center_w=3.9)
picture(s, f"{FIG}/nb1_denoiser.png", 0.6, 3.35, box=(12.15, 3.2))
source(s, "Camino lineal con x₀ ~ 𝒩(0, I). Relaciones análogas para cualquier (α_t, σ_t). Figura: Notebook 1, sección 8.")
notes(s, """
[P2 · 22:00–24:00]
Con el camino lineal y ruido gaussiano, condicionado a x₁ tenemos x_t ~ N(t x₁, (1−t)² I): exactamente un modelo de difusión gaussiano.
Despejando de v = E[x₁ − x₀ | x_t] obtenemos: la predicción de ruido (lo que aprende DDPM), la predicción del dato limpio (un denoiser) y el score ∇log p_t (lo que aprende score matching).
Son transformaciones afines unas de otras: una red que predice v sirve para todo. La figura muestra el denoiser x̂₁ = E[x₁|x_t] calculado con nuestro modelo de velocidad: con t pequeño apunta a la media (borroso), con t cercano a 1 apunta a los datos.
""")

# 17 — ODE vs SDE ------------------------------------------------------------------------
s = content("03 · Conexión con difusión", "ODE o SDE: mismas marginales, distinto muestreador")
rect(s, 0.6, 1.65, 12.15, 1.0, fill=TINT)
eq(s, "sde", 0.6, 1.8, scale=0.8, center_w=12.15)
picture(s, f"{FIG}/nb1_ode_vs_sde.png", 0.6, 2.85, box=(12.15, 2.75))
rect(s, 0.6, 5.75, 12.15, 0.8, fill=NAVY)
text(s, 0.9, 5.88, 11.6, 0.6,
     "Lo que realmente distingue los métodos: **camino (α_t, σ_t)** · **parametrización** · **ponderación en t** · **muestreador**",
     size=17, color=WHITE, hl=GOLD)
notes(s, """
[P2 · 24:00–26:00]
Usando el score derivado de la velocidad podemos muestrear con una SDE (estilo difusión). Para cualquier g_t ≥ 0, esta SDE tiene las mismas marginales p_t que la ODE: el término de deriva extra compensa exactamente la difusión (Fokker-Planck).
Lo comprobamos en el notebook: ODE y SDE generan la misma distribución.
Conclusión de esta sección: 'difusión vs flow matching' no es una dicotomía. Son el mismo objeto; las decisiones de diseño reales son las cuatro del recuadro. Flow Matching hace explícita la elección del camino, y el camino lineal resulta ser especialmente bueno… y se puede mejorar todavía.
""")

# 18 — Divider RF -------------------------------------------------------------------------
s = divider("04", "Rectified Flow", "Enderezar las trayectorias para muestrear en uno o dos pasos")
notes(s, "[P2 · 26:00] 'Hasta ahora: entrenar es fácil. Ahora: muestrear rápido.'")

# 19 — Curvas = muchos pasos ----------------------------------------------------------------
s = content("04 · Rectified Flow", "Trayectorias curvas obligan a dar muchos pasos")
picture(s, f"{FIG}/nb1_nfe_grid.png", 0.6, 1.6, box=(12.15, 2.75))
rect(s, 0.6, 4.6, 5.95, 1.9, fill=TINT)
text(s, 0.85, 4.75, 5.5, 0.35, "Un paso de Euler", size=15, color=TEAL, bold=True)
eq(s, "euler", 0.6, 5.25, scale=0.8, center_w=5.95)
rect(s, 6.8, 4.6, 5.95, 1.9, fill=TINT2)
text(s, 7.05, 4.75, 5.5, 1.7,
     ["Con 1 paso: x₁ ≈ x₀ + v(x₀, 0) = x₀ + 𝔼[x₁ − x₀ | x₀] → **la media**: una mancha borrosa.",
      "Si las trayectorias fueran **rectas**, Euler sería **exacto con un solo paso**."], size=15, hl=CRIMSON, spacing=6)
notes(s, """
[P2 · 26:15–27:30]
Mismo modelo del notebook 1, muestreado con 1, 2, 4, 8 y 64 pasos de Euler. Con 1 paso sale una mancha: el primer paso apunta a la media condicional de los datos, que está en el centro.
La razón es geométrica: un paso de Euler sigue la tangente. Si la trayectoria es curva, el error es grande; si es recta, la tangente lleva exactamente al destino.
Pregunta natural: ¿podemos conseguir un flujo con trayectorias rectas?
""")

# 20 — Por qué curvas ------------------------------------------------------------------------
s = content("04 · Rectified Flow", "¿Por qué son curvas? Porque los cruces se promedian")
picture(s, f"{FIG}/nb2_toy_reflow.png", 0.6, 1.6, box=(12.15, 4.1))
caps = [("Acoplamiento independiente", "Las rectas x₀ → x₁ se cruzan en el centro"),
        ("1-Rectified Flow", "La red promedia los cruces: el flujo los «desenreda», pero se curva"),
        ("2-Rectified Flow", "Reentrenado con sus propios pares: ya no hay cruces → rectas")]
for i, (a_, b_) in enumerate(caps):
    x = 0.6 + i * 4.1
    text(s, x + 0.1, 5.8, 3.8, 0.35, a_, size=15, color=[MUTED, TEAL, CRIMSON][i], bold=True)
    text(s, x + 0.1, 6.15, 3.8, 0.6, b_, size=13, color=INK)
notes(s, """
[P2 · 27:30–29:30]
Ejemplo canónico del paper de Rectified Flow, reproducido en el notebook 2. Ruido: dos gaussianas a la izquierda; datos: dos a la derecha.
Izquierda: con pares independientes, la mitad de las rectas son diagonales y se cruzan. En el cruce, la red sólo puede aprender el promedio de velocidades (apunta a la derecha).
Centro: el flujo resultante (1-Rectified Flow, que es exactamente CFM) no cruza trayectorias: los puntos de arriba van arriba y los de abajo, abajo. Ha construido un nuevo emparejamiento ruido–dato, determinista y sin cruces… pero curvo.
Derecha: si reentrenamos con esos nuevos pares, las rectas condicionales ya no se cruzan, el promedio no deforma nada, y las trayectorias son rectas. Medimos la rectitud S: de 5.2 a 0.0002.
""")

# 21 — Reflow algoritmo -----------------------------------------------------------------------
s = content("04 · Rectified Flow", "Reflow: reentrenar con los pares que genera el propio flujo")
steps = [("Entrenar 1-RF", "CFM con pares (x₀, x₁) independientes"),
         ("Generar pares", "Z₀ ~ p₀,  Z₁ = ODE(Z₀) con un solver preciso"),
         ("Entrenar 2-RF", "Misma pérdida con los pares (Z₀, Z₁). Repetir → k-RF")]
for i, (a_, b_) in enumerate(steps):
    x = 0.6 + i * 4.2
    rect(s, x, 1.7, 3.75, 1.65, fill=TINT)
    circle_num(s, x + 0.2, 1.88, 0.5, i + 1, size=15)
    text(s, x + 0.85, 1.93, 2.8, 0.4, a_, size=17, color=INK, bold=True, font=HEAD)
    text(s, x + 0.2, 2.55, 3.4, 0.75, b_, size=14, color=INK)
    if i < 2: arrow(s, x + 3.8, 2.33)
props = [("Conserva las marginales", None, "Los pares (Z₀, Z₁) siguen siendo un acoplamiento de p₀ y q"),
         ("No aumenta el coste de transporte", "cost", None),
         ("Rectitud (S = 0 ⇔ rectas)", "straight", None),
         ("Endereza progresivamente", "bound", None)]
y = 3.65
for i, (a_, e_, d_) in enumerate(props):
    x = 0.6 + (i % 2) * 6.2; yy = y + (i // 2) * 1.5
    rect(s, x, yy, 5.95, 1.35, fill="F4F6F8")
    text(s, x + 0.2, yy + 0.1, 5.5, 0.35, a_, size=15, color=TEAL, bold=True)
    if e_: eq(s, e_, x, yy + 0.5, scale=0.6, center_w=5.95)
    if d_: text(s, x + 0.2, yy + 0.55, 5.5, 0.7, d_, size=14)
source(s, "Liu, Gong & Liu (2023). Flow Straight and Fast: Learning to Generate and Transfer Data with Rectified Flow. ICLR. Teoremas 3.3, 3.5 y 3.7.")
notes(s, """
[P2 · 29:30–31:30]
Algoritmo de reflow: (1) entrenar un flujo con CFM; (2) generar un dataset de pares simulando la ODE del flujo entrenado; (3) reentrenar con esos pares. Se puede repetir.
Garantías teóricas del paper:
- Cada reflow conserva las marginales: seguimos generando la misma distribución.
- El coste de transporte E[c(Z₁ − Z₀)] no aumenta para NINGÚN coste convexo c: el nuevo emparejamiento es 'más óptimo'.
- La rectitud S mide cuánto se desvía la velocidad de la recta entre extremos; tras K reflows el mínimo de S decae como 1/K.
Precio: generar los pares exige simular la ODE (caro) y los errores del primer modelo se heredan.
""")

# 22 — OT-CFM --------------------------------------------------------------------------------------
s = content("04 · Rectified Flow", "Alternativa: transporte óptimo en el minibatch")
rect(s, 0.6, 1.65, 12.15, 1.0, fill=TINT)
eq(s, "otcfm", 0.6, 1.8, scale=0.75, center_w=12.15)
picture(s, f"{FIG}/nb2_trajectories.png", 0.6, 2.85, box=(12.15, 3.0))
bullets(s, 0.6, 5.95, 12.1, 0.7, [
    "Dentro de cada minibatch se reordenan los pares con el algoritmo húngaro: **sin fase extra** de simulación, coste O(B³) por batch",
    "Endereza mucho, pero en alta dimensión el OT de un minibatch se aleja del OT global",
], size=14, gap=3)
notes(s, """
[P2 · 31:30–33:00]
Otra forma de conseguir pares que no se crucen: en vez de independientes, emparejar ruido y datos del minibatch resolviendo un problema de transporte óptimo (asignación con coste cuadrático). Propuesto simultáneamente por Tong et al. (2023, OT-CFM) y Pooladian et al. (2023, multisample FM).
Figura del notebook 2 con 8 gaussianas: 1-RF (CFM) curvo, S = 2.2; reflow casi recto; OT-CFM bastante recto sin necesitar una segunda fase.
Limitación: con batches pequeños en dimensión alta, el OT del minibatch es una aproximación pobre del OT global.
""")

# 23 — Nuestro experimento -----------------------------------------------------------------------------
NFES = ["1", "2", "4", "8", "16", "32", "64"]
RES = {  # Sliced-Wasserstein (↓) — salida del Notebook 2, sección 3 (8 gaussianas)
    "1-RF (CFM)": [2.0537, 0.5595, 0.1550, 0.1017, 0.0885, 0.0837, 0.0822],
    "OT-CFM": [0.1774, 0.1318, 0.1212, 0.1115, 0.1066, 0.1045, 0.1032],
    "2-RF (reflow)": [0.1168, 0.1138, 0.0997, 0.0911, 0.0862, 0.0838, 0.0831],
}
DISTILL_1 = 0.0839
s = content("04 · Rectified Flow", "Nuestro experimento: calidad frente a número de pasos")
cd = CategoryChartData(); cd.categories = NFES
for k, v in RES.items(): cd.add_series(k, v)
cd.add_series("2-RF + destilación (1 paso)", [DISTILL_1] + [None] * (len(NFES) - 1))
g = s.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, Inches(0.6), Inches(1.6), Inches(6.9), Inches(4.95), cd)
ch = g.chart; style_chart(ch, 12)
ch.has_legend = True; ch.legend.position = XL_LEGEND_POSITION.TOP; ch.legend.include_in_layout = False; ch.legend.font.size = Pt(11)
for ser, c in zip(ch.plots[0].series, [TEAL, GOLD, CRIMSON, BLUE]):
    ser.smooth = False
    ser.format.line.color.rgb = rgb(c); ser.format.line.width = Pt(2.5)
    ser.marker.style = XL_MARKER_STYLE.CIRCLE; ser.marker.size = 8
    ser.marker.format.fill.solid(); ser.marker.format.fill.fore_color.rgb = rgb(c); ser.marker.format.line.color.rgb = rgb(c)
last = ch.plots[0].series[3]; last.marker.style = XL_MARKER_STYLE.DIAMOND; last.marker.size = 14
va = ch.value_axis; log_axis(va); gridlines(va); va.format.line.fill.background()
va.crosses = XL_AXIS_CROSSES.MINIMUM
va.has_title = True; va.axis_title.text_frame.text = "Sliced Wasserstein (↓ mejor, escala log)"
va.axis_title.text_frame.paragraphs[0].runs[0].font.size = Pt(12); va.axis_title.text_frame.paragraphs[0].runs[0].font.bold = False
va.tick_labels.font.size = Pt(11); va.tick_labels.number_format = "0.0#"; va.tick_labels.number_format_is_linked = False
ca = ch.category_axis; ca.has_title = True; ca.axis_title.text_frame.text = "NFE (pasos de Euler)"
ca.axis_title.text_frame.paragraphs[0].runs[0].font.size = Pt(12); ca.axis_title.text_frame.paragraphs[0].runs[0].font.bold = False
ca.tick_labels.font.size = Pt(11); ca.format.line.color.rgb = rgb("C9D1D9")
panels = [crop(f"{FIG}/nb2_few_step_grid.png", (0.024 + 0.245 * i, 0.0, 0.256 + 0.245 * i, 0.345), f"onestep_{i}") for i in range(4)]
ims = [Image.open(q) for q in panels]; pw_, ph_ = min(i.size[0] for i in ims), min(i.size[1] for i in ims)
comp = Image.new("RGB", (2 * pw_ + 20, 2 * ph_ + 20), "white")
for i, im in enumerate(ims):
    comp.paste(im.crop((0, 0, pw_, ph_)), ((i % 2) * (pw_ + 20), (i // 2) * (ph_ + 20)))
comp.save(f"{AS}/crops/onestep_2x2.png")
text(s, 7.85, 1.65, 4.9, 0.35, "Muestras con **un solo paso**", size=15, color=INK, hl=INK)
picture(s, f"{AS}/crops/onestep_2x2.png", 7.85, 2.05, box=(4.9, 2.75), align="top")
bullets(s, 7.85, 5.0, 4.9, 1.6, [
    "1 paso: 1-RF colapsa en la media; **2-RF ya genera los 8 modos**",
    "2-RF con **2 pasos** ≈ 1-RF con 8; destilado, **1 paso** ≈ 32",
    "Con muchos pasos convergen, salvo **OT-CFM**, que se estanca algo por encima",
], size=14, gap=5)
source(s, "Notebook 2 · datos: 8 gaussianas · MLP 3×256 · métrica: sliced Wasserstein con 5 000 muestras.")
notes(s, """
[P2 · 33:00–34:30]
Resultado cuantitativo del notebook 2. Eje x: número de pasos de Euler; eje y: distancia (sliced Wasserstein) entre generados y datos, en escala logarítmica.
1-RF (CFM estándar) necesita ~8 pasos para ser razonable. Reflow (2-RF) es bueno desde 1-2 pasos. OT-CFM queda en medio en pocos pasos.
La estrella: destilación del 2-RF, entrenando sólo en t = 0 para que un paso aterrice en el destino: 1 paso ≈ 32 pasos del modelo original.
Con muchos pasos todos llegan a calidad similar: reflow no mejora el modelo, mejora su velocidad de muestreo.
OT-CFM se estanca algo por encima con muchos pasos. Explicación plausible: el emparejamiento OT de cada minibatch reparte el ruido entre los 8 modos con fronteras bruscas que la red suaviza (deja masa entre modos); además el OT del minibatch sólo aproxima el OT global. Con batch 512 no mejoró.
""")

# 24 — Resultados paper RF + InstaFlow ------------------------------------------------------------------
s = content("04 · Rectified Flow", "En imágenes: de ruido a FID 4.85 con un solo paso")
cd = CategoryChartData()
cd.categories = ["VP ODE (difusión)", "1-RF", "2-RF", "3-RF", "2-RF + destilación"]
cd.add_series("FID", (451, 378, 12.21, 8.15, 4.85))
g = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(0.6), Inches(1.6), Inches(7.2), Inches(4.6), cd)
ch = g.chart; style_chart(ch, 13); ch.has_legend = False
ch.has_title = True; ch.chart_title.text_frame.text = "CIFAR-10 · FID con 1 paso de Euler (↓ mejor, escala log)"
ch.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(15); ch.chart_title.text_frame.paragraphs[0].runs[0].font.color.rgb = rgb(INK)
pl = ch.plots[0]; pl.gap_width = 60; pl.has_data_labels = True
pl.data_labels.number_format = "General"; pl.data_labels.number_format_is_linked = False
pl.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END; pl.data_labels.font.size = Pt(13); pl.data_labels.font.bold = True
for i, c in enumerate([MUTED, TEAL, CRIMSON, CRIMSON, BLUE]):
    pt = pl.series[0].points[i]; pt.format.fill.solid(); pt.format.fill.fore_color.rgb = rgb(c)
va = ch.value_axis; log_axis(va); va.visible = False; va.has_major_gridlines = False
va.maximum_scale = 3000; va.minimum_scale = 1
ca = ch.category_axis; ca.reverse_order = True; ca.tick_labels.font.size = Pt(13); ca.format.line.color.rgb = rgb("C9D1D9")
stat(s, 8.15, 1.6, 4.6, 2.15, "FID 2.58", "1-RF con solver adaptativo (127 NFE): competitivo con difusión de la época", color=TEAL, big_size=34)
rect(s, 8.15, 3.95, 4.6, 2.25, fill=NAVY)
text(s, 8.4, 4.1, 4.1, 0.4, "InstaFlow (ICLR 2024)", size=16, color=GOLD, bold=True)
text(s, 8.4, 4.55, 4.1, 1.6, ["Reflow + destilación aplicados a **Stable Diffusion**: texto → imagen en **un paso**.",
                              "FID 23.3 en MS COCO 2017-5k · 0.09 s por imagen"], size=14, color=WHITE, hl=GOLD, spacing=6)
source(s, "Liu, Gong & Liu (2023), Tabla 1 · Liu et al. (2024) InstaFlow: One Step is Enough for High-Quality Diffusion-Based Text-to-Image Generation.")
notes(s, """
[P2 · 34:30–36:00]
Resultados del paper de Rectified Flow en CIFAR-10 con UN paso. La difusión (VP ODE) y el 1-RF dan ruido (FID ~400). Tras un reflow, el 2-RF baja a 12.2 sin destilar; con destilación, 4.85: en su momento el mejor resultado de un paso con U-Net (TDPM tenía 8.91).
Con muchos pasos, 1-RF llega a FID 2.58.
InstaFlow (mismos autores) aplica la receta a Stable Diffusion: genera imágenes 512×512 en un paso, en 0.09 s.
Ahora pasamos a la demo con los notebooks.
""")

# 25 — Demo ---------------------------------------------------------------------------------------------
s = content("Demo", "Demo en vivo: los notebooks")
for i, (n_, t_, items_) in enumerate([
    ("01", "flow_matching_2d.ipynb", ["CFM en ~30 líneas de PyTorch", "Animación del transporte y campo de velocidades",
                                      "Calidad vs. NFE (Euler / punto medio)", "Velocidad → denoiser y score; muestreo SDE"]),
    ("02", "rectified_flow_reflow.ipynb", ["Dos gaussianas → dos gaussianas: cruces vs. reflow", "1-RF vs. 2-RF vs. OT-CFM: rectitud S",
                                           "Generación en 1–4 pasos y destilación", "Ejercicios: 3-RF, tamaño de batch OT, coste"])]):
    x = 0.6 + i * 6.2
    rect(s, x, 1.65, 5.95, 3.9, fill=TINT if i == 0 else TINT2)
    circle_num(s, x + 0.25, 1.9, 0.7, n_, fill=TEAL if i == 0 else CRIMSON, size=18)
    text(s, x + 1.15, 2.05, 4.6, 0.45, t_, size=17, color=INK, bold=True, font=CODE)
    bullets(s, x + 0.3, 2.85, 5.4, 2.6, items_, size=15, gap=8, bullet_color=TEAL if i == 0 else CRIMSON)
rect(s, 0.6, 5.8, 12.15, 0.75, fill="F4F6F8")
text(s, 0.85, 5.95, 11.7, 0.5, "Se ejecutan en Colab o en local, **en CPU**: ~1 min el notebook 1 y ~3 min el 2. Todas las figuras de datos 2D de estas diapositivas salen de aquí.",
     size=15)
notes(s, """
[P2 · 36:00–41:00 · DEMO ~5 min]
Guion de la demo (tener los notebooks ya ejecutados por si falla algo):
1) Notebook 1: enseñar la celda de entrenamiento (6 líneas), la animación GIF del flujo y la figura calidad-vs-NFE. Cambiar en directo el dataset a 'checkerboard' si sobra tiempo.
2) Notebook 2: ejecutar la celda de las dos gaussianas (20 s) y comentar 1-RF vs 2-RF; enseñar la tabla final de sliced Wasserstein y la rejilla de 1/2/4 pasos.
Si no hay conexión: usar las figuras ya guardadas en notebooks/figuras.
""")

# 26 — SD3 ------------------------------------------------------------------------------------------------
s = content("05 · Estado del arte", "Stable Diffusion 3: Rectified Flow a escala")
text(s, 0.6, 1.6, 6.3, 0.8, "Esser et al. comparan **61 formulaciones** (difusión y flujos) con el mismo Transformer (MM-DiT) en ImageNet y CC12M.",
     size=16)
tbl_rows = [("Formulación", "Rango medio ↓"),
            ("RF + t logit-normal (0, 1)", "1.54"),
            ("RF + t logit-normal (1, 0.6)", "2.08"),
            ("ε-pred. lineal (LDM, SD 1.x/2.x)", "2.88"),
            ("RF + t uniforme", "5.67"),
            ("v-pred. coseno", "9.17")]
tshape = s.shapes.add_table(len(tbl_rows), 2, Inches(0.6), Inches(2.55), Inches(6.3), Inches(2.7))
tb = tshape.table; tb.columns[0].width = Inches(4.5); tb.columns[1].width = Inches(1.8)
for r, (a_, b_) in enumerate(tbl_rows):
    for c, v_ in enumerate((a_, b_)):
        cell = tb.cell(r, c); cell.text = ""
        cell.fill.solid(); cell.fill.fore_color.rgb = rgb(TEAL if r == 0 else (TINT if r in (1, 2) else WHITE))
        p = cell.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER
        _runs(p, v_, 14, WHITE if r == 0 else INK, bold=(r == 0 or r == 1))
        cell.margin_left = Inches(0.12); cell.vertical_anchor = MSO_ANCHOR.MIDDLE
text(s, 0.6, 5.35, 6.3, 0.5, "Rango medio sobre 24 configuraciones (dataset, muestreador, nº de pasos). Tabla 1 del paper, selección de filas.",
     size=11, color=MUTED, italic=True)
rect(s, 7.2, 1.6, 5.55, 1.9, fill=TINT)
text(s, 7.45, 1.72, 5.1, 0.35, "Muestreo de t logit-normal", size=15, color=TEAL, bold=True)
eq(s, "lognorm", 7.2, 2.15, scale=0.66, center_w=5.55)
text(s, 7.45, 2.95, 5.1, 0.5, "Más peso a los tiempos intermedios, donde la tarea es más difícil.", size=13, color=INK)
stat(s, 7.2, 3.75, 2.65, 2.1, "8B", "parámetros: escalan de 0.8B a 8B (MM-DiT)", color=CRIMSON, fill=TINT2, big_size=36)
stat(s, 10.1, 3.75, 2.65, 2.1, "0.74", "GenEval (8B, 1024², DPO) vs. 0.67 de DALL·E 3", color=CRIMSON, fill=TINT2, big_size=36)
source(s, "Esser et al. (2024). Scaling Rectified Flow Transformers for High-Resolution Image Synthesis. ICML. Flux (Black Forest Labs) sigue la misma receta.")
notes(s, """
[P2 · 41:00–42:30]
¿Se usa esto de verdad? Sí: Stable Diffusion 3 es un 'rectified flow transformer'. Su objetivo de entrenamiento es exactamente el de Rectified Flow / CFM con camino lineal (sin reflow).
Hicieron un estudio grande: 61 formulaciones (distintos caminos, parametrizaciones y distribuciones de t). Gana RF con t muestreado de una logit-normal, que concentra el entrenamiento en tiempos intermedios. Las formulaciones clásicas (ε lineal de SD 1.x/2.x, v-pred coseno) quedan por detrás, sobre todo con pocos pasos.
Escalan hasta 8B parámetros y superan a DALL·E 3 en GenEval (0.74 vs 0.67) y en preferencia humana. Flux, de los autores originales de Stable Diffusion, usa la misma receta.
""")

# 27 — Más allá ------------------------------------------------------------------------------------------------
s = content("05 · Estado del arte", "Hacia adelante: un solo paso y nuevos dominios")
rect(s, 0.6, 1.6, 6.0, 4.95, fill=TINT)
text(s, 0.85, 1.75, 5.5, 0.4, "Generar en un paso, sin destilar", size=17, color=TEAL, bold=True)
bullets(s, 0.85, 2.3, 5.5, 2.0, [
    "**Consistency models** (Song et al., 2023): mapear cualquier x_t directamente a x₁",
    "**Shortcut models** (Frans et al., 2025): condicionar también en el tamaño de paso",
    "**MeanFlow** (Geng et al., 2025): aprender la velocidad **media** en un intervalo",
], size=14, gap=6)
eq(s, "meanflow", 0.6, 4.35, scale=0.6, center_w=6.0)
text(s, 0.85, 5.4, 5.5, 1.0, "MeanFlow: **FID 3.43 con 1 NFE** en ImageNet 256×256, entrenado desde cero", size=14, color=INK)
rect(s, 6.85, 1.6, 5.9, 4.95, fill=TINT2)
text(s, 7.1, 1.75, 5.4, 0.4, "Flow Matching más allá de las imágenes", size=17, color=CRIMSON, bold=True)
bullets(s, 7.1, 2.3, 5.4, 4.2, [
    "**Vídeo**: Movie Gen (Meta, 2024) se entrena con Flow Matching",
    "**Audio y voz**: Voicebox (Le et al., 2023)",
    "**Geometría**: Riemannian FM (Chen & Lipman, 2024) — proteínas, moléculas, datos en variedades",
    "**Datos discretos**: Discrete FM (Gat et al., 2024) — texto y código",
    "**Biblioteca de referencia**: flow\\_matching (Meta) y la guía Lipman et al. (2024)",
], size=14, gap=7, bullet_color=CRIMSON)
notes(s, """
[P2 · 42:30–43:30]
Dos líneas activas. Primera: generar en un paso SIN la fase extra de reflow/destilación. Consistency models, shortcut models y, muy reciente, MeanFlow, que en lugar de la velocidad instantánea aprende la velocidad media entre dos instantes r y t; con ella, un único paso salta de r = 0 a t = 1. Consigue FID 3.43 con un paso en ImageNet 256 entrenando desde cero.
Segunda: el marco es tan general que se ha llevado a vídeo, audio, variedades (proteínas) y datos discretos. Si queréis profundizar: la guía 'Flow Matching Guide and Code' de Lipman et al. (2024) y su librería.
""")

# 28 — Conclusiones -----------------------------------------------------------------------------------------------
s = prs.slides.add_slide(BLANK); bg(s, NAVY)
text(s, 0.8, 0.6, 11, 0.4, "CONCLUSIONES", size=13, color=GOLD, bold=True)
text(s, 0.8, 1.0, 11.5, 0.9, "Tres ideas para llevarse", size=36, color=WHITE, bold=True, font=HEAD)
take = [("Flow Matching", "Entrenar un flujo continuo es una **regresión**: interpolar ruido y dato y predecir x₁ − x₀. Sin simular la ODE."),
        ("Difusión ⊂ Flow Matching", "Velocidad, ruido, dato y score son la misma información. Lo que importa es el **camino**, la **ponderación en t** y el **muestreador**."),
        ("Rectas = rápidas", "Reflow y OT **enderezan** trayectorias → 1–2 pasos. El camino lineal es hoy el estándar (SD3, Flux).")]
for i, (a_, b_) in enumerate(take):
    x = 0.8 + i * 4.0
    rect(s, x, 2.3, 3.75, 3.6, fill="1E2F52")
    circle_num(s, x + 0.3, 2.6, 0.75, i + 1, fill=CRIMSON, size=22)
    text(s, x + 0.3, 3.55, 3.2, 0.5, a_, size=19, color=WHITE, bold=True, font=HEAD)
    text(s, x + 0.3, 4.15, 3.2, 1.7, b_, size=15, color="DCE3EA", hl=GOLD)
text(s, 0.8, 6.25, 11.7, 0.4, "Gracias · ¿Preguntas?", size=22, color=GOLD, bold=True, font=HEAD)
footer(s, dark=True)
notes(s, """
[P2 · 43:30–44:30]
Resumen en tres ideas. (1) Flow Matching convierte el entrenamiento de flujos continuos en una regresión simple. (2) La difusión gaussiana es un caso particular: las parametrizaciones son equivalentes y las verdaderas decisiones son camino, ponderación y muestreador. (3) Trayectorias rectas permiten muestrear en muy pocos pasos: reflow, OT, destilación… y el camino lineal es lo que usan hoy SD3 y Flux.
Preguntas.
""")

# 29 — Referencias ----------------------------------------------------------------------------------------------------
s = content("Referencias", "Referencias principales")
refs = [
    "Chen, Rubanova, Bettencourt & Duvenaud (2018). Neural Ordinary Differential Equations. NeurIPS.",
    "Grathwohl et al. (2019). FFJORD: Free-form Continuous Dynamics for Scalable Reversible Generative Models. ICLR.",
    "Ho, Jain & Abbeel (2020). Denoising Diffusion Probabilistic Models. NeurIPS.",
    "**Lipman, Chen, Ben-Hamu, Nickel & Le (2023). Flow Matching for Generative Modeling. ICLR.**",
    "**Liu, Gong & Liu (2023). Flow Straight and Fast: Learning to Generate and Transfer Data with Rectified Flow. ICLR.**",
    "Albergo & Vanden-Eijnden (2023). Building Normalizing Flows with Stochastic Interpolants. ICLR.",
    "Tong et al. (2024). Improving and Generalizing Flow-Based Generative Models with Minibatch Optimal Transport. TMLR.",
    "Liu et al. (2024). InstaFlow: One Step is Enough for High-Quality Diffusion-Based Text-to-Image Generation. ICLR.",
    "**Esser et al. (2024). Scaling Rectified Flow Transformers for High-Resolution Image Synthesis. ICML.**",
    "Lipman et al. (2024). Flow Matching Guide and Code. arXiv:2412.06264.",
    "Geng, Deng, Bai, Kolter & He (2025). Mean Flows for One-step Generative Modeling. NeurIPS.",
]
bullets(s, 0.6, 1.6, 12.1, 5.0, refs, size=14, gap=5, hl=INK)
notes(s, "Lectura recomendada para empezar: Lipman et al. (2024), 'Flow Matching Guide and Code', y el blog 'Diffusion Meets Flow Matching' (Gao et al., 2024).")

out = "FlowMatching_RectifiedFlows.pptx"
prs.save(out)
print("guardado", out, "·", len(prs.slides), "diapositivas")
