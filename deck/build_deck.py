#!/usr/bin/env python3
"""
Build the n_TOF EAR2 X17 February & May 2026 beam-test status deck (PowerPoint).
Focus: technical tests to see whether the high-energy measurement is possible;
the gamma flash and DREAM-DAQ saturation; what we did, measured and understood;
why none of the mitigation ideas worked out.  Extensive backup section.
"""
import os
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, "assets")

# ---- palette ----------------------------------------------------------------
NAVY   = RGBColor(0x14, 0x2C, 0x52)
BLUE   = RGBColor(0x2E, 0x6D, 0xB4)
ORANGE = RGBColor(0xE0, 0x78, 0x20)
RED    = RGBColor(0xC0, 0x39, 0x2B)
GREEN  = RGBColor(0x2E, 0x8B, 0x57)
GREY   = RGBColor(0x55, 0x55, 0x55)
LGREY  = RGBColor(0xEE, 0xEE, 0xEE)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
DARK   = RGBColor(0x22, 0x22, 0x22)

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]


def _set(tf_para, text, size, color=DARK, bold=False, italic=False, align=None, font="Calibri"):
    tf_para.text = text
    r = tf_para.runs[0]
    r.font.size = Pt(size); r.font.bold = bold; r.font.italic = italic
    r.font.color.rgb = color; r.font.name = font
    if align is not None:
        tf_para.alignment = align


def add_slide(bg=WHITE):
    s = prs.slides.add_slide(BLANK)
    r = s.shapes.add_shape(1, 0, 0, SW, SH)
    r.fill.solid(); r.fill.fore_color.rgb = bg; r.line.fill.background()
    r.shadow.inherit = False
    s.shapes._spTree.remove(r._element); s.shapes._spTree.insert(2, r._element)
    return s


def add_band(s, color=NAVY, h=Inches(0.16), y=0):
    b = s.shapes.add_shape(1, 0, y, SW, h)
    b.fill.solid(); b.fill.fore_color.rgb = color; b.line.fill.background()
    b.shadow.inherit = False
    return b


def title_box(s, title, kicker=None):
    add_band(s, NAVY, Inches(0.14), 0)
    if kicker:
        k = s.shapes.add_textbox(Inches(0.55), Inches(0.20), Inches(12), Inches(0.35))
        _set(k.text_frame.paragraphs[0], kicker.upper(), 12.5, ORANGE, bold=True)
    ty = Inches(0.52) if kicker else Inches(0.30)
    t = s.shapes.add_textbox(Inches(0.55), ty, Inches(12.3), Inches(0.9))
    tf = t.text_frame; tf.word_wrap = True
    _set(tf.paragraphs[0], title, 27, NAVY, bold=True)
    return t


def footer(s, n):
    f = s.shapes.add_textbox(Inches(11.6), Inches(7.06), Inches(1.6), Inches(0.32))
    _set(f.text_frame.paragraphs[0], f"n_TOF X17  ·  {n}", 9, GREY, align=PP_ALIGN.RIGHT)


def bullets(s, items, x, y, w, h, size=15, gap=6):
    tb = s.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame; tf.word_wrap = True
    first = True
    for it in items:
        lvl = 0; txt = it; color = DARK; bold = False
        if isinstance(it, tuple):
            txt, lvl = it[0], it[1]
            if len(it) > 2: color = it[2]
            if len(it) > 3: bold = it[3]
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        bullet = ("•  " if lvl == 0 else "–  ") if txt else ""
        _set(p, bullet + txt, size - lvl*1.0, color, bold=bold)
        p.level = lvl
        p.space_after = Pt(gap)
    return tb


def fit_image(s, path, x, y, w, h, caption=None, frame=True):
    """Place image fit within box (w,h) preserving aspect, centred."""
    iw, ih = Image.open(path).size
    ar = iw / ih
    bw, bh = w, h
    if bw / bh > ar:
        nh = bh; nw = int(bh * ar)
    else:
        nw = bw; nh = int(bw / ar)
    nx = x + (bw - nw) // 2
    ny = y + (bh - nh) // 2
    if frame:
        bg = s.shapes.add_shape(1, nx-Emu(9525), ny-Emu(9525), nw+Emu(19050), nh+Emu(19050))
        bg.fill.solid(); bg.fill.fore_color.rgb = WHITE
        bg.line.color.rgb = RGBColor(0xCC, 0xCC, 0xCC); bg.line.width = Pt(0.75)
        bg.shadow.inherit = False
    pic = s.shapes.add_picture(path, nx, ny, nw, nh)
    if caption:
        c = s.shapes.add_textbox(x, y + h, w, Inches(0.3))
        _set(c.text_frame.paragraphs[0], caption, 10.5, GREY, italic=True, align=PP_ALIGN.CENTER)
    return pic


def A(name):  # asset path
    return os.path.join(ASSETS, name)


def callout(s, text, x, y, w, h, fill=RGBColor(0xFB, 0xF3, 0xD8), line=ORANGE, size=14, bold=False, color=DARK):
    box = s.shapes.add_shape(5, x, y, w, h)  # rounded rect
    box.fill.solid(); box.fill.fore_color.rgb = fill
    box.line.color.rgb = line; box.line.width = Pt(1.5)
    box.shadow.inherit = False
    tf = box.text_frame; tf.word_wrap = True
    tf.margin_left = Inches(0.15); tf.margin_right = Inches(0.15)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    parts = text.split("\n")
    for i, pt in enumerate(parts):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        _set(p, pt, size, color, bold=bold)
    return box


def section_divider(title, subtitle=None, n=""):
    s = add_slide(NAVY)
    add_band(s, ORANGE, Inches(0.16), Inches(3.0))
    t = s.shapes.add_textbox(Inches(0.8), Inches(2.0), Inches(11.7), Inches(1.0))
    _set(t.text_frame.paragraphs[0], title, 36, WHITE, bold=True)
    if subtitle:
        st = s.shapes.add_textbox(Inches(0.8), Inches(3.25), Inches(11.7), Inches(0.8))
        _set(st.text_frame.paragraphs[0], subtitle, 17, RGBColor(0xCF, 0xDB, 0xEC))
    return s


def table(s, data, x, y, w, h, header=True, fs=11, col_w=None):
    rows, cols = len(data), len(data[0])
    gt = s.shapes.add_table(rows, cols, x, y, w, h).table
    if col_w:
        for j, cw in enumerate(col_w):
            gt.columns[j].width = cw
    for i, row in enumerate(data):
        for j, val in enumerate(row):
            cell = gt.cell(i, j)
            cell.margin_left = Inches(0.06); cell.margin_right = Inches(0.06)
            cell.margin_top = Inches(0.02); cell.margin_bottom = Inches(0.02)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            cell.text_frame.word_wrap = True
            if i == 0 and header:
                _set(p, str(val), fs, WHITE, bold=True)
                cell.fill.solid(); cell.fill.fore_color.rgb = NAVY
            else:
                _set(p, str(val), fs, DARK)
                cell.fill.solid()
                cell.fill.fore_color.rgb = WHITE if (i % 2) else LGREY
    return gt


# =============================================================================
#  TITLE
# =============================================================================
s = add_slide(NAVY)
add_band(s, ORANGE, Inches(0.18), Inches(2.55))
t = s.shapes.add_textbox(Inches(0.8), Inches(1.15), Inches(11.7), Inches(1.4))
_set(t.text_frame.paragraphs[0], "Searching for the X17 at n_TOF EAR2", 40, WHITE, bold=True)
sub = s.shapes.add_textbox(Inches(0.8), Inches(2.75), Inches(11.7), Inches(1.1))
tf = sub.text_frame; tf.word_wrap = True
_set(tf.paragraphs[0], "February & May 2026 Beam Tests — Status, Results & What We Learned", 21, RGBColor(0xCF,0xDB,0xEC), bold=False)
p2 = tf.add_paragraph()
_set(p2, "Technical tests of the high-energy measurement against the gamma flash", 15, RGBColor(0xAEC3DC,0xAE if False else 0xC3, 0xDC) if False else RGBColor(0xAE,0xC3,0xDC), italic=True)
meta = s.shapes.add_textbox(Inches(0.8), Inches(6.3), Inches(11.7), Inches(0.8))
_set(meta.text_frame.paragraphs[0], "n_TOF Collaboration  ·  EAR2, CERN  ·  Internal status — June 2026", 13, RGBColor(0x9A,0xAE,0xC8))

# =============================================================================
#  1. CONTEXT / MOTIVATION
# =============================================================================
s = add_slide(); title_box(s, "The measurement and its central challenge", kicker="Context")
bullets(s, [
    ("Goal: search for the hypothetical X17 boson via the e⁺e⁻ pairs it would emit,", 0),
    ("looking for a peak in the pair opening angle near 120°.", 1),
    ("EAR2 sits 20 m directly above the lead spallation target. Each PS proton bunch", 0),
    ("on the target produces a broad neutron spectrum — ~80% of the expected X17 yield", 1),
    ("is in the first ~10 µs (the ~1 MeV neutron region).", 1),
    ("The catch — the “gamma flash”: every relativistic particle from the bunch hitting", 0, RED, True),
    ("the lead target arrives in EAR2 essentially all at once.", 1, RED, False),
    ("It dumps enormous charge in the Micromegas and saturates (“blinds”) the DREAM DAQ", 0),
    ("exactly when the high-energy X17 signal would appear.", 1),
], Inches(0.55), Inches(1.6), Inches(7.0), Inches(5.2), size=16)
fit_image(s, A("fig_yield_tof.png"), Inches(7.7), Inches(1.7), Inches(5.3), Inches(3.4),
          caption="Expected X17 yield vs neutron flight time. Flash at t≈0; DREAM dead-time follows.")
callout(s, "Feb & May were technical tests: can we survive the flash and still\nbe efficient to e⁺e⁻ tracks in the high-energy window?",
        Inches(7.7), Inches(5.45), Inches(5.3), Inches(1.1), size=14, bold=True, color=NAVY)
footer(s, "Context")

# =============================================================================
#  2. The gamma flash & DREAM saturation
# =============================================================================
s = add_slide(); title_box(s, "The gamma flash blinds the DREAM DAQ", kicker="The core problem")
fit_image(s, A("fig_dream_progression.png"), Inches(0.5), Inches(1.55), Inches(12.3), Inches(3.2),
          caption="Helium/Ethane, 30 mm drift. Raising resist HV (360→510 V): flash fills the readout, then ALL channels go flat — DREAM saturates for the full 10 µs window.")
bullets(s, [
    ("At low HV the detector is not yet efficient to ~MeV electrons/positrons.", 0),
    ("As HV is raised toward efficiency, the flash charge grows until the DREAM front-end", 0),
    ("locks at baseline and recovers only after the entire 10 µs readout window.", 1),
    ("Confirmed to be a DAQ limit (not the detector) — see backup: single MM strip read on", 0, GREEN, True),
    ("the n_TOF 1 GHz scope shows clear signals where DREAM reports nothing.", 1, GREEN, False),
], Inches(0.55), Inches(5.25), Inches(12.2), Inches(2.0), size=14)
footer(s, "Problem")

# =============================================================================
#  SECTION: FEBRUARY
# =============================================================================
section_divider("February 2026", "Single Micromegas + SiPM wall — first characterisation in the n_TOF beam")

# Feb setup
s = add_slide(); title_box(s, "February setup — overhead view", kicker="February · Setup")
fit_image(s, A("setup_february.png"), Inches(0.4), Inches(1.5), Inches(6.6), Inches(5.6))
bullets(s, [
    ("One prototype MX17 resistive-strip Micromegas (40×40 cm) ~20 cm from the beam axis.", 0),
    ("One SiPM wall (50×50 cm, 20 × 2.5 cm bars) directly behind it.", 0),
    ("No ³He target — samples (empty, lead, B4C, carbon, Timepix) swapped on the beam line.", 0),
    ("Read out with the DREAM DAQ; PS-pickup trigger aligned just before the flash.", 0),
    ("Variables scanned: gas (Ar- and He-based), resist HV, drift gap (30 → 11 → 6 mm),", 0),
    ("aluminium vs carbon frame.", 1),
    ("On beam ~1 month — the workhorse dataset for understanding the flash.", 0, NAVY, True),
], Inches(7.2), Inches(1.7), Inches(5.9), Inches(5.2), size=14.5)
footer(s, "February")

# Feb key result: saturate before efficient
s = add_slide(); title_box(s, "DREAM saturates BEFORE the detector is efficient", kicker="February · Key result")
fit_image(s, A("fig_sat_vs_eff.png"), Inches(0.45), Inches(1.55), Inches(8.0), Inches(5.3),
          caption="He/Ethane, empty carbon capsule. Solid: hits/event in the 1.5–5.4 µs signal window vs HV (rises, peaks, then →0 as DREAM saturates). Dashed: e⁻ efficiency calibrations.")
bullets(s, [
    ("Primary metric: mean strip hits/event vs resist HV,", 0),
    ("in the ~1 MeV signal window (1.5–5.4 µs).", 1),
    ("Curve rises as the detector turns on, then collapses", 0),
    ("to zero as the flash saturates DREAM.", 1),
    ("The collapse (“turn-off”) happens at LOWER HV than", 0, RED, True),
    ("the electron-efficiency plateau (Cs-137 & scint.).", 1, RED, False),
    ("⇒ No usable operating window for the high-energy", 0, NAVY, True),
    ("measurement in this configuration.", 1, NAVY, True),
], Inches(8.6), Inches(1.7), Inches(4.6), Inches(5.0), size=13.5)
footer(s, "February")

# Feb: drift gap helps
s = add_slide(); title_box(s, "Smaller drift gap → DREAM survives to higher HV", kicker="February · What helped")
fit_image(s, A("fig_driftgap_compare.png"), Inches(0.4), Inches(1.55), Inches(7.7), Inches(4.2))
fit_image(s, A("fig_turnoff_vs_drift.png"), Inches(8.2), Inches(1.55), Inches(4.9), Inches(3.0))
bullets(s, [
    ("Reducing the active volume (30 → 11 → 6 mm drift) lowers the flash charge integrated by DREAM,", 0),
    ("so the turn-off shifts up: ~430 V (30 mm) → ~470 V (11 mm) → ~500 V (6 mm) in He/Ethane.", 1),
    ("Firm, repeatable conclusion. But extrapolation suggests we'd need ≲3 mm drift to reach efficiency —", 0, NAVY, True),
    ("not demonstrated, and not robust without a Saclay cosmic calibration in the same gas.", 1),
], Inches(0.5), Inches(5.95), Inches(12.6), Inches(1.4), size=13.5)
footer(s, "February")

# Feb: the cost
s = add_slide(); title_box(s, "…but smaller drift gap costs us the tracking", kicker="February · The trade-off")
fit_image(s, A("fig_angres_vs_drift.png"), Inches(0.45), Inches(1.6), Inches(7.4), Inches(5.0),
          caption="Max resolvable opening angle vs drift gap (µ-TPC mode).")
bullets(s, [
    ("The 30 mm gap enables µ-TPC mode: full 3-D track angle", 0),
    ("reconstruction (≈2° resolution on the Saclay bench).", 1),
    ("Shrinking the gap to survive the flash sacrifices the", 0, RED, True),
    ("longitudinal lever arm → no 3-D angle, only projected hits", 1, RED, False),
    ("(“hit mode”).", 1, RED, False),
    ("Fewer readout strips for July (channel budget) further", 0),
    ("degrades the minimum resolvable angle.", 1),
    ("So: a flash-surviving config and a track-reconstructing", 0, NAVY, True),
    ("config pull in opposite directions.", 1, NAVY, True),
], Inches(8.1), Inches(1.7), Inches(5.0), Inches(5.0), size=13.5)
footer(s, "February")

# Feb: best case overlap
s = add_slide(); title_box(s, "Best case seen: a marginal window in He/Ethane at 6 mm", kicker="February · Best case")
fit_image(s, A("fig_he_overlap_run94.png"), Inches(0.5), Inches(1.6), Inches(6.2), Inches(4.6),
          caption="Run 94 — SiPM-coincidence beam scan, filtered vs all events.")
bullets(s, [
    ("With He/Ethane + 6 mm drift, the efficiency turn-on (~525 V)", 0),
    ("and the DREAM turn-off (~520 V) very nearly overlap.", 1),
    ("Hints at a narrow ~515–530 V window where the detector may", 0),
    ("be approaching efficiency while DREAM is not yet fully dead.", 1),
    ("Helium also brings high neutron sensitivity → large point-like", 0, RED, True),
    ("background in the same window.", 1, RED, False),
    ("Encouraging but unproven; not a robust working point.", 0, NAVY, True),
], Inches(7.0), Inches(1.7), Inches(6.1), Inches(5.0), size=13.5)
footer(s, "February")

# Feb summary
s = add_slide(); title_box(s, "February — what we understood", kicker="February · Summary")
bullets(s, [
    ("Confirmed: the DREAM DAQ (not the detector) is the limiting element — it saturates from", 0, NAVY, True),
    ("the flash at operating HV and stays blind for the full readout window.", 1),
    ("All Argon-based gases saturate at or below the efficiency threshold → no window.", 0),
    ("Helium/Ethane survives the flash more cleanly, but is highly neutron-sensitive.", 0),
    ("Reducing the drift gap reliably buys HV headroom — but sacrifices µ-TPC tracking.", 0),
    ("Single-strip scope test proves signal exists in the DREAM-blind region.", 0),
    ("Carbon frame: only ~10–20% hit-rate reduction; not worth the complexity.", 0),
    ("Grounding the MX connector cards almost entirely removed flash-induced sinusoidal noise.", 0),
    ("Net: no demonstrated high-energy working point; clear levers and limits identified for May.", 0, GREEN, True),
], Inches(0.6), Inches(1.7), Inches(12.4), Inches(5.2), size=15.5)
footer(s, "February")

# =============================================================================
#  SECTION: MAY
# =============================================================================
section_divider("May 2026", "Two Micromegas at 90° — DAQ fixes, flash-mitigation attempts, thermal trigger")

# May setup
s = add_slide(); title_box(s, "May setup — overhead view", kicker="May · Setup")
fit_image(s, A("setup_may.png"), Inches(0.4), Inches(1.5), Inches(7.0), Inches(5.7))
bullets(s, [
    ("Two MX17 Micromegas at 90° to each other, ~20 cm from the beam axis:", 0),
    ("MX17-3 (Det-3): 16 mm drift, full M1 readout — the reliable detector.", 1),
    ("MX17-4 (Det-4): 30 mm drift, multiplexed M2 — found largely inefficient.", 1),
    ("SiPM wall (50×50 cm) behind Det-3.", 0),
    ("One GDD trigger scintillator (30×30 cm) behind each detector.", 0),
    ("6–18 May; four gas mixtures; scintillator-coincidence trigger commissioned", 0),
    ("for the first time.", 1),
], Inches(7.6), Inches(1.7), Inches(5.5), Inches(5.2), size=14.5)
footer(s, "May")

# May: DREAM 4096 bug
s = add_slide(); title_box(s, "Found & fixed: the DREAM “4096” lock-up", kicker="May · DAQ")
bullets(s, [
    ("First beam runs hit a show-stopper: random DREAM channels stuck at exactly 4096", 0, RED, True),
    ("(max ADC) for entire subruns.", 1),
    ("Diagnosed with Simone's noise generator + Irakli's hint to cut the sample count.", 0),
    ("A clean sweep (Runs 21–30) bracketed it: lock-up appears at ≳500 samples;", 0),
    ("≤450 samples → all channels behave.", 1),
    ("Adopted 450 samples = 9 µs window (20 ns/sample) as the standard.", 0, GREEN, True),
    ("Likely a buffer/timing overflow in the DREAM ASIC at long readout length —", 0),
    ("to be confirmed with Irakli (firmware fix vs hard limit).", 1),
    ("Consumed ~1.5 days of beam, but unblocked all subsequent running.", 0, GREY, False),
], Inches(0.6), Inches(1.7), Inches(7.4), Inches(5.2), size=15)
callout(s, "Cost the high-energy window 80 → 9 µs?\nNo — 450 samples still spans the\nfull ~1 MeV signal region.",
        Inches(8.4), Inches(2.0), Inches(4.5), Inches(1.6), size=14, color=NAVY)
fit_image(s, A("fig_neon_event.png"), Inches(8.3), Inches(3.9), Inches(4.7), Inches(2.9),
          caption="Clean event readout after the fix (Ne/Iso).")
footer(s, "May")

# May: mitigation attempts
s = add_slide(); title_box(s, "Three flash-mitigation ideas — none worked out", kicker="May · Mitigation attempts")
table(s, [
    ["Idea", "What we tried", "Outcome"],
    ["2 mm Al plate", "Shield Det-4 against low-energy X-rays in the flash",
     "No effect — flash is dominated by MeV photons/electrons, not keV X-rays"],
    ["Mesh charge injection", "Inject positive charge on the floating mesh to cancel flash charge",
     "Tentative early promise, but circuit killed by 400–800 V discharges from the resistive layer; could not iterate"],
    ["Neon / Isobutane 95/5", "Lower-Z gas: predicted ‘goldilocks’ between Ar and He",
     "Efficiency onset ~490 V, flash-survival peak ~420 V — ~70 V gap, not usable; neon also costly"],
], Inches(0.55), Inches(1.7), Inches(12.25), Inches(3.4), fs=13.5,
   col_w=[Inches(2.4), Inches(4.3), Inches(5.55)])
callout(s, "Conclusion: no viable path to a high-energy measurement was demonstrated. "
           "The mesh-injection circuit is the only surviving concept, but needs a discharge-protected redesign. "
           "The high-energy measurement is not planned for July.",
        Inches(0.55), Inches(5.35), Inches(12.25), Inches(1.5), size=15, bold=True, color=NAVY)
footer(s, "May")

# May: thermal trigger
s = add_slide(); title_box(s, "Pivot: commissioning the thermal-neutron trigger", kicker="May · Thermal path")
bullets(s, [
    ("From 13 May, focus shifted to the thermal window (t > ~1 ms): low occupancy, clean topology,", 0),
    ("no flash in the signal region — the ~20% of yield we can realistically reach.", 1),
    ("Trigger = SiPM wall (behind Det-3) in coincidence with the GDD scintillator behind it.", 0),
    ("The coincidence rejects single Compton gammas and selects charged particles through both layers.", 1),
    ("Measured coincidence rates per PS pulse (after monostable correction):", 0, NAVY, True),
], Inches(0.55), Inches(1.6), Inches(12.4), Inches(2.2), size=15)
table(s, [
    ["Coincidence configuration", "Non-flash triggers / PS pulse"],
    ["Det-4 scint + Det-3 wall", "≈ 2.68"],
    ["Det-3 scint + Det-3 wall", "≈ 6.87"],
    ["Det-4 scint + Det-3 scint + Det-3 wall (3-way)", "≈ 0.11  (very rare)"],
], Inches(0.7), Inches(4.0), Inches(7.6), Inches(1.7), fs=13.5, col_w=[Inches(5.2), Inches(2.4)])
bullets(s, [
    ("Sr-90 source test: the ~1 mm Al backing on the SiPM wall blocks 1–2 MeV electrons —", 0, RED, True),
    ("must be removed/replaced with thin foil for July.", 1, RED, False),
    ("Lesson: NIM trigger logic must live inside EAR2 — the long cable run drops signals", 0),
    ("below the N1081B 10 mV threshold.", 1),
], Inches(8.5), Inches(3.95), Inches(4.6), Inches(2.8), size=13)
footer(s, "May")

# May: gas comparison
s = add_slide(); title_box(s, "Four gases on beam — comparison", kicker="May · Gas study")
table(s, [
    ["Gas", "Op. V", "Flash behaviour / notes"],
    ["Ar/CF4 90/10", "~635 V", "February standard; flash saturation at op. V; early beam work"],
    ["Ne/Iso 95/5", "~490 V", "No discharges on ramp (advantage); flash peak ~420 V → gap not useful; not viable"],
    ["He/Eth 96.5/3.5", "~515 V", "Lower discharge rate; survives flash best; good thermal data; Feb calibration exists"],
    ["Ar/CF4/Iso 88/10/2", "~800 V", "Very high V; heavy sparking on ramp; flash saturation confirmed (Feb benchmark)"],
], Inches(0.55), Inches(1.7), Inches(12.25), Inches(3.0), fs=13,
   col_w=[Inches(2.6), Inches(1.4), Inches(8.25)])
callout(s, "July decision: an Argon/Isobutane mix — operating voltage similar to Ne/Iso, "
           "with the higher primary ionisation of an Argon-based gas for track quality. "
           "(Note: ZS mode was left on from Run 36 — all Ne/Iso, He/Eth, Ar/CF4/Iso data are zero-suppressed; analysis must account for it.)",
        Inches(0.55), Inches(5.0), Inches(12.25), Inches(1.7), size=14, color=NAVY)
footer(s, "May")

# =============================================================================
#  SUMMARY
# =============================================================================
s = add_slide(); title_box(s, "Summary — what these technical tests established", kicker="Bottom line")
bullets(s, [
    ("These were technical tests, not physics runs — the question was feasibility of the", 0, GREY, True),
    ("high-energy X17 measurement against the gamma flash.", 1, GREY, False),
    ("The DREAM DAQ saturates from the flash before the detector is efficient to e⁺e⁻ — confirmed", 0, RED, True),
    ("in both Feb and May, across all gases tested.", 1, RED, False),
    ("Levers we quantified: smaller drift gap buys HV headroom (but kills tracking); He/Ethane", 0),
    ("survives the flash best; reducing flash charge is the whole game.", 1),
    ("Mitigation ideas tried — Al shielding, mesh charge injection, neon gas — none gave a usable window.", 0, RED, True),
    ("DAQ understood and hardened: 4096 lock-up fixed (450 samples); grounding, ZS, RunCtrl issues mapped.", 0, GREEN, True),
    ("The thermal-neutron measurement is the viable path: coincidence trigger commissioned, ~3 triggers/pulse,", 0, GREEN, True),
    ("architecture for July substantially defined.", 1, GREEN, False),
    ("July: full 4-detector + ³He-target run, targeting the thermal region.", 0, NAVY, True),
], Inches(0.6), Inches(1.6), Inches(12.5), Inches(5.4), size=15)
footer(s, "Summary")

# =============================================================================
#  BACKUP
# =============================================================================
section_divider("Backup", "Supporting figures, tables and detail")

def img_slide(title, kicker, asset, caption=None, bullet_items=None):
    s = add_slide(); title_box(s, title, kicker=kicker)
    if bullet_items:
        fit_image(s, A(asset), Inches(0.45), Inches(1.6), Inches(8.0), Inches(5.2), caption=caption)
        bullets(s, bullet_items, Inches(8.6), Inches(1.7), Inches(4.6), Inches(5.0), size=13)
    else:
        fit_image(s, A(asset), Inches(0.7), Inches(1.6), Inches(11.9), Inches(5.1), caption=caption)
    footer(s, "Backup")
    return s

img_slide("Motivation — the ATOMKI X17 anomaly", "Backup · Physics",
          "fig_atomki_peak.png",
          caption="ATOMKI: excess e⁺e⁻ pairs peaking near ~120° opening angle — consistent with a 17 MeV boson.")
img_slide("Gamma flash on the n_TOF 1 GHz DAQ", "Backup · Flash",
          "fig_flash_scope.png",
          caption="Single MM strip on the n_TOF scope: ~ -15 mV, ~300 ns → ~90 pC/strip, already enough to saturate DREAM. Signals are visible here where DREAM reports nothing.")
img_slide("Flash saturates all DREAM channels (waveforms)", "Backup · Flash",
          "fig_flash_waveforms.png",
          caption="Run 18, Ar/CF4/Iso. All X/Y channels driven to the 4096 rail by the flash, then flat baseline — no recovery within the readout window.")
img_slide("High occupancy in the ~1 MeV neutron peak", "Backup · Flash",
          "fig_high_occupancy.png",
          caption="Even when not fully saturated, the post-flash hit density makes individual e⁺e⁻ track reconstruction extremely hard (Helium gas).")
img_slide("What we CAN see — isolated thermal-region track", "Backup · Thermal",
          "fig_thermal_track.png",
          caption="He/Ethane, 30 mm drift, µ-TPC mode: a clean, well-separated track in the thermal window — exactly the regime targeted for July.")
img_slide("He/Ethane 30 mm — target comparison", "Backup · February",
          "fig_he_targets.png",
          caption="Hits/event vs HV in three time windows for different samples. Carbon turns on/off earliest (more material → more scattering); B4C loudest at late times.")
img_slide("Cost of fewer strips / smaller acceptance", "Backup · February",
          "fig_deadzone_cost.png",
          caption="Geometry of minimum resolvable angle and dead-zone fractions for the readout/channel-budget trade-offs.")
img_slide("Neon/Isobutane event display", "Backup · May",
          "fig_neon_event.png",
          caption="Run 47, Ne/Iso, resist 530 V: waveforms + reconstructed hits. Neon shows cleaner low-energy tracks than Ar/CF4, but no usable flash window.")
img_slide("Sr-90 beta spectrum (source test)", "Backup · May",
          "fig_sr90_spectrum.png",
          caption="Sr-90 endpoint ~2 MeV with substantial spectrum near 1 MeV — used to probe SiPM-wall material budget; the ~1 mm Al backing blocks these electrons.")
img_slide("EAR2 setup photograph", "Backup · Setup",
          "fig_setup_photo.png",
          caption="The MX17 detector and DREAM rack installed in EAR2.")

# Backup: February gas/sample table
s = add_slide(); title_box(s, "February — gases & samples tested", kicker="Backup · February")
table(s, [
    ["Gas mixture", "Period", "Key characteristics"],
    ["Ar/CF4/Iso 88/10/2", "Feb 2–7, 9–11", "Strong flash saturation; op. V ~490 V; high spark rate"],
    ["He/Ethane 96.5/3.5", "Feb 3–7, 20–Mar 2", "Survives flash at low HV; very neutron-sensitive; op. V ~500–540 V"],
    ["Ar/CF4/CO2 45/40/15", "Feb 9–11", "Brief; op. V ~680–800 V; outside usable flash window"],
    ["Ar/CF4 90/10", "Feb 10–16", "± B4C target; op. V region incompatible with flash survival"],
    ["Ar/CO2 70/30", "Feb 16–Mar 1", "Op. V ~680–720 V; saturation well below op. V"],
], Inches(0.55), Inches(1.65), Inches(12.25), Inches(2.5), fs=12.5,
   col_w=[Inches(3.0), Inches(2.3), Inches(6.95)])
bullets(s, [
    ("Samples: empty / lead (gammas) / B4C (strong n absorber) / carbon (≈ July capsule) / Timepix.", 0),
    ("Drift gaps: 30 mm (µ-TPC) → 6 mm (Feb 18) → 11 mm (final days). Frame: aluminium → carbon (Feb 26).", 0),
    ("Feb 19: single-strip scope test conclusively pins the blindness on DREAM, not the detector.", 0, GREEN, True),
], Inches(0.6), Inches(4.4), Inches(12.4), Inches(2.0), size=13.5)
footer(s, "Backup")

# Backup: February run summary (selected)
s = add_slide(); title_box(s, "February — selected run summary", kicker="Backup · February")
table(s, [
    ["Run", "Date", "Gas", "Target", "Drift", "Frame", "Note"],
    ["18", "Feb 3", "Ar/CF4/Iso", "Timepix", "30", "Al", "First detailed flash study"],
    ["26", "Feb 4", "He/Eth", "None", "30", "Al", "Overnight HV scan"],
    ["30", "Feb 5", "He/Eth", "B4C", "30", "Al", "B4C vs no-target"],
    ["33", "Feb 6", "He/Eth", "Carbon", "30", "Al", "Carbon target scan"],
    ["52–54", "Feb 11–12", "Ar/CF4", "None", "30", "Al", "SiPM-triggered cosmics"],
    ["71", "Feb 19", "Ar/CO2", "B4C", "6", "Al", "0 V vs 1000 V drift — 6 mm alive"],
    ["97", "Feb 24", "He/Eth", "Cs-137", "6", "Al", "Key efficiency calibration"],
    ["136", "Feb 28", "He/Eth", "Lead", "6", "Carbon", "Al vs carbon frame (~10–20%)"],
    ["143", "Mar 1", "He/Eth", "Carbon", "11", "Carbon", "Final 11 mm dataset"],
], Inches(0.7), Inches(1.65), Inches(11.9), Inches(5.0), fs=12,
   col_w=[Inches(1.0), Inches(1.4), Inches(2.0), Inches(1.6), Inches(1.0), Inches(1.4), Inches(3.5)])
footer(s, "Backup")

# Backup: May issues
s = add_slide(); title_box(s, "May — issues identified", kicker="Backup · May")
table(s, [
    ["Issue", "Status / action"],
    ["DREAM 4096 lock-up at ≳470 samples", "RESOLVED — limit to 450 samples (9 µs window)"],
    ["Zero-suppression left on from Run 36", "All Ne/Iso, He/Eth, Ar/CF4/Iso data are ZS; reprocess with flag"],
    ["Det-4 largely inefficient (M2 mux)", "Physics analysis relies on Det-3 only"],
    ["SiPM wall ~1 mm Al backing", "Blocks 1–2 MeV e⁻; remove / thin-foil for July"],
    ["NIM trigger outside EAR2", "Cable attenuation < N1081B threshold; move logic in-area"],
    ["B4C target misaligned (Runs 35–44)", "B4C data from those runs suspect"],
    ["RunCtrl manual ‘G’ keypress", "Needs software watchdog for overnight runs"],
    ["Disk I/O bottleneck on DAQ PC", "Clone OS to SSD before July"],
], Inches(0.55), Inches(1.65), Inches(12.25), Inches(4.6), fs=12.5,
   col_w=[Inches(5.0), Inches(7.25)])
footer(s, "Backup")

# Backup: July trigger concept
s = add_slide(); title_box(s, "July concept — thermal trigger architecture", kicker="Backup · Outlook")
bullets(s, [
    ("Full four-detector geometry around a pressurised ³He carbon capsule.", 0, NAVY, True),
    ("Per side: SiPM wall (16/20 bars → 8 channels) summed & discriminated → one logic pulse.", 0),
    ("Coincide with the first layer of a new liquid-scintillator calorimeter (PMT readout) →", 0),
    ("a “sector hit” that rejects single Compton gammas but accepts e⁺/e⁻.", 1),
    ("Singles = any one sector; Doubles = majority ≥2 (for a ~120° pair).", 0),
    ("Optional hardware PS-flash veto; scaledown to Doubles if singles rate too high.", 0),
    ("Liquid scintillator must contain up to 10 MeV e± within ~5 cm (GEANT4 to verify acceptance).", 0),
    ("Rate scaling vs empty-capsule ~3/pulse: area ×1.8, no Al ×~2, ³He fill ×2–3, four walls ×4", 0),
    ("→ O(50–200) singles/pulse; DREAM dead-time vs rate must be modelled to choose Singles vs Doubles.", 1),
], Inches(0.6), Inches(1.65), Inches(12.5), Inches(5.2), size=14.5)
footer(s, "Backup")

# Backup: roadmap
s = add_slide(); title_box(s, "Roadmap to July", kicker="Backup · Outlook")
table(s, [
    ["Period", "Milestone", "Key items"],
    ["June 2026", "Production", "Build 4 final MMs (30 mm, Al frame); upgrade/obtain DREAM FEUs & 64 cables; build liquid-scint modules & 3 more SiPM walls"],
    ["June 2026", "DAQ & trigger", "In-area NIM trigger (N1081B, veto); SSD clone; auto-ZS detection; RunCtrl watchdog; GEANT4 acceptance + dead-time model"],
    ["July 2026", "Physics run", "Four-detector thermal-region X17 search with ³He target"],
], Inches(0.55), Inches(1.7), Inches(12.25), Inches(3.4), fs=13,
   col_w=[Inches(1.6), Inches(2.2), Inches(8.45)])
bullets(s, [
    ("Open detector-quality question: Det-4 inefficiency and Det-3 discharges — root cause before building the four.", 0, RED, True),
], Inches(0.6), Inches(5.4), Inches(12.4), Inches(0.9), size=14)
footer(s, "Backup")

out = os.path.join(HERE, "X17_nTOF_Feb_May_2026_Status.pptx")
prs.save(out)
print("Saved", out, "with", len(prs.slides.__iter__.__self__._sldIdLst), "slides")
