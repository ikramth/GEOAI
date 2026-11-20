#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Generate GeoAI_Presentation.pptx for USB submission.

Usage:
    python generate_pptx.py
Requirements:
    pip install python-pptx
"""
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

ROOT = Path(__file__).resolve().parent
PUBLIC = ROOT / "public"
OUT = ROOT / "GeoAI_Presentation.pptx"

NAVY = RGBColor(0x0B, 0x25, 0x45)
NAVY2 = RGBColor(0x13, 0x40, 0x74)
CYAN = RGBColor(0x00, 0x66, 0x99)
LIGHT = RGBColor(0x8F, 0xE3, 0xFF)
CHAR = RGBColor(0x1E, 0x29, 0x3B)
GREY = RGBColor(0x47, 0x55, 0x69)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
SLATE = RGBColor(0xF1, 0xF5, 0xF9)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]


def set_bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_rect(slide, left, top, width, height, fill, radius=0.15, line=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top),
                                  Inches(width), Inches(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
    # adjust corner radius
    try:
        shape.adjustments[0] = radius
    except Exception:
        pass
    return shape


def add_text(slide, text, left, top, width, height, size=18, bold=False,
             color=CHAR, align=None, italic=False, name="Calibri"):
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    p = tf.paragraphs[0]
    p.text = text
    if align is not None:
        p.alignment = align
    r = p.runs[0]
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = color
    r.font.name = name
    return box


def add_bullets(slide, items, left, top, width, height, size=15, color=CHAR, bullet="•"):
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = f"{bullet} {item}"
        p.space_after = Pt(6)
        r = p.runs[0]
        r.font.size = Pt(size)
        r.font.color.rgb = color
        r.font.name = "Calibri"
    return box


def header(slide, title, subtitle=None, dark=True):
    # top header bar
    bar = add_rect(slide, 0, 0, 13.333, 1.15, NAVY if dark else SLATE, radius=0)
    add_text(slide, title, 0.6, 0.22, 12.1, 0.6, size=26, bold=True,
             color=WHITE if dark else NAVY)
    if subtitle:
        add_text(slide, subtitle, 0.6, 0.78, 12.1, 0.4, size=12, italic=True,
                 color=LIGHT if dark else GREY)


def footer(slide, text="GeoAI · Master 1 Smart Cities · Université Yahia Farès de Médéa · 2026-2027"):
    add_text(slide, text, 0.6, 7.05, 12.1, 0.3, size=9, color=GREY)


def add_picture(slide, name, left, top, width=None, height=None, contain=True):
    path = PUBLIC / name
    if not path.exists():
        add_text(slide, f"[image manquante: {name}]", left, top, 4, 0.4, size=12, color=GREY)
        return
    kwargs = {}
    if width is not None:
        kwargs["width"] = Inches(width)
    if height is not None:
        kwargs["height"] = Inches(height)
    pic = slide.shapes.add_picture(str(path), Inches(left), Inches(top), **kwargs)
    return pic


def card(slide, left, top, width, height, title, body_lines, fill=WHITE, title_color=NAVY):
    add_rect(slide, left, top, width, height, fill, radius=0.12)
    add_text(slide, title, left + 0.25, top + 0.18, width - 0.5, 0.35, size=13, bold=True, color=title_color)
    y = top + 0.55
    for line in body_lines:
        add_text(slide, line, left + 0.25, y, width - 0.5, 0.3, size=11, color=CHAR)
        y += 0.32


# ---------- Slide 1 : Couverture ----------
slide = prs.slides.add_slide(BLANK)
set_bg(slide, SLATE)
header(slide, "GeoAI : L'Intelligence Artificielle au Service de l'Information Géospatiale",
       "Transformer la donnée satellite en connaissance spatiale — modèles de vision par ordinateur intégrés aux SIG.", dark=True)
add_text(slide, "Projet CC · Module GIS / SIG", 0.7, 1.5, 6, 0.4, size=12, bold=True, color=CYAN)
add_text(slide, "Réalisé par", 0.7, 2.0, 4, 0.4, size=14, bold=True, color=NAVY)
add_text(slide, "1. TAHRAOUI IKRAM", 0.9, 2.45, 6, 0.4, size=17, bold=True, color=NAVY)
add_text(slide, "2. BENKORTBI OUSSAMA", 0.9, 2.9, 6, 0.4, size=17, bold=True, color=NAVY)
add_bullets(slide, [
    "Master 1 Smart Cities Engineering (Métiers de la Ville)",
    "Département d'Architecture · Faculté de Technologie",
    "Université Yahia Farès de Médéa · 2026-2027",
    "Encadrant : Dr. Amel KELLOU",
], 0.7, 3.6, 6.2, 2.6, size=13)
footer(slide)

# ---------- Slide 2 : Fondements ----------
slide = prs.slides.add_slide(BLANK)
set_bg(slide, WHITE)
header(slide, "Trois disciplines, une intersection : la GeoAI",
       "Chaque pilier apporte une contrainte que les deux autres rendent féconde.")
card(slide, 0.7, 1.6, 3.9, 4.6, "01 · Géodésie", [
    "Le cadre planétaire.",
    "Positionner juste tout objet sur la Terre.",
    "WGS84 · Projections · Coordonnées",
    "Rôle : justesse spatiale",
])
card(slide, 4.75, 1.6, 3.9, 4.6, "02 · SIG", [
    "La base relationnelle.",
    "Stocker, relier et interroger les couches.",
    "Tables attributaires · SQL · Jointures",
    "Rôle : mémoire du territoire",
])
card(slide, 8.8, 1.6, 3.9, 4.6, "03 · Intelligence Artificielle", [
    "La puissance cognitive.",
    "Apprendre les patterns et inférer.",
    "Apprentissage profond · Inférence",
    "Rôle : moteur d'interprétation",
])
add_text(slide, "Géodésie + SIG + IA = GEOAI", 0.7, 6.4, 12, 0.4, size=15, bold=True, color=CYAN, align=PP_ALIGN.CENTER)
footer(slide)

# ---------- Slide 3 : Pipeline ----------
slide = prs.slides.add_slide(BLANK)
set_bg(slide, SLATE)
header(slide, "De l'Espace à la Carte : Comment la GeoAI Analyse un Territoire",
       "Acquisition satellitaire (Médéa) → station sol ASAL (Ouargla) → prétraitement COG → CNN 3×3 → Segmentation finale.")
card(slide, 0.7, 1.6, 6.0, 4.9, "Flux en 4 étapes", [
    "01 Capture satellitaire — ville de Médéa, orbite et balayage.",
    "02 Transmission ASAL — station sol de Ouargla, prétraitement COG.",
    "03 Pixellisation CNN — noyau 3×3, matrices de caractéristiques.",
    "04 Segmentation finale — extraction routes et bâtiments en vecteurs.",
], fill=WHITE)
card(slide, 6.9, 1.6, 5.8, 4.9, "Sortie du modèle", [
    "U-Net · masque sémantique",
    "Résolution : 10 m / pixel · COG",
    "IoU moyen : 0.94",
    "Classes : route, bâtiment, végétation, eau",
], fill=WHITE, title_color=CYAN)
footer(slide)

# ---------- Slide 3b : Apple (full-bleed image) ----------
slide = prs.slides.add_slide(BLANK)
set_bg(slide, WHITE)
header(slide, "How AI reads layers and recognizes objects? — green apple",
       "Spectral signature · geometric edges · spatial relationships · 3D/AI processing")
add_picture(slide, "apple-layers.jpg", 1.0, 1.35, width=11.3)
footer(slide)

# ---------- Slide 3c : Layers (full-bleed image) ----------
slide = prs.slides.add_slide(BLANK)
set_bg(slide, WHITE)
header(slide, "Layers (Bands) — satellite bands / NDVI",
       "Visible Red · NIR · SWIR · Pixel · NDVI formula · GeoAI classification")
add_picture(slide, "geoai-bands.jpg", 1.0, 1.35, width=11.3)
footer(slide)

# ---------- Slide 4 : Rupture ----------
slide = prs.slides.add_slide(BLANK)
set_bg(slide, SLATE)
header(slide, "Du SIG classique à la GeoAI : changer de paradigme", None)
card(slide, 0.7, 1.6, 5.9, 4.9, "SIG classique — avant", [
    "Numérisation manuelle, couche par couche.",
    "Calibration et co-registrement pas à pas.",
    "Règles statiques écrites par l'expert.",
    "Cycle long, non parallélisable.",
], fill=WHITE)
card(slide, 6.75, 1.6, 5.9, 4.9, "GeoAI — maintenant", [
    "Inférence cognitive automatisée.",
    "Alignement appris par vision par ordinateur.",
    "Patterns extraits des données.",
    "Mise à l'échelle GPU.",
], fill=RGBColor(0xE8, 0xF4, 0xFB), title_color=CYAN)
footer(slide)

# ---------- Slide 5 : Données ----------
slide = prs.slides.add_slide(BLANK)
set_bg(slide, WHITE)
header(slide, "Avant tout modèle : une donnée propre, calée, lisible", None)
add_bullets(slide, [
    "Imagerie satellitaire HD : plurispectre, Sentinel-2 → Pléiades, séries temporelles.",
    "Formats optimisés : GeoTIFF (géoréférence), COG (tuiles, HTTP partiel), STAC.",
    "Co-registrement automatique : points-clés + homographies, SuperGlue.",
    "Pipeline : acquisition → correction nuages → co-registrement → tuilage COG → prêt pour le modèle.",
], 0.7, 1.6, 12.0, 5.0, size=16)
footer(slide)

# ---------- Slide 6 : CNN ----------
slide = prs.slides.add_slide(BLANK)
set_bg(slide, SLATE)
header(slide, "CNN : des filtres qui apprennent à voir", None)
add_bullets(slide, [
    "Inspiré du cortex visuel : convolutions transforment les pixels en concepts.",
    "01 Extraction — noyaux 3×3 détectent bords, textures, formes.",
    "02 Hiérarchie — bords → motifs → objets.",
    "03 Compression — pooling, réduction de dimension.",
    "Architecture : Image → Conv 3×3·32 → ReLU → MaxPool → Conv 3×3·64 → Feature maps → Dense → Sortie.",
], 0.7, 1.6, 12.0, 5.0, size=15)
footer(slide)

# ---------- Slide 7 : Segmentation ----------
slide = prs.slides.add_slide(BLANK)
set_bg(slide, WHITE)
header(slide, "Segmentation & Détection : U-Net et YOLO", None)
card(slide, 0.7, 1.6, 5.9, 4.9, "U-Net — segmentation sémantique", [
    "Chaque pixel reçoit une classe.",
    "Bâtiments, routes, eau, végétation.",
    "Encodeur → bottleneck → décodeur.",
    "Skip-connections pour les contours.",
], fill=RGBColor(0xE8, 0xF4, 0xFB), title_color=CYAN)
card(slide, 6.75, 1.6, 5.9, 4.9, "YOLO — détection d'objets", [
    "Localiser + classifier en une passe.",
    "bâtiment 0.94 · route 0.89",
    "eau 0.97 · végétation 0.86",
    "Échelle industrielle, temps réel.",
], fill=SLATE)
footer(slide)

# ---------- Slide 8 : Segmentation finale (rich) ----------
slide = prs.slides.add_slide(BLANK)
set_bg(slide, SLATE)
header(slide, "Segmentation finale — sortie vectorielle", "Masque sémantique → polygones, parcelles, bâtiments, routes.")
add_rect(slide, 0.7, 1.5, 7.4, 4.9, WHITE, radius=0.12)
add_text(slide, "Carte segmentée (raster → vecteur)", 0.95, 1.65, 7.0, 0.35, size=13, bold=True, color=NAVY)
# simple legend blocks
add_rect(slide, 1.0, 2.1, 1.6, 0.45, CYAN, radius=0.08)
add_text(slide, "Bâtiments 96.7%", 1.0, 2.1, 1.6, 0.45, size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
add_rect(slide, 2.8, 2.1, 1.6, 0.45, NAVY2, radius=0.08)
add_text(slide, "Routes 98.4%", 2.8, 2.1, 1.6, 0.45, size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
add_rect(slide, 4.6, 2.1, 1.6, 0.45, RGBColor(0x7F, 0xA3, 0x7A), radius=0.08)
add_text(slide, "Végétation", 4.6, 2.1, 1.6, 0.45, size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
add_rect(slide, 6.4, 2.1, 1.5, 0.45, RGBColor(0x1D, 0x4E, 0xD8), radius=0.08)
add_text(slide, "Eau", 6.4, 2.1, 1.5, 0.45, size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
add_text(slide, "Confiance moyenne : 0.94 · Résolution : 10 m/pixel · Format COG", 1.0, 6.0, 7.0, 0.4, size=11, color=GREY)
card(slide, 8.3, 1.5, 4.4, 4.9, "Classes détectées", [
    "Bâtiments — 96.7%",
    "Routes — 98.4%",
    "Végétation — 92.1%",
    "Eau — 97.3%",
    "IoU moyen : 0.94",
], fill=WHITE, title_color=CYAN)
footer(slide)

# ---------- Slide 9 : Prédictive ----------
slide = prs.slides.add_slide(BLANK)
set_bg(slide, WHITE)
header(slide, "Anticiper : du changement observé au risque prédit", None)
add_picture(slide, "cartographie_drone.jpg", 7.2, 1.5, width=5.5)
add_bullets(slide, [
    "Modélisation spatiotemporelle : dynamique du territoire.",
    "Change detection : soustraire deux dates.",
    "Cartographie des risques : crues, érosion.",
    "Polygones cyan = changements détectés.",
], 0.7, 1.6, 6.2, 4.8, size=15)
footer(slide)

# ---------- Slide 10 : Workflow ----------
slide = prs.slides.add_slide(BLANK)
set_bg(slide, SLATE)
header(slide, "Le pipeline complet, de l'image à la carte", None)
add_bullets(slide, [
    "01 Acquisition — satellite, cadastre, catalogues (STAC).",
    "02 Prétraitement — nettoyage, nuages, co-registrement, COG (Python, Rasterio, GDAL).",
    "03 Inférence IA — CNN / U-Net / YOLO sur GPU (PyTorch, ONNX).",
    "04 Cartographie SIG — vectorisation, publication (QGIS, PostGIS, GeoServer).",
], 0.7, 1.6, 12.0, 5.0, size=16)
footer(slide)

# ---------- Slide 11 : Défis ----------
slide = prs.slides.add_slide(BLANK)
set_bg(slide, WHITE)
header(slide, "Trois défis pour une GeoAI de confiance", None)
card(slide, 0.7, 1.6, 3.9, 4.6, "Qualité des données", [
    "Annotation coûteuse.",
    "Biais d'échantillonnage.",
    "Généralisation entre régions.",
], fill=SLATE)
card(slide, 4.75, 1.6, 3.9, 4.6, "Explicabilité (XAI)", [
    "Cartes de saillance.",
    "Confiance calibrée.",
    "Comprendre la frontière.",
], fill=SLATE)
card(slide, 8.8, 1.6, 3.9, 4.6, "Cartographie autonome", [
    "Auto-labelling.",
    "Modèles de fond.",
    "Humain dans la boucle.",
], fill=SLATE)
footer(slide)

# ---------- Slide 12 : Fiche officielle / closing ----------
slide = prs.slides.add_slide(BLANK)
set_bg(slide, SLATE)
header(slide, "Fiche du Projet — GeoAI & SIG", None)
add_bullets(slide, [
    "Réalisé par : 1. TAHRAOUI IKRAM — 2. BENKORTBI OUSSAMA.",
    "Encadré par : Dr. Amel KELLOU.",
    "Module : GIS (Geographic Information Systems and Spatial Data).",
    "Domaine & Filière : Architecture, Urbanisme et Métiers de la Ville — Master 1 Smart Cities Engineering.",
    "Structure : Département d'Architecture | Faculté de Technologie | Université Yahia Farès de Médéa (2026-2027).",
    "Cadre : Évaluation Contrôle Continu (CC) — Semestre 1.",
], 0.7, 1.6, 12.0, 5.0, size=15)
footer(slide)

prs.save(str(OUT))
print(f"Saved {OUT}")