"""
Power BI Theme Visual Editor — Streamlit Application
Focused version: dropdown visual selector, live preview from default JSON,
customization limited to data colors, fonts, x-axis, y-axis, legend.
Now includes predefined theme presets.
"""

import streamlit as st
import json
import copy
import math
import random

# ── PAGE CONFIG ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Power BI Theme Editor",
    page_icon="🎨",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── INJECT CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@300;400;500&family=Fraunces:ital,opsz,wght@0,9..144,300;0,9..144,600;1,9..144,300&display=swap');

html, body, .stApp { background:#080b10 !important; font-family:'IBM Plex Mono',monospace; }
.block-container { padding:0 !important; max-width:100% !important; }
section[data-testid="stSidebar"] { display:none; }
#MainMenu, footer, header { visibility:hidden; }

.topbar {
  display:flex; align-items:center; justify-content:space-between;
  padding:14px 28px; background:#060810;
  border-bottom:1px solid #151c2e;
  position:sticky; top:0; z-index:100;
}
.topbar-logo { font-family:'Fraunces',serif; font-weight:600; font-size:1.25rem; color:#e2e8f8; letter-spacing:-0.03em; }
.topbar-logo em { color:#4f8ef7; font-style:italic; }
.topbar-tag { font-size:9px; letter-spacing:.18em; text-transform:uppercase; color:#2d3a58; background:#0d1220; border:1px solid #1a2440; border-radius:20px; padding:3px 10px; }

.preview-stage {
  padding:24px;
  background: radial-gradient(ellipse at 30% 20%, rgba(79,142,247,0.04) 0%, transparent 60%),
              radial-gradient(ellipse at 80% 80%, rgba(139,92,246,0.03) 0%, transparent 60%),
              #0a0d16;
  min-height:calc(100vh - 57px);
}
.preview-label {
  font-size:9px; letter-spacing:.18em; text-transform:uppercase;
  color:#2a3555; margin-bottom:14px; display:flex; align-items:center; gap:10px;
}
.preview-label::after { content:''; flex:1; height:1px; background:#111827; }

.vis-card {
  background:#fff; border-radius:8px;
  box-shadow: 0 2px 12px rgba(0,0,0,.4), 0 0 0 1px rgba(255,255,255,.05);
  overflow:hidden; margin-bottom:18px;
}

.json-block {
  background:#060810; border:1px solid #111827; border-radius:8px;
  padding:14px; font-size:11px; line-height:1.8; color:#5a7aaa;
  overflow-x:auto; font-family:'IBM Plex Mono',monospace;
  max-height:280px; overflow-y:auto;
}

/* Theme preset cards */
.theme-grid {
  display: grid;
  grid-template-columns: repeat(2, 1fr);
  gap: 8px;
  margin-bottom: 6px;
}
.theme-card {
  border-radius: 8px;
  border: 1px solid #111827;
  overflow: hidden;
  cursor: pointer;
  transition: all 0.18s;
  background: #070b15;
}
.theme-card:hover { border-color: #2a4a8a; transform: translateY(-1px); box-shadow: 0 4px 14px rgba(79,142,247,.15); }
.theme-card.active { border-color: #4f8ef7; box-shadow: 0 0 0 1px #4f8ef7, 0 4px 14px rgba(79,142,247,.2); }
.theme-palette { display:flex; height:8px; }
.theme-palette span { flex:1; }
.theme-card-inner { padding: 7px 9px; }
.theme-name { font-family:'IBM Plex Mono',monospace; font-size:9.5px; color:#8aaad8; letter-spacing:.04em; }
.theme-desc { font-family:'IBM Plex Mono',monospace; font-size:8px; color:#2d3a58; margin-top:2px; }

/* Widget label overrides */
div[data-testid="stSelectbox"] > label,
div[data-testid="stColorPicker"] > label,
div[data-testid="stSlider"] > label,
div[data-testid="stToggle"] > label,
div[data-testid="stTextInput"] > label {
  font-family:'IBM Plex Mono',monospace !important;
  font-size:10px !important; letter-spacing:.1em;
  text-transform:uppercase; color:#3a4d72 !important;
}

div[data-testid="stSelectbox"] > div > div,
div[data-testid="stTextInput"] > div > div > input {
  background:#090d18 !important; border:1px solid #141e34 !important;
  color:#c8d8f0 !important; font-family:'IBM Plex Mono',monospace !important;
  font-size:12px !important; border-radius:6px !important;
}
div[data-testid="stSelectbox"] > div > div:focus-within,
div[data-testid="stTextInput"] > div > div > input:focus {
  border-color:#4f8ef7 !important;
  box-shadow:0 0 0 2px rgba(79,142,247,.15) !important;
}

div[data-testid="stButton"] > button {
  background:#0d1626; border:1px solid #1a2844; color:#8aaad8;
  font-family:'IBM Plex Mono',monospace; font-size:11px;
  border-radius:6px; padding:8px 16px; width:100%; transition:all .18s;
}
div[data-testid="stButton"] > button:hover {
  background:#132040; border-color:#4f8ef7; color:#7eb8ff;
  transform:translateY(-1px); box-shadow:0 4px 12px rgba(79,142,247,.18);
}

div[data-testid="stDownloadButton"] > button {
  background:linear-gradient(135deg,#1a3a6e,#0f2550);
  border:1px solid #2a5aae; color:#7eb8ff;
  font-family:'IBM Plex Mono',monospace; font-size:12px; font-weight:500;
  border-radius:7px; padding:10px 20px; width:100%; transition:all .2s;
}
div[data-testid="stDownloadButton"] > button:hover {
  background:linear-gradient(135deg,#1e4a90,#13305e);
  transform:translateY(-1px); box-shadow:0 6px 20px rgba(79,142,247,.25);
}

div[data-testid="stExpander"] {
  background:#070b15; border:1px solid #0f1828 !important;
  border-radius:7px !important; margin:4px 0;
}
div[data-testid="stExpander"] summary {
  font-family:'IBM Plex Mono',monospace; font-size:11px;
  color:#4a6090; padding:10px 14px;
}
div[data-testid="stExpander"] summary:hover { color:#c8d8f0; }

hr { border-color:#0f1628 !important; margin:8px 0 !important; }

::-webkit-scrollbar { width:4px; height:4px; }
::-webkit-scrollbar-track { background:transparent; }
::-webkit-scrollbar-thumb { background:#141e34; border-radius:4px; }

div[data-testid="stHorizontalBlock"] > div:nth-child(2) {
  background: #0a0d16 !important;
  padding: 0 !important;
}
div[data-testid="stHorizontalBlock"] > div:nth-child(2) > div {
  background: #0a0d16 !important;
}
iframe { border:none !important; }
</style>
""", unsafe_allow_html=True)

# ── PREDEFINED THEMES ─────────────────────────────────────────────────────────
PRESET_THEMES = {
    "Microsoft Default": {
        "desc": "Official PBI palette",
        "dataColors": ["#118DFF","#12239E","#E66C37","#6B007B","#E044A7","#744EC2","#D9B300","#D64550","#197278","#1AAB40"],
        "foreground": "#252423", "foregroundNeutralSecondary": "#605E5C", "foregroundNeutralTertiary": "#B3B0AD",
        "background": "#FFFFFF", "backgroundLight": "#F3F2F1", "backgroundNeutral": "#C8C6C4",
        "tableAccent": "#118DFF", "good": "#1AAB40", "neutral": "#D9B300", "bad": "#D64554",
        "maximum": "#118DFF", "center": "#D9B300", "minimum": "#DEEFFF",
        "fontFace": "Segoe UI", "fontSize": 10, "fontColor": "#252423",
        "xAxisGridStyle": "dotted", "yAxisGridStyle": "dotted",
        "xAxisGridColor": "#B3B0AD", "yAxisGridColor": "#B3B0AD",
    },
    "Midnight Ocean": {
        "desc": "Deep navy & cyan",
        "dataColors": ["#00B4D8","#0077B6","#90E0EF","#023E8A","#48CAE4","#ADE8F4","#CAF0F8","#0096C7","#00B4D8","#03045E"],
        "foreground": "#1A1A2E", "foregroundNeutralSecondary": "#4A4A6A", "foregroundNeutralTertiary": "#8A8AAA",
        "background": "#F0F8FF", "backgroundLight": "#E0F0FA", "backgroundNeutral": "#B0C8E0",
        "tableAccent": "#0077B6", "good": "#00B4D8", "neutral": "#90E0EF", "bad": "#E63946",
        "maximum": "#023E8A", "center": "#90E0EF", "minimum": "#CAF0F8",
        "fontFace": "Segoe UI", "fontSize": 10, "fontColor": "#1A1A2E",
        "xAxisGridStyle": "dashed", "yAxisGridStyle": "dashed",
        "xAxisGridColor": "#B0C8E0", "yAxisGridColor": "#B0C8E0",
    },
    "Ember Forge": {
        "desc": "Warm terracotta tones",
        "dataColors": ["#E76F51","#F4A261","#E9C46A","#264653","#2A9D8F","#A8DADC","#457B9D","#1D3557","#E63946","#F1FAEE"],
        "foreground": "#1D2B1D", "foregroundNeutralSecondary": "#5C4A3A", "foregroundNeutralTertiary": "#A08070",
        "background": "#FFFBF5", "backgroundLight": "#FFF3E0", "backgroundNeutral": "#D4B896",
        "tableAccent": "#E76F51", "good": "#2A9D8F", "neutral": "#E9C46A", "bad": "#E63946",
        "maximum": "#E76F51", "center": "#F4A261", "minimum": "#FFF3E0",
        "fontFace": "Georgia", "fontSize": 10, "fontColor": "#1D2B1D",
        "xAxisGridStyle": "dotted", "yAxisGridStyle": "dotted",
        "xAxisGridColor": "#D4B896", "yAxisGridColor": "#D4B896",
    },
    "Neon Pulse": {
        "desc": "Cyberpunk & vivid",
        "dataColors": ["#FF2D78","#00F5FF","#7B2FBE","#39FF14","#FF6B00","#FFD700","#FF69B4","#00BFFF","#9400D3","#FF4500"],
        "foreground": "#F0F0FF", "foregroundNeutralSecondary": "#A0A0C0", "foregroundNeutralTertiary": "#606080",
        "background": "#0D0D1A", "backgroundLight": "#15152A", "backgroundNeutral": "#252540",
        "tableAccent": "#FF2D78", "good": "#39FF14", "neutral": "#FFD700", "bad": "#FF2D78",
        "maximum": "#00F5FF", "center": "#7B2FBE", "minimum": "#15152A",
        "fontFace": "Trebuchet MS", "fontSize": 10, "fontColor": "#F0F0FF",
        "xAxisGridStyle": "dashed", "yAxisGridStyle": "dashed",
        "xAxisGridColor": "#252540", "yAxisGridColor": "#252540",
    },
    "Sage & Stone": {
        "desc": "Earthy & minimal",
        "dataColors": ["#6B8F71","#AAC0AA","#D4C5A9","#8B7355","#C4A882","#E8DCC8","#4A7C59","#9CAF88","#BFA98A","#5C6B4A"],
        "foreground": "#2C2A1E", "foregroundNeutralSecondary": "#6B6050", "foregroundNeutralTertiary": "#A09080",
        "background": "#FAF8F2", "backgroundLight": "#F0EDE4", "backgroundNeutral": "#D8D0C0",
        "tableAccent": "#6B8F71", "good": "#4A7C59", "neutral": "#C4A882", "bad": "#A05050",
        "maximum": "#6B8F71", "center": "#D4C5A9", "minimum": "#F0EDE4",
        "fontFace": "Verdana", "fontSize": 10, "fontColor": "#2C2A1E",
        "xAxisGridStyle": "dotted", "yAxisGridStyle": "dotted",
        "xAxisGridColor": "#D8D0C0", "yAxisGridColor": "#D8D0C0",
    },
    "Royal Ink": {
        "desc": "Deep purple & gold",
        "dataColors": ["#6A0DAD","#9B59B6","#F1C40F","#2C3E50","#8E44AD","#D4AC0D","#1ABC9C","#E74C3C","#3498DB","#F39C12"],
        "foreground": "#1A0A2E", "foregroundNeutralSecondary": "#5A4A7A", "foregroundNeutralTertiary": "#9A8AB0",
        "background": "#FDFAFF", "backgroundLight": "#F5EEFF", "backgroundNeutral": "#D8C8EE",
        "tableAccent": "#6A0DAD", "good": "#1ABC9C", "neutral": "#F1C40F", "bad": "#E74C3C",
        "maximum": "#6A0DAD", "center": "#F1C40F", "minimum": "#F5EEFF",
        "fontFace": "Calibri", "fontSize": 10, "fontColor": "#1A0A2E",
        "xAxisGridStyle": "solid", "yAxisGridStyle": "solid",
        "xAxisGridColor": "#D8C8EE", "yAxisGridColor": "#D8C8EE",
    },
    "Arctic Frost": {
        "desc": "Crisp icy blues",
        "dataColors": ["#5B9BD5","#ED7D31","#A5A5A5","#FFC000","#4472C4","#70AD47","#255E91","#9E480E","#636363","#997300"],
        "foreground": "#1F2937", "foregroundNeutralSecondary": "#6B7280", "foregroundNeutralTertiary": "#D1D5DB",
        "background": "#F8FAFC", "backgroundLight": "#EDF2F7", "backgroundNeutral": "#CBD5E0",
        "tableAccent": "#5B9BD5", "good": "#70AD47", "neutral": "#FFC000", "bad": "#ED7D31",
        "maximum": "#4472C4", "center": "#FFC000", "minimum": "#EDF2F7",
        "fontFace": "Calibri", "fontSize": 10, "fontColor": "#1F2937",
        "xAxisGridStyle": "dotted", "yAxisGridStyle": "dotted",
        "xAxisGridColor": "#CBD5E0", "yAxisGridColor": "#CBD5E0",
    },
    "Obsidian Dark": {
        "desc": "Dark mode analytics",
        "dataColors": ["#5C6BC0","#26C6DA","#66BB6A","#FFA726","#EF5350","#AB47BC","#29B6F6","#D4E157","#FF7043","#8D6E63"],
        "foreground": "#E8EAED", "foregroundNeutralSecondary": "#9AA0A6", "foregroundNeutralTertiary": "#5F6368",
        "background": "#202124", "backgroundLight": "#292A2D", "backgroundNeutral": "#3C4043",
        "tableAccent": "#5C6BC0", "good": "#66BB6A", "neutral": "#FFA726", "bad": "#EF5350",
        "maximum": "#26C6DA", "center": "#FFA726", "minimum": "#292A2D",
        "fontFace": "Segoe UI", "fontSize": 10, "fontColor": "#E8EAED",
        "xAxisGridStyle": "dashed", "yAxisGridStyle": "dashed",
        "xAxisGridColor": "#3C4043", "yAxisGridColor": "#3C4043",
    },
}

# ── BASE THEME ────────────────────────────────────────────────────────────────
BASE_THEME = {
    "name": "CY26SU02",
    "dataColors": [
        "#118DFF","#12239E","#E66C37","#6B007B","#E044A7",
        "#744EC2","#D9B300","#D64550","#197278","#1AAB40"
    ],
    "foreground": "#252423",
    "foregroundNeutralSecondary": "#605E5C",
    "foregroundNeutralTertiary": "#B3B0AD",
    "background": "#FFFFFF",
    "backgroundLight": "#F3F2F1",
    "backgroundNeutral": "#C8C6C4",
    "tableAccent": "#118DFF",
    "good": "#1AAB40",
    "neutral": "#D9B300",
    "bad": "#D64554",
    "maximum": "#118DFF",
    "center": "#D9B300",
    "minimum": "#DEEFFF",
    "null": "#FF7F48",
    "hyperlink": "#0078d4",
    "visitedHyperlink": "#0078d4",
    "textClasses": {
        "callout": {"fontSize": 24, "fontFace": "DIN",               "color": "#252423"},
        "title":   {"fontSize": 12, "fontFace": "DIN",               "color": "#252423"},
        "header":  {"fontSize": 12, "fontFace": "Segoe UI Semibold", "color": "#252423"},
        "label":   {"fontSize": 10, "fontFace": "Segoe UI",          "color": "#252423"},
    },
    "visualStyles": {
        "*": {"*": {
            "*":            [{"wordWrap": True}],
            "line":         [{"transparency": 0}],
            "outline":      [{"transparency": 0}],
            "plotArea":     [{"transparency": 0}],
            "categoryAxis": [{"showAxisTitle": True, "gridlineStyle": "dotted", "concatenateLabels": False}],
            "valueAxis":    [{"showAxisTitle": True, "gridlineStyle": "dotted"}],
            "title":        [{"titleWrap": True}],
            "lineStyles":   [{"strokeWidth": 3}],
            "background":   [{"show": True, "transparency": 0}],
            "border":       [{"width": 1}],
        }},
        "scatterChart":   {"*": {"bubbles":    [{"bubbleSize": -10}], "fillPoint": [{"show": True}]}},
        "lineChart":      {"*": {"general":    [{"responsive": True}]}},
        "pieChart":       {"*": {"legend":     [{"show": True, "position": "RightCenter"}], "labels": [{"labelStyle": "Data value, percent of total"}]}},
        "donutChart":     {"*": {"legend":     [{"show": True, "position": "RightCenter"}]}},
        "pivotTable":     {"*": {"rowHeaders": [{"showExpandCollapseButtons": True}]}},
        "columnChart":    {"*": {"general":    [{"responsive": True}], "legend": [{"showGradientLegend": True}]}},
        "barChart":       {"*": {"general":    [{"responsive": True}], "legend": [{"showGradientLegend": True}]}},
        "kpi":            {"*": {"trendline":  [{"transparency": 20}]}},
        "cardVisual":     {"*": {"layout":     [{"maxTiles": 3}]}},
        "slicer":         {"*": {"general":    [{"responsive": True}]}},
        "waterfallChart": {"*": {"general":    [{"responsive": True}]}},
        "ribbonChart":    {"*": {"general":    [{"responsive": True}]}},
        "areaChart":      {"*": {"general":    [{"responsive": True}]}},
    }
}

# ── VISUAL REGISTRY ───────────────────────────────────────────────────────────
VISUALS = [
    {"type": "columnChart",          "label": "Column Chart"},
    {"type": "clusteredColumnChart", "label": "Clustered Column Chart"},
    {"type": "barChart",             "label": "Bar Chart"},
    {"type": "clusteredBarChart",    "label": "Clustered Bar Chart"},
    {"type": "lineChart",            "label": "Line Chart"},
    {"type": "areaChart",            "label": "Area Chart"},
    {"type": "stackedAreaChart",     "label": "Stacked Area Chart"},
    {"type": "pieChart",             "label": "Pie Chart"},
    {"type": "donutChart",           "label": "Donut Chart"},
    {"type": "scatterChart",         "label": "Scatter Plot"},
    {"type": "kpi",                  "label": "KPI Card"},
    {"type": "cardVisual",           "label": "Card Visual"},
    {"type": "pivotTable",           "label": "Matrix / Table"},
    {"type": "slicer",               "label": "Slicer"},
    {"type": "waterfallChart",       "label": "Waterfall Chart"},
    {"type": "ribbonChart",          "label": "Ribbon Chart"},
    {"type": "treemap",              "label": "Treemap"},
    {"type": "funnel",               "label": "Funnel Chart"},
    {"type": "gauge",                "label": "Gauge"},
]
VISUAL_TYPES  = [v["type"]  for v in VISUALS]
VISUAL_LABELS = [v["label"] for v in VISUALS]
TYPE_TO_LABEL = {v["type"]: v["label"] for v in VISUALS}

FONT_OPTIONS     = ["Segoe UI","DIN","Segoe UI Semibold","Segoe UI Light","Calibri","Arial","Verdana","Trebuchet MS","Roboto","Source Sans Pro","IBM Plex Sans","DM Sans","Georgia"]
GRIDLINE_STYLES  = ["dotted","dashed","solid","none"]
LEGEND_POSITIONS = ["Bottom","Top","Right","Left","RightCenter","TopCenter","BottomCenter","TopLeft"]

# ── SESSION STATE ─────────────────────────────────────────────────────────────
def _init():
    if "theme"            not in st.session_state: st.session_state.theme = copy.deepcopy(BASE_THEME)
    if "vis_custom"       not in st.session_state: st.session_state.vis_custom = {}
    if "selected_type"    not in st.session_state: st.session_state.selected_type = "columnChart"
    if "selected_types"   not in st.session_state: st.session_state.selected_types = ["columnChart"]
    if "theme_name"       not in st.session_state: st.session_state.theme_name = BASE_THEME["name"]
    if "active_preset"    not in st.session_state: st.session_state.active_preset = "Microsoft Default"
    if "preset_applied"   not in st.session_state: st.session_state.preset_applied = None
    if "global_font_face" not in st.session_state: st.session_state.global_font_face = "Segoe UI"
    if "global_font_size" not in st.session_state: st.session_state.global_font_size = 10
    if "global_font_color" not in st.session_state: st.session_state.global_font_color = "#252423"
_init()

def apply_preset(preset_name):
    """Apply a preset theme — updates base theme fields and resets all per-visual customizations."""
    preset = PRESET_THEMES[preset_name]
    st.session_state.active_preset = preset_name
    st.session_state.theme_name = preset_name.replace(" ", "_")

    # Patch BASE_THEME-level keys in session theme
    t = st.session_state.theme
    t["dataColors"] = preset["dataColors"][:]
    t["foreground"] = preset["foreground"]
    t["foregroundNeutralSecondary"] = preset["foregroundNeutralSecondary"]
    t["foregroundNeutralTertiary"] = preset["foregroundNeutralTertiary"]
    t["background"] = preset["background"]
    t["backgroundLight"] = preset["backgroundLight"]
    t["backgroundNeutral"] = preset["backgroundNeutral"]
    t["tableAccent"] = preset["tableAccent"]
    t["good"] = preset["good"]
    t["neutral"] = preset["neutral"]
    t["bad"] = preset["bad"]
    t["maximum"] = preset["maximum"]
    t["center"] = preset["center"]
    t["minimum"] = preset["minimum"]

    # Update global font defaults from preset
    st.session_state.global_font_face  = preset.get("fontFace", "Segoe UI")
    st.session_state.global_font_size  = preset.get("fontSize", 10)
    st.session_state.global_font_color = preset.get("fontColor", preset["foreground"])

    # Reset all per-visual customizations so they pick up new palette
    st.session_state.vis_custom = {}

def get_vis_custom(vtype):
    if vtype not in st.session_state.vis_custom:
        t = st.session_state.theme
        preset = PRESET_THEMES.get(st.session_state.active_preset, {})
        dc = t.get("dataColors", BASE_THEME["dataColors"])
        ff  = st.session_state.global_font_face
        fc  = st.session_state.global_font_color
        xgs = preset.get("xAxisGridStyle", "dotted")
        ygs = preset.get("yAxisGridStyle", "dotted")
        xgc = preset.get("xAxisGridColor", t.get("foregroundNeutralTertiary","#B3B0AD"))
        ygc = preset.get("yAxisGridColor", t.get("foregroundNeutralTertiary","#B3B0AD"))
        st.session_state.vis_custom[vtype] = {
            "color0": dc[0], "color1": dc[1] if len(dc)>1 else "#12239E",
            "color2": dc[2] if len(dc)>2 else "#E66C37", "color3": dc[3] if len(dc)>3 else "#6B007B",
            "fontFace": ff, "fontSize": st.session_state.global_font_size, "fontColor": fc,
            "xAxisShow": True, "xAxisTitle": True, "xAxisTitleText": "Category",
            "xAxisGridStyle": xgs, "xAxisGridColor": xgc,
            "yAxisShow": True, "yAxisTitle": True, "yAxisTitleText": "Value",
            "yAxisGridStyle": ygs, "yAxisGridColor": ygc,
            "legendShow": True, "legendPosition": "Bottom", "legendFontSize": 9,
        }
    return st.session_state.vis_custom[vtype]

# ── SVG HELPERS ───────────────────────────────────────────────────────────────
def _title(label, fg, w):
    return f'<text x="{w//2}" y="18" text-anchor="middle" font-size="11" font-weight="600" fill="{fg}" font-family="Segoe UI,sans-serif">{label}</text>'

def _hgrids(p, pl, pb, w, h):
    if p["yAxisGridStyle"] == "none": return ""
    dash = {"dotted":"2,3","dashed":"6,3"}.get(p["yAxisGridStyle"],"")
    da   = f' stroke-dasharray="{dash}"' if dash else ""
    return "".join(
        f'<line x1="{pl}" y1="{h-pb-lv*(h-pb-28):.1f}" x2="{w-8}" y2="{h-pb-lv*(h-pb-28):.1f}" stroke="{p["yAxisGridColor"]}" stroke-width="0.6"{da}/>'
        for lv in [0.25,0.5,0.75,1.0]
    )

def _axes(p, w, h, pl, pb):
    ax = p["xAxisGridColor"]; fg = p["fontColor"]; out=""
    if p["xAxisShow"]:
        out += f'<line x1="{pl}" y1="{h-pb}" x2="{w-8}" y2="{h-pb}" stroke="{ax}" stroke-width="1"/>'
        if p["xAxisTitle"]:
            out += f'<text x="{(pl+w-8)//2}" y="{h-2}" text-anchor="middle" font-size="8.5" fill="{fg}" opacity=".5" font-family="Segoe UI,sans-serif">{p["xAxisTitleText"]}</text>'
    if p["yAxisShow"]:
        out += f'<line x1="{pl}" y1="26" x2="{pl}" y2="{h-pb}" stroke="{ax}" stroke-width="1"/>'
        if p["yAxisTitle"]:
            mid=(26+h-pb)//2
            out += f'<text x="9" y="{mid}" text-anchor="middle" font-size="8.5" fill="{fg}" opacity=".5" font-family="Segoe UI,sans-serif" transform="rotate(-90,9,{mid})">{p["yAxisTitleText"]}</text>'
    return out

def _legend(p, w, h, cols):
    if not p["legendShow"]: return ""
    pos   = p["legendPosition"]
    items = [("Series A",cols[0]),("Series B",cols[1]),("Series C",cols[2])]
    if pos in ("Bottom","BottomCenter"):
        return "".join(f'<rect x="{12+i*90}" y="{h-14}" width="8" height="8" rx="1" fill="{c}"/><text x="{23+i*90}" y="{h-7}" font-size="{p["legendFontSize"]}" fill="{p["fontColor"]}" opacity=".7">{l}</text>' for i,(l,c) in enumerate(items))
    elif pos in ("Top","TopCenter"):
        return "".join(f'<rect x="{12+i*90}" y="22" width="8" height="8" rx="1" fill="{c}"/><text x="{23+i*90}" y="29" font-size="{p["legendFontSize"]}" fill="{p["fontColor"]}" opacity=".7">{l}</text>' for i,(l,c) in enumerate(items))
    elif pos in ("Right","RightCenter"):
        return "".join(f'<rect x="{w-70}" y="{30+i*18}" width="8" height="8" rx="1" fill="{c}"/><text x="{w-59}" y="{37+i*18}" font-size="{p["legendFontSize"]}" fill="{p["fontColor"]}" opacity=".7">{l}</text>' for i,(l,c) in enumerate(items))
    return ""

# ── SVG RENDERERS ─────────────────────────────────────────────────────────────
def render_column(p, w=600, h=330, label="Column Chart"):
    bg = st.session_state.theme.get("background","#FFFFFF")
    vals=[0.55,0.82,0.42,0.95,0.68,0.50]; cats=["Jan","Feb","Mar","Apr","May","Jun"]
    cols=[p["color0"],p["color1"],p["color2"],p["color1"],p["color0"],p["color2"]]
    pl,pb=34,32; ch=h-pb-30; bw=(w-pl-16)/len(vals)*0.65; gap=(w-pl-16)/len(vals)
    bars="".join(
        f'<rect x="{pl+i*gap+gap*0.18:.1f}" y="{30+ch*(1-v):.1f}" width="{bw:.1f}" height="{ch*v:.1f}" rx="2" fill="{cols[i]}" opacity=".92"/>'
        f'<text x="{pl+i*gap+gap*0.18+bw/2:.1f}" y="{h-pb+11}" text-anchor="middle" font-size="8" fill="{p["fontColor"]}" opacity=".5" font-family="Segoe UI,sans-serif">{cats[i]}</text>'
        for i,v in enumerate(vals)
    )
    return f'<svg viewBox="0 0 {w} {h}" preserveAspectRatio="xMidYMid meet" xmlns="http://www.w3.org/2000/svg"><rect width="{w}" height="{h}" fill="{bg}" rx="6"/>{_title(label,p["fontColor"],w)}{_hgrids(p,pl,pb,w,h)}{bars}{_axes(p,w,h,pl,pb)}{_legend(p,w,h,[p["color0"],p["color1"],p["color2"]])}</svg>'

def render_bar(p, w=600, h=330, label="Bar Chart"):
    bg = st.session_state.theme.get("background","#FFFFFF")
    vals=[0.72,0.88,0.45,0.95,0.61]; cats=["Alpha","Beta","Gamma","Delta","Epsilon"]
    cols=[p["color0"],p["color1"],p["color2"],p["color3"],p["color0"]]
    pl,pr,pb=62,16,16; cw=w-pl-pr; bh=(h-40)/len(vals)*0.6; gap=(h-40)/len(vals)
    ax=p["xAxisGridColor"]
    bars="".join(
        f'<rect x="{pl}" y="{28+i*gap+gap*0.2:.1f}" width="{v*cw:.1f}" height="{bh:.1f}" rx="2" fill="{cols[i]}" opacity=".92"/>'
        f'<text x="{pl-5}" y="{28+i*gap+gap*0.2+bh/2+3:.1f}" text-anchor="end" font-size="8.5" fill="{p["fontColor"]}" opacity=".65" font-family="Segoe UI,sans-serif">{cats[i]}</text>'
        for i,v in enumerate(vals)
    )
    dash={"dotted":"2,3","dashed":"6,3"}.get(p["xAxisGridStyle"],""); da=f' stroke-dasharray="{dash}"' if dash else ""
    vg="" if p["xAxisGridStyle"]=="none" else "".join(
        f'<line x1="{pl+lv*cw:.1f}" y1="26" x2="{pl+lv*cw:.1f}" y2="{h-pb}" stroke="{ax}" stroke-width="0.6"{da}/>'
        for lv in [0.25,0.5,0.75,1.0]
    )
    pb=16
    return f'<svg viewBox="0 0 {w} {h}" preserveAspectRatio="xMidYMid meet" xmlns="http://www.w3.org/2000/svg"><rect width="{w}" height="{h}" fill="{bg}" rx="6"/>{_title(label,p["fontColor"],w)}{vg}{bars}<line x1="{pl}" y1="26" x2="{pl}" y2="{h-pb}" stroke="{ax}" stroke-width="1"/>{_legend(p,w,h,[p["color0"],p["color1"],p["color2"]])}</svg>'

def render_line(p, w=600, h=330, label="Line Chart", filled=False):
    bg = st.session_state.theme.get("background","#FFFFFF")
    pl,pb=34,32; ch=h-pb-30
    s1=[0.48,0.65,0.38,0.88,0.72,0.58,0.80,0.70]; s2=[0.30,0.52,0.60,0.42,0.75,0.50,0.65,0.45]
    xs=[pl+i*(w-pl-16)/7 for i in range(8)]
    paths=""
    for pts,col in [(s1,p["color0"]),(s2,p["color1"])]:
        ys=[30+ch*(1-v) for v in pts]; coords=list(zip(xs,ys))
        path="M"+" L".join(f"{x:.1f},{y:.1f}" for x,y in coords)
        if filled:
            paths+=f'<path d="{path} L{xs[-1]:.1f},{h-pb} L{xs[0]:.1f},{h-pb} Z" fill="{col}" opacity=".12"/>'
        paths+=f'<path d="{path}" fill="none" stroke="{col}" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/>'
        paths+="".join(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="3" fill="{col}" stroke="{bg}" stroke-width="1.5"/>' for x,y in coords)
    months=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug"]
    xlbls="".join(f'<text x="{x:.1f}" y="{h-pb+12}" text-anchor="middle" font-size="8" fill="{p["fontColor"]}" opacity=".5" font-family="Segoe UI,sans-serif">{m}</text>' for x,m in zip(xs,months))
    return f'<svg viewBox="0 0 {w} {h}" preserveAspectRatio="xMidYMid meet" xmlns="http://www.w3.org/2000/svg"><rect width="{w}" height="{h}" fill="{bg}" rx="6"/>{_title(label,p["fontColor"],w)}{_hgrids(p,pl,pb,w,h)}{paths}{xlbls}{_axes(p,w,h,pl,pb)}{_legend(p,w,h,[p["color0"],p["color1"],p["color2"]])}</svg>'

def render_pie(p, w=600, h=330, donut=False, label="Pie Chart"):
    bg=st.session_state.theme.get("background","#FFFFFF"); fg=p["fontColor"]
    slices=[0.35,0.25,0.20,0.12,0.08]
    dc=st.session_state.theme.get("dataColors",BASE_THEME["dataColors"])
    colors=[p["color0"],p["color1"],p["color2"],p["color3"],dc[4] if len(dc)>4 else "#E044A7"]
    cx,cy=w//2-50,h//2+10; r=min(cx-10,cy-30)-4; inner=r*0.52 if donut else 0
    angle=-math.pi/2; paths=[]
    for i,s in enumerate(slices):
        a1,a2=angle,angle+s*2*math.pi
        x1,y1=cx+r*math.cos(a1),cy+r*math.sin(a1)
        x2,y2=cx+r*math.cos(a2),cy+r*math.sin(a2)
        large=1 if s>.5 else 0
        if donut:
            ix1,iy1=cx+inner*math.cos(a1),cy+inner*math.sin(a1)
            ix2,iy2=cx+inner*math.cos(a2),cy+inner*math.sin(a2)
            d=f"M{x1:.1f},{y1:.1f} A{r},{r} 0 {large},1 {x2:.1f},{y2:.1f} L{ix2:.1f},{iy2:.1f} A{inner:.1f},{inner:.1f} 0 {large},0 {ix1:.1f},{iy1:.1f} Z"
        else:
            d=f"M{cx},{cy} L{x1:.1f},{y1:.1f} A{r},{r} 0 {large},1 {x2:.1f},{y2:.1f} Z"
        paths.append(f'<path d="{d}" fill="{colors[i]}" stroke="{bg}" stroke-width="2"/>')
        angle=a2
    cats=["Product A","Product B","Product C","Product D","Other"]
    leg=("".join(f'<rect x="{w-115}" y="{28+i*20}" width="9" height="9" rx="2" fill="{colors[i]}"/><text x="{w-102}" y="{37+i*20}" font-size="9" fill="{fg}" opacity=".75" font-family="Segoe UI,sans-serif">{cats[i]} {round(s*100)}%</text>' for i,s in enumerate(slices)) if p["legendShow"] else "")
    center=(f'<text x="{cx}" y="{cy+5}" text-anchor="middle" font-size="18" font-weight="700" fill="{p["color0"]}" font-family="Segoe UI,sans-serif">87%</text>' if donut else "")
    return f'<svg viewBox="0 0 {w} {h}" preserveAspectRatio="xMidYMid meet" xmlns="http://www.w3.org/2000/svg"><rect width="{w}" height="{h}" fill="{bg}" rx="6"/>{"".join(paths)}{center}{leg}{_title(label,fg,w)}</svg>'

def render_scatter(p, w=600, h=330, label="Scatter Plot"):
    bg=st.session_state.theme.get("background","#FFFFFF"); random.seed(7)
    pl,pb=36,32
    dots="".join(f'<circle cx="{pl+random.random()*(w-pl-20):.1f}" cy="{28+random.random()*(h-pb-30):.1f}" r="5.5" fill="{[p["color0"],p["color1"],p["color2"]][i%3]}" opacity=".72"/>' for i in range(28))
    return f'<svg viewBox="0 0 {w} {h}" preserveAspectRatio="xMidYMid meet" xmlns="http://www.w3.org/2000/svg"><rect width="{w}" height="{h}" fill="{bg}" rx="6"/>{_title(label,p["fontColor"],w)}{_hgrids(p,pl,pb,w,h)}{dots}{_axes(p,w,h,pl,pb)}{_legend(p,w,h,[p["color0"],p["color1"],p["color2"]])}</svg>'

def render_kpi(p, w=600, h=330, label="KPI Card"):
    bg=st.session_state.theme.get("background","#FFFFFF"); fg=p["fontColor"]
    trend=[(w*0.12,h*0.80),(w*0.27,h*0.67),(w*0.42,h*0.72),(w*0.58,h*0.57),(w*0.73,h*0.62),(w*0.88,h*0.50)]
    tpath="M"+" L".join(f"{x:.1f},{y:.1f}" for x,y in trend)
    tdots="".join(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="3" fill="{p["color0"]}"/>' for x,y in trend)
    return f'<svg viewBox="0 0 {w} {h}" preserveAspectRatio="xMidYMid meet" xmlns="http://www.w3.org/2000/svg"><rect width="{w}" height="{h}" fill="{bg}" rx="6"/>{_title(label,fg,w)}<text x="{w//2}" y="{h*0.43:.0f}" text-anchor="middle" font-size="54" font-weight="700" fill="{p["color0"]}" font-family="Segoe UI,sans-serif">87.4%</text><text x="{w//2}" y="{h*0.58:.0f}" text-anchor="middle" font-size="13" fill="{fg}" opacity=".4" font-family="Segoe UI,sans-serif">▲ +7.4 pp vs target 80%</text><path d="{tpath}" fill="none" stroke="{p["color0"]}" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/>{tdots}</svg>'

def render_card(p, w=600, h=330, label="Card Visual"):
    bg=st.session_state.theme.get("background","#FFFFFF"); bgl=st.session_state.theme.get("backgroundLight","#F3F2F1"); fg=p["fontColor"]
    cards=[("Total Revenue","$2.41M",p["color0"]),("Growth Rate","+14.2%",p["color1"]),("Churn Risk","3.8%",p["color2"])]
    cw=(w-32)//3
    tiles="".join(
        f'<rect x="{8+i*(cw+8)}" y="30" width="{cw}" height="{h-46}" rx="6" fill="{bgl}"/>'
        f'<text x="{8+i*(cw+8)+cw//2}" y="{30+(h-46)*0.45:.0f}" text-anchor="middle" font-size="22" font-weight="700" fill="{col}" font-family="Segoe UI,sans-serif">{val}</text>'
        f'<text x="{8+i*(cw+8)+cw//2}" y="{30+(h-46)*0.68:.0f}" text-anchor="middle" font-size="9" fill="{fg}" opacity=".45" font-family="Segoe UI,sans-serif">{lbl}</text>'
        for i,(lbl,val,col) in enumerate(cards)
    )
    return f'<svg viewBox="0 0 {w} {h}" preserveAspectRatio="xMidYMid meet" xmlns="http://www.w3.org/2000/svg"><rect width="{w}" height="{h}" fill="{bg}" rx="6"/>{_title(label,fg,w)}{tiles}</svg>'

def render_matrix(p, w=600, h=330, label="Matrix / Table"):
    bg=st.session_state.theme.get("background","#FFFFFF"); bgl=st.session_state.theme.get("backgroundLight","#F3F2F1"); fg=p["fontColor"]; ac=p["color0"]
    rows=["Alpha","Beta","Gamma","Delta","Epsilon","Zeta"]; rh=(h-50)/(len(rows)+1)
    cx=[16,w*0.35,w*0.52,w*0.68,w*0.84]; headers=["Category","Q1","Q2","Q3","Δ%"]
    hdr=f'<rect x="8" y="28" width="{w-16}" height="{rh:.1f}" fill="{ac}" opacity=".18" rx="3"/>'
    hdr+="".join(f'<text x="{x:.1f}" y="{28+rh*0.7:.1f}" font-size="9" font-weight="600" fill="{ac}" font-family="Segoe UI Semibold,sans-serif">{h2}</text>' for x,h2 in zip(cx,headers))
    random.seed(5); data=""
    for i,row in enumerate(rows):
        y=28+(i+1)*rh
        if i%2==0: data+=f'<rect x="8" y="{y:.1f}" width="{w-16}" height="{rh:.1f}" fill="{bgl}" opacity=".4"/>'
        data+=f'<text x="{cx[0]}" y="{y+rh*0.7:.1f}" font-size="8.5" fill="{fg}" font-family="Segoe UI,sans-serif">{row}</text>'
        q=[random.randint(200,900) for _ in range(3)]; pct=round((q[2]-q[0])/q[0]*100,1)
        pc=st.session_state.theme.get("good","#1AAB40") if pct>=0 else st.session_state.theme.get("bad","#D64554")
        for j,v in enumerate(q): data+=f'<text x="{cx[j+1]:.1f}" y="{y+rh*0.7:.1f}" font-size="8.5" fill="{fg}" text-anchor="middle">{v}</text>'
        data+=f'<text x="{cx[4]:.1f}" y="{y+rh*0.7:.1f}" font-size="8" fill="{pc}" text-anchor="middle">{"+" if pct>=0 else ""}{pct}%</text>'
    sep="".join(f'<line x1="{x:.1f}" y1="28" x2="{x:.1f}" y2="{28+(len(rows)+1)*rh:.1f}" stroke="{bgl}" stroke-width="0.6"/>' for x in cx[1:])
    return f'<svg viewBox="0 0 {w} {h}" preserveAspectRatio="xMidYMid meet" xmlns="http://www.w3.org/2000/svg"><rect width="{w}" height="{h}" fill="{bg}" rx="6"/>{_title(label,fg,w)}{hdr}{data}{sep}</svg>'

def render_slicer(p, w=600, h=220, label="Slicer"):
    bg=st.session_state.theme.get("background","#FFFFFF"); bgl=st.session_state.theme.get("backgroundLight","#F3F2F1"); fg=p["fontColor"]
    items=["All","2021","2022","2023","2024","2025"]; iw=(w-24)/len(items)
    chips="".join(
        f'<rect x="{12+i*iw:.1f}" y="38" width="{iw-6:.1f}" height="44" rx="6" fill="{p["color0"] if i==0 else bgl}" opacity="{1 if i==0 else 0.8}"/>'
        f'<text x="{12+i*iw+(iw-6)/2:.1f}" y="65" text-anchor="middle" font-size="10" fill="{"white" if i==0 else fg}" font-weight="{"600" if i==0 else "400"}" font-family="Segoe UI,sans-serif">{it}</text>'
        for i,it in enumerate(items)
    )
    return f'<svg viewBox="0 0 {w} {h}" preserveAspectRatio="xMidYMid meet" xmlns="http://www.w3.org/2000/svg"><rect width="{w}" height="{h}" fill="{bg}" rx="6"/>{_title(label,fg,w)}{chips}</svg>'

def render_waterfall(p, w=600, h=330, label="Waterfall Chart"):
    bg=st.session_state.theme.get("background","#FFFFFF"); fg=p["fontColor"]
    good=st.session_state.theme.get("good","#1AAB40"); bad=st.session_state.theme.get("bad","#D64554")
    pl,pb=34,36; ch=h-pb-30
    steps=[("Start",None,0.25),("Sales",True,0.52),("COGS",False,-0.18),("Returns",False,-0.08),("Other",True,0.12),("End",None,0)]
    bw=(w-pl-16)/len(steps)*0.6; gap=(w-pl-16)/len(steps); level=0.25; bars=""
    for i,(lbl,pos,delta) in enumerate(steps):
        x=pl+i*gap+gap*0.2
        if lbl in ("Start","End"):
            bh=level*ch; y=h-pb-bh
            bars+=f'<rect x="{x:.1f}" y="{y:.1f}" width="{bw:.1f}" height="{bh:.1f}" rx="2" fill="{p["color0"]}"/>'
        else:
            bh=abs(delta)*ch; col=good if pos else bad
            y=h-pb-(level+delta)*ch if pos else h-pb-level*ch
            bars+=f'<rect x="{x:.1f}" y="{y:.1f}" width="{bw:.1f}" height="{bh:.1f}" rx="2" fill="{col}"/>'
            level+=delta
        bars+=f'<text x="{x+bw/2:.1f}" y="{h-pb+14}" text-anchor="middle" font-size="8" fill="{fg}" opacity=".55" font-family="Segoe UI,sans-serif">{lbl}</text>'
    return f'<svg viewBox="0 0 {w} {h}" preserveAspectRatio="xMidYMid meet" xmlns="http://www.w3.org/2000/svg"><rect width="{w}" height="{h}" fill="{bg}" rx="6"/>{_title(label,fg,w)}{_hgrids(p,pl,pb,w,h)}{bars}<line x1="{pl}" y1="{h-pb}" x2="{w-8}" y2="{h-pb}" stroke="{p["xAxisGridColor"]}" stroke-width="1"/></svg>'

def render_treemap(p, w=600, h=330, label="Treemap"):
    bg=st.session_state.theme.get("background","#FFFFFF"); fg=p["fontColor"]
    cols=[p["color0"],p["color1"],p["color2"],p["color3"]]
    rects=[(8,28,w*0.55-4,h*0.55-4,cols[0],"Product A"),(8,28+h*0.55-4,w*0.35-4,h*0.45-6,cols[1],"Product B"),(w*0.35+4,28+h*0.55-4,w*0.65-12,h*0.25-4,cols[2],"C"),(w*0.35+4,28+h*0.80-8,w*0.65-12,h*0.20-10,cols[3],"D"),(w*0.55+4,28,w*0.45-12,h*0.55-4,cols[1],"E")]
    tiles="".join(f'<rect x="{x:.1f}" y="{y:.1f}" width="{rw:.1f}" height="{rh:.1f}" fill="{col}" rx="2" opacity=".88"/><text x="{x+rw/2:.1f}" y="{y+rh/2:.1f}" text-anchor="middle" dominant-baseline="middle" font-size="{max(8,min(14,int(rw/6)))}" fill="white" font-weight="600" font-family="Segoe UI,sans-serif">{lbl}</text>' for x,y,rw,rh,col,lbl in rects)
    return f'<svg viewBox="0 0 {w} {h}" preserveAspectRatio="xMidYMid meet" xmlns="http://www.w3.org/2000/svg"><rect width="{w}" height="{h}" fill="{bg}" rx="6"/>{_title(label,fg,w)}{tiles}</svg>'

def render_funnel(p, w=600, h=330, label="Funnel Chart"):
    bg=st.session_state.theme.get("background","#FFFFFF"); fg=p["fontColor"]
    stages=[("Leads","1,200",1.0),("Prospects","860",0.72),("Qualified","540",0.45),("Proposals","310",0.26),("Won","140",0.12)]
    cols=[p["color0"],p["color1"],p["color2"],p["color3"],p["color1"]]
    cx=w//2; bh=(h-50)/len(stages); maxw=w*0.75
    bars="".join(f'<rect x="{cx-r*maxw/2:.1f}" y="{28+i*bh:.1f}" width="{r*maxw:.1f}" height="{bh-3:.1f}" rx="3" fill="{cols[i]}" opacity=".9"/><text x="{cx}" y="{28+i*bh+bh/2+1:.1f}" text-anchor="middle" dominant-baseline="middle" font-size="9.5" fill="white" font-weight="600" font-family="Segoe UI,sans-serif">{lbl}: {val}</text>' for i,(lbl,val,r) in enumerate(stages))
    return f'<svg viewBox="0 0 {w} {h}" preserveAspectRatio="xMidYMid meet" xmlns="http://www.w3.org/2000/svg"><rect width="{w}" height="{h}" fill="{bg}" rx="6"/>{_title(label,fg,w)}{bars}</svg>'

def render_gauge(p, w=600, h=330, label="Gauge"):
    bg=st.session_state.theme.get("background","#FFFFFF"); fg=p["fontColor"]; bgl=st.session_state.theme.get("backgroundNeutral","#C8C6C4")
    cx,cy,r=w//2,int(h*0.68),int(min(w,h)*0.38)
    pct=0.72; sweep=math.pi*pct
    track=f'<path d="M{cx-r},{cy} A{r},{r} 0 0,1 {cx+r},{cy}" fill="none" stroke="{bgl}" stroke-width="20" stroke-linecap="round"/>'
    ex=cx+r*math.cos(math.pi-sweep); ey=cy-r*math.sin(sweep)
    fill=f'<path d="M{cx-r},{cy} A{r},{r} 0 {"1" if sweep>math.pi/2 else "0"},1 {ex:.1f},{ey:.1f}" fill="none" stroke="{p["color0"]}" stroke-width="20" stroke-linecap="round"/>'
    na=math.pi-sweep; nx=cx+r*0.85*math.cos(na); ny=cy-r*0.85*math.sin(na)
    needle=f'<line x1="{cx}" y1="{cy}" x2="{nx:.1f}" y2="{ny:.1f}" stroke="{fg}" stroke-width="3" stroke-linecap="round" opacity=".55"/><circle cx="{cx}" cy="{cy}" r="7" fill="{fg}" opacity=".45"/>'
    return f'<svg viewBox="0 0 {w} {h}" preserveAspectRatio="xMidYMid meet" xmlns="http://www.w3.org/2000/svg"><rect width="{w}" height="{h}" fill="{bg}" rx="6"/>{_title(label,fg,w)}{track}{fill}{needle}<text x="{cx}" y="{cy-int(r*0.25)}" text-anchor="middle" font-size="36" font-weight="700" fill="{p["color0"]}" font-family="Segoe UI,sans-serif">72%</text><text x="{cx}" y="{cy-int(r*0.05)}" text-anchor="middle" font-size="10" fill="{fg}" opacity=".4" font-family="Segoe UI,sans-serif">of target 85%</text></svg>'

def render_visual(vtype, p, w=600, h=330):
    lbl=TYPE_TO_LABEL.get(vtype, vtype)
    fn = {
        "columnChart": render_column, "clusteredColumnChart": render_column,
        "barChart": render_bar,        "clusteredBarChart": render_bar,
        "lineChart": lambda p,w,h: render_line(p,w,h,lbl,False),
        "areaChart": lambda p,w,h: render_line(p,w,h,lbl,True),
        "stackedAreaChart": lambda p,w,h: render_line(p,w,h,lbl,True),
        "pieChart":    lambda p,w,h: render_pie(p,w,h,False,lbl),
        "donutChart":  lambda p,w,h: render_pie(p,w,h,True,lbl),
        "scatterChart": render_scatter, "kpi": render_kpi, "cardVisual": render_card,
        "pivotTable": render_matrix,    "slicer": render_slicer,
        "waterfallChart": render_waterfall, "ribbonChart": render_column,
        "treemap": render_treemap,      "funnel": render_funnel, "gauge": render_gauge,
    }.get(vtype, render_column)
    return fn(p, w, h)

# ── JSON EXPORT ───────────────────────────────────────────────────────────────
def build_export():
    out = copy.deepcopy(BASE_THEME)
    out["name"] = st.session_state.theme_name
    t = st.session_state.theme
    # Copy top-level theme fields
    for key in ["dataColors","foreground","foregroundNeutralSecondary","foregroundNeutralTertiary",
                "background","backgroundLight","backgroundNeutral","tableAccent",
                "good","neutral","bad","maximum","center","minimum"]:
        if key in t:
            out[key] = t[key]

    # ── GLOBAL FONT — write to textClasses and visualStyles["*"]["*"]
    gff  = st.session_state.global_font_face
    gfs  = st.session_state.global_font_size
    gfc  = st.session_state.global_font_color
    out["textClasses"] = {
        "callout": {"fontSize": max(gfs + 14, 18), "fontFace": gff, "color": gfc},
        "title":   {"fontSize": max(gfs + 2, 10),  "fontFace": gff, "color": gfc},
        "header":  {"fontSize": max(gfs + 2, 10),  "fontFace": gff, "color": gfc},
        "label":   {"fontSize": gfs,                "fontFace": gff, "color": gfc},
    }
    # Inject into global visual wildcard so every visual inherits it
    global_star = out["visualStyles"]["*"]["*"]
    global_star["labels"] = [{"fontSize": gfs, "fontFamily": gff, "color": {"solid": {"color": gfc}}}]
    global_star["title"]  = [{"titleWrap": True, "fontFamily": gff, "fontSize": gfs, "fontColor": {"solid": {"color": gfc}}}]

    # Per-visual overrides
    for vtype, p in st.session_state.vis_custom.items():
        if vtype not in out["visualStyles"]:
            out["visualStyles"][vtype] = {"*": {}}
        vs = out["visualStyles"][vtype]["*"]
        vs["dataPoint"]    = [{"fill": {"solid": {"color": p[f"color{i}"]}} } for i in range(4)]
        vs["labels"]       = [{"show": True, "fontSize": p["fontSize"], "fontFamily": p["fontFace"], "color": {"solid": {"color": p["fontColor"]}}}]
        vs["title"]        = [{"show": True, "fontFamily": p["fontFace"], "fontSize": p["fontSize"], "fontColor": {"solid": {"color": p["fontColor"]}}}]
        vs["categoryAxis"] = [{"show": p["xAxisShow"], "showAxisTitle": p["xAxisTitle"], "title": p["xAxisTitleText"], "gridlineStyle": p["xAxisGridStyle"], "gridlineColor": {"solid": {"color": p["xAxisGridColor"]}}}]
        vs["valueAxis"]    = [{"show": p["yAxisShow"], "showAxisTitle": p["yAxisTitle"], "title": p["yAxisTitleText"], "gridlineStyle": p["yAxisGridStyle"], "gridlineColor": {"solid": {"color": p["yAxisGridColor"]}}}]
        vs["legend"]       = [{"show": p["legendShow"], "position": p["legendPosition"], "fontSize": p["legendFontSize"]}]
    return out

# ══════════════════════════════════════════════════════════════════════════════
# RENDER UI
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="topbar"><div class="topbar-logo">Power BI <em>Theme</em> Editor</div><div class="topbar-tag">Visual Designer</div></div>', unsafe_allow_html=True)

left, right = st.columns([1, 2.1], gap="small")

# ── LEFT PANEL ────────────────────────────────────────────────────────────────
with left:
    st.markdown('<div style="padding:16px 12px 0">', unsafe_allow_html=True)

    # ── PRESET THEMES SECTION ─────────────────────────────────────────────────
    st.markdown('<p style="font-size:9px;letter-spacing:.15em;text-transform:uppercase;color:#2a3555;margin-bottom:8px">Preset Themes</p>', unsafe_allow_html=True)

    # Build the theme grid as an HTML block with clickable buttons via Streamlit columns
    preset_names = list(PRESET_THEMES.keys())
    rows = [preset_names[i:i+2] for i in range(0, len(preset_names), 2)]

    for row_presets in rows:
        cols_ui = st.columns(2, gap="small")
        for col_ui, pname in zip(cols_ui, row_presets):
            with col_ui:
                preset_data = PRESET_THEMES[pname]
                dc = preset_data["dataColors"]
                is_active = (st.session_state.active_preset == pname)

                # Build swatch HTML
                swatch_html = "".join(
                    f'<span style="flex:1;background:{c}"></span>' for c in dc[:6]
                )
                active_style = "border:1px solid #4f8ef7;box-shadow:0 0 0 1px #4f8ef7,0 4px 14px rgba(79,142,247,.2);" if is_active else "border:1px solid #111827;"
                active_name_color = "#c8d8f0" if is_active else "#8aaad8"

                st.markdown(f"""
                <div style="border-radius:8px;overflow:hidden;background:#070b15;{active_style}margin-bottom:2px">
                  <div style="display:flex;height:6px">
                    {swatch_html}
                  </div>
                  <div style="padding:6px 8px">
                    <div style="font-family:'IBM Plex Mono',monospace;font-size:9px;color:{active_name_color};letter-spacing:.04em">{pname}</div>
                    <div style="font-family:'IBM Plex Mono',monospace;font-size:8px;color:#2d3a58;margin-top:2px">{preset_data['desc']}</div>
                  </div>
                </div>
                """, unsafe_allow_html=True)

                btn_label = "✓ Applied" if is_active else "Apply"
                if st.button(btn_label, key=f"preset_{pname}", use_container_width=True):
                    apply_preset(pname)
                    st.rerun()

    st.divider()

    # ── GLOBAL FONT ───────────────────────────────────────────────────────────
    st.markdown('<p style="font-size:9px;letter-spacing:.15em;text-transform:uppercase;color:#2a3555;margin-bottom:6px">🌐  Global Font — Report Level</p>', unsafe_allow_html=True)
    st.markdown('<p style="font-size:8px;color:#1e2e50;margin-bottom:8px;font-family:IBM Plex Mono,monospace">Applies to all visuals via textClasses &amp; visualStyles[*]</p>', unsafe_allow_html=True)
    with st.expander("🔤  Global Font Settings", expanded=True):
        new_gff = st.selectbox("Font Family", FONT_OPTIONS,
            index=FONT_OPTIONS.index(st.session_state.global_font_face) if st.session_state.global_font_face in FONT_OPTIONS else 0,
            key="global_ff")
        new_gfs = st.slider("Font Size", 7, 20, st.session_state.global_font_size, key="global_fs")
        new_gfc = st.color_picker("Font Color", st.session_state.global_font_color, key="global_fc")
        # Detect changes and push to all existing vis_custom entries
        if new_gff != st.session_state.global_font_face or new_gfs != st.session_state.global_font_size or new_gfc != st.session_state.global_font_color:
            st.session_state.global_font_face  = new_gff
            st.session_state.global_font_size  = new_gfs
            st.session_state.global_font_color = new_gfc
            # Propagate to all already-initialized vis_custom entries
            for vtype in st.session_state.vis_custom:
                st.session_state.vis_custom[vtype]["fontFace"]  = new_gff
                st.session_state.vis_custom[vtype]["fontSize"]  = new_gfs
                st.session_state.vis_custom[vtype]["fontColor"] = new_gfc

    st.divider()

    # ── THEME NAME ────────────────────────────────────────────────────────────
    st.markdown('<p style="font-size:9px;letter-spacing:.15em;text-transform:uppercase;color:#2a3555;margin-bottom:5px">Theme Name</p>', unsafe_allow_html=True)
    st.session_state.theme_name = st.text_input("_tn", value=st.session_state.theme_name, label_visibility="collapsed", key="tname_inp")

    # ── MULTI-SELECT VISUAL ───────────────────────────────────────────────────
    st.markdown('<p style="font-size:9px;letter-spacing:.15em;text-transform:uppercase;color:#2a3555;margin:14px 0 5px">Select Visuals for Preview</p>', unsafe_allow_html=True)
    st.markdown('<p style="font-size:8px;color:#1e2e50;margin-bottom:6px;font-family:IBM Plex Mono,monospace">Select one or more — all will render in the preview pane</p>', unsafe_allow_html=True)
    sel_labels_multi = st.multiselect(
        "_vsm", VISUAL_LABELS,
        default=[TYPE_TO_LABEL.get(t2, "Column Chart") for t2 in st.session_state.selected_types if t2 in TYPE_TO_LABEL.values() or TYPE_TO_LABEL.get(t2)],
        label_visibility="collapsed", key="vis_multidd"
    )
    # Fallback: keep at least one selected
    if not sel_labels_multi:
        sel_labels_multi = [TYPE_TO_LABEL.get(st.session_state.selected_types[0], "Column Chart")]
    sel_types_multi = [VISUAL_TYPES[VISUAL_LABELS.index(lbl)] for lbl in sel_labels_multi]
    st.session_state.selected_types = sel_types_multi

    # Active visual for customization panel = first selected
    sel_type = sel_types_multi[0]
    sel_label = TYPE_TO_LABEL.get(sel_type, sel_type)
    st.session_state.selected_type = sel_type
    p = get_vis_custom(sel_type)

    st.divider()
    if len(sel_types_multi) > 1:
        st.markdown(f'<p style="font-size:8px;color:#2a4a7a;margin-bottom:6px;font-family:IBM Plex Mono,monospace">⚙ Customizing: <span style="color:#6a9adf">{sel_label}</span> (first selected)</p>', unsafe_allow_html=True)

    # 1. DATA COLORS
    with st.expander("🎨  Data Colors", expanded=True):
        ca, cb = st.columns(2)
        with ca:
            p["color0"] = st.color_picker("Series 1", p["color0"], key=f"c0_{sel_type}")
            p["color2"] = st.color_picker("Series 3", p["color2"], key=f"c2_{sel_type}")
        with cb:
            p["color1"] = st.color_picker("Series 2", p["color1"], key=f"c1_{sel_type}")
            p["color3"] = st.color_picker("Series 4", p["color3"], key=f"c3_{sel_type}")

    # 2. FONT
    with st.expander("🔤  Font", expanded=True):
        p["fontFace"]  = st.selectbox("Family", FONT_OPTIONS, index=FONT_OPTIONS.index(p["fontFace"]) if p["fontFace"] in FONT_OPTIONS else 0, key=f"ff_{sel_type}")
        p["fontSize"]  = st.slider("Size", 7, 20, p["fontSize"], key=f"fs_{sel_type}")
        p["fontColor"] = st.color_picker("Color", p["fontColor"], key=f"fc_{sel_type}")

    # 3. X-AXIS
    no_x = {"kpi","cardVisual","pieChart","donutChart","slicer","treemap","funnel","gauge"}
    if sel_type not in no_x:
        with st.expander("↔  X-Axis"):
            p["xAxisShow"]      = st.toggle("Show Axis",  p["xAxisShow"],  key=f"xas_{sel_type}")
            p["xAxisTitle"]     = st.toggle("Show Title", p["xAxisTitle"], key=f"xat_{sel_type}")
            p["xAxisTitleText"] = st.text_input("Title Text", p["xAxisTitleText"], key=f"xatt_{sel_type}")
            p["xAxisGridStyle"] = st.selectbox("Grid Style", GRIDLINE_STYLES, index=GRIDLINE_STYLES.index(p["xAxisGridStyle"]), key=f"xags_{sel_type}")
            p["xAxisGridColor"] = st.color_picker("Grid Color", p["xAxisGridColor"], key=f"xagc_{sel_type}")

    # 4. Y-AXIS
    no_y = {"kpi","cardVisual","pieChart","donutChart","slicer","treemap","funnel","gauge","pivotTable"}
    if sel_type not in no_y:
        with st.expander("↕  Y-Axis"):
            p["yAxisShow"]      = st.toggle("Show Axis",  p["yAxisShow"],  key=f"yas_{sel_type}")
            p["yAxisTitle"]     = st.toggle("Show Title", p["yAxisTitle"], key=f"yat_{sel_type}")
            p["yAxisTitleText"] = st.text_input("Title Text", p["yAxisTitleText"], key=f"yatt_{sel_type}")
            p["yAxisGridStyle"] = st.selectbox("Grid Style", GRIDLINE_STYLES, index=GRIDLINE_STYLES.index(p["yAxisGridStyle"]), key=f"yags_{sel_type}")
            p["yAxisGridColor"] = st.color_picker("Grid Color", p["yAxisGridColor"], key=f"yagc_{sel_type}")

    # 5. LEGEND
    no_leg = {"kpi","cardVisual","slicer","gauge","funnel"}
    if sel_type not in no_leg:
        with st.expander("📌  Legend"):
            p["legendShow"]     = st.toggle("Show Legend", p["legendShow"], key=f"ls_{sel_type}")
            p["legendPosition"] = st.selectbox("Position", LEGEND_POSITIONS, index=LEGEND_POSITIONS.index(p["legendPosition"]) if p["legendPosition"] in LEGEND_POSITIONS else 0, key=f"lp_{sel_type}")
            p["legendFontSize"] = st.slider("Font Size", 7, 16, p["legendFontSize"], key=f"lfs_{sel_type}")

    st.divider()

    export_data = build_export()
    export_json = json.dumps(export_data, indent=2)
    fname = f"{st.session_state.theme_name.replace(' ','_')}.json"
    st.download_button(f"⬇  Download {fname}", data=export_json, file_name=fname, mime="application/json", use_container_width=True)
    st.markdown(f'<p style="font-size:9px;color:#2a3555;text-align:center;margin-top:6px">{len(st.session_state.vis_custom)} visual(s) customized · {len(export_data["visualStyles"])} style blocks</p>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

# ── RIGHT PANEL ───────────────────────────────────────────────────────────────
with right:
    import streamlit.components.v1 as components

    # All SVGs use viewBox and scale to fill their container — use a consistent large internal size
    VIS_W, VIS_H = 600, 330

    vis_cards_html = ""
    for vt in sel_types_multi:
        vp = get_vis_custom(vt)
        svg_item = render_visual(vt, vp, w=VIS_W, h=VIS_H)
        vt_label = TYPE_TO_LABEL.get(vt, vt)
        vis_cards_html += f'''
        <div class="vis-card-wrap">
          <div class="vis-label-small">{vt_label}</div>
          <div class="vis-card">{svg_item}</div>
        </div>'''

    # JSON snippet — global font block + first selected visual block
    export_data = build_export()
    global_snippet = {
        "textClasses": export_data.get("textClasses", {}),
        "visualStyles_global": export_data["visualStyles"]["*"]["*"].get("labels", []),
    }
    vis_snippet = {sel_type: export_data["visualStyles"].get(sel_type, BASE_THEME["visualStyles"].get(sel_type, {}))}
    combined_snippet = {"_globalFont": global_snippet, **vis_snippet}
    json_txt = json.dumps(combined_snippet, indent=2)

    dc_list  = st.session_state.theme.get("dataColors", BASE_THEME["dataColors"])
    swatches = "".join(
        f'<span title="{c}" style="display:inline-block;width:24px;height:24px;border-radius:5px;background:{c};margin:3px;border:1px solid rgba(0,0,0,.15)"></span>'
        for c in dc_list[:10]
    )

    active_preset_name = st.session_state.active_preset
    active_preset_desc = PRESET_THEMES[active_preset_name]["desc"]

    # Global font badge values
    gff_disp = st.session_state.global_font_face
    gfs_disp = st.session_state.global_font_size
    gfc_disp = st.session_state.global_font_color

    n_visuals = len(sel_types_multi)
    grid_cols = "1fr" if n_visuals == 1 else ("1fr 1fr" if n_visuals <= 4 else "1fr 1fr 1fr")
    preview_label_txt = f"Live Preview — {', '.join(TYPE_TO_LABEL.get(vt,vt) for vt in sel_types_multi)}" if n_visuals <= 3 else f"Live Preview — {n_visuals} visuals selected"

    preview_html = f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@300;400;500&display=swap" rel="stylesheet">
<style>
  * {{ box-sizing:border-box; margin:0; padding:0; }}
  body {{
    background: #0a0d16;
    font-family: 'IBM Plex Mono', monospace;
    padding: 24px;
  }}
  .section-label {{
    font-size: 9px;
    letter-spacing: .18em;
    text-transform: uppercase;
    color: #2a3555;
    margin-bottom: 12px;
    display: flex;
    align-items: center;
    gap: 10px;
  }}
  .section-label::after {{
    content: '';
    flex: 1;
    height: 1px;
    background: #111827;
  }}
  .vis-grid {{
    display: grid;
    grid-template-columns: {grid_cols};
    gap: 16px;
    margin-bottom: 24px;
  }}
  .vis-card-wrap {{ display: flex; flex-direction: column; gap: 6px; }}
  .vis-label-small {{
    font-size: 8.5px;
    letter-spacing: .12em;
    text-transform: uppercase;
    color: #2d4070;
  }}
  .vis-card {{
    background: #fff;
    border-radius: 10px;
    overflow: hidden;
    box-shadow: 0 4px 24px rgba(0,0,0,.45), 0 0 0 1px rgba(255,255,255,.06);
    width: 100%;
    line-height: 0;
  }}
  .vis-card svg {{
    display: block;
    width: 100%;
    height: auto;
    aspect-ratio: 600 / 330;
  }}
  .json-block {{
    background: #060810;
    border: 1px solid #111827;
    border-radius: 8px;
    padding: 16px;
    font-size: 11px;
    line-height: 1.8;
    color: #5a7aaa;
    overflow-x: auto;
    white-space: pre;
    max-height: 260px;
    overflow-y: auto;
    font-family: 'IBM Plex Mono', monospace;
    margin-bottom: 20px;
  }}
  .palette-row {{
    display: flex;
    flex-wrap: wrap;
    gap: 4px;
    padding: 12px;
    background: #060810;
    border: 1px solid #111827;
    border-radius: 8px;
    margin-bottom: 16px;
  }}
  .badge-row {{
    display: flex;
    flex-wrap: wrap;
    gap: 8px;
    margin-bottom: 16px;
    align-items: center;
  }}
  .preset-badge {{
    display: inline-flex;
    align-items: center;
    gap: 8px;
    background: #070e1f;
    border: 1px solid #1a2e55;
    border-radius: 20px;
    padding: 5px 12px;
  }}
  .font-badge {{
    display: inline-flex;
    align-items: center;
    gap: 8px;
    background: #07100f;
    border: 1px solid #1a3530;
    border-radius: 20px;
    padding: 5px 12px;
  }}
  .preset-dot {{ width:8px; height:8px; border-radius:50%; }}
  .preset-badge-name {{ font-size: 10px; color: #5a80c0; letter-spacing:.05em; }}
  .preset-badge-desc {{ font-size: 9px; color: #2a3555; }}
  .font-badge-name {{ font-size: 10px; color: #5aaa80; letter-spacing:.05em; }}
  .font-badge-meta {{ font-size: 9px; color: #1a3530; }}
  ::-webkit-scrollbar {{ width:4px; height:4px; }}
  ::-webkit-scrollbar-track {{ background:transparent; }}
  ::-webkit-scrollbar-thumb {{ background:#1e2d50; border-radius:4px; }}
</style>
</head>
<body>
  <div class="section-label">Active Preset &amp; Global Font</div>
  <div class="badge-row">
    <div class="preset-badge">
      <div class="preset-dot" style="background:{dc_list[0] if dc_list else '#118DFF'}"></div>
      <span class="preset-badge-name">{active_preset_name}</span>
      <span class="preset-badge-desc">— {active_preset_desc}</span>
    </div>
    <div class="font-badge">
      <div class="preset-dot" style="background:{gfc_disp}"></div>
      <span class="font-badge-name">{gff_disp}</span>
      <span class="font-badge-meta">{gfs_disp}px · {gfc_disp}</span>
    </div>
  </div>

  <div class="section-label">{preview_label_txt}</div>
  <div class="vis-grid">
    {vis_cards_html}
  </div>

  <div class="section-label">Theme JSON — Global Font + Active Visual Block</div>
  <div class="json-block">{json_txt}</div>

  <div class="section-label">Default Data Palette</div>
  <div class="palette-row">{swatches}</div>
</body>
</html>
"""
    # Dynamic height: taller when more visuals
    preview_height = 960 if n_visuals <= 2 else (1100 if n_visuals <= 4 else 1300)
    components.html(preview_html, height=preview_height, scrolling=True)