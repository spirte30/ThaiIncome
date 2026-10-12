"""Thai Household Income Dashboard (Streamlit + Plotly).

Data: data.xlsx (sheet "data") from the National Statistical Office of Thailand.
Each row is one income category x one socio-economic sub-group x one province x one
year. Income categories are nested (sub-totals and leaves live side by side), so
values are always shown for ONE selected category and never summed across
categories.
"""

import re
from html import escape
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio
import streamlit as st

from icons import icon

st.set_page_config(
    page_title="Thai Household Income Dashboard",
    page_icon=":material/dashboard:",
    layout="wide",
    initial_sidebar_state="auto",
)

# ─────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────
DATA_FILE = Path(__file__).parent / "data.xlsx"
SHEET_NAME = "data"
REQUIRED_COLUMNS = [
    "year", "province", "source_income1", "source_income2",
    "source_income3", "soc_eco_class1", "soc_eco_class2",
    "value", "unit", "attribute", "source",
]
TEXT_COLUMNS = [
    "province", "source_income1", "source_income2", "source_income3",
    "soc_eco_class1", "soc_eco_class2", "unit", "attribute", "source",
]

TOTAL_INCOME = "รายได้ทั้งสิ้นต่อเดือน"
# Income hierarchy (source_income3 -> parent), verified against data.xlsx:
# children sum to their parent within ±2 baht (rounding). Order = display order.
INCOME_PARENT = {
    TOTAL_INCOME: None,
    "รายได้ประจำ": TOTAL_INCOME,
    "รายได้ที่เป็นตัวเงิน": "รายได้ประจำ",
    "ค่าจ้างและเงินเดือน": "รายได้ที่เป็นตัวเงิน",
    "กำไรสุทธิจากการทำธุรกิจ": "รายได้ที่เป็นตัวเงิน",
    "กำไรสุทธิจากการทำการเกษตร": "รายได้ที่เป็นตัวเงิน",
    "เงินที่ได้รับเป็นการช่วยเหลือ": "รายได้ที่เป็นตัวเงิน",
    "รายได้จากทรัพย์สิน": "รายได้ที่เป็นตัวเงิน",
    "รายได้ที่ไม่เป็นตัวเงิน": "รายได้ประจำ",
    "รายได้ไม่ประจำ (ที่เป็นตัวเงิน)": TOTAL_INCOME,
}
INCOME_LEAVES = [c for c in INCOME_PARENT if c not in INCOME_PARENT.values()]

ALL_PROVINCES = "ทุกจังหวัด"
ALL_GROUPS = "ทุกกลุ่ม"
UNIT_LABEL = "บาท/เดือน"
# Verified: table SFD_SPB0802_66 in this dataset matches data.xlsx row-for-row (15,400 rows).
SOURCE_URL = "https://data.go.th/dataset/os_08_00007"
SOURCE_TITLE = "รายได้เฉลี่ยต่อเดือนของครัวเรือน จำแนกตามแหล่งที่มาของรายได้ และสถานะทางเศรษฐสังคมของครัวเรือน"

THAI_COLS = {
    "year": "ปี (พ.ศ.)",
    "province": "จังหวัด",
    "source_income1": "หมวดรายได้ ระดับ 1",
    "source_income2": "หมวดรายได้ ระดับ 2",
    "source_income3": "หมวดรายได้ ระดับ 3",
    "soc_eco_class1": "กลุ่มเศรษฐกิจและสังคม (หลัก)",
    "soc_eco_class2": "กลุ่มเศรษฐกิจและสังคม (ย่อย)",
    "value": f"มูลค่า ({UNIT_LABEL})",
    "unit": "หน่วย",
    "attribute": "หมายเหตุค่าข้อมูล (attribute)",
    "source": "แหล่งข้อมูล (source)",
}
# Short labels for the attribute flags found in the data (original text kept in details).
ATTRIBUTE_SHORT = {
    "ศูนย์ยกกำลังดับเบิลยู": "0 จากการสำรวจตัวอย่าง",
    "ลบศูนย์ยกกำลังอาร์ดี": "ปัดเศษเป็น 0 (ติดลบ)",
    "ศูนย์ยกกำลังอาร์ดี": "ปัดเศษเป็น 0",
}

# Theme (single source: the CSS custom properties below are generated from these)
NAVY = "#0B1F3A"
TEAL = "#0F9D8A"        # bars, icons, borders (non-text: 3.4:1 on white)
TEAL_INK = "#0B7A6B"    # teal for text/links/widgets (5.2:1 on white, WCAG AA)
ACCENT = "#C27C0E"      # the one accent: warnings and "look here" marks
ACCENT_INK = "#A15C07"  # accent for text (5.2:1 on white)
POSITIVE = "#0F7B5F"    # status: increase
NEGATIVE = "#B42318"    # status: decrease / error
TEXT = "#0F1B2D"
MUTED = "#5B6B7F"       # 5.1:1 on the page background
BG = "#F5F7FA"
SURFACE = "#FFFFFF"
BORDER = "#DCE3EC"
GRID = "#E6EBF1"
# Charts use teal only: one base, darker/lighter steps, and a pale tint for faded marks.
TEAL_DARK = "#0B6157"
TEAL_FADED = "#CDE8E3"
TEAL_SHADES = ["#0B6157", "#0F9D8A", "#4DBDAC", "#93D6CB"]  # dark -> light, one per group
# Validated categorical palette (light surface): lightness, chroma, CVD, contrast pass.
YEAR_COLORS = ["#7FCDC0", TEAL_DARK]  # earlier year light, latest year dark
FONT_FAMILY = "Sarabun, 'Noto Sans Thai', sans-serif"

# Summary statistic when several rows (provinces x groups) collapse into one bar/KPI.
# Median is the default: one extreme sub-group can otherwise dominate a province mean.
STATS = {"median": "มัธยฐาน", "mean": "ค่าเฉลี่ย"}

STATE_DEFAULTS = {
    "f_stat": "median",
    "f_province": ALL_PROVINCES,
    "f_income": TOTAL_INCOME,
    "f_class1": ALL_GROUPS,
    "f_class2": ALL_GROUPS,
}

# ─────────────────────────────────────────────
# Styles: all CSS lives here. Our own classes use the tokens in :root; the page-frame
# block at the end is the only one that targets Streamlit containers.
# ─────────────────────────────────────────────
CSS = f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;500;600;700&display=swap');
:root {{
  --navy:{NAVY}; --teal:{TEAL}; --teal-ink:{TEAL_INK}; --teal-50:#E7F5F2;
  --accent:{ACCENT}; --accent-ink:{ACCENT_INK}; --pos:{POSITIVE}; --neg:{NEGATIVE};
  --text:{TEXT}; --text-2:#334155; --muted:{MUTED}; --bg:{BG}; --surface:{SURFACE};
  --border:{BORDER}; --border-soft:#E6EBF1;
  --on-navy:#FFFFFF; --on-navy-2:#C8D3E1; --on-navy-3:#A9B7C9; --on-navy-accent:#7FD9CB;

  --s1:4px; --s2:8px; --s3:12px; --s4:16px; --s5:24px; --s6:32px;
  --r-sm:8px; --r-md:12px; --r-lg:16px; --r-pill:999px;
  --shadow-sm:0 1px 2px rgba(11,31,58,.05);
  --shadow-md:0 2px 8px rgba(11,31,58,.07);

  --font:{FONT_FAMILY};
  --fs-xs:.75rem; --fs-sm:.8125rem; --fs-base:.9375rem; --fs-md:1rem;
  --fs-lg:1.25rem; --fs-xl:1.6rem; --fs-2xl:clamp(1.5rem, 2.6vw, 2.1rem);
}}
html, body, .stApp, .stMarkdown, button, input, select, textarea {{
  font-family:var(--font);
}}
.lc-icon {{ display:inline-block; vertical-align:-0.18em; flex-shrink:0; }}
.num {{ font-variant-numeric:tabular-nums; }}

/* Header */
.app-header {{
  background:var(--navy); color:var(--on-navy); border-radius:var(--r-lg);
  padding:var(--s5) var(--s6) var(--s4); margin:0 0 var(--s4); box-shadow:var(--shadow-md);
}}
.app-header .eyebrow {{
  display:flex; align-items:center; gap:var(--s2); color:var(--on-navy-accent);
  font-size:var(--fs-xs); font-weight:600; letter-spacing:.08em; text-transform:uppercase;
}}
.app-header h1 {{
  color:var(--on-navy); font-size:var(--fs-2xl); font-weight:700;
  margin:var(--s2) 0 var(--s1); padding:0; line-height:1.25;
}}
.app-header .lead {{
  color:var(--on-navy-2); font-size:var(--fs-md); margin:0 0 var(--s3); line-height:1.6;
}}
.app-header .meta {{
  display:flex; flex-wrap:wrap; gap:var(--s2) var(--s5); color:var(--on-navy-3);
  font-size:var(--fs-sm); border-top:1px solid rgba(255,255,255,.12); padding-top:var(--s3);
  font-variant-numeric:tabular-nums;
}}
.app-header .meta span {{ display:inline-flex; align-items:center; gap:6px; }}
.app-header .meta b {{ color:var(--on-navy); font-weight:600; }}
.app-header .meta a {{
  color:var(--on-navy-accent); text-decoration:underline; text-underline-offset:3px;
}}

/* Chips (filter summary, row counts) */
.chips {{
  display:flex; flex-wrap:wrap; gap:var(--s2); margin:0 0 var(--s4); align-items:center;
}}
.chips .lbl {{ color:var(--muted); font-size:var(--fs-sm); margin-right:var(--s1); }}
.chip .c-note {{ color:var(--muted); }}
.chip {{
  display:inline-flex; align-items:center; gap:6px; background:var(--surface);
  border:1px solid var(--border); border-radius:var(--r-pill); padding:var(--s1) var(--s3);
  font-size:var(--fs-sm); color:var(--text); box-shadow:var(--shadow-sm);
  font-variant-numeric:tabular-nums; max-width:100%;
}}

/* KPI cards: label -> value + unit -> sub. Rows share one subgrid so labels, values
   and subs line up across a row and every card in a row has the same height.
   Phones (content < 600px) get one column; otherwise all three sit in one row. */
.kpi-wrap {{ container-type:inline-size; margin:0 0 var(--s2); }}
.kpi-grid {{
  display:grid; gap:var(--s3); grid-template-columns:minmax(0, 1fr); grid-auto-rows:auto;
}}
.kpi {{
  display:grid; grid-row:span 3; grid-template-rows:subgrid; row-gap:6px;
  align-content:start; container-type:inline-size;
  background:var(--surface); border:1px solid var(--border); border-radius:var(--r-md);
  padding:var(--s4); min-width:0; box-shadow:var(--shadow-sm);
}}
@container (min-width: 600px) {{
  .kpi-grid {{ grid-template-columns:minmax(0, .8fr) minmax(0, 1.1fr) minmax(0, 1.1fr); }}
}}
.kpi .k-label {{
  display:flex; align-items:flex-start; gap:6px; color:var(--muted);
  font-size:var(--fs-sm); font-weight:500; line-height:1.4;
}}
.kpi .k-label .lc-icon {{ margin-top:2px; }}
.kpi .k-value {{
  display:flex; align-items:baseline; gap:6px; white-space:nowrap;
  color:var(--navy); font-weight:700; line-height:1.25;
  font-size:clamp(1.2rem, 13.5cqi, var(--fs-xl)); font-variant-numeric:tabular-nums;
}}
.kpi .k-value .unit {{ font-size:var(--fs-sm); font-weight:500; color:var(--muted); }}
.kpi .k-sub {{
  color:var(--muted); font-size:var(--fs-sm); line-height:1.5; overflow-wrap:anywhere;
  font-variant-numeric:tabular-nums;
}}
.kpi .k-sub .nw {{ white-space:nowrap; }}
.kpi.accent {{ border-top:3px solid var(--teal); padding-top:calc(var(--s4) - 2px); }}

/* Section headings: icon + title + one-line description */
.section {{
  display:flex; align-items:center; gap:var(--s3); margin:var(--s6) 0 var(--s1);
  padding-bottom:var(--s2); border-bottom:1px solid var(--border);
}}
.section .s-icon {{
  display:inline-flex; align-items:center; justify-content:center; width:32px; height:32px;
  border-radius:var(--r-sm); background:var(--teal-50); flex-shrink:0;
}}
.section h3 {{
  font-size:var(--fs-lg); font-weight:600; color:var(--navy); margin:0; padding:0;
  line-height:1.35;
}}
.section-desc {{
  color:var(--muted); font-size:var(--fs-base); margin:var(--s2) 0 var(--s3); line-height:1.6;
}}

.chart-title {{
  color:var(--navy); font-size:var(--fs-md); font-weight:600; line-height:1.45;
  margin:var(--s4) 0 0;
}}
.chart-sub {{ color:var(--muted); font-size:var(--fs-sm); line-height:1.5; margin:0 0 var(--s1); }}

/* Charts with a narrow-screen variant (labels above bars below 640px) */
.st-key-group_chart_narrow, .st-key-year_chart_narrow {{ display:none; }}
@media (max-width: 640px) {{
  .st-key-group_chart_wide, .st-key-year_chart_wide {{ display:none; }}
  .st-key-group_chart_narrow, .st-key-year_chart_narrow {{ display:block; }}
}}

/* Sidebar */
.side-title {{
  display:flex; align-items:center; gap:var(--s2); font-weight:600; color:var(--navy);
  font-size:var(--fs-md); margin:0 0 var(--s1);
}}
.side-group {{
  color:var(--muted); font-size:var(--fs-xs); font-weight:600; letter-spacing:.06em;
  text-transform:uppercase; margin:var(--s4) 0 var(--s1);
}}

/* Callouts */
.note {{
  display:flex; gap:var(--s3); align-items:flex-start; background:var(--surface);
  border:1px solid var(--border); border-left:3px solid var(--teal); border-radius:var(--r-sm);
  padding:var(--s3) var(--s4); color:var(--text-2); font-size:var(--fs-base); line-height:1.6;
  margin:var(--s2) 0 var(--s3); box-shadow:var(--shadow-sm);
}}
.note.warn {{ border-left-color:var(--accent); }}
.note.error {{ border-left-color:var(--neg); }}
.note .lc-icon {{ margin-top:3px; }}

/* Empty state */
.empty {{
  text-align:center; background:var(--surface); border:1px dashed var(--border);
  border-radius:var(--r-lg); padding:var(--s6) var(--s5); margin:var(--s5) 0;
}}
.empty .e-icon {{
  display:inline-flex; align-items:center; justify-content:center; width:48px; height:48px;
  border-radius:var(--r-pill); background:var(--teal-50); margin-bottom:var(--s3);
}}
.empty h4 {{ color:var(--navy); font-size:var(--fs-lg); font-weight:600; margin:0 0 var(--s2); padding:0; }}
.empty p {{ color:var(--text-2); font-size:var(--fs-base); margin:0 auto; max-width:46ch; line-height:1.6; }}
.empty ul {{
  display:inline-block; text-align:left; color:var(--text-2); font-size:var(--fs-base);
  margin:var(--s3) 0 0; padding-left:var(--s5); line-height:1.7;
}}

/* Footer */
.app-footer {{
  margin-top:var(--s6); padding:var(--s4); background:var(--surface);
  border:1px solid var(--border); border-radius:var(--r-md);
  color:var(--muted); font-size:var(--fs-sm); line-height:1.7;
  display:grid; gap:var(--s2);
}}
.app-footer .row {{ display:flex; gap:var(--s2); align-items:flex-start; }}
.app-footer .row .lc-icon {{ margin-top:4px; }}
.app-footer a {{ color:var(--teal-ink); overflow-wrap:anywhere; }}

/* Page frame: the only rules on Streamlit's own containers (stable data-testids).
   Caps line length on wide screens and tightens the default padding. */
[data-testid="stMainBlockContainer"] {{
  max-width:1200px; padding:4rem var(--s6) var(--s6);
}}
@media (max-width: 640px) {{
  [data-testid="stMainBlockContainer"] {{ padding:3.75rem var(--s4) var(--s5); }}
  .app-header {{ padding:var(--s5) var(--s4) var(--s4); }}
}}
</style>
"""


# ─────────────────────────────────────────────
# HTML helpers (all dynamic text is escaped)
# ─────────────────────────────────────────────
def html(markup: str) -> None:
    st.markdown(markup, unsafe_allow_html=True)


def note(text: str, kind: str = "info") -> None:
    """Callout with a Lucide icon. `text` must be trusted or pre-escaped."""
    name = {"info": "info", "warn": "triangle-alert", "error": "circle-alert",
            "ok": "circle-check"}[kind]
    color = {"info": TEAL_INK, "warn": ACCENT_INK, "error": NEGATIVE, "ok": TEAL_INK}[kind]
    html(f'<div class="note {kind}">{icon(name, 16, color)}<div>{text}</div></div>')


def section(icon_name: str, title: str, desc: str | None = None) -> None:
    html(
        f'<div class="section"><span class="s-icon">{icon(icon_name, 18, TEAL_INK)}</span>'
        f'<h3>{escape(title)}</h3></div>'
        + (f'<div class="section-desc">{desc}</div>' if desc else "")
    )


def fmt_baht(v: float) -> str:
    return f"{v:,.0f}"


# ─────────────────────────────────────────────
# Data: load, validate, clean
# ─────────────────────────────────────────────
def validate_schema(df: pd.DataFrame) -> None:
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"ไม่พบคอลัมน์ที่จำเป็น: {missing}")
    found = set(df["source_income3"].dropna().astype(str).str.strip())
    unknown = sorted(found - set(INCOME_PARENT))
    absent = sorted(set(INCOME_PARENT) - found)
    if unknown or absent:
        raise ValueError(
            "หมวดรายได้ใน source_income3 ไม่ตรงกับโครงสร้างที่ตรวจสอบไว้ "
            f"(ไม่รู้จัก: {unknown}, ไม่พบ: {absent})"
        )


def clean_data(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    df = df.copy()
    for col in TEXT_COLUMNS:
        df[col] = df[col].astype("string").str.strip()
    raw_value = df["value"]
    df["value"] = pd.to_numeric(raw_value, errors="coerce")
    df["year"] = pd.to_numeric(df["year"], errors="coerce")
    issues = {
        "non_numeric_value": int((df["value"].isna() & raw_value.notna()).sum()),
        "missing_value": int(raw_value.isna().sum()),
        "missing_year": int(df["year"].isna().sum()),
    }
    df = df.dropna(subset=["year"])
    df["year"] = df["year"].astype(int)
    return df, issues


def check_hierarchy(df: pd.DataFrame) -> int | None:
    """Max |sum(children) - parent| over all province/group/year cells (baht)."""
    keys = ["year", "province", "soc_eco_class2"]
    wide = df.pivot_table(index=keys, columns="source_income3", values="value", aggfunc="first")
    worst = 0.0
    for parent in set(p for p in INCOME_PARENT.values() if p):
        children = [c for c, p in INCOME_PARENT.items() if p == parent]
        diff = (wide[children].sum(axis=1, min_count=len(children)) - wide[parent]).abs().max()
        if pd.isna(diff):
            return None
        worst = max(worst, float(diff))
    return int(worst)


@st.cache_data(show_spinner="กำลังโหลดข้อมูล...")
def load_data(path: str) -> tuple[pd.DataFrame, dict]:
    if not Path(path).exists():
        raise FileNotFoundError(f"ไม่พบไฟล์ข้อมูล: {Path(path).name}")
    xl = pd.ExcelFile(path, engine="openpyxl")
    if SHEET_NAME not in xl.sheet_names:
        raise ValueError(f"ไม่พบ Sheet '{SHEET_NAME}' (พบ: {xl.sheet_names})")
    df = xl.parse(SHEET_NAME)
    validate_schema(df)
    df, issues = clean_data(df)
    key = ["year", "province", "source_income3", "soc_eco_class2"]
    issues.update(
        rows=len(df),
        columns=len(df.columns),
        duplicate_rows=int(df.duplicated().sum()),
        duplicate_keys=int(df.duplicated(key).sum()),
        negative_values=int((df["value"] < 0).sum()),
        zero_values=int((df["value"] == 0).sum()),
        attribute_counts=df["attribute"].value_counts().to_dict(),
        hierarchy_max_diff=check_hierarchy(df),
    )
    return df, issues


# ─────────────────────────────────────────────
# Filters
# ─────────────────────────────────────────────
def reset_filters(latest_year: int) -> None:
    st.session_state["f_year"] = latest_year
    for k, v in STATE_DEFAULTS.items():
        st.session_state[k] = v


def on_class1_change() -> None:
    st.session_state["f_class2"] = ALL_GROUPS


def income_label(name: str) -> str:
    depth, parent = 0, INCOME_PARENT[name]
    while parent:
        depth, parent = depth + 1, INCOME_PARENT[parent]
    suffix = "" if name in INCOME_LEAVES else "  (ยอดรวม)"
    return " " * depth + ("› " if depth else "") + name + suffix


def render_sidebar(df: pd.DataFrame) -> dict:
    years = sorted(df["year"].unique().tolist(), reverse=True)
    provinces = [ALL_PROVINCES] + sorted(df["province"].dropna().unique().tolist())
    class1_opts = [ALL_GROUPS] + sorted(df["soc_eco_class1"].dropna().unique().tolist())

    st.session_state.setdefault("f_year", years[0])
    for k, v in STATE_DEFAULTS.items():
        st.session_state.setdefault(k, v)

    with st.sidebar:
        html(f'<div class="side-title">{icon("sliders-horizontal", 18, NAVY)}ตัวกรองข้อมูล</div>')
        st.caption("ทุกตัวกรองมีผลกับ KPI กราฟ และตาราง")

        html('<div class="side-group">ข้อมูลที่แสดง</div>')
        year = st.selectbox("ปี (พ.ศ.)", years, key="f_year")
        province = st.selectbox("จังหวัด", provinces, key="f_province")
        income = st.selectbox(
            "หมวดรายได้", list(INCOME_PARENT), key="f_income", format_func=income_label,
            help="แสดงค่าของหมวดที่เลือกเพียงหมวดเดียว หมวดที่มีคำว่า (ยอดรวม) "
                 "รวมหมวดย่อยที่อยู่ใต้หมวดนั้นไว้แล้ว",
        )
        html('<div class="side-group">กลุ่มเศรษฐกิจและสังคม</div>')
        class1 = st.selectbox(
            "กลุ่มหลัก", class1_opts, key="f_class1",
            on_change=on_class1_change,
        )
        sub = df if class1 == ALL_GROUPS else df[df["soc_eco_class1"] == class1]
        class2_opts = [ALL_GROUPS] + sorted(sub["soc_eco_class2"].dropna().unique().tolist())
        if st.session_state["f_class2"] not in class2_opts:
            st.session_state["f_class2"] = ALL_GROUPS
        class2 = st.selectbox("กลุ่มย่อย", class2_opts, key="f_class2")

        html('<div class="side-group">วิธีคำนวณ</div>')
        stat = st.radio(
            "ค่าสรุปเมื่อมีหลายรายการ", list(STATS), key="f_stat", format_func=STATS.get,
            horizontal=True,
            help="มัธยฐานไม่ไวต่อกลุ่มที่มีค่าสูงหรือต่ำผิดปกติ ค่าเฉลี่ยไวต่อค่าโดด "
                 "ทั้งสองแบบไม่ถ่วงน้ำหนักด้วยจำนวนครัวเรือน เพราะไฟล์ไม่มีข้อมูลนี้",
        )

        st.button("รีเซ็ตตัวกรอง", on_click=reset_filters, args=(years[0],),
                  width="stretch", type="secondary", icon=":material/restart_alt:")

        st.divider()
        html(f'<div class="side-title">{icon("info", 16, MUTED)}วิธีอ่านค่า</div>')
        st.caption(
            "ข้อมูลมีทั้งหมวดยอดรวมและหมวดย่อยอยู่ในไฟล์เดียวกัน แดชบอร์ดจึงแสดงค่าของ"
            "หมวดที่เลือกเท่านั้น ไม่นำหลายหมวดมารวมกัน เมื่อผลลัพธ์มีหลายรายการ "
            "ค่าที่แสดงคือมัธยฐานหรือค่าเฉลี่ยของรายการตามที่เลือก ไม่ใช่ค่าของประชากรทั้งจังหวัด "
            "เพราะไฟล์ไม่มีจำนวนครัวเรือนสำหรับถ่วงน้ำหนัก"
        )

    return dict(year=year, province=province, income=income, class1=class1, class2=class2,
                stat=stat)


def filter_data(df: pd.DataFrame, f: dict, *, skip: tuple[str, ...] = ()) -> pd.DataFrame:
    """Apply filters; names in `skip` are not applied (for comparison charts)."""
    mask = pd.Series(True, index=df.index)
    if "year" not in skip:
        mask &= df["year"] == f["year"]
    if "province" not in skip and f["province"] != ALL_PROVINCES:
        mask &= df["province"] == f["province"]
    if "income" not in skip:
        mask &= df["source_income3"] == f["income"]
    if "class1" not in skip and f["class1"] != ALL_GROUPS:
        mask &= df["soc_eco_class1"] == f["class1"]
    if "class2" not in skip and f["class2"] != ALL_GROUPS:
        mask &= df["soc_eco_class2"] == f["class2"]
    return df[mask]


# ─────────────────────────────────────────────
# KPIs
# ─────────────────────────────────────────────
def compute_kpis(dff: pd.DataFrame, f: dict, years: list[int]) -> dict:
    vals = dff["value"].dropna()
    k = dict(
        year=f["year"],
        other_years=[y for y in years if y != f["year"]],
        n_provinces=dff["province"].nunique(),
        n_groups=dff["soc_eco_class2"].nunique(),
        n_records=len(dff),
        value=None, max_value=None, max_where=None,
    )
    if not vals.empty:
        k["value"] = float(vals.iloc[0]) if len(vals) == 1 else float(vals.agg(f["stat"]))
        k["mean"], k["median"] = float(vals.mean()), float(vals.median())
        top = dff.loc[vals.idxmax()]
        k["max_value"] = float(top["value"])
        k["max_where"] = f'{top["province"]} · {top["soc_eco_class2"]}'
    return k


def kpi_card(icon_name: str, label: str, value: str, sub: str, *, unit: str | None = None,
             accent: bool = False) -> str:
    """One KPI card. `value` and `sub` must be trusted or pre-escaped."""
    classes = "kpi" + (" accent" if accent else "")
    unit_html = f'<span class="unit">{escape(unit)}</span>' if unit else ""
    return (
        f'<div class="{classes}">'
        f'<div class="k-label">{icon(icon_name, 15, MUTED)}<span>{escape(label)}</span></div>'
        f'<div class="k-value"><span>{value}</span>{unit_html}</div>'
        f'<div class="k-sub">{sub}</div></div>'
    )


def render_kpis(k: dict, f: dict) -> None:
    """Three cards; the selected year and province are in the filter chips above."""
    if k["value"] is None:
        val = dict(value="—", sub="ไม่มีค่าตัวเลข")
    elif k["n_records"] == 1:
        val = dict(value=fmt_baht(k["value"]), unit=UNIT_LABEL, sub="ค่าจริงของรายการที่เลือก")
    else:
        other_stat = "mean" if f["stat"] == "median" else "median"
        val = dict(value=fmt_baht(k["value"]), unit=UNIT_LABEL,
                   sub=(f'<span class="nw">{STATS[f["stat"]]}ของ {k["n_records"]:,} รายการ</span> · '
                        f'<span class="nw">{STATS[other_stat]} {fmt_baht(k[other_stat])}</span>'))

    if k["max_value"] is not None:
        top = dict(value=fmt_baht(k["max_value"]), unit=UNIT_LABEL)
    else:
        top = dict(value="—")
    cards = [
        kpi_card("list-checks", "รายการที่ผ่านตัวกรอง", f'{k["n_records"]:,}',
                 f'1 รายการ = จังหวัด × กลุ่มย่อย ({k["n_groups"]} กลุ่มย่อย)'),
        kpi_card("wallet", "รายได้หมวดที่เลือก", val["value"],
                 f'{escape(f["income"])}<br>{val["sub"]}', unit=val.get("unit"), accent=True),
        kpi_card("trophy", "ค่าสูงสุดในผลลัพธ์", top["value"],
                 escape(k["max_where"]) if k["max_where"] else "—", unit=top.get("unit")),
    ]
    html(f'<div class="kpi-wrap"><div class="kpi-grid">{"".join(cards)}</div></div>')


# ─────────────────────────────────────────────
# Charts
# ─────────────────────────────────────────────
PLOTLY_CONFIG = {"displaylogo": False, "displayModeBar": "hover",
                 "modeBarButtonsToRemove": ["lasso2d", "select2d"]}

# One template for every chart: font, surfaces, hairline grid, tight margins, hover.
pio.templates["thai_income"] = go.layout.Template(
    layout=dict(
        font=dict(family=FONT_FAMILY, color=TEXT, size=13),
        paper_bgcolor=SURFACE,
        plot_bgcolor=SURFACE,
        colorway=[TEAL] + TEAL_SHADES,
        margin=dict(l=8, r=16, t=16, b=8),
        separators=".,",
        hoverlabel=dict(bgcolor=SURFACE, bordercolor=BORDER, align="left",
                        font=dict(family=FONT_FAMILY, color=TEXT, size=13)),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0,
                    title_text="", font=dict(size=12, color=TEXT)),
        bargap=0.28,
        xaxis=dict(gridcolor=GRID, gridwidth=1, zerolinecolor="#AEB9C7", linecolor=GRID,
                   tickformat="~s", ticks="", automargin=True,
                   title=dict(font=dict(size=12, color=MUTED), standoff=8),
                   tickfont=dict(color=MUTED, size=12)),
        yaxis=dict(gridcolor=GRID, zerolinecolor="#AEB9C7", linecolor=GRID, ticks="",
                   automargin=True, title=dict(font=dict(size=12, color=MUTED)),
                   tickfont=dict(color=TEXT, size=12)),
    )
)


def style_fig(fig, height: int, *, legend: bool = False):
    # Plotly Express and Streamlit's frontend both write explicit layout values (margin,
    # legend title, background, font) that outrank a template, so the template's layout
    # is applied explicitly as well. The template stays the single source of truth.
    fig.update_layout(template="thai_income")
    fig.update_layout(pio.templates["thai_income"].layout.to_plotly_json())
    fig.update_layout(height=height, showlegend=legend, legend_title_text="")
    # Value at the end of every bar, thousands-separated like the hover.
    fig.update_traces(selector=dict(type="bar"), texttemplate="%{x:,.0f}",
                      textposition="outside", cliponaxis=False,
                      textfont=dict(family=FONT_FAMILY, size=11, color=TEXT))
    # Headroom so the outside labels fit even on a phone-width plot.
    xs = [v for t in fig.data if t.type == "bar" for v in t.x if pd.notna(v)]
    if xs:
        lo, hi = min(min(xs), 0), max(max(xs), 0)
        fig.update_xaxes(range=[lo * 1.18, hi * 1.18 if hi else 1])
    return fig


def show_chart(fig, narrow_fig=None, key: str | None = None) -> None:
    """Render the chart title as HTML (wraps on narrow screens) and the figure below it.

    theme=None keeps Streamlit's chart theme from overriding the shared template.
    With `narrow_fig`, both figures are rendered and CSS shows the narrow one below
    640px (the server cannot know the viewport width).
    """
    title = fig.layout.title.text
    if title:
        html(f'<div class="chart-title">{escape(title)}</div>')
        fig.update_layout(title_text="")
    # The value-axis title (measure + unit) can't wrap inside Plotly, so it moves here too.
    measure = fig.layout.xaxis.title.text
    if measure:
        html(f'<div class="chart-sub">{escape(measure)}</div>')
        fig.update_xaxes(title_text="")
    if narrow_fig is None:
        st.plotly_chart(fig, width="stretch", config=PLOTLY_CONFIG, theme=None)
        return
    narrow_fig.update_layout(title_text="")
    narrow_fig.update_xaxes(title_text="")
    with st.container(key=f"{key}_wide"):
        st.plotly_chart(fig, width="stretch", config=PLOTLY_CONFIG, theme=None)
    with st.container(key=f"{key}_narrow"):
        st.plotly_chart(narrow_fig, width="stretch", config=PLOTLY_CONFIG, theme=None)


def labels_above_bars(fig, categories: list[str], row_px: int = 56):
    """Copy of a horizontal bar chart with each category label above its bar.

    For narrow screens, where axis labels would leave the bars no room. Same traces
    and data; only the label placement and height change.
    """
    narrow = go.Figure(fig)
    narrow.update_yaxes(showticklabels=False, tickmode="auto")
    narrow.update_layout(
        meta="narrow-labels", bargap=0.6, height=row_px * len(categories) + 150,
        margin=dict(l=4, r=12, t=16, b=8),
        annotations=[
            dict(text=escape(c), x=0, xref="paper", xanchor="left", y=c, yref="y",
                 yanchor="bottom", yshift=row_px * (1 - 0.6) / 2 - 1, showarrow=False,
                 font=dict(size=12, color=TEXT), align="left")
            for c in categories
        ],
    )
    return narrow


def wrap_label(text: str, width: int = 24) -> str:
    """Split a long axis label onto two lines at the separator nearest its middle.

    Display only (used as ticktext); the category values themselves are untouched.
    """
    if len(text) <= width:
        return text
    mid = len(text) / 2
    cuts = [i + 1 for i, ch in enumerate(text) if ch in "_/ "]
    cuts += [m.end() for m in re.finditer("และ", text)]
    cuts = [c for c in cuts if 0 < c < len(text)]
    if not cuts:
        return text
    cut = min(cuts, key=lambda c: abs(c - mid))
    return text[:cut].rstrip("_ ") + "<br>" + text[cut:].lstrip("_ ")


def wrap_category_ticks(fig, categories: list[str]) -> None:
    fig.update_yaxes(tickmode="array", tickvals=categories,
                     ticktext=[wrap_label(c) for c in categories])


def province_chart(df: pd.DataFrame, f: dict, view: str):
    """Value of the selected category by province (all provinces; selection highlighted)."""
    data = filter_data(df, f, skip=("province",))
    agg = (data.groupby("province", as_index=False)
               .agg(median=("value", "median"), mean=("value", "mean"), n=("value", "size")))
    top = data.loc[data.groupby("province")["value"].idxmax(), ["province", "soc_eco_class2", "value"]]
    agg = agg.merge(top.rename(columns={"soc_eco_class2": "top_group", "value": "top_value"}),
                    on="province")
    agg["value"] = agg[f["stat"]]
    agg = agg.sort_values("value", ascending=False).reset_index(drop=True)
    agg["rank"] = agg.index + 1
    if view == "top":
        show = agg.head(15)
    elif view == "bottom":
        show = agg.tail(15)
    else:
        show = agg
    selected = f["province"]
    if selected != ALL_PROVINCES and selected not in show["province"].values:
        show = pd.concat([show, agg[agg["province"] == selected]])
    show = show.sort_values("value", ascending=True)
    show["สถานะ"] = (show["province"] == selected).map({True: "จังหวัดที่เลือก", False: "จังหวัดอื่น"})

    many = agg["n"].max() > 1
    measure = f'{STATS[f["stat"]]}ของกลุ่มย่อย' if many else "มูลค่า"
    highlight = selected != ALL_PROVINCES
    fig = px.bar(
        show, x="value", y="province", orientation="h",
        color="สถานะ" if highlight else None,
        color_discrete_map={"จังหวัดที่เลือก": TEAL_DARK, "จังหวัดอื่น": TEAL_FADED},
        custom_data=["rank", "n", "median", "mean", "top_group", "top_value"],
        title=f"{f['income']} รายจังหวัด · พ.ศ. {f['year']}",
        labels={"value": f"{measure} ({UNIT_LABEL})", "province": ""},
    )
    if not highlight:
        fig.update_traces(marker_color=TEAL)
    fig.update_traces(
        marker_line_width=0,
        hovertemplate=(
            "<b>%{y}</b><br>" + measure + ": %{x:,.0f} " + UNIT_LABEL
            + "<br>อันดับ %{customdata[0]} จาก " + str(len(agg)) + " จังหวัด"
            + ("<br><br>จาก %{customdata[1]} กลุ่มย่อย"
               "<br>มัธยฐาน %{customdata[2]:,.0f} · ค่าเฉลี่ย %{customdata[3]:,.0f}"
               "<br>สูงสุด: %{customdata[4]} (%{customdata[5]:,.0f})" if many else "")
            + "<extra></extra>"
        ),
    )
    fig.update_yaxes(categoryorder="array", categoryarray=show["province"].tolist())
    style_fig(fig, max(360, 26 * len(show) + 90), legend=highlight)
    return fig, many


def group_chart(df: pd.DataFrame, f: dict):
    """Value of the selected category by socio-economic sub-group (selection highlighted)."""
    data = filter_data(df, f, skip=("class1", "class2"))
    agg = (data.groupby(["soc_eco_class1", "soc_eco_class2"], as_index=False)
               .agg(median=("value", "median"), mean=("value", "mean"), n=("value", "size")))
    agg["value"] = agg[f["stat"]]
    agg = agg.sort_values("value", ascending=True)
    class1_order = sorted(df["soc_eco_class1"].dropna().unique())
    many = agg["n"].max() > 1
    measure = f'{STATS[f["stat"]]}ของจังหวัด' if many else "มูลค่า"
    fig = px.bar(
        agg, x="value", y="soc_eco_class2", orientation="h", color="soc_eco_class1",
        category_orders={"soc_eco_class1": class1_order},
        color_discrete_sequence=TEAL_SHADES,
        custom_data=["soc_eco_class1", "n", "median", "mean"],
        title=f"{f['income']} ตามกลุ่มเศรษฐกิจและสังคม",
        labels={"value": f"{measure} ({UNIT_LABEL})", "soc_eco_class2": ""},
    )
    fig.update_traces(
        marker_line_width=0,
        hovertemplate=(
            "<b>%{y}</b><br>กลุ่มหลัก: %{customdata[0]}<br>"
            + measure + ": %{x:,.0f} " + UNIT_LABEL
            + ("<br>จาก %{customdata[1]} จังหวัด"
               "<br>มัธยฐาน %{customdata[2]:,.0f} · ค่าเฉลี่ย %{customdata[3]:,.0f}" if many else "")
            + "<extra></extra>"
        ),
    )
    # With a group filter, the selected bars turn dark and the rest pale (nothing is
    # hidden); the shade legend no longer applies then, the main group is in the hover.
    selecting = f["class1"] != ALL_GROUPS or f["class2"] != ALL_GROUPS
    if selecting:
        for trace in fig.data:
            trace.marker.color = [
                TEAL_DARK if (f["class1"] in (ALL_GROUPS, trace.name)
                              and f["class2"] in (ALL_GROUPS, c2)) else TEAL_FADED
                for c2 in trace.y
            ]
    fig.update_yaxes(categoryorder="array", categoryarray=agg["soc_eco_class2"].tolist())
    wrap_category_ticks(fig, agg["soc_eco_class2"].tolist())
    style_fig(fig, 44 * len(agg) + 150, legend=not selecting)
    return fig, many


def composition_chart(df: pd.DataFrame, f: dict):
    """The 7 non-overlapping leaf categories (they add up to total income)."""
    data = filter_data(df, f, skip=("income",))
    data = data[data["source_income3"].isin(INCOME_LEAVES)]
    agg = (data.groupby("source_income3", as_index=False)
               .agg(value=("value", "mean"), n=("value", "size")))

    def under_selected(leaf: str) -> bool:
        node = leaf
        while node:
            if node == f["income"]:
                return True
            node = INCOME_PARENT[node]
        return False

    agg["อยู่ในหมวดที่เลือก"] = agg["source_income3"].map(
        lambda c: "อยู่ในหมวดที่เลือก" if under_selected(c) else "หมวดอื่น")
    agg["parent"] = agg["source_income3"].map(INCOME_PARENT)
    agg = agg.sort_values("value", ascending=True)
    many = agg["n"].max() > 1
    measure = "ค่าเฉลี่ย" if many else "มูลค่า"
    # Dark only when it singles something out; for total income every leaf is selected.
    partial = (agg["อยู่ในหมวดที่เลือก"] == "หมวดอื่น").any()
    fig = px.bar(
        agg, x="value", y="source_income3", orientation="h",
        color="อยู่ในหมวดที่เลือก",
        color_discrete_map={"อยู่ในหมวดที่เลือก": TEAL_DARK if partial else TEAL,
                            "หมวดอื่น": TEAL_FADED},
        custom_data=["parent", "n"],
        title="องค์ประกอบรายได้ 7 หมวดย่อย (ไม่ซ้อนกัน)",
        labels={"value": f"{measure} ({UNIT_LABEL})", "source_income3": ""},
    )
    fig.update_traces(
        marker_line_width=0,
        hovertemplate=(
            "<b>%{y}</b><br>อยู่ภายใต้: %{customdata[0]}<br>"
            + measure + ": %{x:,.0f} " + UNIT_LABEL
            + ("<br>เฉลี่ยจาก %{customdata[1]} รายการ" if many else "")
            + "<extra></extra>"
        ),
    )
    fig.update_yaxes(categoryorder="array", categoryarray=agg["source_income3"].tolist())
    # A one-item legend says nothing when every leaf falls under the selection.
    style_fig(fig, 40 * len(agg) + 130, legend=len(fig.data) > 1)
    return fig, many


def year_chart(df: pd.DataFrame, f: dict):
    """Selected province & category across years, by sub-group."""
    data = filter_data(df, f, skip=("year",))
    agg = (data.groupby(["year", "soc_eco_class2"], as_index=False)["value"].mean())
    if agg.empty:
        return None
    order = (agg[agg["year"] == agg["year"].max()]
             .sort_values("value", ascending=False)["soc_eco_class2"].tolist())
    agg["ปี"] = "พ.ศ. " + agg["year"].astype(str)
    fig = px.bar(
        agg, x="value", y="soc_eco_class2", color="ปี", orientation="h", barmode="group",
        category_orders={"ปี": [f"พ.ศ. {y}" for y in sorted(agg["year"].unique())],
                         "soc_eco_class2": order},
        color_discrete_sequence=YEAR_COLORS,
        title=f"{f['income']} · {f['province']} เปรียบเทียบระหว่างปี",
        labels={"value": f"มูลค่า ({UNIT_LABEL})", "soc_eco_class2": ""},
    )
    fig.update_traces(
        marker_line_width=0,
        hovertemplate="<b>%{y}</b><br>%{fullData.name}: %{x:,.0f} " + UNIT_LABEL + "<extra></extra>",
    )
    wrap_category_ticks(fig, order)
    style_fig(fig, max(320, 52 * agg["soc_eco_class2"].nunique() + 120), legend=True)
    fig.update_layout(bargap=0.22, bargroupgap=0.06)
    return fig


# ─────────────────────────────────────────────
# Table
# ─────────────────────────────────────────────
def short_attribute(text) -> str:
    if pd.isna(text):
        return ""
    for prefix, label in ATTRIBUTE_SHORT.items():
        if str(text).startswith(prefix):
            return label
    return str(text)


def table_column_config() -> dict:
    """Display-only settings: widths, number format, alignment. Values are untouched."""
    return {
        THAI_COLS["year"]: st.column_config.NumberColumn("ปี พ.ศ.", format="%d", width=64),
        THAI_COLS["province"]: st.column_config.TextColumn("จังหวัด", width=124),
        THAI_COLS["value"]: st.column_config.NumberColumn(
            f"มูลค่า ({UNIT_LABEL})", format="%,d", width=136,
            help="ค่าจริงจากไฟล์ต้นฉบับ หน่วยบาทต่อเดือน"),
        THAI_COLS["soc_eco_class2"]: st.column_config.TextColumn("กลุ่มย่อย", width=345),
        THAI_COLS["soc_eco_class1"]: st.column_config.TextColumn("กลุ่มหลัก", width=220),
        THAI_COLS["source_income3"]: st.column_config.TextColumn("หมวดรายได้", width=190),
    }


def render_table(dff: pd.DataFrame, f: dict) -> None:
    # Value right after the identifying columns so it stays visible on narrow screens.
    cols = ["year", "province", "value", "soc_eco_class2", "soc_eco_class1", "source_income3"]
    table = dff[cols].copy()
    table["flag"] = dff["attribute"].map(short_attribute)
    table = table.sort_values("value", ascending=False).rename(
        columns={**{c: THAI_COLS[c] for c in cols}, "flag": "หมายเหตุ"})
    html(
        f'<div class="chips"><span class="chip">{icon("table", 14, TEAL)}'
        f'พบ <b>{len(table):,}</b> รายการ</span>'
        f'<span class="lbl">เรียงจากมูลค่ามากไปน้อย คลิกหัวคอลัมน์เพื่อเรียงใหม่</span></div>'
    )
    st.dataframe(
        table, width="stretch", height=420, hide_index=True,
        column_config={**table_column_config(),
                       "หมายเหตุ": st.column_config.TextColumn(width=170)},
        column_order=[THAI_COLS[c] for c in cols] + ["หมายเหตุ"],
    )
    with st.expander("รายละเอียดทุกคอลัมน์ รวม source และ attribute ต้นฉบับ"):
        st.caption("แสดงข้อมูลตามไฟล์ต้นฉบับทุกคอลัมน์ ไม่มีการแก้ไขค่า")
        st.dataframe(
            dff[REQUIRED_COLUMNS].rename(columns=THAI_COLS), width="stretch",
            height=320, hide_index=True,
            column_config={
                THAI_COLS["year"]: st.column_config.NumberColumn(format="%d"),
                THAI_COLS["value"]: st.column_config.NumberColumn(format="%,d"),
            },
        )


def render_quality(issues: dict) -> None:
    with st.expander("ผลตรวจสอบคุณภาพข้อมูล (ทั้งไฟล์)"):
        diff = issues["hierarchy_max_diff"]
        rows = [
            ("จำนวนแถว × คอลัมน์", f'{issues["rows"]:,} × {issues["columns"]}'),
            ("แถวซ้ำทั้งแถว", f'{issues["duplicate_rows"]:,}'),
            ("แถวซ้ำตามคีย์ (ปี, จังหวัด, หมวดรายได้, กลุ่มย่อย)", f'{issues["duplicate_keys"]:,}'),
            ("ค่า value ที่ว่าง / แปลงเป็นตัวเลขไม่ได้",
             f'{issues["missing_value"]:,} / {issues["non_numeric_value"]:,}'),
            ("ค่าติดลบ (เป็นค่าจริงตามข้อมูล เช่น ขาดทุนสุทธิ)", f'{issues["negative_values"]:,}'),
            ("ค่าเท่ากับ 0", f'{issues["zero_values"]:,}'),
            ("ผลรวมหมวดย่อยเทียบยอดรวม (คลาดเคลื่อนสูงสุด)",
             f"{diff} บาท (จากการปัดเศษ)" if diff is not None else "ตรวจสอบไม่ได้"),
        ]
        rows += [(f"attribute: {a}", f"{n:,} แถว") for a, n in issues["attribute_counts"].items()]
        st.dataframe(pd.DataFrame(rows, columns=["รายการตรวจสอบ", "ผลลัพธ์"]),
                     width="stretch", hide_index=True)


# ─────────────────────────────────────────────
# Layout
# ─────────────────────────────────────────────
def render_header(df: pd.DataFrame | None) -> None:
    meta = [
        f'<span>{icon("database", 15, "#7FD9CB")}แหล่งข้อมูล <b>สำนักงานสถิติแห่งชาติ</b></span>',
        f'<span>{icon("link", 15, "#7FD9CB")}<a href="{SOURCE_URL}" target="_blank" '
        f'rel="noopener noreferrer">ชุดข้อมูลบน data.go.th</a></span>',
        f'<span>{icon("wallet", 15, "#7FD9CB")}หน่วย <b>{UNIT_LABEL}</b></span>',
    ]
    if df is not None:
        years = ", ".join(str(y) for y in sorted(df["year"].unique()))
        meta += [
            f'<span>{icon("calendar-days", 15, "#7FD9CB")}ปี พ.ศ. <b>{years}</b></span>',
            f'<span>{icon("map-pin", 15, "#7FD9CB")}<b>{df["province"].nunique()}</b> จังหวัด</span>',
        ]
    html(
        '<div class="app-header">'
        f'<div class="eyebrow">{icon("layout-dashboard", 16, "#7FD9CB")}Household Income Analytics</div>'
        '<h1>Thai Household Income Dashboard</h1>'
        '<p class="lead">แดชบอร์ดวิเคราะห์ข้อมูลรายได้ครัวเรือนไทย '
        'แยกตามจังหวัดและกลุ่มเศรษฐกิจและสังคม</p>'
        f'<div class="meta">{"".join(meta)}</div></div>'
    )


def render_filter_chips(f: dict, k: dict) -> None:
    """Active filters. Also carries what the former year/province KPI cards said."""
    group = f["class2"] if f["class2"] != ALL_GROUPS else f["class1"]
    other = ", ".join(str(y) for y in k["other_years"])
    year_note = f"ปีอื่นในชุดข้อมูล: {other}" if other else "ปีเดียวในชุดข้อมูล"
    prov_note = (f'{k["n_provinces"]} จังหวัดในผลลัพธ์' if f["province"] == ALL_PROVINCES
                 else "จังหวัดที่เลือก")
    chips = [
        ("calendar-days", f"พ.ศ. {f['year']}", year_note),
        ("map-pin", f["province"], prov_note),
        ("wallet", f["income"], None),
        ("users", group, None),
    ]
    html('<div class="chips"><span class="lbl">ตัวกรองปัจจุบัน</span>'
         + "".join(f'<span class="chip">{icon(n, 14, TEAL)}{escape(t)}'
                   + (f'<span class="c-note">· {escape(note)}</span>' if note else "")
                   + "</span>" for n, t, note in chips)
         + "</div>")


def render_empty_state(f: dict, latest_year: int) -> None:
    """Shown instead of KPIs-derived charts when the filters leave no rows."""
    tips = []
    if f["province"] != ALL_PROVINCES:
        tips.append(f'เปลี่ยนจังหวัดเป็น “{ALL_PROVINCES}” (ตอนนี้: {escape(f["province"])})')
    if f["class2"] != ALL_GROUPS:
        tips.append(f'เปลี่ยนกลุ่มย่อยเป็น “{ALL_GROUPS}” (ตอนนี้: {escape(f["class2"])})')
    elif f["class1"] != ALL_GROUPS:
        tips.append(f'เปลี่ยนกลุ่มหลักเป็น “{ALL_GROUPS}” (ตอนนี้: {escape(f["class1"])})')
    tips.append(f'ลองเลือกปีอื่น (ตอนนี้: พ.ศ. {f["year"]})')
    html(
        '<div class="empty">'
        f'<div class="e-icon">{icon("search-x", 24, TEAL_INK)}</div>'
        '<h4>ไม่พบข้อมูลตามตัวกรองที่เลือก</h4>'
        '<p>ชุดตัวกรองนี้ไม่มีรายการในไฟล์ข้อมูล ลองปรับตัวกรองในแถบด้านข้าง '
        'หรือกดรีเซ็ตเพื่อกลับไปค่าเริ่มต้น</p>'
        f'<ul>{"".join(f"<li>{t}</li>" for t in tips)}</ul></div>'
    )
    _, mid, _ = st.columns([1, 1, 1])
    with mid:
        st.button("รีเซ็ตตัวกรองทั้งหมด", on_click=reset_filters, args=(latest_year,),
                  key="reset_empty", width="stretch", type="primary",
                  icon=":material/restart_alt:")


def render_footer() -> None:
    html(
        '<div class="app-footer">'
        f'<div class="row">{icon("database", 15, MUTED)}<div>แหล่งข้อมูล: '
        f'สำนักงานสถิติแห่งชาติ · {SOURCE_TITLE} · '
        f'<a href="{SOURCE_URL}" target="_blank" rel="noopener noreferrer">{SOURCE_URL}</a>'
        '</div></div>'
        f'<div class="row">{icon("triangle-alert", 15, MUTED)}<div>ข้อจำกัด: ไฟล์มีทั้งหมวดยอดรวม'
        'และหมวดย่อย การรวมทุกแถวจะนับรายได้ซ้ำ มัธยฐานและค่าเฉลี่ยในแดชบอร์ด'
        'คำนวณจากรายการ (จังหวัด × กลุ่มย่อย) ไม่ใช่ค่าของประชากร เพราะไม่มีจำนวนครัวเรือนสำหรับถ่วงน้ำหนัก</div></div>'
        f'<div class="row">{icon("chart-column", 15, MUTED)}<div>พัฒนาด้วย Python · Streamlit · Plotly '
        '· ไอคอนจาก Lucide</div></div></div>'
    )


def main() -> None:
    html(CSS)
    try:
        df, issues = load_data(str(DATA_FILE))
    except FileNotFoundError as e:
        render_header(None)
        note(f"<b>ไม่พบไฟล์ข้อมูล</b><br>{escape(str(e))} กรุณาวาง data.xlsx "
             "ไว้ในโฟลเดอร์เดียวกับ app.py", "error")
        st.stop()
    except ValueError as e:
        render_header(None)
        note(f"<b>โครงสร้างข้อมูลไม่ถูกต้อง</b><br>{escape(str(e))}", "error")
        st.stop()

    render_header(df)
    f = render_sidebar(df)
    years = sorted(df["year"].unique().tolist())
    dff = filter_data(df, f)

    kpis = compute_kpis(dff, f, years)
    render_filter_chips(f, kpis)
    render_kpis(kpis, f)
    if issues["non_numeric_value"] or issues["missing_value"]:
        note(f'พบค่า value ที่ว่างหรือแปลงเป็นตัวเลขไม่ได้ '
             f'{issues["missing_value"] + issues["non_numeric_value"]:,} แถว '
             'แถวเหล่านี้ไม่ถูกนำไปคำนวณ', "warn")

    if dff.empty:
        render_empty_state(f, years[-1])
        render_footer()
        return

    # Section 1: provinces
    section("chart-column", "เปรียบเทียบระหว่างจังหวัด",
            "แสดงทุกจังหวัดเพื่อเทียบอันดับ จังหวัดที่เลือกในตัวกรองจะถูกเน้นสี")
    view = st.segmented_control(
        "มุมมอง", options=["top", "bottom", "all"], default="top", key="province_view",
        format_func={"top": "สูงสุด 15 จังหวัด", "bottom": "ต่ำสุด 15 จังหวัด",
                     "all": "ทั้งหมด"}.get,
        label_visibility="collapsed",
    ) or "top"
    fig, many = province_chart(df, f, view)
    show_chart(fig)
    if many:
        note(f'แต่ละจังหวัดมีหลายกลุ่มเศรษฐกิจและสังคม ค่าในกราฟคือ<b>{STATS[f["stat"]]}</b>'
             "ของกลุ่มย่อย (ไม่ถ่วงน้ำหนักด้วยจำนวนครัวเรือน) ไม่ใช่รายได้เฉลี่ยของประชากรทั้งจังหวัด "
             "เลือกกลุ่มย่อยในตัวกรองเพื่อเทียบจังหวัดแบบกลุ่มต่อกลุ่ม "
             "หรือชี้ที่แท่งกราฟเพื่อดูทั้งสองค่าและกลุ่มที่มีค่าสูงสุด")

    # Section 2: groups & composition
    section("chart-no-axes-combined", "กลุ่มเศรษฐกิจและสังคม และองค์ประกอบรายได้",
            "กราฟแรกแสดงทุกกลุ่มย่อยโดยจางกลุ่มที่อยู่นอกตัวกรอง กราฟที่สองแสดง 7 หมวดย่อยที่ไม่ซ้อนกัน "
            "ผลรวมของทั้ง 7 หมวดเท่ากับรายได้ทั้งสิ้น (คลาดเคลื่อนไม่เกิน ±2 บาทจากการปัดเศษ)")
    fig, many = group_chart(df, f)
    show_chart(fig, labels_above_bars(fig, list(fig.layout.yaxis.categoryarray)),
               key="group_chart")
    if many:
        st.caption(f'{STATS[f["stat"]]}ของทุกจังหวัดในกลุ่มนั้น (ไม่ถ่วงน้ำหนัก)')
    fig, many = composition_chart(df, f)
    show_chart(fig)
    st.caption("แถบสีเขียวคือหมวดย่อยที่อยู่ภายใต้หมวดรายได้ที่เลือก"
               + (" · ใช้ค่าเฉลี่ยเสมอ เพราะค่าเฉลี่ยของหมวดย่อยรวมกันได้เท่ากับยอดรวม"
                  " (มัธยฐานไม่มีคุณสมบัตินี้)" if many else ""))

    # Section 3: year comparison
    section("calendar-days", "เปรียบเทียบระหว่างปี",
            "ค่าของหมวดที่เลือกในจังหวัดที่เลือก แยกตามกลุ่มย่อย เทียบทุกปีในชุดข้อมูล")
    if f["province"] == ALL_PROVINCES:
        note("เลือกจังหวัดในแถบด้านข้างเพื่อดูการเปรียบเทียบระหว่างปีของจังหวัดนั้น")
    else:
        fig = year_chart(df, f)
        if fig is None:
            note("ไม่มีข้อมูลสำหรับการเปรียบเทียบระหว่างปี", "warn")
        else:
            show_chart(fig, labels_above_bars(fig, list(fig.layout.yaxis.tickvals), row_px=64),
                       key="year_chart")

    # Section 4: table
    section("table", "ตารางข้อมูล", "ข้อมูลที่ผ่านตัวกรองทั้งหมด ค่าเป็นค่าจริงจากไฟล์ต้นฉบับ")
    render_table(dff, f)
    render_quality(issues)
    render_footer()


main()
