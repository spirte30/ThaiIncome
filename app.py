"""Thai Household Income Dashboard (Streamlit + Plotly).

Data: data.xlsx (sheet "data") from the National Statistical Office of Thailand.
Each row is one income category x one socio-economic sub-group x one province x one
year. Income categories are nested (sub-totals and leaves live side by side), so
values are always shown for ONE selected category and never summed across
categories.
"""

from html import escape
from pathlib import Path

import pandas as pd
import plotly.express as px
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

# Theme
NAVY = "#0B1F3A"
TEAL = "#0F9D8A"
TEXT = "#0F1B2D"
MUTED = "#5B6B7F"
GRID = "#E6EBF1"
DIM = "#C9D3DF"
# Validated categorical palette (light surface): lightness, chroma, CVD, contrast pass.
CATEGORICAL = ["#2F6BB3", "#0F9D8A", "#C27C0E", "#8A5CC2"]
YEAR_COLORS = ["#2F6BB3", "#0F9D8A"]
FONT_FAMILY = "Sarabun, 'Noto Sans Thai', sans-serif"

STATE_DEFAULTS = {
    "f_province": ALL_PROVINCES,
    "f_income": TOTAL_INCOME,
    "f_class1": ALL_GROUPS,
    "f_class2": ALL_GROUPS,
}

# ─────────────────────────────────────────────
# Styles (only our own classes; no Streamlit-internal selectors)
# ─────────────────────────────────────────────
CSS = f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;500;600;700&display=swap');
html, body, .stApp, .stMarkdown, button, input, select, textarea {{
  font-family: {FONT_FAMILY};
}}
.lc-icon {{ display:inline-block; vertical-align:-0.18em; flex-shrink:0; }}

.app-header {{
  background:{NAVY}; color:#FFFFFF; border-radius:14px;
  padding:26px 30px 20px; margin:0 0 18px;
}}
.app-header .eyebrow {{
  display:flex; align-items:center; gap:8px; color:#7FD9CB;
  font-size:.8rem; font-weight:600; letter-spacing:.08em; text-transform:uppercase;
}}
.app-header h1 {{
  color:#FFFFFF; font-size:clamp(1.5rem, 2.6vw, 2.1rem); font-weight:700;
  margin:6px 0 4px; padding:0; line-height:1.25;
}}
.app-header .lead {{ color:#C8D3E1; font-size:1rem; margin:0 0 14px; line-height:1.6; }}
.app-header .meta {{
  display:flex; flex-wrap:wrap; gap:6px 22px; color:#A9B7C9; font-size:.85rem;
  border-top:1px solid rgba(255,255,255,.12); padding-top:12px;
}}
.app-header .meta span {{ display:inline-flex; align-items:center; gap:6px; }}
.app-header .meta b {{ color:#FFFFFF; font-weight:600; }}

.chips {{ display:flex; flex-wrap:wrap; gap:8px; margin:0 0 16px; align-items:center; }}
.chips .lbl {{ color:{MUTED}; font-size:.85rem; margin-right:2px; }}
.chip {{
  display:inline-flex; align-items:center; gap:6px; background:#FFFFFF;
  border:1px solid #DCE3EC; border-radius:999px; padding:4px 12px;
  font-size:.85rem; color:{TEXT};
}}

.kpi-grid {{
  display:grid; grid-template-columns:repeat(auto-fit, minmax(min(100%, 170px), 1fr));
  gap:14px; margin:0 0 8px;
}}
.kpi {{
  background:#FFFFFF; border:1px solid #E1E7EF; border-radius:12px;
  padding:16px 18px; min-width:0;
}}
.kpi .k-label {{
  display:flex; align-items:center; gap:7px; color:{MUTED};
  font-size:.86rem; font-weight:500;
}}
.kpi .k-value {{
  color:{NAVY}; font-size:1.55rem; font-weight:700; line-height:1.3; margin-top:6px;
  overflow-wrap:anywhere;
}}
.kpi .k-value small {{ font-size:.85rem; font-weight:500; color:{MUTED}; margin-left:4px; }}
.kpi .k-sub {{ color:{MUTED}; font-size:.8rem; margin-top:4px; line-height:1.45; }}
.kpi.accent {{ border-top:3px solid {TEAL}; }}

.section {{
  display:flex; align-items:center; gap:10px; margin:28px 0 2px;
  padding-bottom:8px; border-bottom:1px solid #DCE3EC;
}}
.section h3 {{ font-size:1.15rem; font-weight:600; color:{NAVY}; margin:0; padding:0; }}
.section-desc {{ color:{MUTED}; font-size:.9rem; margin:6px 0 10px; line-height:1.55; }}

.side-title {{
  display:flex; align-items:center; gap:8px; font-weight:600; color:{NAVY};
  font-size:1.02rem; margin:0 0 4px;
}}
.note {{
  display:flex; gap:10px; align-items:flex-start; background:#FFFFFF;
  border:1px solid #DCE3EC; border-left:3px solid {TEAL}; border-radius:8px;
  padding:10px 14px; color:#334155; font-size:.88rem; line-height:1.6; margin:6px 0 10px;
}}
.note.warn {{ border-left-color:#C27C0E; }}
.note.error {{ border-left-color:#C2410C; }}
.note .lc-icon {{ margin-top:3px; }}

.app-footer {{
  margin-top:34px; padding:16px 0 4px; border-top:1px solid #DCE3EC;
  color:{MUTED}; font-size:.84rem; line-height:1.7;
}}
.app-footer .row {{ display:flex; gap:8px; align-items:flex-start; }}
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
    color = {"info": TEAL, "warn": "#C27C0E", "error": "#C2410C", "ok": TEAL}[kind]
    html(f'<div class="note {kind}">{icon(name, 16, color)}<div>{text}</div></div>')


def section(icon_name: str, title: str, desc: str | None = None) -> None:
    html(
        f'<div class="section">{icon(icon_name, 20, TEAL)}<h3>{escape(title)}</h3></div>'
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

        year = st.selectbox("ปี (พ.ศ.)", years, key="f_year")
        province = st.selectbox("จังหวัด", provinces, key="f_province")
        income = st.selectbox(
            "หมวดรายได้", list(INCOME_PARENT), key="f_income", format_func=income_label,
            help="แสดงค่าของหมวดที่เลือกเพียงหมวดเดียว หมวดที่มีคำว่า (ยอดรวม) "
                 "รวมหมวดย่อยที่อยู่ใต้หมวดนั้นไว้แล้ว",
        )
        class1 = st.selectbox(
            "กลุ่มเศรษฐกิจและสังคม (หลัก)", class1_opts, key="f_class1",
            on_change=on_class1_change,
        )
        sub = df if class1 == ALL_GROUPS else df[df["soc_eco_class1"] == class1]
        class2_opts = [ALL_GROUPS] + sorted(sub["soc_eco_class2"].dropna().unique().tolist())
        if st.session_state["f_class2"] not in class2_opts:
            st.session_state["f_class2"] = ALL_GROUPS
        class2 = st.selectbox("กลุ่มเศรษฐกิจและสังคม (ย่อย)", class2_opts, key="f_class2")

        st.button("รีเซ็ตตัวกรอง", on_click=reset_filters, args=(years[0],),
                  width="stretch", type="secondary")

        st.divider()
        html(f'<div class="side-title">{icon("info", 16, MUTED)}วิธีอ่านค่า</div>')
        st.caption(
            "ข้อมูลมีทั้งหมวดยอดรวมและหมวดย่อยอยู่ในไฟล์เดียวกัน แดชบอร์ดจึงแสดงค่าของ"
            "หมวดที่เลือกเท่านั้น ไม่นำหลายหมวดมารวมกัน เมื่อผลลัพธ์มีหลายรายการ "
            "ค่าที่แสดงคือค่าเฉลี่ยอย่างง่ายของรายการ (ไม่ถ่วงน้ำหนักด้วยจำนวนครัวเรือน "
            "เพราะไฟล์ไม่มีข้อมูลจำนวนครัวเรือน)"
        )

    return dict(year=year, province=province, income=income, class1=class1, class2=class2)


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
        k["value"] = float(vals.iloc[0]) if len(vals) == 1 else float(vals.mean())
        top = dff.loc[vals.idxmax()]
        k["max_value"] = float(top["value"])
        k["max_where"] = f'{top["province"]} · {top["soc_eco_class2"]}'
    return k


def kpi_card(icon_name: str, label: str, value: str, sub: str, accent: bool = False) -> str:
    return (
        f'<div class="kpi{" accent" if accent else ""}">'
        f'<div class="k-label">{icon(icon_name, 15, MUTED)}{escape(label)}</div>'
        f'<div class="k-value">{value}</div>'
        f'<div class="k-sub">{sub}</div></div>'
    )


def render_kpis(k: dict, f: dict) -> None:
    other = ", ".join(str(y) for y in k["other_years"])
    if f["province"] != ALL_PROVINCES:
        prov_value, prov_sub = escape(f["province"]), "จังหวัดที่เลือก"
    else:
        prov_value, prov_sub = f'{k["n_provinces"]}<small>จังหวัด</small>', "จังหวัดที่อยู่ในผลลัพธ์"

    if k["value"] is None:
        val_html, val_sub = "—", "ไม่มีค่าตัวเลข"
    elif k["n_records"] == 1:
        val_html = f'{fmt_baht(k["value"])}<small>{UNIT_LABEL}</small>'
        val_sub = "ค่าจริงของรายการที่เลือก"
    else:
        val_html = f'{fmt_baht(k["value"])}<small>{UNIT_LABEL}</small>'
        val_sub = f'ค่าเฉลี่ยอย่างง่ายของ {k["n_records"]:,} รายการ (ไม่ถ่วงน้ำหนัก)'

    max_html = (f'{fmt_baht(k["max_value"])}<small>{UNIT_LABEL}</small>'
                if k["max_value"] is not None else "—")
    cards = [
        kpi_card("calendar-days", "ปีที่เลือก", f'พ.ศ. {k["year"]}',
                 f"ปีอื่นในชุดข้อมูล: {other}" if other else "ปีเดียวในชุดข้อมูล"),
        kpi_card("map-pin", "จังหวัด", prov_value, prov_sub),
        kpi_card("list-checks", "รายการที่ผ่านตัวกรอง", f'{k["n_records"]:,}',
                 f'1 รายการ = จังหวัด × กลุ่มย่อย ({k["n_groups"]} กลุ่มย่อย)'),
        kpi_card("wallet", "รายได้หมวดที่เลือก", val_html,
                 f'{escape(f["income"])}<br>{val_sub}', accent=True),
        kpi_card("trophy", "ค่าสูงสุดในผลลัพธ์", max_html,
                 escape(k["max_where"]) if k["max_where"] else "—"),
    ]
    html(f'<div class="kpi-grid">{"".join(cards)}</div>')


# ─────────────────────────────────────────────
# Charts
# ─────────────────────────────────────────────
PLOTLY_CONFIG = {"displaylogo": False, "modeBarButtonsToRemove": ["lasso2d", "select2d"]}


def style_fig(fig, height: int, *, legend: bool = False):
    fig.update_layout(
        template="plotly_white",
        height=height,
        font=dict(family=FONT_FAMILY, color=TEXT, size=13),
        title=dict(font=dict(size=15, color=NAVY), x=0, xanchor="left"),
        margin=dict(l=8, r=24, t=56, b=8),
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF",
        hoverlabel=dict(bgcolor="#FFFFFF", bordercolor="#DCE3EC",
                        font=dict(family=FONT_FAMILY, color=TEXT, size=13)),
        showlegend=legend,
        legend=dict(orientation="h", yanchor="bottom", y=1.0, xanchor="left", x=0,
                    title_text="", font=dict(size=12)),
        bargap=0.28,
    )
    fig.update_xaxes(gridcolor=GRID, zerolinecolor="#AEB9C7", linecolor=GRID,
                     title_font=dict(size=12, color=MUTED), tickfont=dict(color=MUTED))
    fig.update_yaxes(gridcolor=GRID, zerolinecolor="#AEB9C7", linecolor=GRID,
                     title_font=dict(size=12, color=MUTED), tickfont=dict(color=TEXT))
    return fig


def province_chart(df: pd.DataFrame, f: dict, view: str):
    """Value of the selected category by province (all provinces; selection highlighted)."""
    data = filter_data(df, f, skip=("province",))
    agg = (data.groupby("province", as_index=False)
               .agg(value=("value", "mean"), n=("value", "size"))
               .sort_values("value", ascending=False)
               .reset_index(drop=True))
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
    measure = "ค่าเฉลี่ยของกลุ่มย่อย" if many else "มูลค่า"
    highlight = selected != ALL_PROVINCES
    fig = px.bar(
        show, x="value", y="province", orientation="h",
        color="สถานะ" if highlight else None,
        color_discrete_map={"จังหวัดที่เลือก": TEAL, "จังหวัดอื่น": DIM},
        custom_data=["rank", "n"],
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
            + ("<br>เฉลี่ยจาก %{customdata[1]} กลุ่มย่อย" if many else "")
            + "<extra></extra>"
        ),
    )
    fig.update_yaxes(categoryorder="array", categoryarray=show["province"].tolist(),
                     tickfont=dict(size=12))
    style_fig(fig, max(360, 26 * len(show) + 90), legend=highlight)
    return fig, many


def group_chart(df: pd.DataFrame, f: dict):
    """Value of the selected category by socio-economic sub-group (selection highlighted)."""
    data = filter_data(df, f, skip=("class1", "class2"))
    agg = (data.groupby(["soc_eco_class1", "soc_eco_class2"], as_index=False)
               .agg(value=("value", "mean"), n=("value", "size"))
               .sort_values("value", ascending=True))
    class1_order = sorted(df["soc_eco_class1"].dropna().unique())
    many = agg["n"].max() > 1
    measure = "ค่าเฉลี่ยของจังหวัด" if many else "มูลค่า"
    fig = px.bar(
        agg, x="value", y="soc_eco_class2", orientation="h", color="soc_eco_class1",
        category_orders={"soc_eco_class1": class1_order},
        color_discrete_sequence=CATEGORICAL, custom_data=["soc_eco_class1", "n"],
        title=f"{f['income']} ตามกลุ่มเศรษฐกิจและสังคม",
        labels={"value": f"{measure} ({UNIT_LABEL})", "soc_eco_class2": ""},
    )
    fig.update_traces(
        marker_line_width=0,
        hovertemplate=(
            "<b>%{y}</b><br>กลุ่มหลัก: %{customdata[0]}<br>"
            + measure + ": %{x:,.0f} " + UNIT_LABEL
            + ("<br>เฉลี่ยจาก %{customdata[1]} จังหวัด" if many else "")
            + "<extra></extra>"
        ),
    )
    # Dim groups outside the group filter instead of hiding them.
    for trace in fig.data:
        keep = []
        for c1, c2 in zip([trace.name] * len(trace.y), trace.y):
            ok = (f["class1"] in (ALL_GROUPS, c1)) and (f["class2"] in (ALL_GROUPS, c2))
            keep.append(1.0 if ok else 0.25)
        trace.marker.opacity = keep
    fig.update_yaxes(categoryorder="array", categoryarray=agg["soc_eco_class2"].tolist(),
                     tickfont=dict(size=11))
    style_fig(fig, 470, legend=True)
    fig.update_layout(legend=dict(y=-0.32, yanchor="top"), margin=dict(b=8))
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
    fig = px.bar(
        agg, x="value", y="source_income3", orientation="h",
        color="อยู่ในหมวดที่เลือก",
        color_discrete_map={"อยู่ในหมวดที่เลือก": TEAL, "หมวดอื่น": DIM},
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
    fig.update_yaxes(categoryorder="array", categoryarray=agg["source_income3"].tolist(),
                     tickfont=dict(size=11))
    # A one-item legend says nothing when every leaf falls under the selection.
    style_fig(fig, 470, legend=len(fig.data) > 1)
    fig.update_layout(legend=dict(y=-0.2, yanchor="top"))
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
    style_fig(fig, max(320, 46 * agg["soc_eco_class2"].nunique() + 110), legend=True)
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


def render_table(dff: pd.DataFrame, f: dict) -> None:
    # Value next to the identifying columns so it stays visible without horizontal scroll.
    cols = ["year", "province", "soc_eco_class2", "value", "source_income3", "soc_eco_class1"]
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
        column_config={
            THAI_COLS["year"]: st.column_config.NumberColumn(format="%d", width="small"),
            THAI_COLS["value"]: st.column_config.NumberColumn(format="localized", width="small"),
        },
        column_order=[THAI_COLS[c] for c in cols] + ["หมายเหตุ"],
    )
    with st.expander("รายละเอียดทุกคอลัมน์ รวม source และ attribute ต้นฉบับ"):
        st.caption("แสดงข้อมูลตามไฟล์ต้นฉบับทุกคอลัมน์ ไม่มีการแก้ไขค่า")
        st.dataframe(
            dff[REQUIRED_COLUMNS].rename(columns=THAI_COLS), width="stretch",
            height=320, hide_index=True,
            column_config={THAI_COLS["year"]: st.column_config.NumberColumn(format="%d")},
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


def render_filter_chips(f: dict) -> None:
    group = f["class2"] if f["class2"] != ALL_GROUPS else f["class1"]
    chips = [
        ("calendar-days", f"พ.ศ. {f['year']}"),
        ("map-pin", f["province"]),
        ("wallet", f["income"]),
        ("users", group),
    ]
    html('<div class="chips"><span class="lbl">ตัวกรองปัจจุบัน</span>'
         + "".join(f'<span class="chip">{icon(n, 14, TEAL)}{escape(t)}</span>' for n, t in chips)
         + "</div>")


def render_footer() -> None:
    html(
        '<div class="app-footer">'
        f'<div class="row">{icon("database", 15, MUTED)}<div>แหล่งข้อมูล: '
        'สำนักงานสถิติแห่งชาติ (ตามคอลัมน์ source ในไฟล์ data.xlsx)</div></div>'
        f'<div class="row">{icon("triangle-alert", 15, MUTED)}<div>ข้อจำกัด: ไฟล์มีทั้งหมวดยอดรวม'
        'และหมวดย่อย การรวมทุกแถวจะนับรายได้ซ้ำ ค่าเฉลี่ยในแดชบอร์ดเป็นค่าเฉลี่ยอย่างง่าย'
        'ของรายการ ไม่ใช่ค่าเฉลี่ยของประชากร เพราะไม่มีจำนวนครัวเรือนสำหรับถ่วงน้ำหนัก</div></div>'
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

    render_filter_chips(f)
    render_kpis(compute_kpis(dff, f, years), f)
    if issues["non_numeric_value"] or issues["missing_value"]:
        note(f'พบค่า value ที่ว่างหรือแปลงเป็นตัวเลขไม่ได้ '
             f'{issues["missing_value"] + issues["non_numeric_value"]:,} แถว '
             'แถวเหล่านี้ไม่ถูกนำไปคำนวณ', "warn")

    if dff.empty:
        note("<b>ไม่พบข้อมูลตามตัวกรองที่เลือก</b><br>"
             "ลองเปลี่ยนจังหวัดหรือกลุ่มเศรษฐกิจและสังคม หรือกดรีเซ็ตตัวกรองในแถบด้านข้าง", "warn")
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
    st.plotly_chart(fig, width="stretch", config=PLOTLY_CONFIG)
    if many:
        note("แต่ละจังหวัดมีหลายกลุ่มเศรษฐกิจและสังคม ค่าในกราฟคือค่าเฉลี่ยอย่างง่ายของกลุ่มย่อย "
             "(ไม่ถ่วงน้ำหนัก) เลือกกลุ่มย่อยในตัวกรองเพื่อดูค่าจริงของกลุ่มนั้น")

    # Section 2: groups & composition
    section("chart-no-axes-combined", "กลุ่มเศรษฐกิจและสังคม และองค์ประกอบรายได้",
            "กราฟซ้ายแสดงทุกกลุ่มย่อยโดยจางกลุ่มที่อยู่นอกตัวกรอง กราฟขวาแสดง 7 หมวดย่อยที่ไม่ซ้อนกัน "
            "ผลรวมของทั้ง 7 หมวดเท่ากับรายได้ทั้งสิ้น (คลาดเคลื่อนไม่เกิน ±2 บาทจากการปัดเศษ)")
    left, right = st.columns(2, gap="large")
    with left:
        fig, many = group_chart(df, f)
        st.plotly_chart(fig, width="stretch", config=PLOTLY_CONFIG)
        if many:
            st.caption("ค่าเฉลี่ยอย่างง่ายของทุกจังหวัดในกลุ่มนั้น (ไม่ถ่วงน้ำหนัก)")
    with right:
        fig, many = composition_chart(df, f)
        st.plotly_chart(fig, width="stretch", config=PLOTLY_CONFIG)
        st.caption("แถบสีเขียวคือหมวดย่อยที่อยู่ภายใต้หมวดรายได้ที่เลือก"
                   + (" · ค่าเฉลี่ยอย่างง่ายของรายการที่ผ่านตัวกรอง" if many else ""))

    # Section 3: year comparison
    section("calendar-days", "เปรียบเทียบระหว่างปี")
    if f["province"] == ALL_PROVINCES:
        note("เลือกจังหวัดในแถบด้านข้างเพื่อดูการเปรียบเทียบระหว่างปีของจังหวัดนั้น")
    else:
        fig = year_chart(df, f)
        if fig is None:
            note("ไม่มีข้อมูลสำหรับการเปรียบเทียบระหว่างปี", "warn")
        else:
            st.plotly_chart(fig, width="stretch", config=PLOTLY_CONFIG)

    # Section 4: table
    section("table", "ตารางข้อมูล", "ข้อมูลที่ผ่านตัวกรองทั้งหมด ค่าเป็นค่าจริงจากไฟล์ต้นฉบับ")
    render_table(dff, f)
    render_quality(issues)
    render_footer()


main()
