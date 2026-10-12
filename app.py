import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os

# ─────────────────────────────────────────────
# Page Configuration
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Thai Household Income Dashboard",
    page_icon="🏠",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# Custom CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"] {
    font-family: 'Sarabun', sans-serif;
}

.stApp {
    background: linear-gradient(135deg, #0a1628 0%, #0d2040 50%, #0a1628 100%);
    color: #e2e8f0;
}

.dashboard-header {
    background: linear-gradient(135deg, #0f3460 0%, #1a4a7a 50%, #0f3460 100%);
    border: 1px solid rgba(32, 178, 170, 0.3);
    border-radius: 16px;
    padding: 28px 36px;
    margin-bottom: 24px;
    box-shadow: 0 8px 32px rgba(0,0,0,0.4), inset 0 1px 0 rgba(255,255,255,0.05);
}
.dashboard-title {
    font-size: 2.1rem;
    font-weight: 700;
    color: #20b2aa;
    margin: 0 0 6px 0;
    letter-spacing: 0.5px;
    text-shadow: 0 0 20px rgba(32,178,170,0.4);
}
.dashboard-subtitle {
    font-size: 1.05rem;
    color: #94a3b8;
    margin: 0 0 12px 0;
}
.dashboard-source {
    font-size: 0.82rem;
    color: #64748b;
    border-top: 1px solid rgba(32,178,170,0.15);
    padding-top: 10px;
    margin-top: 10px;
}

.kpi-grid {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 16px;
    margin-bottom: 24px;
}
.kpi-card {
    background: linear-gradient(145deg, #0f2a4a, #163356);
    border: 1px solid rgba(32,178,170,0.25);
    border-radius: 14px;
    padding: 20px 22px;
    box-shadow: 0 4px 20px rgba(0,0,0,0.3);
    transition: transform 0.2s, box-shadow 0.2s;
}
.kpi-card:hover {
    transform: translateY(-2px);
    box-shadow: 0 8px 28px rgba(32,178,170,0.15);
}
.kpi-icon { font-size: 1.5rem; margin-bottom: 8px; display: block; }
.kpi-label {
    font-size: 0.78rem;
    color: #64748b;
    font-weight: 500;
    letter-spacing: 0.5px;
    text-transform: uppercase;
    margin-bottom: 6px;
}
.kpi-value { font-size: 1.6rem; font-weight: 700; color: #20b2aa; line-height: 1.2; }
.kpi-sub { font-size: 0.75rem; color: #475569; margin-top: 4px; }

.section-header {
    font-size: 1.0rem;
    font-weight: 600;
    color: #20b2aa;
    border-left: 3px solid #20b2aa;
    padding-left: 12px;
    margin: 20px 0 12px 0;
}

[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #061020 0%, #0a1a30 100%) !important;
    border-right: 1px solid rgba(32,178,170,0.15) !important;
}

.footer {
    background: rgba(15,42,74,0.6);
    border: 1px solid rgba(32,178,170,0.1);
    border-radius: 10px;
    padding: 16px 20px;
    margin-top: 24px;
    font-size: 0.78rem;
    color: #475569;
    line-height: 1.6;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────
DATA_FILE = "data.xlsx"
REQUIRED_COLUMNS = [
    "year", "province", "source_income1", "source_income2",
    "source_income3", "soc_eco_class1", "soc_eco_class2",
    "value", "unit", "attribute", "source"
]

THAI_COLS = {
    "year": "ปี (พ.ศ.)",
    "province": "จังหวัด",
    "source_income1": "หมวดรายได้ระดับ 1",
    "source_income2": "หมวดรายได้ระดับ 2",
    "source_income3": "หมวดรายได้ระดับ 3",
    "soc_eco_class1": "กลุ่มเศรษฐกิจ-สังคม (หลัก)",
    "soc_eco_class2": "กลุ่มเศรษฐกิจ-สังคม (ย่อย)",
    "value": "มูลค่า (บาท)",
    "unit": "หน่วย",
    "attribute": "หมายเหตุ",
    "source": "แหล่งข้อมูล",
}

PLOTLY_LAYOUT = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(family="Sarabun, sans-serif", color="#cbd5e1"),
    margin=dict(l=20, r=20, t=45, b=20),
)
AXES_STYLE = dict(gridcolor="rgba(255,255,255,0.06)", linecolor="rgba(255,255,255,0.1)")
AXES_STYLE_SMALL = dict(gridcolor="rgba(255,255,255,0.06)", linecolor="rgba(255,255,255,0.1)", tickfont=dict(size=10))


# ─────────────────────────────────────────────
# Data Loading
# ─────────────────────────────────────────────
@st.cache_data(show_spinner="⏳ กำลังโหลดข้อมูล...")
def load_data(filepath: str) -> pd.DataFrame:
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"ไม่พบไฟล์ข้อมูล: {filepath}")
    try:
        xl = pd.ExcelFile(filepath, engine="openpyxl")
    except Exception as e:
        raise RuntimeError(f"เปิดไฟล์ Excel ไม่สำเร็จ: {e}")

    if "data" not in xl.sheet_names:
        raise ValueError(
            f"ไม่พบ Sheet 'data' ใน {filepath} (พบ Sheet: {xl.sheet_names})"
        )

    df = xl.parse("data")
    missing_cols = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing_cols:
        raise ValueError(f"ไม่พบคอลัมน์ที่จำเป็น: {missing_cols}")

    # Clean string columns
    str_cols = [
        "province", "source_income1", "source_income2", "source_income3",
        "soc_eco_class1", "soc_eco_class2", "unit", "attribute", "source"
    ]
    for col in str_cols:
        df[col] = df[col].astype(str).str.strip().replace("nan", pd.NA)

    df["value"] = pd.to_numeric(df["value"], errors="coerce")
    df["year"] = pd.to_numeric(df["year"], errors="coerce")
    df = df.dropna(subset=["year"])
    df["year"] = df["year"].astype(int)
    return df


# ─────────────────────────────────────────────
# Header
# ─────────────────────────────────────────────
st.markdown("""
<div class="dashboard-header">
    <div class="dashboard-title">🏠 Thai Household Income Dashboard</div>
    <div class="dashboard-subtitle">
        แดชบอร์ดวิเคราะห์ข้อมูลรายได้ครัวเรือนไทย แยกตามจังหวัดและกลุ่มเศรษฐกิจและสังคม
    </div>
    <div class="dashboard-source">
        📊 แหล่งข้อมูล: <strong>สำนักงานสถิติแห่งชาติ (NSO)</strong> &nbsp;|&nbsp;
        หน่วยข้อมูล: <strong>บาท/เดือน/ครัวเรือน</strong> &nbsp;|&nbsp;
        🔗 ข้อมูลเปิด: <a href="https://catalog.data.go.th/organization/nso" target="_blank"
        style="color:#20b2aa">data.go.th</a>
        <span style="color:#374151"> (ยังไม่ได้ยืนยัน URL ชุดข้อมูลเฉพาะ)</span>
    </div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# Load Data + Error Handling
# ─────────────────────────────────────────────
load_error = None
df_raw = None
try:
    df_raw = load_data(DATA_FILE)
except FileNotFoundError as e:
    load_error = ("file_not_found", str(e))
except ValueError as e:
    load_error = ("value_error", str(e))
except Exception as e:
    load_error = ("unknown", str(e))

if load_error:
    etype, emsg = load_error
    if etype == "file_not_found":
        st.error(
            f"❌ **ไม่พบไฟล์ข้อมูล**\n\n{emsg}\n\n"
            "กรุณาวาง `data.xlsx` ไว้ในโฟลเดอร์เดียวกับ `app.py`"
        )
    elif etype == "value_error":
        st.error(f"❌ **โครงสร้างข้อมูลไม่ถูกต้อง**\n\n{emsg}")
    else:
        st.error(f"❌ **เกิดข้อผิดพลาดในการโหลดข้อมูล**\n\n{emsg}")
    st.stop()

# ─────────────────────────────────────────────
# Sidebar Filters
# ─────────────────────────────────────────────
years_all = sorted(df_raw["year"].dropna().unique().tolist(), reverse=True)
provinces_all = ["ทุกจังหวัด"] + sorted(df_raw["province"].dropna().unique().tolist())
soc_eco1_all = ["ทุกกลุ่ม"] + sorted(df_raw["soc_eco_class1"].dropna().unique().tolist())

# Determine leaf vs roll-up income3 values
leaf_mask_raw = df_raw["source_income3"] != df_raw["source_income1"]
income3_leaf = sorted(
    df_raw.loc[leaf_mask_raw, "source_income3"].dropna().unique().tolist()
)
income3_rollup = sorted(
    df_raw.loc[~leaf_mask_raw, "source_income3"].dropna().unique().tolist()
)
# Build grouped options list
income3_opts = (
    ["── ยอดรวม ──"] + income3_rollup +
    ["── แยกประเภท ──"] + income3_leaf
)

with st.sidebar:
    st.markdown("### 🔎 ตัวกรองข้อมูล")
    st.markdown("---")

    sel_year = st.selectbox(
        "📅 ปี (พ.ศ.)",
        options=years_all,
        index=0,
        help="เลือกปีพุทธศักราชที่ต้องการวิเคราะห์",
    )

    sel_province = st.selectbox(
        "📍 จังหวัด",
        options=provinces_all,
        index=0,
        help="'ทุกจังหวัด' = แสดงข้อมูลทุกจังหวัดในกราฟเปรียบเทียบ",
    )

    sel_income3_raw = st.selectbox(
        "💰 หมวดรายได้",
        options=income3_opts,
        index=1,  # default: first rollup = รายได้ทั้งสิ้นต่อเดือน
        help=(
            "ยอดรวม = ค่าที่คำนวณรวมไว้แล้วในชุดข้อมูล\n"
            "แยกประเภท = แสดงเฉพาะหมวดย่อยนั้น\n"
            "⚠️ ห้ามนำหลายหมวดมารวมกันเอง"
        ),
    )

    sel_soc_eco1 = st.selectbox(
        "👥 กลุ่มเศรษฐกิจ-สังคม",
        options=soc_eco1_all,
        index=0,
        help="กรองตามกลุ่มอาชีพ/เศรษฐกิจของหัวหน้าครัวเรือน",
    )

    st.markdown("---")
    st.markdown("""
    <div style='font-size:0.75rem; color:#475569; line-height:1.7;'>
    <strong style='color:#64748b;'>⚠️ หมายเหตุการตีความ</strong><br>
    • ชุดข้อมูลมีทั้งหมวดรวมและหมวดย่อย<br>
    • ห้ามนำทุกแถวมารวมกัน<br>
      (จะเกิดการนับรายได้ซ้ำ)<br>
    • กราฟแสดงค่าของ<em>หมวดที่เลือก</em>เท่านั้น<br>
    • ค่าติดลบหมายถึงรายได้ที่ติดลบจริง<br>
      ตามคำอธิบายในชุดข้อมูล
    </div>
    """, unsafe_allow_html=True)

# Resolve income3 (skip header separators)
if sel_income3_raw.startswith("──"):
    sel_income3 = "รายได้ทั้งสิ้นต่อเดือน"
else:
    sel_income3 = sel_income3_raw

# ─────────────────────────────────────────────
# Apply Filters
# ─────────────────────────────────────────────
df_filtered = df_raw.copy()
df_filtered = df_filtered[df_filtered["year"] == sel_year]
if sel_province != "ทุกจังหวัด":
    df_filtered = df_filtered[df_filtered["province"] == sel_province]
df_filtered = df_filtered[df_filtered["source_income3"] == sel_income3]
if sel_soc_eco1 != "ทุกกลุ่ม":
    df_filtered = df_filtered[df_filtered["soc_eco_class1"] == sel_soc_eco1]

n_records = len(df_filtered)
n_provinces_result = df_filtered["province"].nunique() if n_records > 0 else 0

# ─────────────────────────────────────────────
# KPI Cards
# ─────────────────────────────────────────────
if n_records > 0 and not df_filtered["value"].isna().all():
    max_val = df_filtered["value"].max()
    max_prov = df_filtered.loc[df_filtered["value"].idxmax(), "province"]
else:
    max_val = None
    max_prov = "—"

province_display = (
    sel_province if sel_province != "ทุกจังหวัด"
    else f"{n_provinces_result} จังหวัด"
)
max_val_str = f"{max_val:,.0f}" if max_val is not None else "—"

st.markdown(f"""
<div class="kpi-grid">
    <div class="kpi-card">
        <span class="kpi-icon">📅</span>
        <div class="kpi-label">ปีที่วิเคราะห์</div>
        <div class="kpi-value">พ.ศ.&nbsp;{sel_year}</div>
        <div class="kpi-sub">ปีพุทธศักราช</div>
    </div>
    <div class="kpi-card">
        <span class="kpi-icon">📍</span>
        <div class="kpi-label">จังหวัดในผลลัพธ์</div>
        <div class="kpi-value">{province_display}</div>
        <div class="kpi-sub">จากตัวกรองที่เลือก</div>
    </div>
    <div class="kpi-card">
        <span class="kpi-icon">📋</span>
        <div class="kpi-label">จำนวนรายการ</div>
        <div class="kpi-value">{n_records:,}</div>
        <div class="kpi-sub">รายการที่ผ่านตัวกรอง</div>
    </div>
    <div class="kpi-card">
        <span class="kpi-icon">🏆</span>
        <div class="kpi-label">จังหวัดค่าสูงสุด</div>
        <div class="kpi-value" style="font-size:1.1rem;">{max_prov}</div>
        <div class="kpi-sub">{max_val_str} บาท/เดือน</div>
    </div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# Empty data guard
# ─────────────────────────────────────────────
if n_records == 0:
    st.warning(
        "⚠️ **ไม่พบข้อมูลตามตัวกรองที่เลือก** "
        "กรุณาปรับเงื่อนไขแล้วลองอีกครั้ง"
    )
    st.stop()

# ─────────────────────────────────────────────
# Chart 1: Income by Province
# ─────────────────────────────────────────────
st.markdown(
    '<div class="section-header">📊 รายได้ตามจังหวัด</div>',
    unsafe_allow_html=True
)

# When multiple soc_eco groups per province → use mean to avoid artificial inflation
chart1_agg = df_filtered.groupby("province", as_index=False)["value"].mean()
chart1_agg = chart1_agg.sort_values("value", ascending=False)
multi_group = (n_records > len(chart1_agg)) and sel_soc_eco1 == "ทุกกลุ่ม"

col_slider, _ = st.columns([1, 3])
with col_slider:
    max_prov_show = st.slider(
        "แสดงสูงสุด (จังหวัด)",
        min_value=5,
        max_value=min(77, len(chart1_agg)),
        value=min(20, len(chart1_agg)),
        step=5,
    )

if multi_group:
    st.info(
        "ℹ️ มีหลายกลุ่มเศรษฐกิจ-สังคมต่อจังหวัด — "
        "กราฟแสดง **ค่าเฉลี่ย** ของทุกกลุ่ม "
        "เลือกกลุ่มเฉพาะในตัวกรองซ้ายเพื่อดูค่ารายกลุ่ม"
    )

chart1_top = chart1_agg.head(max_prov_show)
fig1 = px.bar(
    chart1_top,
    x="province",
    y="value",
    title=(
        f"หมวด '{sel_income3}' — มูลค่าต่อจังหวัด (ปี พ.ศ. {sel_year})"
        + (" [เฉลี่ยทุกกลุ่มเศรษฐกิจ]" if multi_group else "")
    ),
    labels={"province": "จังหวัด", "value": "มูลค่า (บาท/เดือน)"},
    color="value",
    color_continuous_scale=[[0, "#0f3460"], [0.5, "#1565c0"], [1, "#20b2aa"]],
    text="value",
)
fig1.update_traces(
    texttemplate="%{text:,.0f}",
    textposition="outside",
    textfont=dict(size=8, color="#cbd5e1"),
    hovertemplate="<b>%{x}</b><br>มูลค่า: %{y:,.0f} บาท/เดือน<extra></extra>",
)
fig1.update_layout(
    **PLOTLY_LAYOUT,
    height=460,
    xaxis_tickangle=-45,
    coloraxis_showscale=False,
    title_font=dict(size=13, color="#94a3b8"),
)
fig1.update_xaxes(**AXES_STYLE)
fig1.update_yaxes(**AXES_STYLE)
st.plotly_chart(fig1, width='stretch')

# ─────────────────────────────────────────────
# Chart 2 & 3 side by side
# ─────────────────────────────────────────────
st.markdown(
    '<div class="section-header">🔍 วิเคราะห์ตามกลุ่มเศรษฐกิจ-สังคม และประเภทรายได้ย่อย</div>',
    unsafe_allow_html=True
)
col_c2, col_c3 = st.columns(2)

# Chart 2: By soc_eco_class2
with col_c2:
    chart2_data = (
        df_filtered
        .groupby("soc_eco_class2", as_index=False)["value"]
        .mean()
        .sort_values("value", ascending=True)
    )
    if chart2_data.empty:
        st.warning("ℹ️ ไม่มีข้อมูลกลุ่มเศรษฐกิจ-สังคม")
    else:
        prov_lbl = sel_province if sel_province != "ทุกจังหวัด" else "ทุกจังหวัด"
        fig2 = px.bar(
            chart2_data,
            x="value",
            y="soc_eco_class2",
            orientation="h",
            title=f"กลุ่มเศรษฐกิจ-สังคม — {prov_lbl} / ปี พ.ศ. {sel_year}",
            labels={
                "soc_eco_class2": "กลุ่มเศรษฐกิจ-สังคม",
                "value": "มูลค่า (บาท/เดือน)"
            },
            color="value",
            color_continuous_scale=[[0, "#0f3460"], [0.5, "#1a5276"], [1, "#20b2aa"]],
            text="value",
        )
        fig2.update_traces(
            texttemplate="%{text:,.0f}",
            textposition="outside",
            textfont=dict(size=9, color="#cbd5e1"),
            hovertemplate="<b>%{y}</b><br>มูลค่า: %{x:,.0f} บาท/เดือน<extra></extra>",
        )
        fig2.update_layout(
            **PLOTLY_LAYOUT,
            height=430,
            coloraxis_showscale=False,
            title_font=dict(size=12, color="#94a3b8"),
        )
        fig2.update_xaxes(**AXES_STYLE)
        fig2.update_yaxes(**AXES_STYLE_SMALL)
        st.plotly_chart(fig2, width='stretch')
        if multi_group or sel_province == "ทุกจังหวัด":
            st.caption("ℹ️ ค่าที่แสดงคือค่าเฉลี่ยของจังหวัดในกลุ่มนั้น")

# Chart 3: Leaf-level income breakdown (no double-count)
with col_c3:
    df_c3 = df_raw.copy()
    df_c3 = df_c3[df_c3["year"] == sel_year]
    if sel_province != "ทุกจังหวัด":
        df_c3 = df_c3[df_c3["province"] == sel_province]
    if sel_soc_eco1 != "ทุกกลุ่ม":
        df_c3 = df_c3[df_c3["soc_eco_class1"] == sel_soc_eco1]
    # Only leaf rows
    df_c3 = df_c3[df_c3["source_income3"] != df_c3["source_income1"]]

    chart3_data = (
        df_c3
        .groupby("source_income3", as_index=False)["value"]
        .mean()
        .sort_values("value", ascending=True)
    )
    if chart3_data.empty:
        st.warning("ℹ️ ไม่มีข้อมูลประเภทรายได้ย่อย")
    else:
        prov_lbl2 = (
            sel_province if sel_province != "ทุกจังหวัด"
            else "ทุกจังหวัด (เฉลี่ย)"
        )
        fig3 = px.bar(
            chart3_data,
            x="value",
            y="source_income3",
            orientation="h",
            title=f"ประเภทรายได้ย่อย — {prov_lbl2} / ปี พ.ศ. {sel_year}",
            labels={
                "source_income3": "ประเภทรายได้",
                "value": "มูลค่า (บาท/เดือน)"
            },
            color="value",
            color_continuous_scale=[[0, "#1a1a5e"], [0.5, "#7c5cbf"], [1, "#b39ddb"]],
            text="value",
        )
        fig3.update_traces(
            texttemplate="%{text:,.0f}",
            textposition="outside",
            textfont=dict(size=9, color="#cbd5e1"),
            hovertemplate=(
                "<b>%{y}</b><br>มูลค่า: %{x:,.0f} บาท/เดือน<extra></extra>"
            ),
        )
        fig3.update_layout(
            **PLOTLY_LAYOUT,
            height=430,
            coloraxis_showscale=False,
            title_font=dict(size=12, color="#94a3b8"),
        )
        fig3.update_xaxes(**AXES_STYLE)
        fig3.update_yaxes(**AXES_STYLE_SMALL)
        st.plotly_chart(fig3, width='stretch')
        st.caption(
            "ℹ️ กราฟนี้แสดงเฉพาะหมวดย่อย (ไม่รวมยอดรวม) "
            "เพื่อป้องกันการนับรายได้ซ้ำ"
        )

# ─────────────────────────────────────────────
# Chart 4: Year Comparison (if multiple years & specific province)
# ─────────────────────────────────────────────
years_count = len(years_all)
if years_count > 1 and sel_province != "ทุกจังหวัด":
    st.markdown(
        '<div class="section-header">📈 เปรียบเทียบระหว่างปี</div>',
        unsafe_allow_html=True
    )
    df_yc = df_raw[
        (df_raw["province"] == sel_province) &
        (df_raw["source_income3"] == sel_income3)
    ].copy()
    if sel_soc_eco1 != "ทุกกลุ่ม":
        df_yc = df_yc[df_yc["soc_eco_class1"] == sel_soc_eco1]

    yc_agg = (
        df_yc
        .groupby(["year", "soc_eco_class2"], as_index=False)["value"]
        .mean()
    )
    yc_agg["year"] = yc_agg["year"].astype(str)

    if not yc_agg.empty:
        fig4 = px.bar(
            yc_agg,
            x="soc_eco_class2",
            y="value",
            color="year",
            barmode="group",
            title=(
                f"เปรียบเทียบรายได้ระหว่างปี — {sel_province} / {sel_income3}"
            ),
            labels={
                "soc_eco_class2": "กลุ่มเศรษฐกิจ-สังคม",
                "value": "มูลค่า (บาท/เดือน)",
                "year": "ปี (พ.ศ.)"
            },
            color_discrete_sequence=["#20b2aa", "#7c5cbf", "#e07b39"],
        )
        fig4.update_traces(
            hovertemplate=(
                "<b>%{x}</b><br>มูลค่า: %{y:,.0f} บาท/เดือน<extra></extra>"
            ),
        )
        fig4.update_layout(
            **PLOTLY_LAYOUT,
            height=380,
            xaxis_tickangle=-30,
            title_font=dict(size=13, color="#94a3b8"),
            legend=dict(font=dict(color="#94a3b8")),
        )
        fig4.update_xaxes(**AXES_STYLE)
        fig4.update_yaxes(**AXES_STYLE)
        st.plotly_chart(fig4, width='stretch')
    else:
        st.info("ℹ️ ไม่มีข้อมูลสำหรับการเปรียบเทียบระหว่างปี")

# ─────────────────────────────────────────────
# Data Table
# ─────────────────────────────────────────────
st.markdown(
    '<div class="section-header">📋 ตารางข้อมูลรายละเอียด</div>',
    unsafe_allow_html=True
)

display_cols = [
    "year", "province", "source_income1", "source_income2",
    "source_income3", "soc_eco_class1", "soc_eco_class2", "value", "unit"
]
df_table = df_filtered[display_cols].copy()
df_table = df_table.rename(columns={k: THAI_COLS[k] for k in display_cols})

# Format value column for display (keep numeric for sorting)
st.markdown(
    f"**พบ {len(df_table):,} รายการ** "
    f"(ปี พ.ศ. {sel_year} / {sel_province} / {sel_income3})"
)

if df_table.empty:
    st.info(
        "ℹ️ ไม่พบข้อมูลตามตัวกรองที่เลือก "
        "กรุณาปรับเงื่อนไขแล้วลองอีกครั้ง"
    )
else:
    st.dataframe(
        df_table,
        width='stretch',
        height=380,
        hide_index=True,
        column_config={
            THAI_COLS["value"]: st.column_config.NumberColumn(
                THAI_COLS["value"],
                format="฿%,.0f",
            )
        },
    )

with st.expander("🔍 ดูรายละเอียดเพิ่มเติม (attribute และ source)"):
    df_full = df_filtered.rename(columns=THAI_COLS)
    st.dataframe(df_full, width='stretch', height=300, hide_index=True)

# ─────────────────────────────────────────────
# Footer
# ─────────────────────────────────────────────
st.markdown("""
<div class="footer">
    <strong>📌 แหล่งข้อมูล:</strong>
    สำนักงานสถิติแห่งชาติ (National Statistical Office of Thailand)<br>
    <strong>🗂️ ชื่อชุดข้อมูล:</strong>
    รายได้ครัวเรือนจำแนกตามจังหวัดและกลุ่มเศรษฐกิจ-สังคม<br>
    <strong>⚠️ ข้อจำกัดการตีความ:</strong>
    ชุดข้อมูลมีทั้งหมวดรายได้รวมและหมวดย่อยอยู่พร้อมกัน
    การนำทุกแถวมารวมกันโดยตรงจะทำให้เกิดการนับซ้ำ
    กราฟแสดงค่าของหมวดที่เลือกเท่านั้น ไม่ใช่รายได้รวมทั้งหมดของครัวเรือน<br>
    <strong>📅 ข้อมูลในชุดนี้ครอบคลุมปี พ.ศ.:</strong> 2566, 2568 &nbsp;|&nbsp;
    พัฒนาด้วย Python · Streamlit · Plotly
</div>
""", unsafe_allow_html=True)
