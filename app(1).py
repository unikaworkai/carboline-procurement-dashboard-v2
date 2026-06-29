import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import warnings
import os
import io
 
warnings.filterwarnings("ignore")
 
# ── BRAND COLORS ──────────────────────────────────────────────
RED    = "#C41230"
BLUE   = "#003087"
NAVY   = "#1F3864"
STEEL  = "#2E75B6"
LBLUE  = "#D6E4F0"
AMBER  = "#B8860B"
GREY   = "#595959"
ORANGE = "#C55A11"
GREEN  = "#375623"
WHITE  = "#FFFFFF"
BG     = "#F8F9FA"
 
# ── PAGE CONFIG ───────────────────────────────────────────────
st.set_page_config(
    page_title="Carboline Procurement Engine",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded"
)
 
st.markdown(f"""
<style>
  html, body, [class*="css"] {{ font-family: Arial, sans-serif !important; }}
  .main {{ background-color: {BG}; }}
  .main .block-container {{
    padding-top: 2rem !important;
    padding-bottom: 1rem !important;
    margin-top: 0 !important;
  }}
  section[data-testid="stSidebar"] > div {{
    padding-top: 1rem !important;
  }}
  .kpi {{
    background:{WHITE}; border-radius:10px; padding:18px 14px 12px;
    border-top:4px solid {RED}; box-shadow:0 2px 8px rgba(0,0,0,0.10);
    text-align:center; margin-bottom:6px;
  }}
  .kpi .lbl {{ font-size:11px; color:{GREY}; font-weight:700;
    text-transform:uppercase; letter-spacing:0.5px; margin-bottom:4px; }}
  .kpi .val {{ font-size:26px; font-weight:800; color:{NAVY}; line-height:1.1; }}
  .kpi .sub {{ font-size:10px; color:{GREY}; margin-top:2px; }}
  .kpi-b  {{ border-top-color:{BLUE}  !important; }}
  .kpi-g  {{ border-top-color:{GREEN} !important; }}
  .kpi-a  {{ border-top-color:{AMBER} !important; }}
  .alert  {{
    background:{RED}; color:{WHITE}; border-radius:8px;
    padding:14px 20px; font-size:15px; font-weight:700;
    margin-bottom:12px; border-left:8px solid #800000;
  }}
  .fbox {{
    background:{LBLUE}; border-left:4px solid {BLUE}; border-radius:6px;
    padding:10px 16px; font-size:13px; color:{NAVY}; margin-bottom:12px;
  }}
  .foot {{ color:{GREY}; font-size:11px; text-align:center;
    margin-top:16px; padding-top:8px; border-top:1px solid #ddd; }}
  h2 {{ color:{NAVY}; }}
  h3 {{ color:{NAVY}; font-size:16px; }}
</style>
""", unsafe_allow_html=True)
 
CHART = dict(
    paper_bgcolor=WHITE, plot_bgcolor=WHITE,
    font=dict(family="Arial", color=NAVY),
    margin=dict(l=10, r=10, t=40, b=10),
    legend=dict(orientation="h", y=-0.18, x=0.5, xanchor="center"),
)
 
MONTH = {1:"Jan",2:"Feb",3:"Mar",4:"Apr",5:"May",6:"Jun",
         7:"Jul",8:"Aug",9:"Sep",10:"Oct",11:"Nov",12:"Dec"}
 

SITE_LN = {
    "Lake Charles": "L3021",
    "Green Bay":    "L3023",
    "Dayton":       "L303T",
    "Louisa":       "L303S",
}

FLAG_CLR = {
    "🔴 TOP PRIORITY — A-item understocked": RED,
    "🟡 REVIEW — B-item understocked": ORANGE,
    "🟠 MONITOR — C-item understocked": "#ED7D31",
    "🔵 REVIEW — Overstocked": STEEL,
    "✅ OK": GREEN,
    "⭐ GOOD APPLE — UNDERSTOCKED (priority stock)": AMBER,
    "⭐ GOOD APPLE — stocked OK": "#DAA520",
    "🗑 BAD APPLE — remove from LN": GREY,
    "D-tier — no SS needed": "#7F7F7F",
    "Out of scope": "#BFBFBF",
}
 
# ── HELPERS ───────────────────────────────────────────────────
def n(v, d=0):
    try: return f"{float(v):,.{d}f}"
    except: return "—"
 
def d(v, d=0):
    try: return f"${float(v):,.{d}f}"
    except: return "—"
 
def p(v, d=1):
    try: return f"{float(v):.{d}f}%"
    except: return "—"
 
def kpi(label, value, sub="", cls=""):
    st.markdown(f"""
    <div class="kpi {cls}">
      <div class="lbl">{label}</div>
      <div class="val">{value}</div>
      <div class="sub">{sub}</div>
    </div>""", unsafe_allow_html=True)
 
def empty(msg="No data for current selection"):
    fig = go.Figure()
    fig.add_annotation(text=msg, xref="paper", yref="paper",
                       x=0.5, y=0.5, showarrow=False,
                       font=dict(size=14, color=GREY))
    fig.update_layout(height=350, paper_bgcolor=WHITE,
                      plot_bgcolor=WHITE,
                      xaxis_visible=False, yaxis_visible=False)
    return fig
 
def dl_excel(df, filename):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False)
    st.download_button(f"⬇️ Download {filename}",
                       data=buf.getvalue(), file_name=filename,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
 
# ── EXCEL READER ──────────────────────────────────────────────
def read_sheet(path, sheet, id_col, skip=3):
    try:
        df = pd.read_excel(path, sheet_name=sheet, skiprows=skip)
        df = df[df[id_col].notna()]
        df = df[df[id_col].astype(str).str.strip() != id_col]
        df = df[~df[id_col].astype(str).str.startswith("NaN")]
        return df.reset_index(drop=True)
    except Exception as e:
        st.error(f"Could not read {sheet} from {path}: {e}")
        return pd.DataFrame()
 
# ── DATA LOADING ──────────────────────────────────────────────

def clean_yn(series):
    """Normalize YES/NO columns with emoji prefixes like '⭐ YES' or '🗑 YES'."""
    import re
    return series.astype(str).str.strip().apply(
        lambda v: "YES" if re.search(r'\bYES\b', v, re.IGNORECASE) else
                  ("NO"  if re.search(r'\bNO\b',  v, re.IGNORECASE) else v.upper()))

@st.cache_data(show_spinner="Loading Carboline data...")
def load():
    MASTER = "Carboline_MasterSS_Summary_v1(3).xlsx"
 
    # Check which files exist
    needed = {
        "rop_ss_moq.xlsx": "Safety stock calculations",
        "clean_consumption.xlsx": "Consumption data",
        "clean_po.xlsx": "Purchase orders",
        "clean_lead_time.xlsx": "Lead time data",
        "clean_cost.xlsx": "Standard costs",
        "clean_inventory.xlsx": "Inventory on-hand",
        "clean_supplier.xlsx": "Supplier master",
        "clean_item_master.xlsx": "Item master",
        "abc_classification.xlsx": "ABC classification",
        MASTER: "Master summary",
    }
    missing = [f for f in needed if not os.path.exists(f)]
    if missing:
        return None, missing
 
    D = {}
 
    # ── SS file ───────────────────────────────────────────────
    df_ss = read_sheet("rop_ss_moq.xlsx", "Full_SS_Results", "audit_flag_v4", skip=3)
    # Clean column names that have formula annotations
    df_ss.columns = [str(c).split("\n")[0].split("▶")[0].strip() for c in df_ss.columns]
    # Fix specific renamed formula cols
    col_map = {
        "ss_dollars F=OL": "ss_dollars",
        "rop_dollars F=PL": "rop_dollars",
        "avg_inv_ F=(O+P)/2L": "avg_inv_dollars",
        "curr_rop_ F=ML": "curr_rop_dollars",
        "ss_dollars": "ss_dollars",
        "rop_dollars": "rop_dollars",
    }
    df_ss.rename(columns=col_map, inplace=True)
 
    num_cols = ["avg_daily_usage_lbs", "lead_time_used_v4", "buffer_factor_v4",
                "standard_cost_usd", "current_rop_in_ln", "moq_order_increment",
                "new_ss_lbs_v4", "new_rop_lbs_v4", "rop_to_enter_ln",
                "rop_gap_lbs_v4", "financial_risk_v4"]
    for c in num_cols:
        if c in df_ss.columns:
            df_ss[c] = pd.to_numeric(df_ss[c], errors="coerce").fillna(0)
 
    # Always recompute dollar values cleanly
    df_ss["ss_dollars"]      = df_ss["new_ss_lbs_v4"]  * df_ss["standard_cost_usd"]
    df_ss["rop_dollars"]     = df_ss["new_rop_lbs_v4"] * df_ss["standard_cost_usd"]
    df_ss["avg_inv_dollars"] = (df_ss["new_ss_lbs_v4"] + df_ss["new_rop_lbs_v4"]) / 2 * df_ss["standard_cost_usd"]
    df_ss["curr_rop_dollars"]= df_ss["current_rop_in_ln"] * df_ss["standard_cost_usd"]
 
    # Ensure flag cols exist and normalize emoji-prefixed YES/NO values
    if "is_good_apple" not in df_ss.columns: df_ss["is_good_apple"] = "NO"
    if "is_bad_apple"  not in df_ss.columns: df_ss["is_bad_apple"]  = "NO"
    df_ss["is_good_apple"] = clean_yn(df_ss["is_good_apple"])
    df_ss["is_bad_apple"]  = clean_yn(df_ss["is_bad_apple"])
    D["ss"] = df_ss
 
    # Good / Bad apple sheets
    df_bad = read_sheet("rop_ss_moq.xlsx", "🗑 Bad_Apple_Remove", "Site", skip=3)
    D["bad_apple"] = df_bad
 
    df_gs = read_sheet("rop_ss_moq.xlsx", "⭐ Good_Apple_Stock", "Site", skip=3)
    D["good_stock"] = df_gs
 
    # Good apple missing — header is at row 2 (0-indexed)
    df_gm = pd.read_excel("rop_ss_moq.xlsx",
                           sheet_name="⭐ Good_Apple_Missing_ROP",
                           skiprows=2, header=0)
    df_gm = df_gm[df_gm.iloc[:,0].notna()]
    df_gm.columns = ["item_code","std_cost","destination","lead_time_days","action_raw"][:len(df_gm.columns)]
    df_gm["action"] = "Add Item Ordering parameters in LN"
    D["good_missing"] = df_gm
 
    # ── Consumption ───────────────────────────────────────────
    df_c = read_sheet("clean_consumption.xlsx", "Clean_Consumption_Data", "site_group", skip=3)
    df_c.columns = [str(c).split("▶")[0].strip().replace(" ▶ FORMULA","") for c in df_c.columns]
    df_c["qty_issued"]   = pd.to_numeric(df_c.get("qty_issued",   0), errors="coerce").fillna(0)
    df_c["year"]         = pd.to_numeric(df_c.get("year",         0), errors="coerce")
    df_c["period_month"] = pd.to_numeric(df_c.get("period_month", 0), errors="coerce")
    df_c = df_c[df_c["qty_issued"] > 0]
    D["cons"] = df_c
 
    # ── PO ────────────────────────────────────────────────────
    df_po = read_sheet("clean_po.xlsx", "Clean_PO_Data", "item_code", skip=3)
    df_po.columns = [str(c).replace(" ▶ FORMULA","").replace(" ▶ F","").strip() for c in df_po.columns]
    df_po["unit_price"]           = pd.to_numeric(df_po.get("unit_price",  0), errors="coerce").fillna(0)
    df_po["ordered_qty"]          = pd.to_numeric(df_po.get("ordered_qty", 0), errors="coerce").fillna(0)
    df_po["order_date"]           = pd.to_datetime(df_po.get("order_date"),           errors="coerce")
    df_po["planned_receipt_date"] = pd.to_datetime(df_po.get("planned_receipt_date"), errors="coerce")
    df_po["actual_receipt_date"]  = pd.to_datetime(df_po.get("actual_receipt_date"),  errors="coerce")
    D["po"] = df_po
 
    # ── Lead time ─────────────────────────────────────────────
    df_lt = read_sheet("clean_lead_time.xlsx", "Lead_Time_Data", "po_number", skip=3)
    df_lt.columns = [str(c).replace(" ▶ FORMULA","").replace(" ▶ F","").strip() for c in df_lt.columns]
    df_lt["lead_time_days_winsorized"] = pd.to_numeric(
        df_lt.get("lead_time_days_winsorized", 0), errors="coerce").fillna(0)
    df_lt["planned_receipt_date"] = pd.to_datetime(df_lt.get("planned_receipt_date"), errors="coerce")
    df_lt["actual_receipt_date"]  = pd.to_datetime(df_lt.get("actual_receipt_date"),  errors="coerce")
    df_lt["order_date"]           = pd.to_datetime(df_lt.get("order_date"),            errors="coerce")
    recv = df_lt[df_lt.get("po_status", pd.Series(dtype=str)) == "RECEIVED"].copy()
    recv["on_time_2day"] = (
        recv["actual_receipt_date"] <=
        recv["planned_receipt_date"] + pd.Timedelta(days=2)
    )
    D["lt"]   = df_lt
    D["recv"] = recv
 
    # ── Cost ──────────────────────────────────────────────────
    df_cost = read_sheet("clean_cost.xlsx", "Standard_Cost_Master", "item_code", skip=3)
    df_cost.columns = [str(c).replace(" ▶ FORMULA","").replace(" ▶ F","").strip() for c in df_cost.columns]
    df_cost["standard_cost_usd"] = pd.to_numeric(df_cost.get("standard_cost_usd",0), errors="coerce").fillna(0)
    D["cost"] = df_cost
 
    # ── Inventory ─────────────────────────────────────────────
    df_inv = read_sheet("clean_inventory.xlsx", "Inventory_OnHand", "item_code", skip=3)
    df_inv.columns = [str(c).replace(" ▶ FORMULA","").replace(" ▶ F","").strip() for c in df_inv.columns]
    for c in ["economic_stock","available_stock","inv_on_hand","inv_on_order","inv_allocated"]:
        if c in df_inv.columns:
            df_inv[c] = pd.to_numeric(df_inv[c], errors="coerce").fillna(0)
    D["inv"] = df_inv
 
    # ── Supplier ──────────────────────────────────────────────
    df_sup = read_sheet("clean_supplier.xlsx", "PO_Suppliers_Only", "supplier_bp_code", skip=3)
    D["sup"] = df_sup
 
    # ── Item master ───────────────────────────────────────────
    df_im = read_sheet("clean_item_master.xlsx", "3_Item_Master", "item_code", skip=3)
    df_im.columns = [str(c).replace(" ▶ FORMULA","").replace(" ▶ F","").strip() for c in df_im.columns]
    D["item_master"] = df_im
 
    # ── ABC ───────────────────────────────────────────────────
    df_abc = pd.read_excel("abc_classification.xlsx",
                           sheet_name="8003_RawMaterials_ABC", skiprows=3)
    df_abc.columns = ["rank","abc_tier","item_code","item_description","item_group",
                      "planning_signal","total_lbs_2024_2025","pct_of_total_lbs",
                      "cumulative_pct","buffer_factor","ss_eligible","notes"
                      ][:len(df_abc.columns)]
    df_abc = df_abc[df_abc["item_code"].notna()]
    D["abc"] = df_abc
 
    # ── Master summary ────────────────────────────────────────
    df_site = pd.read_excel(MASTER, sheet_name="Site_Summary", skiprows=3, nrows=6)
    df_site.columns = ["site","items","a_under","ga_under","bad_in_file",
                       "ss_dol","rop_dol","avg_inv_dol","avg_lt","otd","notes"
                       ][:len(df_site.columns)]
    df_site = df_site[df_site["site"].notna()]
    df_site = df_site[~df_site["site"].astype(str).str.contains("TOTAL|nan|NaN", na=True)]
    D["site_sum"] = df_site
 
    df_d = pd.read_excel(MASTER, sheet_name="D_Tier_Full_List", skiprows=1, header=0)
    df_d = df_d[df_d["item_code"].notna()]
    for c in ["total_lbs_all_years","lbs_2021_2023"]:
        if c in df_d.columns:
            df_d[c] = pd.to_numeric(df_d[c], errors="coerce").fillna(0)
    D["d_tier"] = df_d
 
    df_abc_site = pd.read_excel(MASTER, sheet_name="ABC_by_Site", skiprows=1, header=0)
    df_abc_site = df_abc_site[df_abc_site["site"].notna() &
                               ~df_abc_site["site"].astype(str).str.contains("SUBTOTAL|NaN|nan", na=True)]
    D["abc_site"] = df_abc_site
 
    return D, []
 
 
# ── FILTERS ───────────────────────────────────────────────────
def scope(df):
    """Remove OUT_OF_SCOPE and BAD APPLE from financial analysis."""
    return df[(df["abc_tier"] != "OUT_OF_SCOPE") & (df["is_bad_apple"] == "NO")]
 
def apply_f(df, site, tier, apple):
    d2 = df.copy()
    if site != "All Sites": d2 = d2[d2["site_group"] == site]
    if tier != "All Tiers": d2 = d2[d2["abc_tier"] == tier]
    if apple == "⭐ Good Apple Only":   d2 = d2[d2["is_good_apple"] == "YES"]
    elif apple == "🗑 Bad Apple Only":  d2 = d2[d2["is_bad_apple"]  == "YES"]
    elif apple == "Regular Only":
        d2 = d2[(d2["is_good_apple"] == "NO") & (d2["is_bad_apple"] == "NO")]
    return d2
 
def sidebar(df_ss):
    st.sidebar.markdown(f"""
    <div style="background:{RED};border-radius:8px;padding:16px;text-align:center;margin-bottom:16px;">
      <div style="color:{WHITE};font-size:22px;font-weight:900;letter-spacing:2px;">CARBOLINE</div>
      <div style="color:rgba(255,255,255,0.8);font-size:11px;letter-spacing:1px;">
        Coatings · Linings · Fireproofing</div>
      <div style="color:rgba(255,255,255,0.65);font-size:10px;margin-top:3px;">
        An RPM International Company</div>
    </div>""", unsafe_allow_html=True)
 
    st.sidebar.markdown("## 🏭 Procurement Engine")
    st.sidebar.markdown("---")
    st.sidebar.markdown("### Filters")
 
    sites  = ["All Sites"] + sorted(df_ss["site_group"].dropna().unique().tolist())
    tiers  = ["All Tiers", "A", "B", "C", "D"]
    apples = ["All Items", "⭐ Good Apple Only", "🗑 Bad Apple Only", "Regular Only"]
 
    site  = st.sidebar.selectbox("📍 Site",     sites)
    tier  = st.sidebar.selectbox("📊 ABC Tier", tiers)
    apple = st.sidebar.selectbox("🍎 Category", apples)
 
    st.sidebar.markdown("---")
    ins = scope(df_ss)
    st.sidebar.markdown(f"**In-scope items:** {len(ins):,}")
    st.sidebar.markdown(f"**🔴 A-understocked:** {(ins['audit_flag_v4'].str.startswith('🔴')).sum():,}")
    ga_count  = (df_ss["is_good_apple"] == "YES").sum()
    bad_count = (df_ss["is_bad_apple"]  == "YES").sum()
    st.sidebar.markdown(f"**⭐ Good Apples:** {ga_count:,}")
    st.sidebar.markdown(f"**🗑 Bad Apples (in file):** {bad_count:,}")
    st.sidebar.markdown("---")
    st.sidebar.markdown(
        f"<div style='font-size:10px;color:{GREY};'>Data: June 2026 · SS v4 · OTD ±2 day</div>",
        unsafe_allow_html=True)
    return site, tier, apple
 
 
# ═══════════════════════════════════════════════════
# TAB 1 — EXECUTIVE OVERVIEW
# ═══════════════════════════════════════════════════
def t1_overview(D, filt):
    st.markdown(f"""
    <div style="background:linear-gradient(135deg,{NAVY},{BLUE});
                border-radius:12px;padding:22px 30px;margin-bottom:18px;">
      <div style="color:{WHITE};font-size:22px;font-weight:800;letter-spacing:1px;margin-top:10px;">
        🏭 Procurement Engine Dashboard</div>
      <div style="color:rgba(255,255,255,0.8);font-size:13px;margin-top:4px;">
        Carboline Company · Lake Charles · Green Bay · Dayton · Louisa</div>
      <div style="color:rgba(255,255,255,0.6);font-size:11px;margin-top:4px;">
        Safety Stock v4 · Good Apple / Bad Apple · Bulk Tank Caps Applied</div>
    </div>""", unsafe_allow_html=True)
 
    ins = scope(filt)
    a_under = (ins["audit_flag_v4"].str.startswith("🔴")).sum()
    ga_under = (ins["audit_flag_v4"].str.contains("GOOD APPLE.*UNDER", na=False)).sum()
    risk   = ins[ins["rop_gap_lbs_v4"] < 0]["financial_risk_v4"].sum()
    ss_inv = ins["ss_dollars"].sum()
 
    c1,c2,c3,c4,c5,c6 = st.columns(6)
    with c1: kpi("📦 In-Scope Items",    n(len(ins)),        "Excl. finished goods & bad apples")
    with c2: kpi("🔴 A-Items Underst.",  n(a_under),         "Fix LN reorder point now",   "kpi-b")
    with c3: kpi("⭐ Good Apple Underst.",n(ga_under),        "Priority stock below safe level","kpi-a")
    with c4: kpi("💰 A-Item $ Risk",     d(risk,0),          "Understocked A-item exposure", "kpi-b")
    with c5: kpi("📈 SS Investment",     d(ss_inv,0),        "Recommended SS × standard cost")
    with c6: kpi("🚚 OTD Rate",          "78.0%",             "With ±2 day tolerance", "kpi-g")
 
    st.markdown("<br>", unsafe_allow_html=True)
 
    # Row 1: Status by site | ABC donut
    c_l, c_r = st.columns([3,2])
    with c_l:
        st.markdown("### Inventory Status by Site")
        valid = ins[ins["abc_tier"] != "D"]
        grp = valid.groupby("site_group").agg(
            Understocked=("rop_gap_direction_v4", lambda x: x.str.startswith("UNDER").sum()),
            Overstocked=("rop_gap_direction_v4",  lambda x: x.str.startswith("OVER").sum()),
            OK=("rop_gap_direction_v4",            lambda x: x.str.startswith("OK").sum()),
        ).reset_index()
        if grp.empty:
            st.plotly_chart(empty(), use_container_width=True)
        else:
            fig = go.Figure()
            for col_n, col_c in [("Understocked",RED),("Overstocked",STEEL),("OK",GREEN)]:
                fig.add_trace(go.Bar(
                    name=col_n, x=grp["site_group"], y=grp[col_n],
                    marker_color=col_c,
                    text=grp[col_n], textposition="inside",
                    textfont=dict(color=WHITE, size=12)))
            fig.update_layout(barmode="stack", height=360,
                              title="Item Count by Status per Site",
                              xaxis_title=None, yaxis_title="Items", **CHART)
            st.plotly_chart(fig, use_container_width=True)
 
    with c_r:
        st.markdown("### ABC Tier Distribution")
        tc = ins[ins["abc_tier"].isin(["A","B","C","D"])]["abc_tier"].value_counts()
        if tc.empty:
            st.plotly_chart(empty(), use_container_width=True)
        else:
            fig2 = go.Figure(go.Pie(
                labels=tc.index, values=tc.values, hole=0.55,
                marker_colors=[RED, ORANGE, STEEL, GREY],
                textinfo="label+percent", textfont=dict(size=12)))
            fig2.update_layout(
                height=360, showlegend=True,
                annotations=[dict(text="ABC", x=0.5, y=0.5,
                                   font_size=14, showarrow=False, font_color=NAVY)],
                **CHART)
            st.plotly_chart(fig2, use_container_width=True)
 
    # Row 2: Top 10 risk | SS $ by site
    c_l2, c_r2 = st.columns([3,2])
    with c_l2:
        st.markdown("### Top 10 Items by Financial Risk")
        top10 = ins[ins["rop_gap_lbs_v4"] < -0.01].nlargest(10,"financial_risk_v4").copy()
        if top10.empty:
            st.plotly_chart(empty("No understocked items in selection"), use_container_width=True)
        else:
            top10["lbl"] = top10["item_code"] + " — " + top10["item_description"].str[:28]
            fig3 = px.bar(
                top10, x="financial_risk_v4", y="lbl", orientation="h",
                color="site_group",
                color_discrete_map={"Lake Charles":BLUE,"Green Bay":GREEN,"Dayton":AMBER,"Louisa":ORANGE},
                text=top10["financial_risk_v4"].apply(lambda v: d(v,0)),
                title="Highest Financial Exposure — Understocked Items")
            fig3.update_traces(textposition="outside", textfont=dict(size=10))
            fig3.update_layout(height=400, yaxis={"categoryorder":"total ascending"},
                               xaxis_tickprefix="$", xaxis_tickformat=",.0f",
                               yaxis_title=None, **CHART)
            st.plotly_chart(fig3, use_container_width=True)
 
    with c_r2:
        st.markdown("### Safety Stock $ by Site")
        sd = ins.groupby("site_group").agg(
            SS=("ss_dollars","sum"), ROP=("rop_dollars","sum")).reset_index()
        if sd.empty:
            st.plotly_chart(empty(), use_container_width=True)
        else:
            fig4 = go.Figure()
            fig4.add_trace(go.Bar(name="Safety Stock $", x=sd["site_group"], y=sd["SS"],
                                   marker_color=RED,
                                   text=[d(v,0) for v in sd["SS"]], textposition="outside",
                                   textfont=dict(size=9)))
            fig4.add_trace(go.Bar(name="ROP $", x=sd["site_group"], y=sd["ROP"],
                                   marker_color=STEEL,
                                   text=[d(v,0) for v in sd["ROP"]], textposition="outside",
                                   textfont=dict(size=9)))
            fig4.update_layout(barmode="group", height=400,
                               yaxis_tickprefix="$", yaxis_tickformat=",.0f",
                               title="SS $ vs ROP $ by Site (excl. bad apples)", **CHART)
            st.plotly_chart(fig4, use_container_width=True)
 
    # Site Summary mini-table
    site_sum = D.get("site_sum", pd.DataFrame())
    if not site_sum.empty:
        st.markdown("### \U0001f3ed Site Summary")
        disp_s = site_sum.copy()
        rename_s = {
            "site":"Site","items":"Items","a_under":"\U0001f534 A-Underst.",
            "ga_under":"\u2b50 GA Underst.","ss_dol":"SS Investment ($)",
            "rop_dol":"ROP Value ($)","avg_lt":"Avg Lead Time (d)","otd":"OTD %"
        }
        disp_s = disp_s.rename(columns=rename_s)
        keep = [c for c in rename_s.values() if c in disp_s.columns]
        disp_s = disp_s[keep]
        for c in ["SS Investment ($)","ROP Value ($)"]:
            if c in disp_s.columns:
                disp_s[c] = disp_s[c].apply(lambda v: d(v,0))
        if "Avg Lead Time (d)" in disp_s.columns:
            disp_s["Avg Lead Time (d)"] = disp_s["Avg Lead Time (d)"].apply(
                lambda v: f"{float(v):.1f}d" if pd.notna(v) else "\u2014")
        if "OTD %" in disp_s.columns:
            disp_s["OTD %"] = disp_s["OTD %"].apply(lambda v: p(v) if pd.notna(v) else "\u2014")
        st.dataframe(disp_s, use_container_width=True, hide_index=True)
        st.markdown("<br>", unsafe_allow_html=True)

    # Data freshness indicator
    po_df = D.get("po", pd.DataFrame())
    if not po_df.empty and "order_date" in po_df.columns:
        max_date = po_df["order_date"].max()
        if pd.notna(max_date):
            st.markdown(
                f"<div class='foot'>\U0001f4c5 Most recent PO transaction: <b>{max_date.strftime('%B %d, %Y')}</b> · "
                "Data: June 2026 · Safety Stock v4 · OTD uses \u00b12 day tolerance · "
                "Scope: Item Groups 8003 & 8010</div>",
                unsafe_allow_html=True)
        else:
            st.markdown(
                "<div class='foot'>Data as of June 2026 · Safety Stock v4 · "
                "OTD uses \u00b12 day tolerance · Scope: Item Groups 8003 & 8010</div>",
                unsafe_allow_html=True)
    else:
        st.markdown(
            "<div class='foot'>Data as of June 2026 · Safety Stock v4 · "
            "OTD uses \u00b12 day tolerance · Scope: Item Groups 8003 & 8010</div>",
            unsafe_allow_html=True)
 
 
# ═══════════════════════════════════════════════════
# TAB 2 — PRIORITY AUDIT LIST
# ═══════════════════════════════════════════════════
def t2_audit(D, filt):
    st.markdown(f"## 🔴 Priority Audit List")
    st.markdown(
        "<div class='fbox'><b>Action required:</b> Items below are understocked. "
        "The <b>'Update LN To This Value'</b> column is the exact number to enter in "
        "LN Item Ordering session to fix the reorder point.</div>",
        unsafe_allow_html=True)
 
    ins   = scope(filt)
    audit = ins[ins["rop_gap_lbs_v4"] < -0.01].sort_values("financial_risk_v4", ascending=False).copy()
    audit.insert(0, "Rank", range(1, len(audit)+1))
 
    fa,fb,fc,fd = st.columns([2,2,2,2])
    with fa:
        sites_a = ["All Sites"] + sorted(audit["site_group"].unique().tolist())
        sf = st.selectbox("Site", sites_a, key="a_s")
    with fb:
        tiers_a = ["All Tiers"] + [t for t in ["A","B","C"] if t in audit["abc_tier"].unique()]
        tf = st.selectbox("Tier", tiers_a, key="a_t")
    with fc:
        ga_on = st.toggle("⭐ Good Apple Only", key="a_ga")
    with fd:
        rows = st.selectbox("Rows", [25,50,100,"All"], key="a_r")
 
    if sf != "All Sites": audit = audit[audit["site_group"] == sf]
    if tf != "All Tiers": audit = audit[audit["abc_tier"] == tf]
    if ga_on: audit = audit[audit["is_good_apple"] == "YES"]
    if rows != "All": audit = audit.head(int(rows))
 
    m1,m2,m3,m4 = st.columns(4)
    with m1: st.metric("Items Shown", n(len(audit)))
    with m2: st.metric("Total $ Risk", d(audit["financial_risk_v4"].sum(),0))
    with m3: st.metric("Avg Gap (lbs)", n(audit["rop_gap_lbs_v4"].mean(),0))
    with m4: st.metric("A-Items", n((audit["abc_tier"]=="A").sum()))
    st.markdown("---")
 
    if audit.empty:
        st.info("No understocked items match the current filters.")
        return
 
    disp = audit[[
        "Rank","site_group","item_code","item_description","abc_tier",
        "is_good_apple","current_rop_in_ln","new_ss_lbs_v4","new_rop_lbs_v4",
        "rop_to_enter_ln","rop_gap_lbs_v4","financial_risk_v4",
        "avg_daily_usage_lbs","lead_time_used_v4"
    ]].copy()
    disp.columns = [
        "Rank","Site","Item Code","Description","Tier","Good Apple",
        "Current LN ROP (lbs)","Rec. SS (lbs)","Rec. ROP (lbs)",
        "✅ UPDATE LN TO THIS VALUE","Gap (lbs)","Financial Risk ($)",
        "Daily Usage (lbs/day)","Lead Time (days)"
    ]
    for c in ["Current LN ROP (lbs)","Rec. SS (lbs)","Rec. ROP (lbs)",
              "✅ UPDATE LN TO THIS VALUE","Gap (lbs)","Daily Usage (lbs/day)"]:
        disp[c] = disp[c].apply(lambda v: n(v,1))
    disp["Financial Risk ($)"] = disp["Financial Risk ($)"].apply(lambda v: d(v,2))
    disp["Lead Time (days)"]   = disp["Lead Time (days)"].apply(lambda v: n(v,1))
 
    st.dataframe(disp, use_container_width=True, height=500, hide_index=True)
    dl_excel(disp, "Carboline_Priority_Audit_List.xlsx")

    # Export for LN
    st.markdown("#### 📤 Export for LN Import")
    st.caption("Formatted for direct entry into LN Item Ordering session. Company=3000.")
    ln_exp = audit[["site_group","item_code","moq_order_increment","rop_to_enter_ln"]].copy()
    ln_exp["Company"] = "3000"
    ln_exp["Site"]    = ln_exp["site_group"].map(SITE_LN).fillna(ln_exp["site_group"])
    ln_exp["Order Method"] = "Reorder Point"
    ln_export = ln_exp[["Company","Site","item_code","Order Method",
                         "rop_to_enter_ln","moq_order_increment"]].copy()
    ln_export.columns = ["Company","Site","Item","Order Method","Safety Stock","Order Qty Increment"]
    ln_export["Safety Stock"]        = ln_export["Safety Stock"].round(0).astype(int)
    ln_export["Order Qty Increment"] = ln_export["Order Qty Increment"].round(0).astype(int)
    buf_ln = io.BytesIO()
    with pd.ExcelWriter(buf_ln, engine="openpyxl") as w:
        ln_export.to_excel(w, index=False)
    st.download_button("⬇️ Export for LN (formatted import file)",
                       data=buf_ln.getvalue(), file_name="Carboline_LN_Import.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # Chart
    top20 = audit.head(20).copy()
    top20["lbl"] = top20["item_code"] + " — " + top20["item_description"].str[:25]
    colors = [AMBER if r=="YES" else RED for r in top20["is_good_apple"]]
    fig = go.Figure(go.Bar(
        x=top20["financial_risk_v4"], y=top20["lbl"], orientation="h",
        marker_color=colors,
        text=[d(v,0) for v in top20["financial_risk_v4"]],
        textposition="outside"))
    fig.update_layout(
        height=max(380, len(top20)*22),
        title="Top 20 Items by Financial Risk (Gold = Good Apple priority item)",
        xaxis_tickprefix="$", xaxis_tickformat=",.0f",
        yaxis={"categoryorder":"total ascending"}, **CHART)
    st.plotly_chart(fig, use_container_width=True)
 
 
# ═══════════════════════════════════════════════════
# TAB 3 — SAFETY STOCK CALCULATOR
# ═══════════════════════════════════════════════════
def t3_calc(D, filt):
    st.markdown("## 🧮 Safety Stock Calculator")
    st.markdown("""
    <div class='fbox'>
    <b>Formula:</b> Recommended SS = Average Daily Usage × Lead Time × Buffer Factor
    &nbsp;|&nbsp; <b>ROP</b> = (Daily Usage × Lead Time) + Safety Stock
    &nbsp;|&nbsp; <b>MOQ</b> rounded UP to nearest order increment
    &nbsp;|&nbsp; <b>Buffer:</b> A=1.5× · B=1.0× · C=0.5× · D=0
    </div>""", unsafe_allow_html=True)
 
    srch = st.text_input("🔍 Search item code or description", "")
    ins  = scope(filt)
    if srch:
        mask = (ins["item_code"].str.contains(srch, case=False, na=False) |
                ins["item_description"].str.contains(srch, case=False, na=False))
        ins = ins[mask]
 
    m1,m2,m3,m4,m5 = st.columns(5)
    with m1: st.metric("Items", n(len(ins)))
    with m2: st.metric("Total SS $", d(ins["ss_dollars"].sum(),0))
    with m3: st.metric("Total ROP $", d(ins["rop_dollars"].sum(),0))
    with m4: st.metric("Avg Lead Time", f"{ins['lead_time_used_v4'].mean():.1f}d")
    with m5:
        pct = (ins["rop_gap_lbs_v4"] < -0.01).sum() / max(len(ins),1) * 100
        st.metric("% Understocked", p(pct))
 
    if ins.empty:
        st.info("No items match the current filters.")
        return
 
    cols_need = ["site_group","item_code","item_description","abc_tier","audit_flag_v4",
                 "is_good_apple","is_bad_apple","avg_daily_usage_lbs","lead_time_used_v4",
                 "buffer_factor_v4","standard_cost_usd","current_rop_in_ln",
                 "new_ss_lbs_v4","new_rop_lbs_v4","rop_to_enter_ln",
                 "rop_gap_lbs_v4","financial_risk_v4","ss_dollars","rop_dollars"]
    available = [c for c in cols_need if c in ins.columns]
    disp = ins[available].copy()
    rename = {
        "site_group":"Site","item_code":"Item Code","item_description":"Description",
        "abc_tier":"Tier","audit_flag_v4":"Status",
        "is_good_apple":"Good Apple","is_bad_apple":"Bad Apple",
        "avg_daily_usage_lbs":"Daily Usage (lbs)","lead_time_used_v4":"Lead Time (days)",
        "buffer_factor_v4":"Buffer","standard_cost_usd":"Std Cost ($)",
        "current_rop_in_ln":"Current LN ROP","new_ss_lbs_v4":"Rec. SS (lbs)",
        "new_rop_lbs_v4":"Rec. ROP (lbs)","rop_to_enter_ln":"✅ Enter in LN",
        "rop_gap_lbs_v4":"Gap (lbs)","financial_risk_v4":"$ Risk",
        "ss_dollars":"SS Value ($)","rop_dollars":"ROP Value ($)"
    }
    disp.rename(columns=rename, inplace=True)
    for c in ["Daily Usage (lbs)","Current LN ROP","Rec. SS (lbs)",
              "Rec. ROP (lbs)","✅ Enter in LN","Gap (lbs)"]:
        if c in disp.columns:
            disp[c] = disp[c].apply(lambda v: n(v,1))
    for c in ["Std Cost ($)","$ Risk","SS Value ($)","ROP Value ($)"]:
        if c in disp.columns:
            disp[c] = disp[c].apply(lambda v: d(v,2))
    if "Lead Time (days)" in disp.columns:
        disp["Lead Time (days)"] = disp["Lead Time (days)"].apply(lambda v: n(v,1))
    if "Buffer" in disp.columns:
        disp["Buffer"] = disp["Buffer"].apply(lambda v: n(v,1)+"×")
 
    st.dataframe(disp, use_container_width=True, height=520, hide_index=True)
    dl_excel(disp, "Carboline_SS_Calculator.xlsx")
 
 
# ═══════════════════════════════════════════════════
# TAB 4 — SUPPLIER PERFORMANCE
# ═══════════════════════════════════════════════════
def t4_supplier(D, filt):
    st.markdown("## 🚚 Supplier Performance Scorecard")
    recv  = D["recv"]
    sup   = D["sup"]
    po    = D["po"]
    cost  = D["cost"]
 
    if recv.empty or sup.empty:
        st.warning("Lead time or supplier file unavailable."); return
 
    # OTD per supplier
    sup_grp = (recv.groupby("supplier_bp_code")
               .agg(orders=("po_number","count"),
                    on_time=("on_time_2day","sum"),
                    avg_lt=("lead_time_days_winsorized","mean"))
               .reset_index())
    sup_grp["otd_pct"] = (sup_grp["on_time"] / sup_grp["orders"] * 100).round(1)
    sup_grp["avg_lt"]  = sup_grp["avg_lt"].round(1)
    sup_grp["late"]    = sup_grp["orders"] - sup_grp["on_time"]
    sup_grp = sup_grp.merge(
        sup[["supplier_bp_code","supplier_name","country_name","is_us_supplier"]],
        on="supplier_bp_code", how="left")
    sup_grp["supplier_name"] = sup_grp["supplier_name"].fillna(sup_grp["supplier_bp_code"])
    sup_grp["origin"] = sup_grp["is_us_supplier"].apply(
        lambda v: "Domestic (US)" if str(v)=="YES" else "International")
 
    overall_otd = 78.0
    active      = (sup_grp["orders"] >= 5).sum()
    avg_lt_all  = recv["lead_time_days_winsorized"].mean()
    late_pct    = (1 - recv["on_time_2day"].mean()) * 100
 
    k1,k2,k3,k4 = st.columns(4)
    with k1: kpi("🚚 Overall OTD",    p(overall_otd), "±2 day tolerance", "kpi-g")
    with k2: kpi("🏢 Active Suppliers", n(active),   "With 5+ orders", "kpi-b")
    with k3: kpi("⏱ Avg Lead Time",   f"{avg_lt_all:.1f}d", "Winsorized @ 103 days")
    with k4: kpi("❌ Late Delivery",   p(late_pct),   "Without tolerance", "kpi-b")
 
    st.markdown("---")
    sup5 = sup_grp[sup_grp["orders"] >= 5]
 
    c_l, c_r = st.columns(2)
    with c_l:
        st.markdown("### Supplier Risk Map — OTD vs Volume")
        if not sup5.empty:
            fig = px.scatter(
                sup5, x="otd_pct", y="orders",
                size="orders", color="origin",
                hover_name="supplier_name",
                hover_data={"otd_pct":":.1f","orders":True,"avg_lt":":.1f"},
                color_discrete_map={"Domestic (US)":BLUE,"International":ORANGE},
                labels={"otd_pct":"OTD %","orders":"Total Orders"},
                title="OTD % vs Order Volume (bubble = # orders)")
            fig.add_vline(x=78, line_dash="dash", line_color=RED,
                          annotation_text="Target 78%")
            fig.update_layout(height=400, **CHART)
            st.plotly_chart(fig, use_container_width=True)
 
    with c_r:
        st.markdown("### Lead Time Distribution")
        lt_v = recv["lead_time_days_winsorized"].dropna()
        fig2 = px.histogram(lt_v, nbins=50, title="Lead Time Distribution (days)",
                             color_discrete_sequence=[BLUE])
        fig2.add_vline(x=lt_v.mean(),   line_dash="dash", line_color=RED,
                       annotation_text=f"Mean: {lt_v.mean():.1f}d")
        fig2.add_vline(x=lt_v.median(), line_dash="dot",  line_color=GREEN,
                       annotation_text=f"Median: {lt_v.median():.1f}d")
        fig2.update_layout(height=400, xaxis_title="Days", yaxis_title="Deliveries", **CHART)
        st.plotly_chart(fig2, use_container_width=True)
 
    st.markdown("### Bottom 20 Suppliers by On-Time Delivery")
    bot20 = sup5.nsmallest(20,"otd_pct")[
        ["supplier_name","country_name","orders","on_time","late","otd_pct","avg_lt"]
    ].copy()
    bot20.columns = ["Supplier","Country","Total Orders","On-Time","Late","OTD %","Avg LT (days)"]
    bot20["OTD %"]         = bot20["OTD %"].apply(lambda v: p(v))
    bot20["Avg LT (days)"] = bot20["Avg LT (days)"].apply(lambda v: n(v,1))
    st.dataframe(bot20, use_container_width=True, height=420, hide_index=True)
 
    # PPV
    st.markdown("### Purchase Price Variance (PPV) by Supplier")
    st.caption("PPV = (Actual Price Paid − Standard Cost) × Ordered Qty. "
               "Positive = unfavorable (paid more). Negative = favorable (paid less).")
    po2   = po.copy()
    cost2 = cost[["item_code","standard_cost_usd"]].copy()
    ppv   = po2.merge(cost2, on="item_code", how="left")
    ppv["standard_cost_usd"] = pd.to_numeric(ppv["standard_cost_usd"], errors="coerce").fillna(0)
    ppv["unit_price"]        = pd.to_numeric(ppv.get("unit_price",0),   errors="coerce").fillna(0)
    ppv["ordered_qty"]       = pd.to_numeric(ppv.get("ordered_qty",0),  errors="coerce").fillna(0)
    ppv["ppv"]               = (ppv["unit_price"] - ppv["standard_cost_usd"]) * ppv["ordered_qty"]
    ppv                       = ppv[ppv["standard_cost_usd"] > 0]
    sup_ppv = (ppv.groupby("supplier_bp_code")["ppv"].sum()
               .reset_index().rename(columns={"ppv":"total_ppv"}))
    sup_ppv = sup_ppv.merge(sup[["supplier_bp_code","supplier_name"]], on="supplier_bp_code", how="left")
    sup_ppv["supplier_name"] = sup_ppv["supplier_name"].fillna(sup_ppv["supplier_bp_code"])
 
    # Vendor Scorecard
    st.markdown("### \U0001f3c6 Vendor Scorecard (out of 100)")
    st.caption("OTD Score (40 pts) + Lead Time Score (30 pts, target=30 days) + PPV Score (30 pts)")
    if not sup5.empty:
        sc = sup5.copy()
        sc["otd_score"] = (sc["otd_pct"] / 100 * 40).round(1)
        sc["lt_score"]  = (np.maximum(0, 30 - (sc["avg_lt"] - 30)) / 30 * 30).clip(0,30).round(1)
        sc = sc.merge(sup_ppv[["supplier_bp_code","total_ppv"]], on="supplier_bp_code", how="left")
        sc["total_ppv"] = sc["total_ppv"].fillna(0)
        ppv_min = sc["total_ppv"].min(); ppv_max = sc["total_ppv"].max()
        ppv_range = ppv_max - ppv_min if ppv_max != ppv_min else 1
        sc["ppv_score"]   = ((ppv_max - sc["total_ppv"]) / ppv_range * 30).clip(0,30).round(1)
        sc["total_score"] = (sc["otd_score"] + sc["lt_score"] + sc["ppv_score"]).round(1)
        col_sc_l, col_sc_r = st.columns(2)
        with col_sc_l:
            st.markdown("**Top 20 Suppliers**")
            show_top = sc.nlargest(20,"total_score")[["supplier_name","orders","otd_pct","avg_lt","total_ppv","total_score"]].copy()
            show_top.columns = ["Supplier","Orders","OTD %","Avg LT (d)","PPV ($)","Score /100"]
            show_top["OTD %"]      = show_top["OTD %"].apply(lambda v: p(v))
            show_top["Avg LT (d)"] = show_top["Avg LT (d)"].apply(lambda v: n(v,1))
            show_top["PPV ($)"]    = show_top["PPV ($)"].apply(lambda v: d(v,0))
            st.dataframe(show_top, use_container_width=True, height=420, hide_index=True)
        with col_sc_r:
            st.markdown("**Bottom 20 Suppliers**")
            show_bot = sc.nsmallest(20,"total_score")[["supplier_name","orders","otd_pct","avg_lt","total_ppv","total_score"]].copy()
            show_bot.columns = ["Supplier","Orders","OTD %","Avg LT (d)","PPV ($)","Score /100"]
            show_bot["OTD %"]      = show_bot["OTD %"].apply(lambda v: p(v))
            show_bot["Avg LT (d)"] = show_bot["Avg LT (d)"].apply(lambda v: n(v,1))
            show_bot["PPV ($)"]    = show_bot["PPV ($)"].apply(lambda v: d(v,0))
            st.dataframe(show_bot, use_container_width=True, height=420, hide_index=True)

    # PPV Quarterly Trend
    st.markdown("### \U0001f4c8 PPV Quarterly Trend \u2014 Is Overpayment Getting Better or Worse?")
    st.caption("Positive PPV = paid more than standard cost (unfavorable). Negative = favorable.")
    if "order_date" in ppv.columns and not ppv.empty:
        ppv2 = ppv.copy()
        ppv2["order_date"] = pd.to_datetime(ppv2["order_date"], errors="coerce")
        ppv2 = ppv2[ppv2["order_date"].notna()]
        ppv2["quarter"] = ppv2["order_date"].dt.to_period("Q").astype(str)
        q_trend = ppv2.groupby("quarter")["ppv"].sum().reset_index().sort_values("quarter")
        q_trend["color"] = q_trend["ppv"].apply(lambda v: RED if v > 0 else GREEN)
        fig_ppv = go.Figure()
        fig_ppv.add_trace(go.Bar(
            x=q_trend["quarter"], y=q_trend["ppv"],
            marker_color=q_trend["color"],
            text=q_trend["ppv"].apply(lambda v: d(v,0)),
            textposition="outside"))
        fig_ppv.add_hline(y=0, line_color=NAVY, line_width=1)
        fig_ppv.update_layout(height=380,
                               title="Quarterly PPV Total (Red=Unfavorable, Green=Favorable)",
                               xaxis_title="Quarter", yaxis_tickprefix="$", yaxis_tickformat=",.0f",
                               **CHART)
        st.plotly_chart(fig_ppv, use_container_width=True)

    c_p1, c_p2 = st.columns(2)
    with c_p1:
        st.markdown("**\u26a0 Unfavorable \u2014 Paid More Than Standard**")
        top_u = sup_ppv.nlargest(10,"total_ppv")
        fig_u = px.bar(top_u, x="total_ppv", y="supplier_name", orientation="h",
                        color_discrete_sequence=[RED],
                        text=top_u["total_ppv"].apply(lambda v: d(v,0)),
                        title="Top 10 — Highest Overpayment vs Standard")
        fig_u.update_traces(textposition="outside")
        fig_u.update_layout(height=380, yaxis={"categoryorder":"total ascending"},
                             xaxis_tickprefix="$", **CHART)
        st.plotly_chart(fig_u, use_container_width=True)
 
    with c_p2:
        st.markdown("**✅ Favorable — Paid Less Than Standard**")
        top_f = sup_ppv.nsmallest(10,"total_ppv")
        fig_f = px.bar(top_f, x="total_ppv", y="supplier_name", orientation="h",
                        color_discrete_sequence=[GREEN],
                        text=top_f["total_ppv"].apply(lambda v: d(v,0)),
                        title="Top 10 — Highest Underpayment vs Standard")
        fig_f.update_traces(textposition="outside")
        fig_f.update_layout(height=380, yaxis={"categoryorder":"total descending"},
                             xaxis_tickprefix="$", **CHART)
        st.plotly_chart(fig_f, use_container_width=True)

    # Do-Not-Use supplier flag
    dnu_col = "is_do_not_use" if "is_do_not_use" in sup.columns else None
    if dnu_col:
        dnu_sups = sup[sup[dnu_col].astype(str).str.upper()=="YES"][
            ["supplier_bp_code","supplier_name","country_name","city"]].copy()
        if not dnu_sups.empty:
            st.markdown("### 🚫 Do-Not-Use Suppliers")
            st.markdown(
                f"<div class='alert'>⚠ {len(dnu_sups)} suppliers are flagged DO NOT USE. "
                "Review any open POs with these vendors immediately.</div>",
                unsafe_allow_html=True)
            dnu_sups.columns = ["BP Code","Supplier Name","Country","City"]
            st.dataframe(dnu_sups, use_container_width=True, hide_index=True)


# ═══════════════════════════════════════════════════
# TAB 5 — INVENTORY & COVERAGE
# ═══════════════════════════════════════════════════
def t5_inventory(D, filt):
    st.markdown("## 📦 Inventory & Coverage Analysis")
    inv = D["inv"]
    ins = scope(filt)
 
    neg = inv[inv["economic_stock"] < 0] if "economic_stock" in inv.columns else pd.DataFrame()
    if not neg.empty:
        st.markdown(
            f"<div class='alert'>🚨 SUPPLY ALERT — {len(neg)} items have negative "
            f"economic stock (more committed to production than available). "
            f"Immediate review required.</div>",
            unsafe_allow_html=True)
 
    k1,k2,k3,k4 = st.columns(4)
    with k1: kpi("📦 Total On-Hand",   n(inv.get("inv_on_hand", pd.Series([0])).sum(),0)+" lbs", "All 4 sites")
    with k2: kpi("✅ Items With Stock", n((inv.get("inv_on_hand",pd.Series([0]))>0).sum()), "inv_on_hand > 0","kpi-g")
    with k3: kpi("⚠ Zero Stock",       n((inv.get("inv_on_hand",pd.Series([0]))==0).sum()), "No inventory","kpi-a")
    with k4: kpi("🚨 Negative Eco.",   n(len(neg)), "Crisis — more allocated than available","kpi-b")
    st.markdown("---")
 
    c_l, c_r = st.columns(2)
    with c_l:
        st.markdown("### On-Hand Inventory by Site")
        if "site_group" in inv.columns and "inv_on_hand" in inv.columns:
            site_inv = inv.groupby("site_group")["inv_on_hand"].sum().reset_index()
            fig = px.bar(site_inv, x="site_group", y="inv_on_hand",
                          color="site_group",
                          color_discrete_map={"Lake Charles":BLUE,"Green Bay":GREEN,"Dayton":AMBER,"Louisa":ORANGE},
                          text=site_inv["inv_on_hand"].apply(lambda v: n(v,0)),
                          title="On-Hand Inventory (lbs) by Site")
            fig.update_traces(textposition="outside", showlegend=False)
            fig.update_layout(height=360, yaxis_title="Pounds", xaxis_title=None, **CHART)
            st.plotly_chart(fig, use_container_width=True)
 
    with c_r:
        st.markdown("### 🚨 Negative Economic Stock — Crisis Items")
        st.caption("Economic stock = On-Hand + On-Order − Allocated. Negative = production shortage risk.")
        if not neg.empty and "item_code" in neg.columns:
            show_neg = neg.sort_values("economic_stock")[
                [c for c in ["site_group","item_code","inv_on_hand","inv_on_order",
                              "inv_allocated","economic_stock"] if c in neg.columns]
            ].head(20).copy()
            show_neg.columns = [c.replace("inv_","").replace("_"," ").title()
                                 for c in show_neg.columns]
            for c in show_neg.columns:
                if c not in ["Site Group","Item Code"]:
                    show_neg[c] = show_neg[c].apply(lambda v: n(v,0))
            st.dataframe(show_neg, use_container_width=True, height=340, hide_index=True)
        else:
            st.success("No items with negative economic stock.")
 
    # Coverage analysis
    st.markdown("### Coverage: On-Hand vs Recommended Safety Stock")
    if "item_code" in inv.columns and "inv_on_hand" in inv.columns:
        inv_grp = inv.groupby(["item_code","site_group"])["inv_on_hand"].sum().reset_index()
        ab_items = ins[ins["abc_tier"].isin(["A","B"])]
        if not ab_items.empty:
            cov = ab_items.merge(inv_grp, on=["item_code","site_group"], how="left")
            cov["inv_on_hand"]   = cov["inv_on_hand"].fillna(0)
            cov["shortfall"]     = cov["new_ss_lbs_v4"] - cov["inv_on_hand"]
            cov["days_stockout"] = np.where(
                cov["avg_daily_usage_lbs"] > 0,
                (cov["inv_on_hand"] / cov["avg_daily_usage_lbs"]).round(1), 0)
            below = cov[cov["shortfall"] > 0].sort_values("shortfall", ascending=False).head(25)
            if not below.empty:
                show_b = below[["site_group","item_code","item_description","abc_tier",
                                  "inv_on_hand","new_ss_lbs_v4","shortfall","days_stockout"]].copy()
                show_b.columns = ["Site","Item","Description","Tier",
                                   "On Hand (lbs)","Rec. SS (lbs)","Shortfall (lbs)","Days to Stockout"]
                for c in ["On Hand (lbs)","Rec. SS (lbs)","Shortfall (lbs)"]:
                    show_b[c] = show_b[c].apply(lambda v: n(v,1))
                show_b["Days to Stockout"] = show_b["Days to Stockout"].apply(lambda v: n(v,1))
                st.dataframe(show_b, use_container_width=True, height=400, hide_index=True)
            else:
                st.success("✅ All A/B-tier items meet recommended safety stock levels.")
 
    # Bulk tank utilization
    st.markdown("### 🛢 Bulk Tank Utilization (Lake Charles)")
    LC_TANKS = {
        "T25":40212,"T15":43388,"CM847":51478,"CM969":52138,"CM1115":95933,
        "P10":47942,"T10":73309,"O50":43148,"RS266":51298,"RS883":51598,
        "AP30":52437,"T11":89652,"T18":54225,"RS280":55767,"RS977":103057,
        "RS825":72000,"RS820":71432,"Z86":6000
    }
    if "site_group" in inv.columns and "inv_on_hand" in inv.columns:
        lc_inv = inv[inv["site_group"]=="Lake Charles"].set_index("item_code")["inv_on_hand"].to_dict()
        tank_rows = []
        for item, cap in LC_TANKS.items():
            oh  = float(lc_inv.get(item, 0))
            uti = min(oh/cap*100, 100) if cap > 0 else 0
            tank_rows.append({"Item":item,"On Hand (lbs)":oh,"Tank Cap 80% (lbs)":cap,"Utilization %":uti})
        df_tk = pd.DataFrame(tank_rows)
        fig_tk = px.bar(
            df_tk.sort_values("Utilization %"), x="Utilization %", y="Item",
            orientation="h", color="Utilization %",
            color_continuous_scale=[[0,GREEN],[0.5,"#FFF2CC"],[1.0,RED]],
            range_color=[0,100],
            text=df_tk.sort_values("Utilization %")["Utilization %"].apply(lambda v: f"{v:.0f}%"),
            title="LC Bulk Tank Utilization (% of 80% safety capacity)")
        fig_tk.update_traces(textposition="outside")
        fig_tk.update_layout(height=480, coloraxis_showscale=False, **CHART)
        st.plotly_chart(fig_tk, use_container_width=True)
 
 
# ═══════════════════════════════════════════════════
# TAB 6 — CONSUMPTION & SEASONALITY
# ═══════════════════════════════════════════════════
def t6_seasonality(D, filt):
    st.markdown("## 📈 Consumption & Seasonality Analysis")
    st.markdown("""
    <div class='fbox'>
    <b>Seasonality Index</b> = Average consumption in that month ÷ Overall monthly average.
    Index > 1.2 = peak demand month (stock up early).
    Index < 0.8 = low demand month.
    Helps procurement know WHEN to build inventory ahead of demand spikes.
    </div>""", unsafe_allow_html=True)
 
    cons = D["cons"]
 
    f1,f2,f3 = st.columns(3)
    with f1:
        site_s = st.selectbox("Site", ["All Sites"]+sorted(cons["site_group"].unique().tolist()), key="s_s")
    with f2:
        grp_s  = st.selectbox("Item Group", ["All","Raw Materials (LB)","Packaging (EA)"], key="s_g")
    with f3:
        yrs    = st.multiselect("Years", [2021,2022,2023,2024,2025,2026],
                                 default=[2024,2025], key="s_y")
 
    cf = cons.copy()
    if site_s != "All Sites": cf = cf[cf["site_group"] == site_s]
    if grp_s == "Raw Materials (LB)":  cf = cf[cf.get("unit_of_measure","LB") == "LB"]
    elif grp_s == "Packaging (EA)":    cf = cf[cf.get("unit_of_measure","EA") == "EA"]
    if yrs: cf = cf[cf["year"].isin(yrs)]
 
    # Monthly trend
    c_l, c_r = st.columns(2)
    with c_l:
        st.markdown("### Monthly Consumption Trend (All Sites)")
        monthly = (cons.groupby(["year","period_month"])["qty_issued"].sum().reset_index())
        monthly["date_str"] = (monthly["year"].astype(int).astype(str) + "-" +
                                monthly["period_month"].astype(int).astype(str).str.zfill(2))
        monthly = monthly.sort_values("date_str")
        fig = px.line(monthly, x="date_str", y="qty_issued",
                       title="Total Consumption by Month",
                       labels={"date_str":"Month","qty_issued":"Qty Consumed"},
                       color_discrete_sequence=[BLUE])
        fig.update_traces(line=dict(width=2))
        if len(monthly) > 2:
            xn = list(range(len(monthly)))
            z  = np.polyfit(xn, monthly["qty_issued"].fillna(0), 1)
            tr = np.poly1d(z)(xn)
            fig.add_trace(go.Scatter(x=monthly["date_str"], y=tr,
                                      name="Trend", line=dict(color=RED, dash="dash", width=2)))
        fig.update_layout(height=360, xaxis_tickangle=-45, **CHART)
        st.plotly_chart(fig, use_container_width=True)
 
    with c_r:
        st.markdown("### Seasonality Index by Month")
        if not cf.empty:
            ma = cf.groupby("period_month")["qty_issued"].mean().reset_index()
            oa = ma["qty_issued"].mean()
            ma["idx"] = (ma["qty_issued"] / oa).round(2) if oa > 0 else 1
            ma["mname"] = ma["period_month"].map(MONTH)
            colors_bar  = [RED if v>1.2 else (STEEL if v<0.8 else GREY) for v in ma["idx"]]
            fig2 = go.Figure(go.Bar(
                x=ma["mname"], y=ma["idx"], marker_color=colors_bar,
                text=ma["idx"].apply(lambda v: f"{v:.2f}"), textposition="outside"))
            fig2.add_hline(y=1.2, line_dash="dash", line_color=RED,   annotation_text="Peak (1.2)")
            fig2.add_hline(y=0.8, line_dash="dash", line_color=STEEL, annotation_text="Low (0.8)")
            fig2.update_layout(height=360,
                               title="Seasonality Index (Red=Peak · Blue=Low · Grey=Normal)",
                               yaxis_title="Index", **CHART)
            st.plotly_chart(fig2, use_container_width=True)
 
    # Year-over-Year comparison
    st.markdown("### Year-over-Year: 2023 vs 2024-2025 (overlaid)")
    yoy_23   = cons[cons["year"]==2023].groupby("period_month")["qty_issued"].sum().reset_index()
    yoy_2425 = cons[cons["year"].isin([2024,2025])].groupby("period_month")["qty_issued"].sum().reset_index()
    yoy_23["mname"]   = yoy_23["period_month"].map(MONTH)
    yoy_2425["mname"] = yoy_2425["period_month"].map(MONTH)
    fig_yoy = go.Figure()
    fig_yoy.add_trace(go.Scatter(x=yoy_23["mname"],   y=yoy_23["qty_issued"],
                                  name="2023", line=dict(color=GREY, dash="dot", width=2)))
    fig_yoy.add_trace(go.Scatter(x=yoy_2425["mname"], y=yoy_2425["qty_issued"],
                                  name="2024-2025", line=dict(color=BLUE, width=3)))
    fig_yoy.update_layout(height=360, title="Monthly Consumption: 2023 vs 2024-2025",
                           xaxis_title="Month", yaxis_title="Qty Consumed (lbs)", **CHART)
    st.plotly_chart(fig_yoy, use_container_width=True)

    # Heatmap
    st.markdown("### Seasonality Heatmap \u2014 Top 20 A-Items")
    st.caption("Red = peak demand. Blue = low demand. Green = normal. "
               "Use this to plan inventory build-up before demand spikes.")
    cons_lb = cons[cons["year"].isin([2024,2025])]
    if "unit_of_measure" in cons_lb.columns:
        cons_lb = cons_lb[cons_lb["unit_of_measure"] == "LB"]
    a_items = filt[(filt["abc_tier"]=="A") & (filt["is_bad_apple"]=="NO")]["item_code"].unique()[:20]
    heat = (cons_lb[cons_lb["item_code"].isin(a_items)]
            .groupby(["item_code","period_month"])["qty_issued"].mean().reset_index())
 
    if not heat.empty:
        piv = heat.pivot(index="item_code", columns="period_month", values="qty_issued").fillna(0)
        rm  = piv.mean(axis=1).replace(0, 1)
        for c in piv.columns:
            piv[c] = piv[c] / rm
        piv.columns = [MONTH.get(c, str(c)) for c in piv.columns]
        fig_h = px.imshow(
            piv,
            color_continuous_scale=[[0,BLUE],[0.35,LBLUE],[0.6,BG],[0.8,AMBER],[1,RED]],
            zmin=0.4, zmax=1.8, text_auto=".2f",
            title="Seasonality Index — Top 20 A-Items (2024-2025)",
            labels=dict(x="Month", y="Item Code", color="Index"))
        fig_h.update_layout(height=max(380, len(a_items)*26), **CHART)
        st.plotly_chart(fig_h, use_container_width=True)
    else:
        st.info("Not enough data for the heatmap with current selection.")
 
    # Buying Calendar
    st.markdown("### \U0001f4c5 Buying Calendar \u2014 Order Placement Deadlines for A-Items")
    st.caption("Order must be placed by: Peak Month Start \u2212 Lead Time. Sorted by urgency.")
    from datetime import date as _date, timedelta as _td
    cons24 = cons[cons["year"].isin([2024,2025])]
    if not cons24.empty and not filt.empty:
        am_c = cons24.groupby(["item_code","period_month"])["qty_issued"].mean().reset_index()
        ov_c = am_c.groupby("item_code")["qty_issued"].mean()
        am_c = am_c.merge(ov_c.rename("overall"), on="item_code")
        am_c["idx"] = am_c["qty_issued"] / am_c["overall"].replace(0,1)
        peak_mc = am_c.loc[am_c.groupby("item_code")["idx"].idxmax()][["item_code","period_month"]].copy()
        peak_mc.columns = ["item_code","peak_month_num"]
        a_items_ss = filt[(filt["abc_tier"]=="A") & (filt["is_bad_apple"]=="NO")][
            ["item_code","item_description","site_group","lead_time_used_v4"]].copy()
        cal_df = a_items_ss.merge(peak_mc, on="item_code", how="left")
        cal_df["peak_month_num"]    = cal_df["peak_month_num"].fillna(6).astype(int)
        cal_df["lead_time_used_v4"] = cal_df["lead_time_used_v4"].fillna(30).astype(int)
        today_c = _date.today()
        rows_cal = []
        for _, row in cal_df.iterrows():
            peak_mo = int(row["peak_month_num"]); lt = int(row["lead_time_used_v4"]); yr = today_c.year
            try:
                peak_start = _date(yr, peak_mo, 1)
                if peak_start < today_c: peak_start = _date(yr+1, peak_mo, 1)
                order_by  = peak_start - _td(days=lt)
                days_left = (order_by - today_c).days
            except Exception: continue
            rows_cal.append({
                "Site": row["site_group"], "Item Code": row["item_code"],
                "Description": str(row["item_description"])[:40],
                "Peak Month": MONTH.get(peak_mo, str(peak_mo)),
                "Lead Time (d)": f"{lt}d",
                "Order By Date": order_by.strftime("%b %d, %Y"),
                "Days Until Deadline": days_left,
                "Status": ("\U0001f534 OVERDUE" if days_left < 0 else
                           ("\U0001f7e1 THIS MONTH" if days_left <= 30 else "\u2705 OK")),
            })
        if rows_cal:
            df_cal = pd.DataFrame(rows_cal).sort_values("Days Until Deadline")
            st.dataframe(df_cal, use_container_width=True, height=420, hide_index=True)
            dl_excel(df_cal, "Carboline_Buying_Calendar.xlsx")

    # Seasonal items
    st.markdown("### Items Flagged as Seasonal \u2014 Peak Months")
    cons24 = cons[cons["year"].isin([2024,2025])]
    if not cons24.empty:
        am = cons24.groupby(["item_code","period_month"])["qty_issued"].mean().reset_index()
        ov = am.groupby("item_code")["qty_issued"].mean()
        am = am.merge(ov.rename("overall"), on="item_code")
        am["idx"] = (am["qty_issued"] / am["overall"].replace(0,1)).round(2)
        seasonal = (am[am["idx"] > 1.2]
                    .groupby("item_code")
                    .agg(peak_months=("period_month",
                                      lambda x: ", ".join(MONTH.get(m,"?") for m in sorted(x))),
                         max_idx=("idx","max"))
                    .reset_index().sort_values("max_idx", ascending=False).head(30))
        seasonal.columns = ["Item Code","Peak Months","Max Seasonality Index"]
        seasonal["Max Seasonality Index"] = seasonal["Max Seasonality Index"].apply(lambda v: f"{v:.2f}")
        st.dataframe(seasonal, use_container_width=True, height=380, hide_index=True)
 
    # Growth
    st.markdown("### Top 20 Fastest-Growing Items (2023 vs 2024-2025)")
    c23   = cons[cons["year"]==2023].groupby("item_code")["qty_issued"].sum()
    c2425 = cons[cons["year"].isin([2024,2025])].groupby("item_code")["qty_issued"].sum()
    grw   = pd.concat([c23.rename("y2023"), c2425.rename("y2425")], axis=1).dropna()
    grw   = grw[grw["y2023"] > 1000].copy()
    grw["growth_pct"] = ((grw["y2425"] - grw["y2023"]) / grw["y2023"] * 100).round(1)
    grw   = grw.nlargest(20,"growth_pct").reset_index()
    grw.columns = ["Item Code","Lbs 2023","Lbs 2024-25","Growth %"]
    grw["Lbs 2023"]    = grw["Lbs 2023"].apply(lambda v: n(v,0))
    grw["Lbs 2024-25"] = grw["Lbs 2024-25"].apply(lambda v: n(v,0))
    grw["Growth %"]    = grw["Growth %"].apply(lambda v: f"{v:+.1f}%")
    st.dataframe(grw, use_container_width=True, height=380, hide_index=True)
 
 
# ═══════════════════════════════════════════════════
# TAB 7 — D-TIER & MTO STRATEGY
# ═══════════════════════════════════════════════════
def t7_dtier(D):
    st.markdown("## 🗑 D-Tier Items & MTO Powerhouse Strategy")
 
    df_d  = D["d_tier"]
    df_gm = D["good_missing"]
    df_gs = D["good_stock"]
    df_ba = D["bad_apple"]
    df_ss = D["ss"]
 
    ga_under = (df_ss["audit_flag_v4"].str.contains("GOOD APPLE.*UNDER", na=False)).sum()
 
    c_ga, c_ba = st.columns(2)
    with c_ga:
        st.markdown(f"""
        <div style="background:{AMBER};border-radius:10px;padding:20px;color:{WHITE};">
          <div style="font-size:17px;font-weight:800;">⭐ Good Apple — Always Stock</div>
          <div style="font-size:13px;margin-top:8px;line-height:1.7;">
            Total on priority list: <b>146 items</b><br>
            In LN with parameters: <b>59 items</b><br>
            Missing LN parameters: <b>93 items</b> — action needed<br>
            Currently understocked: <b>{ga_under} items</b>
          </div>
          <div style="font-size:11px;margin-top:8px;opacity:0.9;">
            These raws feed ~90% of Carboline's make-to-order portfolio.
            Always maintain safety stock for these items.
          </div>
        </div>""", unsafe_allow_html=True)
 
    with c_ba:
        st.markdown(f"""
        <div style="background:{GREY};border-radius:10px;padding:20px;color:{WHITE};">
          <div style="font-size:17px;font-weight:800;">🗑 Bad Apple — Phase Out</div>
          <div style="font-size:13px;margin-top:8px;line-height:1.7;">
            Total on phase-out list: <b>1,144 items</b><br>
            Still have LN parameters: <b>94 items</b> — needs removal<br>
            Already inactive in LN: <b>1,050 items</b><br>
            Safety Stock forced to: <b>Zero for all 94</b>
          </div>
          <div style="font-size:11px;margin-top:8px;opacity:0.9;">
            LN ordering parameters should be removed to eliminate
            phantom planned purchase orders.
          </div>
        </div>""", unsafe_allow_html=True)
 
    st.markdown("<br>", unsafe_allow_html=True)
 
    c_l, c_r = st.columns(2)
    with c_l:
        st.markdown("### D-Tier Breakdown")
        never = (df_d["total_lbs_all_years"] == 0).sum()
        slow  = (df_d["lbs_2021_2023"] > 0).sum()
        rop_f = df_d.get("has_rop_in_ln", pd.Series(dtype=str)).astype(str).str.startswith("YES").sum()
        fig = go.Figure(go.Pie(
            labels=["Never Consumed","Slow Movers\n(zero since 2024)","Have ROP in LN"],
            values=[max(never-slow,0), slow, rop_f],
            hole=0.5, marker_colors=[GREY, AMBER, RED],
            textinfo="label+value+percent", textfont=dict(size=12)))
        fig.update_layout(
            height=360,
            annotations=[dict(text="1,760\nD-Tier", x=0.5, y=0.5,
                               font_size=13, showarrow=False, font_color=NAVY)],
            **CHART)
        st.plotly_chart(fig, use_container_width=True)
 
    with c_r:
        st.markdown("### ⚠ Items With ROP Still Active (13 items)")
        st.caption("Zero consumption but still have ordering parameters — "
                   "may generate unnecessary planned purchase orders.")
        if "has_rop_in_ln" in df_d.columns:
            rop_items = df_d[df_d["has_rop_in_ln"].astype(str).str.startswith("YES")][
                ["item_code","item_description","item_group","total_lbs_all_years"]
            ].copy()
            rop_items.columns = ["Item Code","Description","Group","Total LBS All Years"]
            rop_items["Total LBS All Years"] = rop_items["Total LBS All Years"].apply(lambda v: n(v,0))
            st.dataframe(rop_items, use_container_width=True, height=320, hide_index=True)
 
    # Full D-tier
    st.markdown("### Full D-Tier Item List (1,760 items)")
    filt_d = st.radio("Show:", ["All D-Tier","Never Consumed","Slow Movers"], horizontal=True)
    show_d = df_d.copy()
    if filt_d == "Never Consumed":     show_d = df_d[df_d["total_lbs_all_years"] == 0]
    elif filt_d == "Slow Movers":      show_d = df_d[df_d["lbs_2021_2023"] > 0]
 
    cols_d = [c for c in ["item_code","item_description","item_group","item_group_label",
                           "planning_signal","standard_cost_usd","total_lbs_all_years",
                           "lbs_2021_2023","last_year_consumed","has_rop_in_ln","d_tier_reason"]
              if c in show_d.columns]
    disp_d = show_d[cols_d].copy()
    rename_d = {
        "item_code":"Item Code","item_description":"Description",
        "item_group":"Group","item_group_label":"Group Label",
        "planning_signal":"Signal","standard_cost_usd":"Std Cost ($)",
        "total_lbs_all_years":"Total LBS","lbs_2021_2023":"LBS 2021-23",
        "last_year_consumed":"Last Year","has_rop_in_ln":"ROP in LN?",
        "d_tier_reason":"Reason"
    }
    disp_d.rename(columns=rename_d, inplace=True)
    for c in ["Total LBS","LBS 2021-23"]:
        if c in disp_d.columns:
            disp_d[c] = disp_d[c].apply(lambda v: n(v,0))
    if "Std Cost ($)" in disp_d.columns:
        disp_d["Std Cost ($)"] = disp_d["Std Cost ($)"].apply(lambda v: d(v,4))
    st.dataframe(disp_d, use_container_width=True, height=400, hide_index=True)
    dl_excel(disp_d, "Carboline_D_Tier_List.xlsx")
 
    # Good Apple missing
    st.markdown(f"### ⭐ Good Apple Items Missing LN Parameters (93 items)")
    st.caption("On the priority list but have no ordering parameters in LN. "
               "LN will never auto-generate purchase orders for these.")
    if not df_gm.empty:
        gm2 = df_gm[["item_code","std_cost","destination","lead_time_days"]].copy()
        gm2.columns = ["Item Code","Std Cost ($)","Destination Site","Manager Lead Time (days)"]
        gm2["Std Cost ($)"] = gm2["Std Cost ($)"].apply(lambda v: d(v,4))
        gm2["Manager Lead Time (days)"] = gm2["Manager Lead Time (days)"].apply(lambda v: n(v,1))
        gm2["Action"] = "Add Item Ordering parameters in LN"
        st.dataframe(gm2, use_container_width=True, height=320, hide_index=True)
        dl_excel(gm2, "Carboline_Good_Apple_Missing_LN.xlsx")
 
    # Bad Apple removal
    st.markdown("### 🗑 Bad Apple Items — Remove From LN (94 items)")
    st.caption("Safety stock forced to zero. LN ordering parameters should be removed "
               "by the procurement team to eliminate phantom planned orders.")
    if not df_ba.empty:
        ba_cols = [c for c in ["Site","Item Code","Description","ABC Tier",
                                "Current LN ROP\n(was active)","MOQ","Old $ Risk"]
                   if c in df_ba.columns]
        ba2 = df_ba[ba_cols].copy() if ba_cols else df_ba.copy()
        # Clean col names
        ba2.columns = [c.replace("\n"," ") for c in ba2.columns]
        if "Old $ Risk" in ba2.columns:
            ba2["Old $ Risk"] = ba2["Old $ Risk"].apply(lambda v: d(v,2))
        st.dataframe(ba2, use_container_width=True, height=320, hide_index=True)
        dl_excel(ba2, "Carboline_Bad_Apple_Remove.xlsx")
 
 


# ═══════════════════════════════════════════════════
# TAB 1 ENHANCEMENTS — Site Summary + Data Freshness
# ═══════════════════════════════════════════════════
# (Injected inline via patching t1_overview below)

# ═══════════════════════════════════════════════════
# TAB 4 ENHANCEMENTS — PPV Trend + Vendor Scorecard
# ═══════════════════════════════════════════════════
# (Injected inline via patching t4_supplier below)

# ═══════════════════════════════════════════════════
# TAB 6 ENHANCEMENTS — Buying Calendar + YoY chart
# ═══════════════════════════════════════════════════
# (Injected inline via patching t6_seasonality below)


# ═══════════════════════════════════════════════════
# TAB 8 — DEMAND FORECASTING (Prophet)
# ═══════════════════════════════════════════════════
def t8_forecast(D):
    st.markdown("## \U0001f52e Demand Forecasting \u2014 Prophet Model")
    st.markdown("""
    <div class='fbox'>
    <b>Prophet Forecasting Engine</b> \u2014 Predicts future material consumption using historical data.
    Mode 1 uses existing ERP consumption history. Mode 2 accepts your own CSV or Excel file (up to 500\u00a0MB).
    95% confidence interval. Requires minimum 12 months of history.
    </div>""", unsafe_allow_html=True)

    try:
        from prophet import Prophet
    except ImportError:
        st.error("Prophet is not installed. Add prophet==1.3.0 to requirements.txt and redeploy.")
        st.code("pip install prophet==1.3.0")
        return

    cons  = D["cons"]
    df_ss = D["ss"]

    mode = st.radio("Forecasting Mode",
                     ["Mode 1 \u2014 Existing Consumption Data", "Mode 2 \u2014 Upload Your Own Data"],
                     horizontal=True)

    @st.cache_resource(show_spinner=False)
    def fit_prophet(data_key):
        df_train = pd.DataFrame(list(data_key), columns=["ds","y"])
        df_train["ds"] = pd.to_datetime(df_train["ds"])
        df_train["y"]  = df_train["y"].astype(float)
        model = Prophet(
            seasonality_mode='multiplicative',
            yearly_seasonality=True,
            weekly_seasonality=False,
            daily_seasonality=False,
            changepoint_prior_scale=0.1,
            seasonality_prior_scale=10,
            interval_width=0.95
        )
        model.add_seasonality(name='monthly', period=30.5, fourier_order=5)
        model.fit(df_train)
        return model

    def run_forecast(prophet_df, item_label, site_label, horizon_months,
                     lt_days=30.0, buf_factor=1.0, curr_rop=0.0):
        if len(prophet_df) < 12:
            st.warning("\u26a0 Not enough history for reliable forecast. Need at least 12 months of data.")
            return
        prophet_df = prophet_df.sort_values("ds").reset_index(drop=True)
        cutoff   = prophet_df["ds"].max() - pd.DateOffset(months=3)
        train_df = prophet_df[prophet_df["ds"] <= cutoff]
        val_df   = prophet_df[prophet_df["ds"] >  cutoff]

        with st.spinner("\U0001f52e Building forecast model \u2014 this takes about 10 seconds..."):
            data_key = tuple(zip(train_df["ds"].astype(str).tolist(), train_df["y"].tolist()))
            model    = fit_prophet(data_key)

        future = model.make_future_dataframe(periods=horizon_months, freq="MS")
        fc     = model.predict(future)
        today_ts = pd.Timestamp.today()

        mape = None
        if not val_df.empty:
            val_pred = fc[fc["ds"].isin(val_df["ds"])][["ds","yhat"]].merge(val_df, on="ds")
            if not val_pred.empty and val_pred["y"].sum() > 0:
                mape = (np.abs(val_pred["yhat"] - val_pred["y"]) /
                        val_pred["y"].replace(0, np.nan)).dropna().mean() * 100

        st.markdown(f"#### \U0001f4c8 Forecast: {item_label} @ {site_label} \u2014 Next {horizon_months} Months")
        hist = prophet_df[prophet_df["ds"] <= today_ts]
        fut  = fc[fc["ds"] > today_ts]

        fig1 = go.Figure()
        fig1.add_trace(go.Scatter(x=hist["ds"], y=hist["y"], mode="markers+lines",
                                   name="Historical", marker=dict(color=NAVY, size=5),
                                   line=dict(color=NAVY, width=1.5)))
        fig1.add_trace(go.Scatter(x=fut["ds"], y=fut["yhat"],
                                   name="Forecast", line=dict(color=BLUE, width=2.5)))
        fig1.add_trace(go.Scatter(
            x=pd.concat([fut["ds"], fut["ds"].iloc[::-1]]),
            y=pd.concat([fut["yhat_upper"], fut["yhat_lower"].iloc[::-1]]),
            fill="toself", fillcolor="rgba(0,48,135,0.12)",
            line=dict(color="rgba(255,255,255,0)"),
            name="95% CI", hoverinfo="skip"))
        fig1.add_vline(x=today_ts.timestamp()*1000, line_dash="dash",
                        line_color=RED, annotation_text="Today")
        fig1.update_layout(height=420,
            title=f"Demand Forecast \u2014 {item_label} at {site_label} \u2014 Next {horizon_months} Months",
            xaxis_title="Date", yaxis_title="Qty (lbs)", **CHART)
        st.plotly_chart(fig1, use_container_width=True)

        trend_vals = fc["trend"]
        trend_dir  = ("GROWING"  if trend_vals.iloc[-1] > trend_vals.iloc[0] * 1.02 else
                      "DECLINING" if trend_vals.iloc[-1] < trend_vals.iloc[0] * 0.98 else "STABLE")
        trend_clr  = GREEN if trend_dir == "GROWING" else (RED if trend_dir == "DECLINING" else GREY)
        col_t_l, col_t_r = st.columns([3,1])
        with col_t_l:
            fig2 = go.Figure(go.Scatter(x=fc["ds"], y=fc["trend"],
                                         line=dict(color=trend_clr, width=2.5), name="Trend"))
            fig2.update_layout(height=300, title=f"Underlying Trend \u2014 {trend_dir}",
                                xaxis_title="Date", yaxis_title="Trend (lbs)", **CHART)
            st.plotly_chart(fig2, use_container_width=True)
        with col_t_r:
            st.markdown(f"""
            <div style="background:{trend_clr};border-radius:10px;padding:20px;
                        color:{WHITE};text-align:center;margin-top:40px;">
              <div style="font-size:13px;font-weight:700;">Trend</div>
              <div style="font-size:24px;font-weight:900;margin-top:8px;">{trend_dir}</div>
            </div>""", unsafe_allow_html=True)

        seas = None
        try:
            seas = model.predict(pd.DataFrame({
                "ds": pd.date_range("2024-01-01", periods=12, freq="MS")}))
            seas["month"] = seas["ds"].dt.month
            seas["idx"]   = seas["yearly"] + 1
            seas["mname"] = seas["month"].map(MONTH)
            s_colors = [RED if v>1.2 else (BLUE if v<0.8 else GREY) for v in seas["idx"]]
            fig3 = go.Figure(go.Bar(x=seas["mname"], y=seas["idx"],
                                     marker_color=s_colors,
                                     text=seas["idx"].apply(lambda v: f"{v:.2f}"),
                                     textposition="outside"))
            fig3.add_hline(y=1.2, line_dash="dash", line_color=RED,   annotation_text="Peak")
            fig3.add_hline(y=0.8, line_dash="dash", line_color=STEEL, annotation_text="Low")
            fig3.update_layout(height=300,
                                title="Seasonal Pattern \u2014 When Demand Is High vs Low",
                                yaxis_title="Seasonality Index", **CHART)
            st.plotly_chart(fig3, use_container_width=True)
        except Exception:
            pass

        if mape is not None:
            clr_m = GREEN if mape < 15 else (AMBER if mape < 30 else RED)
            grade = "Excellent" if mape < 15 else ("Good" if mape < 30 else "Use with caution")
            st.markdown(
                f"<div style='background:{clr_m};color:{WHITE};border-radius:8px;"
                f"padding:10px 16px;font-weight:700;font-size:14px;margin:8px 0;'>"
                f"Model Accuracy: {mape:.1f}% MAPE (lower is better) \u00b7 "
                f"Last 3 months validation \u00b7 {grade}</div>",
                unsafe_allow_html=True)

        future_fc   = fc[fc["ds"] > today_ts].head(6)
        avg_monthly = future_fc["yhat"].clip(lower=0).mean() if not future_fc.empty else 0
        peak_row    = future_fc.loc[future_fc["yhat"].idxmax()] if not future_fc.empty else None
        peak_month  = MONTH.get(peak_row["ds"].month, "") if peak_row is not None else "N/A"
        peak_val    = float(peak_row["yhat"]) if peak_row is not None else 0
        rec_ss      = peak_val * buf_factor * (lt_days / 30)
        gap         = rec_ss - curr_rop
        action      = ("\U0001f4c8 INCREASE ORDER FREQUENCY" if gap > 0.05 * max(rec_ss,1) else
                       ("\U0001f4c9 REDUCE ORDER QTY" if gap < -0.05 * max(rec_ss,1) else
                        "\u2705 STOCK IS SUFFICIENT"))
        st.markdown(f"""
        <div class='rec-box'>
          <div style="font-size:16px;font-weight:800;margin-bottom:12px;">\U0001f4e6 PROCUREMENT RECOMMENDATION</div>
          <div style="font-size:13px;line-height:2;">
            Forecasted avg monthly demand (next 6 months): <b>{n(avg_monthly,0)} lbs</b><br>
            Peak demand month: <b>{peak_month}</b> \u2014 <b>{n(peak_val,0)} lbs</b><br>
            Recommended safety stock for peak period: <b>{n(rec_ss,0)} lbs</b>
            (peak \u00d7 {buf_factor:.1f}\u00d7 buffer \u00d7 lead time adj.)<br>
            Current LN ROP: <b>{n(curr_rop,0)} lbs</b><br>
            Gap: <b>{"+"+n(gap,0) if gap>=0 else n(gap,0)} lbs</b><br>
            Action: <b>{action}</b>
          </div>
        </div>""", unsafe_allow_html=True)

        dl_df = fc[fc["ds"] > today_ts][["ds","yhat","yhat_lower","yhat_upper"]].copy()
        if seas is not None:
            seas_map = dict(zip(seas["month"], seas["idx"]))
            dl_df["seasonality_index"] = dl_df["ds"].dt.month.map(seas_map)
        else:
            dl_df["seasonality_index"] = np.nan
        dl_df.columns = ["Date","Forecasted Qty (lbs)","Lower Bound (95%)","Upper Bound (95%)","Seasonality Index"]
        dl_df["Forecasted Qty (lbs)"] = dl_df["Forecasted Qty (lbs)"].clip(lower=0).round(0)
        dl_df["Lower Bound (95%)"]    = dl_df["Lower Bound (95%)"].clip(lower=0).round(0)
        dl_df["Upper Bound (95%)"]    = dl_df["Upper Bound (95%)"].round(0)
        safe_label = "".join(c for c in item_label if c.isalnum() or c in "_-")[:30]
        dl_excel(dl_df, f"Carboline_Forecast_{safe_label}.xlsx")

    if "Mode 1" in mode:
        all_items = sorted(cons["item_code"].dropna().unique().tolist())
        all_sites = ["All Sites"] + sorted(cons["site_group"].dropna().unique().tolist())
        col_m1a, col_m1b, col_m1c = st.columns(3)
        with col_m1a:
            sel_item = st.selectbox("Item Code", all_items, key="fc_item")
        with col_m1b:
            sel_site = st.selectbox("Site", all_sites, key="fc_site")
        with col_m1c:
            sel_hor  = st.selectbox("Forecast Horizon (months)", [3,6,12,18,24], index=2, key="fc_hor")

        if st.button("\U0001f52e Build Forecast", key="fc_go"):
            cf = cons[cons["item_code"] == sel_item].copy()
            if sel_site != "All Sites":
                cf = cf[cf["site_group"] == sel_site]
            if cf.empty:
                st.warning("No consumption data for this item/site combination.")
                return
            monthly = cf.groupby(["year","period_month"])["qty_issued"].sum().reset_index()
            monthly["ds"] = pd.to_datetime(
                monthly["year"].astype(int).astype(str) + "-" +
                monthly["period_month"].astype(int).astype(str).str.zfill(2) + "-01")
            all_months = pd.date_range(monthly["ds"].min(), monthly["ds"].max(), freq="MS")
            prophet_df = (pd.DataFrame({"ds": all_months})
                          .merge(monthly[["ds","qty_issued"]].rename(columns={"qty_issued":"y"}),
                                 on="ds", how="left").fillna({"y":0}))
            item_ss  = df_ss[df_ss["item_code"] == sel_item]
            lt_days  = float(item_ss["lead_time_used_v4"].mean()) if not item_ss.empty else 30.0
            buf_fac  = float(item_ss["buffer_factor_v4"].mean())  if not item_ss.empty else 1.0
            curr_rop = float(item_ss["current_rop_in_ln"].mean()) if not item_ss.empty else 0.0
            run_forecast(prophet_df, sel_item, sel_site, sel_hor, lt_days, buf_fac, curr_rop)
    else:
        st.markdown("#### Upload Consumption Data")
        st.caption("CSV or Excel \u00b7 Up to 500\u00a0MB \u00b7 70,000+ rows supported.")
        up_file = st.file_uploader("Upload CSV or Excel file", type=["csv","xlsx","xls"], key="fc_upload")
        if up_file is not None:
            try:
                df_up = pd.read_csv(up_file) if up_file.name.endswith(".csv") else pd.read_excel(up_file)
                all_cols = df_up.columns.tolist()
                col_u1, col_u2, col_u3 = st.columns(3)
                with col_u1:
                    date_col = st.selectbox("Date Column",     all_cols, key="fc_dcol")
                with col_u2:
                    qty_col  = st.selectbox("Quantity Column", all_cols, key="fc_qcol")
                with col_u3:
                    hor_u    = st.selectbox("Horizon (months)", [3,6,12,18,24], index=2, key="fc_hor2")
                df_up["_ds"] = pd.to_datetime(df_up[date_col], errors="coerce")
                df_up["_y"]  = pd.to_numeric(df_up[qty_col],   errors="coerce").fillna(0)
                df_up = df_up[df_up["_ds"].notna()]
                st.success(f"\u2705 {len(df_up):,} rows loaded \u00b7 "
                           f"{df_up['_ds'].min().strftime('%b %Y')} \u2192 {df_up['_ds'].max().strftime('%b %Y')}")
                monthly_u = (df_up.groupby(pd.Grouper(key="_ds", freq="MS"))["_y"].sum()
                             .reset_index().rename(columns={"_ds":"ds","_y":"y"}))
                all_months_u = pd.date_range(monthly_u["ds"].min(), monthly_u["ds"].max(), freq="MS")
                prophet_df_u = (pd.DataFrame({"ds": all_months_u})
                                .merge(monthly_u, on="ds", how="left").fillna({"y":0}))
                if st.button("\U0001f52e Build Forecast from Uploaded Data", key="fc_go2"):
                    run_forecast(prophet_df_u, qty_col, "Uploaded Data", hor_u)
            except Exception as e:
                st.error(f"Failed to read uploaded file: {e}")


# ═══════════════════════════════════════════════════
# TAB 9 — COUNTRY OF ORIGIN & TARIFF RISK
# ═══════════════════════════════════════════════════
def t9_tariff(D):
    st.markdown("## \U0001f30d Country of Origin & Tariff Risk Tracker")
    st.markdown("""
    <div class='fbox'>
    <b>Supply chain risk by geography.</b> Analyze spend concentration by country,
    identify single-source dependencies, and model tariff cost impacts.
    </div>""", unsafe_allow_html=True)

    sup  = D["sup"]
    po   = D["po"]
    cost = D["cost"]

    if sup.empty or po.empty:
        st.warning("Supplier or PO data unavailable."); return

    po_recv = po[po["po_status"]=="RECEIVED"].copy() if "po_status" in po.columns else po.copy()
    po_recv = po_recv.merge(
        sup[["supplier_bp_code","supplier_name","country_name","country_code","is_us_supplier"]],
        on="supplier_bp_code", how="left")
    po_recv = po_recv.merge(cost[["item_code","standard_cost_usd"]], on="item_code", how="left")
    po_recv["standard_cost_usd"] = pd.to_numeric(po_recv["standard_cost_usd"], errors="coerce").fillna(0)
    po_recv["ordered_qty"]       = pd.to_numeric(po_recv["ordered_qty"],        errors="coerce").fillna(0)
    po_recv["spend"]             = po_recv["ordered_qty"] * po_recv["standard_cost_usd"]
    po_recv["country_name"]      = po_recv["country_name"].fillna("Unknown")

    country_grp = po_recv.groupby("country_name").agg(
        suppliers=("supplier_bp_code","nunique"),
        items=("item_code","nunique"),
        total_spend=("spend","sum")
    ).reset_index()
    country_grp["pct_spend"] = (country_grp["total_spend"] /
                                 country_grp["total_spend"].sum() * 100).round(2)
    country_grp["risk_level"] = country_grp.apply(
        lambda r: "HIGH"   if r["suppliers"] <= 1 or r["pct_spend"] >= 30 else
                  ("MEDIUM" if r["pct_spend"] >= 10 else "LOW"), axis=1)
    country_grp = country_grp.sort_values("total_spend", ascending=False)

    top_country = country_grp.iloc[0]["country_name"] if not country_grp.empty else "N/A"
    top_pct     = country_grp.iloc[0]["pct_spend"]    if not country_grp.empty else 0
    high_risk   = (country_grp["risk_level"]=="HIGH").sum()
    intl_spend  = country_grp[country_grp["country_name"]!="United States"]["total_spend"].sum()

    k1,k2,k3,k4 = st.columns(4)
    with k1: kpi("\U0001f30d Countries Sourced",   n(len(country_grp)),  "Active source countries")
    with k2: kpi("\U0001f534 High-Risk Countries", n(high_risk),          "Single-source or >30% spend","kpi-b")
    with k3: kpi("\U0001f3c6 Top Country",         f"{top_pct:.1f}%",     top_country)
    with k4: kpi("\U0001f310 International Spend", d(intl_spend,0),       "Non-US sourcing value","kpi-a")
    st.markdown("---")

    country_iso_map = {
        "United States":"USA","Canada":"CAN","United Kingdom":"GBR","Germany":"DEU",
        "France":"FRA","Italy":"ITA","Spain":"ESP","China":"CHN","Japan":"JPN",
        "South Korea":"KOR","Mexico":"MEX","Brazil":"BRA","Australia":"AUS",
        "India":"IND","Netherlands":"NLD","Belgium":"BEL","Switzerland":"CHE",
        "Norway":"NOR","Colombia":"COL","Taiwan":"TWN","Sweden":"SWE","Poland":"POL",
    }
    cg_map = country_grp.copy()
    cg_map["iso3"] = cg_map["country_name"].map(country_iso_map)
    cg_valid = cg_map[cg_map["iso3"].notna()]

    st.markdown("### \U0001f5fa Spend by Country \u2014 World Map")
    if not cg_valid.empty:
        fig_map = px.choropleth(
            cg_valid, locations="iso3", color="total_spend",
            hover_name="country_name",
            hover_data={"total_spend":":.0f","pct_spend":":.2f","suppliers":True,"items":True,"iso3":False},
            color_continuous_scale=["#FFF0F0","#C41230"],
            title="Total Sourcing Spend by Country (darker red = higher spend)")
        fig_map.update_layout(height=480, geo=dict(showframe=False, showcoastlines=True),
                               coloraxis_colorbar=dict(title="Spend ($)"), **CHART)
        st.plotly_chart(fig_map, use_container_width=True)

    st.markdown("### Spend & Risk by Country")
    disp_cg = country_grp.copy()
    disp_cg.columns = ["Country","Suppliers","Items Sourced","Total Spend ($)","% of Total","Risk Level"]
    disp_cg["Total Spend ($)"] = disp_cg["Total Spend ($)"].apply(lambda v: d(v,0))
    disp_cg["% of Total"]      = disp_cg["% of Total"].apply(lambda v: p(v))
    st.dataframe(disp_cg, use_container_width=True, height=380, hide_index=True)
    dl_excel(disp_cg, "Carboline_Tariff_Risk_Report.xlsx")

    st.markdown("### \U0001f4ca Tariff Scenario Calculator")
    country_list = sorted(country_grp[country_grp["country_name"]!="United States"]["country_name"].tolist())
    if not country_list:
        st.info("No international suppliers found."); return
    col_tc1, col_tc2 = st.columns([2,1])
    with col_tc1:
        sel_country = st.selectbox("Select Country", country_list, key="tariff_country")
    with col_tc2:
        tariff_pct = st.number_input("Tariff Rate (%)", min_value=0.0, max_value=200.0,
                                      value=25.0, step=5.0, key="tariff_pct")
    if sel_country and tariff_pct > 0:
        ci = po_recv[po_recv["country_name"] == sel_country]
        if not ci.empty:
            total_impact = ci["spend"].sum() * (tariff_pct / 100)
            n_items = ci["item_code"].nunique(); n_sups = ci["supplier_bp_code"].nunique()
            impact_color = RED if total_impact > 1_000_000 else AMBER
            st.markdown(f"""
            <div style="background:{impact_color};color:{WHITE};border-radius:10px;
                        padding:18px 24px;font-size:15px;font-weight:700;margin:12px 0;">
              A {tariff_pct:.0f}% tariff on <b>{sel_country}</b> would increase raw material costs
              by <b>{d(total_impact,0)}</b>, affecting <b>{n_items:,} items</b>
              from <b>{n_sups:,} suppliers</b>.
            </div>""", unsafe_allow_html=True)
            item_impact = ci.groupby("item_code").agg(spend=("spend","sum")).reset_index()
            item_impact["tariff_cost"] = item_impact["spend"] * (tariff_pct / 100)
            item_impact = item_impact.nlargest(10,"tariff_cost").copy()
            item_impact.columns = ["Item Code","Spend ($)","Tariff Impact ($)"]
            item_impact["Spend ($)"]         = item_impact["Spend ($)"].apply(lambda v: d(v,0))
            item_impact["Tariff Impact ($)"] = item_impact["Tariff Impact ($)"].apply(lambda v: d(v,0))
            st.markdown("**Top 10 Most Impacted Items:**")
            st.dataframe(item_impact, use_container_width=True, hide_index=True)
        else:
            st.info(f"No received PO data found for {sel_country}.")


# ═══════════════════════════════════════════════════
# TAB 10 — OPEN PO MONITOR
# ═══════════════════════════════════════════════════
def t10_open_po(D):
    st.markdown("## \U0001f4cb Open PO Monitor")
    st.markdown("""
    <div class='fbox'>
    <b>Traffic light status:</b> \U0001f7e2 On track (receipt date not yet passed) \u00b7
    \U0001f7e1 Overdue 1\u201330 days \u00b7 \U0001f534 Seriously overdue 30+ days.
    </div>""", unsafe_allow_html=True)

    po  = D["po"]
    sup = D["sup"]

    open_po = po[po["po_status"]=="OPEN"].copy() if "po_status" in po.columns else pd.DataFrame()
    if open_po.empty:
        st.info("No open purchase orders found."); return

    open_po = open_po.merge(sup[["supplier_bp_code","supplier_name"]], on="supplier_bp_code", how="left")
    open_po["supplier_name"]       = open_po["supplier_name"].fillna(open_po["supplier_bp_code"])
    open_po["planned_receipt_date"] = pd.to_datetime(open_po["planned_receipt_date"], errors="coerce")
    open_po["order_date"]           = pd.to_datetime(open_po["order_date"],            errors="coerce")
    today_ts = pd.Timestamp.today().normalize()
    open_po["days_overdue"] = (today_ts - open_po["planned_receipt_date"]).dt.days.fillna(0).astype(int)
    open_po["status_color"] = open_po["days_overdue"].apply(
        lambda dv: "\U0001f7e2 On Track" if dv <= 0 else
                   ("\U0001f7e1 Overdue (<30d)" if dv <= 30 else "\U0001f534 Seriously Overdue"))
    open_po["estimated_value"] = (
        pd.to_numeric(open_po.get("ordered_qty",0), errors="coerce").fillna(0) *
        pd.to_numeric(open_po.get("unit_price",0),  errors="coerce").fillna(0))

    n_open    = len(open_po)
    n_overdue = (open_po["days_overdue"] > 0).sum()
    total_val = open_po["estimated_value"].sum()
    avg_over  = open_po[open_po["days_overdue"]>0]["days_overdue"].mean()

    k1,k2,k3,k4 = st.columns(4)
    with k1: kpi("\U0001f4cb Total Open POs",   n(n_open),    "Not yet received")
    with k2: kpi("\u26a0 Overdue POs",          n(n_overdue), "Past planned receipt date","kpi-b")
    with k3: kpi("\U0001f4b0 Total Open Value", d(total_val,0), "Ordered \u00d7 unit price","kpi-a")
    with k4: kpi("\U0001f4c5 Avg Days Overdue", f"{avg_over:.0f}d" if pd.notna(avg_over) else "\u2014",
                  "Overdue POs only","kpi-b")
    st.markdown("---")

    fa,fb,fc_col = st.columns(3)
    with fa:
        sites_po = ["All Sites"] + sorted(open_po.get("site_group", pd.Series()).dropna().unique().tolist())
        sf_po = st.selectbox("Site", sites_po, key="po_site")
    with fb:
        status_f = st.selectbox("Status", ["All","Overdue Only","On Track"], key="po_status_f")
    with fc_col:
        rows_po = st.selectbox("Max Rows", [50,100,200,"All"], key="po_rows")

    disp_po = open_po.copy()
    if sf_po != "All Sites" and "site_group" in disp_po.columns:
        disp_po = disp_po[disp_po["site_group"] == sf_po]
    if status_f == "Overdue Only": disp_po = disp_po[disp_po["days_overdue"] > 0]
    elif status_f == "On Track":   disp_po = disp_po[disp_po["days_overdue"] <= 0]
    disp_po = disp_po.sort_values("days_overdue", ascending=False)
    if rows_po != "All": disp_po = disp_po.head(int(rows_po))

    show_cols = [c for c in [
        "status_color","po_number","item_code","item_description",
        "site_group","supplier_name","ordered_qty","order_date",
        "planned_receipt_date","days_overdue","estimated_value"
    ] if c in disp_po.columns]
    show_po = disp_po[show_cols].copy()
    rename_po = {
        "status_color":"Status","po_number":"PO Number","item_code":"Item Code",
        "item_description":"Description","site_group":"Site","supplier_name":"Supplier",
        "ordered_qty":"Ordered Qty","order_date":"Order Date",
        "planned_receipt_date":"Planned Receipt","days_overdue":"Days Overdue",
        "estimated_value":"Est. Value ($)"
    }
    show_po.rename(columns=rename_po, inplace=True)
    if "Ordered Qty"   in show_po.columns: show_po["Ordered Qty"]   = show_po["Ordered Qty"].apply(lambda v: n(v,0))
    if "Est. Value ($)" in show_po.columns: show_po["Est. Value ($)"] = show_po["Est. Value ($)"].apply(lambda v: d(v,0))
    for dc in ["Order Date","Planned Receipt"]:
        if dc in show_po.columns:
            show_po[dc] = pd.to_datetime(show_po[dc], errors="coerce").dt.strftime("%Y-%m-%d").fillna("\u2014")
    st.dataframe(show_po, use_container_width=True, height=480, hide_index=True)
    dl_excel(show_po, "Carboline_Open_PO_Monitor.xlsx")

    c_ch1, c_ch2 = st.columns(2)
    with c_ch1:
        st.markdown("### Open PO Count by Site")
        if "site_group" in open_po.columns:
            site_po   = open_po.groupby(["site_group","status_color"]).size().reset_index(name="count")
            color_map = {"\U0001f7e2 On Track":GREEN,"\U0001f7e1 Overdue (<30d)":AMBER,
                         "\U0001f534 Seriously Overdue":RED}
            fig_s = px.bar(site_po, x="site_group", y="count", color="status_color",
                            color_discrete_map=color_map, title="Open POs by Site & Status",
                            labels={"site_group":"Site","count":"PO Count","status_color":"Status"})
            fig_s.update_layout(height=380, barmode="stack", **CHART)
            st.plotly_chart(fig_s, use_container_width=True)
    with c_ch2:
        st.markdown("### Top 15 Suppliers with Open POs")
        sup_po = (open_po.groupby("supplier_name").agg(
            count=("po_number","count"),
            overdue=("days_overdue", lambda x: (x>0).sum())
        ).reset_index().nlargest(15,"count"))
        fig_sp = px.bar(sup_po, x="count", y="supplier_name", orientation="h",
                         color="overdue", color_continuous_scale=[[0,GREEN],[1,RED]],
                         text="count", title="Open PO Count by Supplier (color = # overdue)",
                         labels={"count":"Open POs","supplier_name":"Supplier","overdue":"Overdue"})
        fig_sp.update_traces(textposition="outside")
        fig_sp.update_layout(height=380, yaxis={"categoryorder":"total ascending"},
                              coloraxis_showscale=False, **CHART)
        st.plotly_chart(fig_sp, use_container_width=True)



# ═══════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════
def main():
    data, missing = load()
 
    if missing:
        st.error(
            "❌ **Missing Excel files.** Please copy these into the same folder as app.py:\n\n" +
            "\n".join(f"• **{f}**" for f in missing)
        )
        st.info(
            "📁 Your folder should contain:\n"
            "- rop_ss_moq.xlsx\n"
            "- clean_consumption.xlsx\n"
            "- clean_po.xlsx\n"
            "- clean_lead_time.xlsx\n"
            "- clean_cost.xlsx\n"
            "- clean_inventory.xlsx\n"
            "- clean_supplier.xlsx\n"
            "- clean_item_master.xlsx\n"
            "- abc_classification.xlsx\n"
            "- Carboline_MasterSS_Summary_v1(3).xlsx"
        )
        st.stop()
 
    df_ss = data["ss"]
    site, tier, apple = sidebar(df_ss)
    filt = apply_f(df_ss, site, tier, apple)
 
    tabs = st.tabs([
        "📊 Executive Overview",
        "🔴 Priority Audit List",
        "🧮 Safety Stock Calculator",
        "🚚 Supplier Performance",
        "📦 Inventory & Coverage",
        "📈 Consumption & Seasonality",
        "🗑 D-Tier & MTO Strategy",
        "🔮 Demand Forecasting",
        "🌍 Tariff Risk Tracker",
        "📋 Open PO Monitor",
    ])

    with tabs[0]: t1_overview(data, filt)
    with tabs[1]: t2_audit(data, filt)
    with tabs[2]: t3_calc(data, filt)
    with tabs[3]: t4_supplier(data, filt)
    with tabs[4]: t5_inventory(data, filt)
    with tabs[5]: t6_seasonality(data, filt)
    with tabs[6]: t7_dtier(data)
    with tabs[7]: t8_forecast(data)
    with tabs[8]: t9_tariff(data)
    with tabs[9]: t10_open_po(data)
 
 
if __name__ == "__main__":
    main()

