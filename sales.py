import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import io
from sklearn.linear_model import LinearRegression
from statsmodels.tsa.holtwinters import ExponentialSmoothing
import os
from datetime import datetime
from prophet import Prophet
import io
import base64

# --- Custom CSS for Visual Enhancements ---
st.markdown(
    """
    <style>
    /* General layout and typography */
    .main {
        background-color: #F8FAFC;
        padding: 30px;
        border-radius: 12px;
        box-shadow: 0 10px 25px rgba(0, 0, 0, 0.05);
        transition: background-color 0.3s;
    }
    h1, h2, h3 {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        color: #0F172A;
        margin-bottom: 15px;
        font-weight: 700;
    }

    /* Buttons */
    .stButton>button {
        background: linear-gradient(135deg, #1565C0, #1E88E5);
        color: white;
        border: 2px solid #0EA5E9;
        border-radius: 12px;
        padding: 12px 24px;
        font-weight: 700;
        letter-spacing: 0.3px;
        transition: all 0.2s ease;
        box-shadow: 0 6px 18px rgba(2,132,199,0.25);
    }
    .stButton>button:hover {
        transform: translateY(-1px);
        box-shadow: 0 10px 20px rgba(2,132,199,0.35);
    }

    /* Dataframe (tables) */
    .dataframe {
        border-radius: 12px !important;
        overflow: hidden !important;
        box-shadow: 0 4px 16px rgba(0, 0, 0, 0.08);
        border: 1px solid #E2E8F0;
    }
    .dataframe th {
        background: #1E3A8A !important; /* Dark blue for headers */
        color: #FFFFFF !important;
        font-weight: 800 !important;
        padding: 12px !important;
        text-transform: uppercase;
        letter-spacing: 0.4px;
        height: 40px !important; /* Fixed header height */
        line-height: 40px !important;
        border: 1px solid #E5E7EB !important;
    }
    .dataframe td {
        background-color: #FFFFFF;
        border: 1px solid #E5E7EB !important;
        padding: 10px !important;
        font-weight: 600;
        color: #0F172A;
        height: 40px !important; /* Fixed row height */
        line-height: 40px !important;
        vertical-align: middle !important;
    }

    /* Sidebar */
    .css-1d391kg {
        background-color: #E2E8F0;
        border-radius: 12px;
        padding: 20px;
        box-shadow: 0 4px 10px rgba(0, 0, 0, 0.05);
    }

    /* Metric card styling with pretty border */
    div[data-testid="stMetric"] {
        border: 2px solid #38BDF8;
        border-radius: 14px;
        padding: 16px 14px;
        background: linear-gradient(180deg, #FFFFFF, #F0F9FF);
        box-shadow: 0 10px 20px rgba(56,189,248,0.25);
        white-space: nowrap;
        overflow: visible;
    }
    div[data-testid="stMetric"] > div {
        color: #0F172A !important;
        font-size: 16px;
    }

    /* Dark mode */
    .dark-mode .main { background-color: #1F2937; }
    .dark-mode h1, .dark-mode h2, .dark-mode h3 { color: #F3F4F6; }
    .dark-mode .dataframe td { background-color: #111827; color: #F3F4F6; }
    .dark-mode .dataframe th { background: #1E3A8A !important; } /* Dark blue headers in dark mode */
    .dark-mode div[data-testid="stMetric"] { background: linear-gradient(180deg,#111827,#0B1220); border-color:#60A5FA; box-shadow: 0 10px 20px rgba(59,130,246,0.25); }

    /* Tooltip */
    .tooltip { position: relative; display: inline-block; cursor: pointer; }
    .tooltip .tooltiptext {
        visibility: hidden; width: 220px; background-color: #0F172A; color: #fff; text-align: center;
        border-radius: 8px; padding: 8px; position: absolute; z-index: 1; bottom: 125%; left: 50%; margin-left: -110px;
        opacity: 0; transition: opacity 0.3s;
    }
    .tooltip:hover .tooltiptext { visibility: visible; opacity: 1; }

    .progress-bar {
        background-color: #e0e0e0;
        border-radius: 10px;
        margin-top: 5px;
    }
    .progress-bar-fill {
        background-color: #4CAF50;
        height: 15px;
        border-radius: 10px;
        text-align: right;
        padding-right: 5px;
        color: white;
        font-weight: bold;
        transition: width 0.5s ease-in-out;
    }

    /* Green caption styling for specific percentage captions */
    .green-caption {
        color: #15803D !important;
        font-weight: 600;
        font-size: 14px;
    }
    .dark-mode .green-caption {
        color: #6EE7B7 !important; /* Lighter green for dark mode */
    }
    </style>
    """,
    unsafe_allow_html=True
)

# --- Page Config ---
st.set_page_config(page_title="📊 Haneef Data Dashboard", layout="wide", page_icon="📈")

# --- Dark Mode Toggle ---
if "dark_mode" not in st.session_state:
    st.session_state.dark_mode = False

st.sidebar.checkbox(
    "🌙 Dark Mode",
    value=st.session_state.dark_mode,
    key="dark_mode_toggle",
    on_change=lambda: setattr(st.session_state, "dark_mode", not st.session_state.dark_mode)
)

if st.session_state.dark_mode:
    st.markdown('<script>document.body.classList.add("dark-mode");</script>', unsafe_allow_html=True)
else:
    st.markdown('<script>document.body.classList.remove("dark-mode");</script>', unsafe_allow_html=True)

# --- Cache Data Loading ---
@st.cache_data
def load_data(file):
    with st.spinner("⏳ Loading Excel data..."):
        try:
            xls = pd.ExcelFile(file)
            required_sheets = ["sales data", "Target", "sales channels"]
            missing = [s for s in required_sheets if s not in xls.sheet_names]
            if missing:
                st.error(f"❌ Excel file must contain sheets: {', '.join(required_sheets)}. Missing: {', '.join(missing)}")
                return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

            sales_df = pd.read_excel(xls, sheet_name="sales data")
            target_df = pd.read_excel(xls, sheet_name="Target")
            channels_df = pd.read_excel(xls, sheet_name="sales channels")
            ytd_df = pd.read_excel(xls, sheet_name="YTD") if "YTD" in xls.sheet_names else pd.DataFrame()

            required_cols = ["Billing Date", "Driver Name EN", "Net Value", "Billing Type", "PY Name 1", "SP Name1"]
            if not all(col in sales_df.columns for col in required_cols):
                st.error(f"❌ Missing required columns: {set(required_cols) - set(sales_df.columns)}")
                return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

            def normalize_series(s):
                try:
                    return s.astype(str).str.strip().str.lower().replace({'nan': ''})
                except Exception:
                    return s

            sales_df["Billing Date"] = pd.to_datetime(sales_df["Billing Date"], errors='coerce')
            if "PY Name 1" in sales_df.columns:
                sales_df["_py_name_norm"] = normalize_series(sales_df["PY Name 1"])
            else:
                sales_df["_py_name_norm"] = ""

            if "PY Name 1" in channels_df.columns:
                channels_df["_py_name_norm"] = normalize_series(channels_df["PY Name 1"])
            else:
                channels_df["_py_name_norm"] = ""
            if "Channels" in channels_df.columns:
                channels_df["_channels_norm"] = normalize_series(channels_df["Channels"])
            else:
                channels_df["_channels_norm"] = ""

            if not ytd_df.empty and "Billing Date" in ytd_df.columns:
                ytd_df["Billing Date"] = pd.to_datetime(ytd_df["Billing Date"], errors='coerce')
            if not ytd_df.empty and "PY Name 1" in ytd_df.columns:
                ytd_df["_py_name_norm"] = normalize_series(ytd_df["PY Name 1"])

            return sales_df, target_df, ytd_df, channels_df
        except Exception as e:
            st.error(f"❌ Error loading Excel file: {e}")
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

# --- Helpers: Downloads ---
@st.cache_data
def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Sheet1", index: bool = True) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=index)
    return output.getvalue()

@st.cache_data
def to_multi_sheet_excel_bytes(dfs, sheet_names) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for df, sn in zip(dfs, sheet_names):
            df.to_excel(writer, sheet_name=sn, index=True)
    return output.getvalue()

# --- PPTX Export ---
def create_pptx(report_df, billing_df, py_table, figs_dict, kpi_data):
    with st.spinner("⏳ Generating PPTX report..."):
        prs = Presentation()
        slide_layout = prs.slide_layouts[0]
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        title.text = "Sales & Targets Report"
        title.text_frame.paragraphs[0].font.size = Pt(32)
        title.text_frame.paragraphs[0].font.name = 'Roboto'
        title.text_frame.paragraphs[0].font.color.rgb = RGBColor(30, 58, 138)
        try:
            subtitle = slide.placeholders[1]
            subtitle.text = f"Generated on {datetime.now().strftime('%Y-%m-%d')}"
            subtitle.text_frame.paragraphs[0].font.size = Pt(18)
            subtitle.text_frame.paragraphs[0].font.name = 'Roboto'
            subtitle.text_frame.paragraphs[0].font.color.rgb = RGBColor(55, 65, 81)
        except Exception:
            pass

        slide_layout = prs.slide_layouts[5]
        slide = prs.slides.add_slide(slide_layout)
        slide.shapes.title.text = "📈 Key Performance Indicators"
        slide.shapes.title.text_frame.paragraphs[0].font.size = Pt(24)
        slide.shapes.title.text_frame.paragraphs[0].font.name = 'Roboto'
        slide.shapes.title.text_frame.paragraphs[0].font.color.rgb = RGBColor(30, 58, 138)

        rows = 4
        cols = 3
        left = Inches(1)
        top = Inches(1.5)
        width = Inches(8)
        height = Inches(4)
        table = slide.shapes.add_table(rows, cols, left, top, width, height).table
        kpi_list = list(kpi_data.items())
        for i in range(rows):
            for j in range(cols):
                index = i * cols + j
                if index < len(kpi_list):
                    label, value = kpi_list[index]
                    cell = table.cell(i, j)
                    cell.text = f"{label}\n{value}"
                    cell.text_frame.paragraphs[0].font.size = Pt(12)
                    cell.text_frame.paragraphs[0].font.name = 'Roboto'
                    cell.text_frame.paragraphs[0].font.bold = True
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor(243, 244, 246)

        def add_table_slide(df, title):
            slide_layout = prs.slide_layouts[5]
            slide = prs.slides.add_slide(slide_layout)
            slide.shapes.title.text = title
            slide.shapes.title.text_frame.paragraphs[0].font.size = Pt(24)
            slide.shapes.title.text_frame.paragraphs[0].font.name = 'Roboto'
            slide.shapes.title.text_frame.paragraphs[0].font.color.rgb = RGBColor(30, 58, 138)
            rows, cols = df.shape
            table = slide.shapes.add_table(rows + 1, cols, Inches(0.5), Inches(1.5), Inches(9), Inches(4)).table
            for j, col in enumerate(df.columns):
                cell = table.cell(0, j)
                cell.text = str(col)
                cell.text_frame.paragraphs[0].font.size = Pt(14)
                cell.text_frame.paragraphs[0].font.name = 'Roboto'
                cell.text_frame.paragraphs[0].font.bold = True
                cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
                cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(30, 58, 138)
            for i, row in enumerate(df.itertuples(index=False), start=1):
                for j, val in enumerate(row):
                    cell = table.cell(i, j)
                    if isinstance(val, (int, float, np.integer, np.floating)):
                        cell.text = f"{val:,.0f}"
                    else:
                        cell.text = str(val)
                    cell.text_frame.paragraphs[0].font.size = Pt(12)
                    cell.text_frame.paragraphs[0].font.name = 'Roboto'
                    cell.fill.solid();
                    cell.fill.fore_color.rgb = RGBColor(243, 244, 246) if i % 2 == 0 else RGBColor(255, 255, 255)

        def add_chart_slide(fig, title):
            slide_layout = prs.slide_layouts[5]
            slide = prs.slides.add_slide(slide_layout)
            slide.shapes.title.text = title
            slide.shapes.title.text_frame.paragraphs[0].font.size = Pt(24)
            slide.shapes.title.text_frame.paragraphs[0].font.name = 'Roboto'
            slide.shapes.title.text_frame.paragraphs[0].font.color.rgb = RGBColor(30, 58, 138)
            img_stream = io.BytesIO()
            try:
                fig.write_image(img_stream, format="png", width=800, height=600)
                img_stream.seek(0)
                slide.shapes.add_picture(img_stream, Inches(0.5), Inches(1.5), width=Inches(9))
            except Exception as e:
                slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(9), Inches(4)).text_frame.text = (
                    f"Chart cannot be embedded: {e}. Install kaleido if missing."
                )

        add_table_slide(report_df.reset_index(), "📋 Sales & Targets Summary")
        add_table_slide(billing_df.reset_index(), "📊 Sales by Billing Type per Salesman")
        add_table_slide(py_table.reset_index(), "🏬 Sales by PY Name 1")
        for key, fig in figs_dict.items(): add_chart_slide(fig, key)
        pptx_stream = io.BytesIO(); prs.save(pptx_stream); pptx_stream.seek(0); return pptx_stream

# --- Positive/Negative Coloring ---
def color_positive_negative(val):
    try:
        v = float(val)
        color = "#15803D" if v > 0 else "#B91C1C" if v < 0 else ""
        return f"color: {color}; font-weight: bold"
    except:
        return ""

def create_progress_bar_html(percentage):
    safe_pct = max(0, min(100, percentage))
    fill_color = "#4CAF50" if safe_pct >= 100 else "#2196F3"
    html = f"""
    <div style="background-color: #f0f0f0; border-radius: 5px; height: 10px; margin-top: 5px;">
        <div style="background-color: {fill_color}; height: 100%; width: {safe_pct}%; border-radius: 5px; text-align: right; font-size: 8px; color: white;">
        </div>
    </div>
    """
    return html

# --- SINGLE SIDEBAR UPLOADER ---
st.sidebar.header("📂 Upload Excel (one-time)")
st.sidebar.markdown('<div class="tooltip">ℹ️<span class="tooltiptext">Upload an Excel file with sheets: sales data, Target, sales channels, and optionally YTD.</span></div>', unsafe_allow_html=True)
uploaded = st.sidebar.file_uploader("Upload Excel file with sheets: 'sales data', 'Target', 'sales channels' (optional 'YTD')", type=["xlsx"], key="single_upload")
if st.sidebar.button("🔁 Clear data"):
    for k in ["sales_df", "target_df", "ytd_df", "channels_df", "data_loaded"]:
        if k in st.session_state: del st.session_state[k]
    st.experimental_rerun()

if uploaded is not None and "data_loaded" not in st.session_state:
    sales_df, target_df, ytd_df, channels_df = load_data(uploaded)
    st.session_state["sales_df"] = sales_df
    st.session_state["target_df"] = target_df
    st.session_state["ytd_df"] = ytd_df
    st.session_state["channels_df"] = channels_df
    st.session_state["data_loaded"] = True
    st.success("✅ File loaded — now use the menu to go to any page.")

# --- Sidebar Menu ---
st.sidebar.title("🧭 Menu")
menu = ["Home", "Sales Tracking", "Year to Date Comparison", "Custom Analysis", "SP/PY Target Allocation"]
choice = st.sidebar.selectbox("Navigate", menu)

# --- Home Page ---
if choice == "Home":
    st.title("🏠 Haneef Data Dashboard")
    with st.container():
        st.markdown(
            """
            **Welcome to your Sales Analytics Hub!**
            - 📈 Track sales & targets by salesman, By Customer Name, By Branch Name
            - 📊 Visualize trends with interactive charts (now with advanced forecasting)
            - 💾 Download reports in PPTX & Excel
            - 📅 Compare sales across custom periods
            - 🎯 Allocate SP/PY targets based on recent performance
            Use the sidebar to navigate and upload data once.
            """,
            unsafe_allow_html=True
        )
    if "data_loaded" in st.session_state: st.success("Data is loaded — choose a page from the menu.")
    else: st.info("Please upload your Excel file in the sidebar to start.")

# --- Sales Tracking Page ---
elif choice == "Sales Tracking":
    st.title("📊 MTD Tracking")
    if "data_loaded" not in st.session_state:
        st.warning("⚠️ Please upload the Excel file in the sidebar (one-time).")
    else:
        sales_df = st.session_state["sales_df"]
        target_df = st.session_state["target_df"]
        ytd_df = st.session_state["ytd_df"]
        channels_df = st.session_state["channels_df"]

        # Filters
        st.sidebar.subheader("🔍 Filters (Sales Tracking)")
        st.sidebar.markdown('<div class="tooltip">ℹ️<span class="tooltiptext">Filter data by salesmen, billing types, PY, SP, and date range.</span></div>', unsafe_allow_html=True)
        salesmen = st.sidebar.multiselect(
            "👥 Select Salesmen",
            options=sorted(sales_df["Driver Name EN"].dropna().unique()),
            default=sorted(sales_df["Driver Name EN"].dropna().unique()),
            key="st_salesmen"
        )
        billing_types = st.sidebar.multiselect(
            "📋 Select Billing Types",
            options=sorted(sales_df["Billing Type"].dropna().unique()),
            default=sorted(sales_df["Billing Type"].dropna().unique()),
            key="st_billing_types"
        )
        py_filter = st.sidebar.multiselect(
            "🏬 Select PY Name",
            options=sorted(sales_df["PY Name 1"].dropna().unique()),
            default=sorted(sales_df["PY Name 1"].dropna().unique()),
            key="st_py_filter"
        )
        sp_filter = st.sidebar.multiselect(
            "🏷️ Select SP Name1",
            options=sorted(sales_df["SP Name1"].dropna().unique()),
            default=sorted(sales_df["SP Name1"].dropna().unique()),
            key="st_sp_filter"
        )

        preset = st.sidebar.radio("📅 Quick Date Presets", ["Custom Range", "Last 7 Days", "This Month", "YTD"], key="st_preset")
        today = pd.Timestamp.today().normalize()
        if preset == "Last 7 Days":
            date_range = [today - pd.Timedelta(days=7), today]
        elif preset == "This Month":
            month_start = today.replace(day=1)
            month_end = month_start + pd.offsets.MonthEnd(0)
            date_range = [month_start, month_end]
        elif preset == "YTD":
            date_range = [today.replace(month=1, day=1), today]
        else:
            date_range = st.sidebar.date_input(
                "📆 Select Date Range",
                [sales_df["Billing Date"].min(), sales_df["Billing Date"].max()],
                key="st_date_range"
            )
            if isinstance(date_range, tuple) and len(date_range) == 2:
                date_range = [pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])]
            else:
                date_range = [sales_df["Billing Date"].min(), sales_df["Billing Date"].max()]

        if date_range[0] > date_range[1]:
            st.error("❌ Start date must be before end date.")
        else:
            top_n = st.sidebar.slider(
                "🏆 Show Top N Salesmen",
                min_value=1,
                max_value=max(1, len(sales_df["Driver Name EN"].dropna().unique())),
                value=min(5, max(1, len(sales_df["Driver Name EN"].dropna().unique()))),
                key="st_topn"
            )

            df_filtered = sales_df[
                (sales_df["Driver Name EN"].isin(salesmen))
                & (sales_df["Billing Type"].isin(billing_types))
                & (sales_df["Billing Date"] >= date_range[0])
                & (sales_df["Billing Date"] <= date_range[1])
                & (sales_df["PY Name 1"].isin(py_filter))
                & (sales_df["SP Name1"].isin(sp_filter))
            ].copy()

            if df_filtered.empty:
                st.warning("⚠️ No data matches the selected filters.")
            else:
                billing_start = df_filtered["Billing Date"].min().normalize()
                billing_end = df_filtered["Billing Date"].max().normalize()
                all_days = pd.date_range(billing_start, billing_end, freq="D")
                days_finish = int(sum(1 for d in all_days if d.weekday() != 4))

                current_month_start = today.replace(day=1)
                current_month_end = current_month_start + pd.offsets.MonthEnd(0)
                month_days = pd.date_range(current_month_start, current_month_end, freq="D")
                working_days_current_month = int(sum(1 for d in month_days if d.weekday() != 4))

                # --- Base aggregates ---
                total_sales = df_filtered.groupby("Driver Name EN")["Net Value"].sum()
                talabat_df = df_filtered[df_filtered["PY Name 1"] == "STORES SERVICES KUWAIT CO."]
                talabat_sales = talabat_df.groupby("Driver Name EN")["Net Value"].sum()

                ka_targets = target_df.set_index("Driver Name EN")["KA Target"] if "KA Target" in target_df.columns else pd.Series(dtype=float)
                talabat_targets = target_df.set_index("Driver Name EN")["Talabat Target"] if "Talabat Target" in target_df.columns else pd.Series(dtype=float)

                all_salesmen_idx = total_sales.index.union(talabat_sales.index).union(ka_targets.index).union(talabat_targets.index)
                total_sales = total_sales.reindex(all_salesmen_idx, fill_value=0).astype(float)
                talabat_sales = talabat_sales.reindex(all_salesmen_idx, fill_value=0).astype(float)
                ka_targets = ka_targets.reindex(all_salesmen_idx, fill_value=0).astype(float)
                talabat_targets = talabat_targets.reindex(all_salesmen_idx, fill_value=0).astype(float)

                ka_gap = (ka_targets - total_sales).clip(lower=0)
                talabat_gap = (talabat_targets - talabat_sales).clip(lower=0)

                top_salesmen = total_sales.sort_values(ascending=False).head(top_n).index
                total_sales_top = total_sales.loc[top_salesmen]
                talabat_sales_top = talabat_sales.loc[top_salesmen]
                ka_gap_top = ka_gap.loc[top_salesmen]
                talabat_gap_top = talabat_gap.loc[top_salesmen]

                total_ka_target_all = float(ka_targets.sum())
                total_tal_target_all = float(talabat_targets.sum())
                per_day_ka_target = (total_ka_target_all / working_days_current_month) if working_days_current_month > 0 else 0
                current_sales_per_day = (total_sales.sum() / days_finish) if days_finish > 0 else 0
                forecast_month_end_ka = current_sales_per_day * working_days_current_month

                # --- Channels mapping: Market (Retail) vs E-com ---
                df_py_sales = df_filtered.groupby("_py_name_norm")["Net Value"].sum().reset_index()
                df_channels_merged = df_py_sales.merge(
                    channels_df[["_py_name_norm", "Channels"]],
                    on="_py_name_norm",
                    how="left"
                )
                df_channels_merged["Channels"] = df_channels_merged["Channels"].str.strip().str.lower().fillna("uncategorized")
                channel_sales = df_channels_merged.groupby("Channels")["Net Value"].sum()
                total_retail_sales = float(channel_sales.get("market", 0.0) + channel_sales.get("uncategorized", 0.0))
                total_ecom_sales = float(channel_sales.get("e-com", 0.0))
                total_channel_sales = total_retail_sales + total_ecom_sales
                retail_sales_pct = (total_retail_sales / total_channel_sales * 100) if total_channel_sales > 0 else 0
                ecom_sales_pct = (total_ecom_sales / total_channel_sales * 100) if total_channel_sales > 0 else 0

                # --- KA & Other E-com Calculation ---
                ka_other_ecom_sales = total_sales.sum() - talabat_sales.sum()
                ka_other_ecom_pct = (ka_other_ecom_sales / total_ka_target_all * 100) if total_ka_target_all > 0 else 0

                # --- KPI Data for PPTX ---
                kpi_data = {
                    "KA Target": f"KD {total_ka_target_all:,.0f}",
                    "Talabat Target": f"KD {total_tal_target_all:,.0f}",
                    "KA Gap": f"KD {(total_ka_target_all - total_sales.sum()):,.0f}",
                    "Total Talabat Gap": f"KD {talabat_gap.sum():,.0f}",
                    "Total KA Sales": f"KD {total_sales.sum():,.0f} ({((total_sales.sum() / total_ka_target_all) * 100):.0f}%)" if total_ka_target_all else f"KD {total_sales.sum():,.0f} (0%)",
                    "Total Talabat Sales": f"KD {talabat_sales.sum():,.0f} ({((talabat_sales.sum() / total_tal_target_all) * 100):.0f}%)" if total_tal_target_all else f"KD {talabat_sales.sum():,.0f} (0%)",
                    "KA & Other E-com": f"KD {ka_other_ecom_sales:,.0f} ({ka_other_ecom_pct:.0f}%)",
                    "Market Sales": f"KD {total_retail_sales:,.0f} ({retail_sales_pct:.0f}%)",
                    "E-com Sales": f"KD {total_ecom_sales:,.0f} ({ecom_sales_pct:.0f}%)",
                    "Days Finished (working)": f"{days_finish}",
                    "Per Day KA Target": f"KD {per_day_ka_target:,.0f}",
                    "Current Sales Per Day": f"KD {current_sales_per_day:,.0f}",
                    "Forecasted Month-End KA Sales": f"KD {forecast_month_end_ka:,.0f}"
                }

                tabs = st.tabs(["📈 KPIs", "📋 Tables", "📊 Charts", "💾 Downloads"])

                # --- KPIs with progress bars ---
                with tabs[0]:
                    st.subheader("🏆 Key Metrics")
                    r1c1 = st.columns(1)[0]
                    with r1c1:
                        st.metric("Total KA Sales", f"KD {total_sales.sum():,.0f}")
                        progress_pct_ka = (total_sales.sum() / total_ka_target_all * 100) if total_ka_target_all > 0 else 0
                        st.markdown(create_progress_bar_html(progress_pct_ka), unsafe_allow_html=True)
                        st.markdown(f'<div class="green-caption">{progress_pct_ka:.0f}% of KA Target Achieved</div>', unsafe_allow_html=True)

                    r2c1, r2c2 = st.columns(2)
                    with r2c1:
                        st.metric("KA & Other E-com", f"KD {ka_other_ecom_sales:,.0f}")
                        st.markdown(create_progress_bar_html(ka_other_ecom_pct), unsafe_allow_html=True)
                        st.markdown(f'<div class="green-caption">{ka_other_ecom_pct:.0f}% of KA Target</div>', unsafe_allow_html=True)
                    with r2c2:
                        st.metric("Talabat Sales", f"KD {talabat_sales.sum():,.0f}")
                        progress_pct_talabat = (talabat_sales.sum() / total_tal_target_all * 100) if total_tal_target_all > 0 else 0
                        st.markdown(create_progress_bar_html(progress_pct_talabat), unsafe_allow_html=True)
                        st.markdown(f'<div class="green-caption">{progress_pct_talabat:.0f}% of Talabat Target Achieved</div>', unsafe_allow_html=True)

                    st.subheader("🎯 Target Overview")
                    r3c1, r3c2, r3c3, r3c4 = st.columns(4)
                    r3c1.metric("KA Target", f"KD {total_ka_target_all:,.0f}")
                    r3c2.metric("Talabat Target", f"KD {total_tal_target_all:,.0f}")
                    r3c3.metric("KA Gap", f"KD {(total_ka_target_all - total_sales.sum()):,.0f}")
                    r3c4.metric(" Talabat Gap", f"KD {talabat_gap.sum():,.0f}")

                    st.subheader("📊 Channel Sales")
                    r4c1, r4c2 = st.columns(2)
                    with r4c1:
                        st.metric("Retail Sales", f"KD {total_retail_sales:,.0f}")
                        retail_contribution_pct = (total_retail_sales / total_sales.sum() * 100) if total_sales.sum() > 0 else 0
                        st.caption(f"{retail_contribution_pct:.0f}% of Total KA Sales")
                    with r4c2:
                        st.metric("E-com Sales", f"KD {total_ecom_sales:,.0f}")
                        ecom_contribution_pct = (total_ecom_sales / total_sales.sum() * 100) if total_sales.sum() > 0 else 0
                        st.caption(f"{ecom_contribution_pct:.0f}% of Total KA Sales")

                    st.subheader("📈 Performance Metrics")
                    r5c1, r5c2, r5c3 = st.columns(3)
                    r5c1.metric("Days Finished (working)", days_finish)
                    r5c2.metric("Current Sales Per Day", f"KD {current_sales_per_day:,.0f}")
                    r5c3.metric("Forecasted Month-End KA Sales", f"KD {forecast_month_end_ka:,.0f}")

                # --- TABLES ---
                with tabs[1]:
                    st.subheader("📋 Sales & Targets Summary")
                    report_df = pd.DataFrame({
                        "Salesman": ka_targets.index,
                        "KA Target": ka_targets.values,
                        "KA Sales": total_sales.values,
                        "KA Remaining": ka_gap.values,
                        "KA % Achieved": np.where(ka_targets.values != 0, (total_sales.values / ka_targets.values * 100).round(0), 0),
                        "Talabat Target": talabat_targets.values,
                        "Talabat Sales": talabat_sales.values,
                        "Talabat Remaining": talabat_gap.values,
                        "Talabat % Achieved": np.where(talabat_targets.values != 0, (talabat_sales.values / talabat_targets.values * 100).round(0), 0)
                    })

                    total_row = report_df.sum(numeric_only=True).to_frame().T
                    total_row.index = ["Total"]
                    total_row["KA % Achieved"] = round(total_row["KA Sales"]/total_row["KA Target"]*100,0) if total_row["KA Target"].iloc[0]!=0 else 0
                    total_row["Talabat % Achieved"] = round(total_row["Talabat Sales"]/total_row["Talabat Target"]*100,0) if total_row["Talabat Target"].iloc[0]!=0 else 0

                    total_row = total_row.reset_index(drop=True)
                    total_row["Salesman"] = "Total"
                    total_row = total_row[report_df.columns]
                    report_df_with_total = pd.concat([report_df, total_row], ignore_index=True)

                    col_to_color = {
                        "Salesman": "background-color: #CCFFE6; color:#0F766E; font-weight:700",
                        "KA Target": "background-color: #CCFFE6; color:#0F766E; font-weight:700",
                        "KA Sales": "background-color: #CCFFE6; color:#0F766E; font-weight:700",
                        "KA Remaining": "background-color: #CCFFE6; color:#0F766E; font-weight:700",
                        "KA % Achieved": "background-color: #CCFFE6; color:#0F766E; font-weight:700",
                        "Talabat Target": "background-color: #FFE4CC; color:#9A3412; font-weight:700",
                        "Talabat Sales": "background-color: #FFE4CC; color:#9A3412; font-weight:700",
                        "Talabat Remaining": "background-color: #FFE4CC; color:#9A3412; font-weight:700",
                        "Talabat % Achieved": "background-color: #FFE4CC; color:#9A3412; font-weight:700"
                    }

                    def highlight_columns(s):
                        return [col_to_color.get(c, "") for c in s.index]

                    def highlight_total_row(row):
                        return ['background-color: #BFDBFE; color: #1E3A8A; font-weight: 900' if row['Salesman'] == "Total" else '' for _ in row]

                    styled_report = (
                        report_df_with_total.style
                        .set_table_styles([
                            {'selector': 'th', 'props': [('background', '#1E3A8A'), ('color', 'white'), ('font-weight', '800'), ('height', '40px'), ('line-height', '40px'), ('border', '1px solid #E5E7EB')]}
                        ])
                        .apply(highlight_columns, axis=1)
                        .apply(highlight_total_row, axis=1)
                        .format("{:,.0f}", subset=["KA Target","KA Sales","KA Remaining","Talabat Target","Talabat Sales","Talabat Remaining"])
                        .format("{:.0f}%", subset=["KA % Achieved","Talabat % Achieved"])
                    )
                    st.dataframe(styled_report, use_container_width=True, hide_index=True)
                    st.download_button(
                        "⬇️ Download Sales & Targets Summary (Excel)",
                        data=to_excel_bytes(report_df_with_total, sheet_name="Sales_Targets_Summary", index=False),
                        file_name=f"Sales_Targets_Summary_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                    st.subheader("📊 Sales by Billing Type per Salesman")
                    billing_wide = df_filtered.pivot_table(
                        index="Driver Name EN",
                        columns="Billing Type",
                        values="Net Value",
                        aggfunc="sum",
                        fill_value=0
                    )

                    required_cols_raw = ["ZFR", "YKF2", "YKRE", "YKS1", "YKS2", "ZCAN", "ZRE"]
                    billing_wide = billing_wide.reindex(columns=required_cols_raw, fill_value=0)
                    display_df = billing_wide.rename(columns={"ZFR": "Presales", "YKF2": "HHT"})
                    display_df["Sales Total"] = billing_wide.sum(axis=1)
                    display_df["Return"] = billing_wide["YKRE"] + billing_wide["ZRE"]
                    display_df["Return %"] = np.where(display_df["Sales Total"] != 0, (display_df["Return"] / display_df["Sales Total"] * 100).round(0), 0)
                    display_df["Cancel Total"] = billing_wide[["YKS1", "YKS2", "ZCAN"]].sum(axis=1)
                    ordered_cols = ["Presales", "HHT", "Sales Total", "YKS1", "YKS2", "ZCAN", "Cancel Total", "YKRE", "ZRE", "Return", "Return %"]
                    display_df = display_df.reindex(columns=ordered_cols, fill_value=0)

                    total_row = pd.DataFrame(display_df.sum(numeric_only=True)).T
                    total_row.index = ["Total"]
                    total_row["Return %"] = round((total_row["Return"] / total_row["Sales Total"] * 100), 0) if total_row["Sales Total"].iloc[0] != 0 else 0
                    billing_df = pd.concat([display_df, total_row])

                    billing_df.index.name = "Salesman"

                    col_to_color = {
                        **{c: "background-color: #CCFFE6; color:#0F766E; font-weight:700" for c in ["Presales", "HHT", "Sales Total"]},
                        **{c: "background-color: #FFE4CC; color:#9A3412; font-weight:700" for c in ["YKS1", "YKS2", "ZCAN", "Cancel Total"]},
                        **{c: "background-color: #FFF2CC; color:#92400E; font-weight:700" for c in ["YKRE", "ZRE", "Return", "Return %"]}
                    }

                    def highlight_columns_billing(s):
                        return [col_to_color.get(c, "") for c in s.index]

                    def highlight_total_row_billing(row):
                        return ['background-color: #BFDBFE; color: #1E3A8A; font-weight: 900' if row.name == "Total" else '' for _ in row]

                    styled_billing = (
                        billing_df.style
                        .set_table_styles([
                            {'selector': 'th', 'props': [('background', '#1E3A8A'), ('color', 'white'), ('font-weight', '800'), ('height', '40px'), ('line-height', '40px'), ('border', '1px solid #E5E7EB')]}
                        ])
                        .apply(highlight_columns_billing, axis=1)
                        .apply(highlight_total_row_billing, axis=1)
                        .format({
                            "Presales": "{:,.0f}", "HHT": "{:,.0f}", "Sales Total": "{:,.0f}",
                            "YKS1": "{:,.0f}", "YKS2": "{:,.0f}", "ZCAN": "{:,.0f}", "Cancel Total": "{:,.0f}",
                            "YKRE": "{:,.0f}", "ZRE": "{:,.0f}", "Return": "{:,.0f}", "Return %": "{:.0f}%"
                        })
                    )
                    st.dataframe(styled_billing, use_container_width=True, hide_index=False)
                    st.download_button(
                        "⬇️ Download Billing Type Table (Excel)",
                        data=to_excel_bytes(billing_df, sheet_name="Billing_Types", index=False),
                        file_name=f"Billing_Types_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                    st.subheader("🏬 Sales by PY Name 1")
                    py_table = df_filtered.groupby("PY Name 1")["Net Value"].sum().sort_values(ascending=False).to_frame(name="Sales")
                    py_table["Contribution %"] = np.where(py_table["Sales"] != 0, (py_table["Sales"]/py_table["Sales"].sum()*100).round(0), 0)

                    total_row = py_table.sum(numeric_only=True).to_frame().T
                    total_row.index = ["Total"]
                    py_table_with_total = pd.concat([py_table, total_row])

                    py_table_with_total.index.name = "PY Name 1"

                    col_to_color = {
                        "Sales": "background-color: #CCFFE6; color:#0F766E; font-weight:700",
                        "Contribution %": "background-color: #FFE4CC; color:#9A3412; font-weight:700"
                    }

                    def highlight_columns_py(s):
                        return [col_to_color.get(c, "") for c in s.index]

                    def highlight_total_row_py(row):
                        return ['background-color: #BFDBFE; color: #1E3A8A; font-weight: 900' if row.name == "Total" else '' for _ in row]

                    styled_py = (
                        py_table_with_total.style
                        .set_table_styles([
                            {'selector': 'th', 'props': [('background', '#1E3A8A'), ('color', 'white'), ('font-weight', '800'), ('height', '40px'), ('line-height', '40px'), ('border', '1px solid #E5E7EB')]}
                        ])
                        .apply(highlight_columns_py, axis=1)
                        .apply(highlight_total_row_py, axis=1)
                        .format("{:,.0f}", subset=["Sales"])
                        .format("{:.0f}%", subset=["Contribution %"])
                    )
                    st.dataframe(styled_py, use_container_width=True, hide_index=False)
                    st.download_button(
                        "⬇️ Download PY Name Table (Excel)",
                        data=to_excel_bytes(py_table_with_total, sheet_name="Sales_by_PY_Name", index=False),
                        file_name=f"Sales_by_PY_Name_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                # --- CHARTS ---
                with tabs[2]:
                    st.subheader("📊 Daily Sales Trend with Prophet Forecast")
                    df_time = df_filtered.groupby(pd.Grouper(key="Billing Date", freq="D"))["Net Value"].sum().reset_index()
                    df_time.rename(columns={"Billing Date": "ds", "Net Value": "y"}, inplace=True)
                    
                    if len(df_time) > 1:
                        m = Prophet()
                        m.fit(df_time)
                        future = m.make_future_dataframe(periods=30)
                        forecast = m.predict(future)
                        
                        fig_trend = go.Figure()
                        fig_trend.add_trace(go.Scatter(x=df_time["ds"], y=df_time["y"], mode='lines+markers', name='Actual Sales', line=dict(color='#1E3A8A', width=3)))
                        fig_trend.add_trace(go.Scatter(x=forecast["ds"], y=forecast["yhat"], mode='lines', name='Prophet Forecast', line=dict(color='#3B82F6', width=2, dash='dash')))
                        
                        df_time['y_mean'] = df_time['y'].rolling(window=7).mean()
                        df_time['y_std'] = df_time['y'].rolling(window=7).std()
                        df_time['upper_bound'] = df_time['y_mean'] + 2 * df_time['y_std']
                        df_time['lower_bound'] = df_time['y_mean'] - 2 * df_time['y_std']
                        df_time['anomaly'] = np.where((df_time['y'] > df_time['upper_bound']) | (df_time['y'] < df_time['lower_bound']), df_time['y'], np.nan)
                        
                        fig_trend.add_trace(go.Scatter(
                            x=df_time['ds'], y=df_time['anomaly'],
                            mode='markers', name='Anomaly',
                            marker=dict(color='red', size=10, symbol='x')
                        ))
                        
                        fig_trend.update_layout(
                            title="Daily Sales Trend, Prophet Forecast & Anomalies",
                            xaxis_title="Date",
                            yaxis_title="Net Value (KD)",
                            font=dict(family="Roboto", size=12),
                            plot_bgcolor="#F3F4F6",
                            paper_bgcolor="#F3F4F6",
                            hovermode="x unified"
                        )
                        st.plotly_chart(fig_trend, use_container_width=True)
                    else:
                        st.info("Not enough data to perform a time-series forecast.")
                        
                    st.subheader("📊 Market vs E-com Sales")
                    market_ecom_df = pd.DataFrame({
                        "Channel": ["Market (In-Store & Other)", "E-com (Talabat)"],
                        "Sales": [total_retail_sales, total_ecom_sales]
                    })
                    fig_channel = px.pie(
                        market_ecom_df,
                        values="Sales",
                        names="Channel",
                        title="Market vs E-com Sales Distribution",
                        hole=0.4,
                        color_discrete_sequence=px.colors.sequential.Oranges
                    )
                    fig_channel.update_traces(textposition='inside', textinfo='percent+label')
                    fig_channel.update_layout(font=dict(family="Roboto", size=12), showlegend=True)
                    st.plotly_chart(fig_channel, use_container_width=True)

                    st.subheader("📊 Daily KA Target vs Actual Sales")
                    df_time_target = df_filtered.groupby(pd.Grouper(key="Billing Date", freq="D"))["Net Value"].sum().reset_index()
                    df_time_target.rename(columns={"Billing Date": "Date", "Net Value": "Sales"}, inplace=True)
                    df_time_target = df_time_target.sort_values("Date").reset_index(drop=True)
                    df_time_target["Daily KA Target"] = per_day_ka_target

                    fig_target_trend = go.Figure()
                    fig_target_trend.add_trace(go.Scatter(
                        x=df_time_target["Date"], y=df_time_target["Sales"],
                        mode='lines+markers', name="Actual Sales", line=dict(color='#16A34A', width=3)
                    ))
                    fig_target_trend.add_trace(go.Scatter(
                        x=df_time_target["Date"], y=df_time_target["Daily KA Target"],
                        mode='lines', name="Daily KA Target", line=dict(color='#F59E0B', width=2, dash='dot')
                    ))
                    fig_target_trend.update_layout(
                        title="Daily KA Target vs Actual Sales",
                        xaxis_title="Date",
                        yaxis_title="Net Value (KD)",
                        font=dict(family="Roboto", size=12),
                        plot_bgcolor="#F3F4F6",
                        paper_bgcolor="#F3F4F6",
                        hovermode="x unified"
                    )
                    st.plotly_chart(fig_target_trend, use_container_width=True)

                    st.subheader("📊 Salesman KA Target vs Actual")
                    salesman_target_df = pd.DataFrame({
                        "Salesman": ka_targets.index,
                        "KA Target": ka_targets.values,
                        "KA Sales": total_sales.values
                    }).reset_index(drop=True)

                    fig_salesman_target = go.Figure()
                    fig_salesman_target.add_trace(go.Bar(
                        x=salesman_target_df["Salesman"],
                        y=salesman_target_df["KA Target"],
                        name="KA Target",
                        marker_color="#F87171"
                    ))
                    fig_salesman_target.add_trace(go.Bar(
                        x=salesman_target_df["Salesman"],
                        y=salesman_target_df["KA Sales"],
                        name="KA Sales",
                        marker_color="#34D399"
                    ))
                    fig_salesman_target.update_layout(
                        title="KA Target vs Actual Sales by Salesman",
                        xaxis_title="Salesman",
                        yaxis_title="Value (KD)",
                        barmode="group",
                        font=dict(family="Roboto", size=12),
                        plot_bgcolor="#F3F4F6",
                        paper_bgcolor="#F3F4F6"
                    )
                    st.plotly_chart(fig_salesman_target, use_container_width=True)

                    st.subheader("📊 Sales Breakdown by PY Name")
                    py_sales = df_filtered.groupby("PY Name 1")["Net Value"].sum().reset_index()
                    fig_py = px.pie(
                        py_sales,
                        values='Net Value',
                        names='PY Name 1',
                        title='Sales Distribution by PY Name',
                        hole=0.4,
                        color_discrete_sequence=px.colors.sequential.RdBu
                    )
                    fig_py.update_traces(textposition='inside', textinfo='percent+label')
                    fig_py.update_layout(font=dict(family="Roboto", size=12), showlegend=True)
                    st.plotly_chart(fig_py, use_container_width=True)

                    st.subheader("📊 Sales by Billing Type (Stacked Bar)")
                    billing_sales = df_filtered.pivot_table(
                        index="Driver Name EN",
                        columns="Billing Type",
                        values="Net Value",
                        aggfunc="sum",
                        fill_value=0
                    ).reset_index()
                    fig_billing = px.bar(
                        billing_sales,
                        x="Driver Name EN",
                        y=billing_sales.columns[1:],
                        title="Sales by Billing Type per Salesman",
                        color_discrete_sequence=px.colors.qualitative.Plotly
                    )
                    fig_billing.update_layout(
                        font=dict(family="Roboto", size=12),
                        xaxis_title="Salesman",
                        yaxis_title="Net Value (KD)",
                        barmode="stack",
                        plot_bgcolor="#F3F4F6",
                        paper_bgcolor="#F3F4F6"
                    )
                    st.plotly_chart(fig_billing, use_container_width=True)

                    st.subheader("📊 Sales Correlation Heatmap")
                    numeric_cols = df_filtered.select_dtypes(include=[np.number]).columns
                    if len(numeric_cols) > 1:
                        corr_matrix = df_filtered[numeric_cols].corr()
                        fig_heatmap = px.imshow(
                            corr_matrix,
                            text_auto=True,
                            title="Correlation Matrix of Numeric Columns",
                            color_continuous_scale="RdBu",
                            aspect="equal"
                        )
                        fig_heatmap.update_layout(
                            font=dict(family="Roboto", size=12),
                            plot_bgcolor="#F3F4F6",
                            paper_bgcolor="#F3F4F6"
                        )
                        st.plotly_chart(fig_heatmap, use_container_width=True)
                    else:
                        st.info("Not enough numeric columns for correlation heatmap.")

                    figs_dict = {
                        "Sales by PY Name": fig_py,
                        "Sales by Billing Type": fig_billing,
                        "Market vs E-com": fig_channel,
                        "Daily KA Target Trend": fig_target_trend,
                        "Salesman Target vs Actual": fig_salesman_target
                    }
                    if len(df_time) > 1:
                        figs_dict["Daily Sales Trend"] = fig_trend
                    if len(numeric_cols) > 1:
                        figs_dict["Correlation Heatmap"] = fig_heatmap

                # --- DOWNLOADS ---
                with tabs[3]:
                    st.subheader("📦 Consolidated Downloads")
                    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                    excel_data_combined = to_multi_sheet_excel_bytes(
                        [report_df_with_total, billing_df, py_table],
                        ["Sales_Targets_Summary", "Billing_Types", "Sales_by_PY_Name"]
                    )
                    st.download_button(
                        "💾 Download Consolidated Excel Report",
                        data=excel_data_combined,
                        file_name=f"Sales_Report_Consolidated_{timestamp}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    @st.cache_data
                    def convert_df_to_csv(df):
                        return df.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        "📥 Download Filtered Raw Data (CSV)",
                        data=convert_df_to_csv(df_filtered),
                        file_name=f"filtered_sales_data_{timestamp}.csv",
                        mime="text/csv"
                    )
                    pptx_data = create_pptx(report_df_with_total.reset_index(), billing_df.reset_index(), py_table.reset_index(), figs_dict, kpi_data)
                    st.download_button("📄 Download PPTX Report", data=pptx_data, file_name=f"sales_report_{timestamp}.pptx")

# --- YTD Comparison Page ---
elif choice == "Year to Date Comparison":
    if "ytd_df" in st.session_state and not st.session_state["ytd_df"].empty:
        ytd_df = st.session_state["ytd_df"]
        ytd_df['Billing Date'] = pd.to_datetime(ytd_df['Billing Date'])

        st.title("📅 Year to Date Comparison")
        st.markdown('<div class="tooltip">ℹ️<span class="tooltiptext">Compare sales across two periods by a selected dimension.</span></div>', unsafe_allow_html=True)

        st.subheader("📊 Choose Dimension")
        dimension = st.selectbox("Choose dimension", ["PY Name 1", "Driver Name EN", "SP Name1"], index=0)

        st.subheader("📆 Select Two Periods")
        col1, col2 = st.columns(2)
        with col1:
            st.write("Period 1")
            period1_range = st.date_input("Select Date Range", value=(ytd_df["Billing Date"].min(), ytd_df["Billing Date"].max()), key="ytd_p1_range")
        with col2:
            st.write("Period 2")
            period2_range = st.date_input("Select Date Range", value=(ytd_df["Billing Date"].min(), ytd_df["Billing Date"].max()), key="ytd_p2_range")

        if period1_range and period2_range and len(period1_range) == 2 and len(period2_range) == 2:
            period1_start, period1_end = period1_range
            period2_start, period2_end = period2_range
            df_p1 = ytd_df[(ytd_df["Billing Date"] >= pd.to_datetime(period1_start)) & (ytd_df["Billing Date"] <= pd.to_datetime(period1_end))]
            df_p2 = ytd_df[(ytd_df["Billing Date"] >= pd.to_datetime(period2_start)) & (ytd_df["Billing Date"] <= pd.to_datetime(period2_end))]
            summary_p1 = df_p1.groupby(dimension)["Net Value"].sum().reset_index().rename(columns={"Net Value": f"{period1_start.strftime('%Y-%m-%d')} to {period1_end.strftime('%Y-%m-%d')} Sales"})
            summary_p2 = df_p2.groupby(dimension)["Net Value"].sum().reset_index().rename(columns={"Net Value": f"{period2_start.strftime('%Y-%m-%d')} to {period2_end.strftime('%Y-%m-%d')} Sales"})
            ytd_comparison = pd.merge(summary_p1, summary_p2, on=dimension, how="outer").fillna(0)
            ytd_comparison["Difference"] = ytd_comparison.iloc[:, 2] - ytd_comparison.iloc[:, 1]
            ytd_comparison.rename(columns={dimension: "Name"}, inplace=True)
            ytd_comparison.loc['Total'] = ytd_comparison.sum(numeric_only=True)
            ytd_comparison.loc['Total', 'Name'] = 'Total'

            st.subheader(f"📋 Comparison by {dimension}")
            styled_ytd = (
                ytd_comparison.style
                .set_table_styles([
                    {'selector': 'th', 'props': [('background', '#1E3A8A'), ('color', 'white'), ('font-weight', '800'), ('height', '40px'), ('line-height', '40px'), ('border', '1px solid #E5E7EB')]}
                ])
                .apply(lambda x: ['background-color: #BFDBFE; color: #1E3A8A; font-weight: 900' if x.name == 'Total' else '' for _ in x], axis=1)
                .format("{:,.0f}", subset=[ytd_comparison.columns[1], ytd_comparison.columns[2], 'Difference'])
            )
            st.dataframe(styled_ytd, use_container_width=True, hide_index=False)

            st.download_button(
                "⬇️ Download YTD Comparison (Excel)",
                data=to_excel_bytes(ytd_comparison, sheet_name="YTD_Comparison", index=False),
                file_name=f"YTD_Comparison_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("⚠️ Please ensure the 'YTD' sheet is present in your uploaded file.")

# --- Custom Analysis Page ---
elif choice == "Custom Analysis":
    st.title("🔍 Custom Analysis")
    if "data_loaded" not in st.session_state:
        st.warning("⚠️ Please upload the Excel file in the sidebar (one-time).")
    else:
        sheet_options = {
            "Sales Data": st.session_state.get("sales_df", pd.DataFrame()),
            "YTD": st.session_state.get("ytd_df", pd.DataFrame()),
            "Target": st.session_state.get("target_df", pd.DataFrame()),
            "Sales Channels": st.session_state.get("channels_df", pd.DataFrame()),
            "Extra sheet": st.session_state.get("Extra_sheet_df", pd.DataFrame())
        }
        selected_sheet_name = st.selectbox("📑 Select Sheet for Analysis", list(sheet_options.keys()))
        df = sheet_options[selected_sheet_name]

        if df.empty:
            st.warning(f"⚠️ The sheet '{selected_sheet_name}' is empty or not available in your file.")
        else:
            st.subheader("💡 Explore your data by multiple columns & compare two periods.")

            available_cols = list(df.columns)

            group_cols = st.multiselect("Group by columns", available_cols)
            value_col = st.selectbox("Value to analyze", available_cols)

            if "Billing Date" in df.columns:
                st.subheader("📆 Select Two Periods")
                col1, col2 = st.columns(2)
                with col1:
                    st.write("Period 1")
                    period1_range = st.date_input(
                        "Select Period 1",
                        value=(df["Billing Date"].min(), df["Billing Date"].max()),
                        key="ca_p1_range"
                    )
                with col2:
                    st.write("Period 2")
                    period2_range = st.date_input(
                        "Select Period 2",
                        value=(df["Billing Date"].min(), df["Billing Date"].max()),
                        key="ca_p2_range"
                    )
            else:
                period1_range = period2_range = None
                st.info("⚠️ No 'Billing Date' column found. Period comparison disabled.")

            if group_cols and value_col and period1_range and period2_range and len(period1_range) == 2 and len(period2_range) == 2:
                p1_start, p1_end = pd.to_datetime(period1_range[0]), pd.to_datetime(period1_range[1])
                df_p1 = df[(df["Billing Date"] >= p1_start) & (df["Billing Date"] <= p1_end)]
                summary_p1 = df_p1.groupby(group_cols)[value_col].sum().reset_index()
                summary_p1.rename(columns={value_col: "Period 1"}, inplace=True)

                p2_start, p2_end = pd.to_datetime(period2_range[0]), pd.to_datetime(period2_range[1])
                df_p2 = df[(df["Billing Date"] >= p2_start) & (df["Billing Date"] <= p2_end)]
                summary_p2 = df_p2.groupby(group_cols)[value_col].sum().reset_index()
                summary_p2.rename(columns={value_col: "Period 2"}, inplace=True)

                comparison_df = pd.merge(summary_p1, summary_p2, on=group_cols, how="outer").fillna(0)
                comparison_df["Difference"] = comparison_df["Period 2"] - comparison_df["Period 1"]

                st.subheader(f"📋 Comparison of {value_col} by {', '.join(group_cols)}")
                styled_custom = (
                    comparison_df.style
                    .set_table_styles([
                        {'selector': 'th', 'props': [('background', '#1E3A8A'), ('color', 'white'), ('font-weight', '800'), ('height', '40px'), ('line-height', '40px'), ('border', '1px solid #E5E7EB')]}
                    ])
                    .format({
                        "Period 1": "{:,.0f}",
                        "Period 2": "{:,.0f}",
                        "Difference": "{:,.0f}"
                    })
                )
                st.dataframe(styled_custom, use_container_width=True)

                fig = px.bar(
                    comparison_df.sort_values(by="Period 2", ascending=False),
                    x=group_cols[0] if len(group_cols) == 1 else comparison_df[group_cols].astype(str).agg(" | ".join, axis=1),
                    y=["Period 1", "Period 2"],
                    barmode="group",
                    title=f"Comparison of {value_col} by {', '.join(group_cols)}",
                    color_discrete_sequence=px.colors.qualitative.Set2
                )
                st.plotly_chart(fig, use_container_width=True)

                st.download_button(
                    "⬇️ Download Comparison (Excel)",
                    data=to_excel_bytes(comparison_df, sheet_name="Custom_Comparison", index=False),
                    file_name=f"Custom_Comparison_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.info("👉 Please select at least one group column, one value column, and valid date ranges.")

# --- SP/PY Target Allocation Page ---
elif choice == "SP/PY Target Allocation":
    st.title("🎯 SP/PY Target Allocation")
    if "data_loaded" not in st.session_state:
        st.warning("⚠️ Please upload the Excel file from the sidebar first.")
        st.stop()

    sales_df = st.session_state["sales_df"]
    ytd_df = st.session_state["ytd_df"]
    target_df = st.session_state.get("target_df", pd.DataFrame())

    st.subheader("Configuration")
    st.markdown('<div class="tooltip">ℹ️<span class="tooltiptext">Allocate targets by branch or customer based on historical sales.</span></div>', unsafe_allow_html=True)
    allocation_type = st.radio("Select Target Allocation Type", ["By Branch (SP Name1)", "Customer (PY Name 1)"])
    group_col = "SP Name1" if allocation_type == "By Branch (SP Name1)" else "PY Name 1"

    target_option = st.radio("Select Target Input Option", ["Manual", "Auto (from 'Target' sheet)"])

    total_target = 0
    if target_option == "Manual":
        total_target = st.number_input("Enter the Total Target to be Allocated for this Month (KD)", min_value=0, value=1000000, step=1000)
    else:
        if target_df.empty or "KA Target" not in target_df.columns:
            st.error("❌ 'Target' sheet or 'KA Target' column not found. Please upload a file with this sheet for 'Auto' mode.")
            st.stop()
        total_target = target_df["KA Target"].sum()
        st.info(f"Using Total Target from 'Target' sheet: KD {total_target:,.0f}")

    if total_target <= 0:
        st.warning("Please ensure the total target is greater than 0.")
        st.stop()

    st.subheader("Historical Data Period")
    today = pd.Timestamp.today().normalize()
    data_period_option = st.radio("Select Historical Data Period", ["Last 6 Months", "Manual Days"], index=1)

    if data_period_option == "Last 6 Months":
        lookback_period = pd.DateOffset(months=6)
        days_label = "6 Months"; months_count = 6
        end_date_selected = today; start_date_selected = today - lookback_period
    else:
        start_date_manual = today - pd.DateOffset(days=180)
        selected_dates = st.date_input("Select date range", value=(start_date_manual, today))
        if len(selected_dates) == 2:
            start_date_selected, end_date_selected = selected_dates
            lookback_period = end_date_selected - start_date_selected
            days_label = f"From {start_date_selected.strftime('%Y-%m-%d')} to {end_date_selected.strftime('%Y-%m-%d')}"
            months_count = lookback_period.days / 30
        else:
            st.warning("Please select both a start and an end date.")
            st.stop()

    historical_df = ytd_df[(ytd_df["Billing Date"] >= pd.Timestamp(start_date_selected)) & (ytd_df["Billing Date"] <= pd.Timestamp(end_date_selected))].copy()
    if historical_df.empty:
        st.warning(f"⚠️ No sales data available in 'YTD' for {days_label}.")
        st.stop()

    historical_sales = historical_df.groupby(group_col)["Net Value"].sum()
    total_historical_sales_value = historical_sales.sum()
    current_month_sales_df = sales_df[(sales_df["Billing Date"].dt.month == today.month) & (sales_df["Billing Date"].dt.year == today.year)].copy()
    current_month_sales = current_month_sales_df.groupby(group_col)["Net Value"].sum()
    total_current_month_sales = current_month_sales.sum()

    target_balance = total_target - total_current_month_sales

    if total_target > 0:
        average_historical_sales = total_historical_sales_value / months_count
        st.subheader("🎯 Target Analysis")
        col1, col2, col3 = st.columns(3)
        col4, col5 = st.columns(2)
        with col1: st.metric("Historical Sales Total", f"KD {total_historical_sales_value:,.0f}")
        with col2: st.metric("Allocated Target Total", f"KD {total_target:,.0f}")
        with col3:
            if average_historical_sales > 0:
                percentage_increase_needed = ((total_target - average_historical_sales) / average_historical_sales) * 100
                delta_value = total_target - average_historical_sales
                st.metric("Increase Needed vs Avg Sales", f"{percentage_increase_needed:.0f}%", delta=f"KD {delta_value:,.0f}")
            else:
                st.metric("Increase Needed vs Avg Sales", "N/A", delta="Historical = 0")
        st.markdown("---")
        with col4: st.metric("Current Month Sales", f"KD {total_current_month_sales:,.0f}")
        with col5: st.metric("Target Balance", f"KD {target_balance:,.0f}")

    allocation_table = pd.DataFrame(index=historical_sales.index.union(current_month_sales.index).unique())
    allocation_table.index.name = "Name"
    allocation_table[f"Last {days_label} Total Sales"] = historical_sales.reindex(allocation_table.index, fill_value=0)
    allocation_table[f"Last {days_label} Average Sales"] = allocation_table[f"Last {days_label} Total Sales"] / months_count
    if total_historical_sales_value > 0:
        allocation_table["This Month Auto-Allocated Target"] = allocation_table[f"Last {days_label} Total Sales"] / total_historical_sales_value * total_target
    else:
        allocation_table["This Month Auto-Allocated Target"] = 0
    allocation_table["Current Month Sales"] = current_month_sales.reindex(allocation_table.index, fill_value=0)
    allocation_table["Target Balance"] = allocation_table["This Month Auto-Allocated Target"] - allocation_table["Current Month Sales"]

    total_row = allocation_table.sum().to_frame().T
    total_row.index = ["Total"]
    total_row["Target Balance"] = total_target - total_current_month_sales
    allocation_table_with_total = pd.concat([allocation_table, total_row])

    def color_target_balance(val):
        if isinstance(val, (int, float)):
            color = 'red' if val > 0 else 'green'
            return f'color: {color}'
        return ''

    st.subheader(f"📊 Auto-Allocated Targets Based on {days_label}")
    styled_allocation = (
        allocation_table_with_total.astype(int).style
        .set_table_styles([
            {'selector': 'th', 'props': [('background', '#1E3A8A'), ('color', 'white'), ('font-weight', '800'), ('height', '40px'), ('line-height', '40px'), ('border', '1px solid #E5E7EB')]}
        ])
        .apply(lambda x: ['background-color: #BFDBFE; color: #1E3A8A; font-weight: 900' if x.name == 'Total' else '' for _ in x], axis=1)
        .format("{:,.0f}")
    )
    st.dataframe(styled_allocation, use_container_width=True)

    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    excel_data = to_excel_bytes(allocation_table, sheet_name="Allocated_Targets")
    st.download_button(
        "💾 Download Target Allocation Table",
        data=excel_data,
        file_name=f"target_allocation_{allocation_type.replace(' ', '_')}_{timestamp}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
