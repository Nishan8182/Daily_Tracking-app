# save as sales_app_single_upload.py
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
import os

# --- Custom CSS for Visual Enhancements ---
st.markdown(
    """
    <style>
    /* General layout and typography */
    .main {
        background-color: #F3F4F6;
        padding: 20px;
        border-radius: 10px;
    }
    h1, h2, h3 {
        font-family: 'Roboto', Arial, sans-serif;
        color: #1E3A8A;
    }
    .stButton>button {
        background-color: #1E3A8A;
        color: white;
        border-radius: 5px;
        padding: 8px 16px;
    }
    .stButton>button:hover {
        background-color: #3B82F6;
    }
    /* Table styling */
    .dataframe th {
        background-color: #1E3A8A;
        color: white;
        font-weight: bold;
    }
    .dataframe td {
        background-color: #FFFFFF;
        border: 1px solid #E5E7EB;
    }
    /* Sidebar styling */
    .css-1d391kg {
        background-color: #E5E7EB;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# --- Page Config ---
st.set_page_config(page_title="📊 Haneef Data Dashboard", layout="wide", page_icon="📈")

# --- Cache Data Loading ---
@st.cache_data
def load_data(file):
    """
    Returns: sales_df, target_df, ytd_df (empty DataFrames if sheets are missing or invalid)
    """
    with st.spinner("⏳ Loading Excel data..."):
        try:
            xls = pd.ExcelFile(file)
            required_sheets = ["sales data", "Target"]
            if not all(sheet in xls.sheet_names for sheet in required_sheets):
                st.error(f"❌ Excel file must contain sheets: {', '.join(required_sheets)}")
                return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

            sales_df = pd.read_excel(xls, sheet_name="sales data")
            target_df = pd.read_excel(xls, sheet_name="Target")
            ytd_df = pd.read_excel(xls, sheet_name="YTD") if "YTD" in xls.sheet_names else pd.DataFrame()

            required_cols = ["Billing Date", "Driver Name EN", "Net Value", "Billing Type", "PY Name 1", "SP Name1"]
            if not all(col in sales_df.columns for col in required_cols):
                st.error(f"❌ Missing required columns: {set(required_cols) - set(sales_df.columns)}")
                return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

            sales_df["Billing Date"] = pd.to_datetime(sales_df["Billing Date"], errors='coerce')
            if not ytd_df.empty and "Billing Date" in ytd_df.columns:
                ytd_df["Billing Date"] = pd.to_datetime(ytd_df["Billing Date"], errors='coerce')

            return sales_df, target_df, ytd_df
        except Exception as e:
            st.error(f"❌ Error loading Excel file: {e}")
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

# --- PPTX Export ---
def create_pptx(report_df, billing_df, py_df, figs_dict, kpi_data):
    with st.spinner("⏳ Generating PPTX report..."):
        prs = Presentation()
        
        # Title Slide
        slide_layout = prs.slide_layouts[0]
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        title.text = "Sales & Targets Report"
        title.text_frame.paragraphs[0].font.size = Pt(32)
        title.text_frame.paragraphs[0].font.name = 'Roboto'
        title.text_frame.paragraphs[0].font.color.rgb = RGBColor(30, 58, 138)  # #1E3A8A
        try:
            subtitle = slide.placeholders[1]
            subtitle.text = "Generated from Sales Data"
            subtitle.text_frame.paragraphs[0].font.size = Pt(18)
            subtitle.text_frame.paragraphs[0].font.name = 'Roboto'
            subtitle.text_frame.paragraphs[0].font.color.rgb = RGBColor(55, 65, 81)  # #374151
        except Exception:
            pass

        # KPI Matrix Slide (NEW)
        slide_layout = prs.slide_layouts[5]
        slide = prs.slides.add_slide(slide_layout)
        slide.shapes.title.text = "📈 Key Performance Indicators"
        slide.shapes.title.text_frame.paragraphs[0].font.size = Pt(24)
        slide.shapes.title.text_frame.paragraphs[0].font.name = 'Roboto'
        slide.shapes.title.text_frame.paragraphs[0].font.color.rgb = RGBColor(30, 58, 138)

        # Create a table for the KPI matrix
        rows = 4
        cols = 3
        left = Inches(1)
        top = Inches(1.5)
        width = Inches(8)
        height = Inches(4)

        table = slide.shapes.add_table(rows, cols, left, top, width, height).table

        # Populate the table with KPIs
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
                    cell.fill.fore_color.rgb = RGBColor(243, 244, 246)  # Light gray background

        def add_table_slide(df, title):
            slide_layout = prs.slide_layouts[5]
            slide = prs.slides.add_slide(slide_layout)
            slide.shapes.title.text = title
            slide.shapes.title.text_frame.paragraphs[0].font.size = Pt(24)
            slide.shapes.title.text_frame.paragraphs[0].font.name = 'Roboto'
            slide.shapes.title.text_frame.paragraphs[0].font.color.rgb = RGBColor(30, 58, 138)
            
            rows, cols = df.shape
            table = slide.shapes.add_table(rows + 1, cols, Inches(0.5), Inches(1.5), Inches(9), Inches(4)).table
            
            # Style header
            for j, col in enumerate(df.columns):
                cell = table.cell(0, j)
                cell.text = str(col)
                cell.text_frame.paragraphs[0].font.size = Pt(14)
                cell.text_frame.paragraphs[0].font.name = 'Roboto'
                cell.text_frame.paragraphs[0].font.bold = True
                cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(30, 58, 138)  # #1E3A8A
            
            # Style data cells
            for i, row in enumerate(df.itertuples(index=False), start=1):
                for j, val in enumerate(row):
                    cell = table.cell(i, j)
                    if isinstance(val, (int, float, np.integer, np.floating)):
                        cell.text = f"{val:,.0f}"
                    else:
                        cell.text = str(val)
                    cell.text_frame.paragraphs[0].font.size = Pt(12)
                    cell.text_frame.paragraphs[0].font.name = 'Roboto'
                    cell.fill.solid()
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
        add_table_slide(py_df.to_frame().reset_index(), "🏬 Sales by PY Name 1")

        # Insights Charts
        for key, fig in figs_dict.items():
            add_chart_slide(fig, key)

        pptx_stream = io.BytesIO()
        prs.save(pptx_stream)
        pptx_stream.seek(0)
        return pptx_stream

# --- Positive/Negative Coloring ---
def color_positive_negative(val):
    try:
        v = float(val)
        color = "#15803D" if v > 0 else "#B91C1C" if v < 0 else ""
        return f"color: {color}; font-weight: bold"
    except:
        return ""

# ---------------------------
# SINGLE SIDEBAR UPLOADER
# ---------------------------
st.sidebar.header("📂 Upload Excel (one-time)")
uploaded = st.sidebar.file_uploader("Upload Excel file with sheets: 'sales data', 'Target' (optional 'YTD')", type=["xlsx"], key="single_upload")
if st.sidebar.button("🔁 Clear data"):
    for k in ["sales_df", "target_df", "ytd_df", "data_loaded"]:
        if k in st.session_state:
            del st.session_state[k]
    st.experimental_rerun()

if uploaded is not None and "data_loaded" not in st.session_state:
    sales_df, target_df, ytd_df = load_data(uploaded)
    st.session_state["sales_df"] = sales_df
    st.session_state["target_df"] = target_df
    st.session_state["ytd_df"] = ytd_df
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
            
            - 📈 Track sales & targets by salesman, PY Name, and SP Name
            - 📊 Visualize trends with interactive charts
            - 💾 Download reports in PPTX & Excel
            - 📅 Compare sales across custom periods
            - 🎯 Allocate SP/PY targets based on recent performance
            
            Use the sidebar to navigate and upload data once.
            """,
            unsafe_allow_html=True
        )
    if "data_loaded" in st.session_state:
        st.success("Data is loaded — choose a page from the menu.")
    else:
        st.info("Please upload your Excel file in the sidebar to start.")

# --- Sales Tracking Page ---
elif choice == "Sales Tracking":
    st.title("📊 Sales Tracking Dashboard")
    if "data_loaded" not in st.session_state:
        st.warning("⚠️ Please upload the Excel file in the sidebar (one-time).")
    else:
        sales_df = st.session_state["sales_df"]
        target_df = st.session_state["target_df"]
        ytd_df = st.session_state["ytd_df"]

        # Filters (moved to sidebar to avoid repeated uploads)
        st.sidebar.subheader("🔍 Filters (Sales Tracking)")
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
            ]

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
                per_day_ka_target = (total_ka_target_all / days_finish) if days_finish > 0 else 0
                current_sales_per_day = (total_sales.sum() / days_finish) if days_finish > 0 else 0
                forecast_month_end_ka = current_sales_per_day * working_days_current_month

                # Prepare KPI data for PPTX
                kpi_data = {
                    "Total KA Sales": f"{total_sales.sum():,.0f} ({((total_sales.sum() / total_ka_target_all) * 100):.2f}%)" if total_ka_target_all else "0",
                    "Total Talabat Sales": f"{talabat_sales.sum():,.0f} ({((talabat_sales.sum() / total_tal_target_all) * 100):.2f}%)" if total_tal_target_all else "0",
                    "Total KA Gap": f"{ka_gap.sum():,.0f} ({(ka_gap.sum() / total_ka_target_all * 100):.2f}%)" if total_ka_target_all else "0",
                    "Total Talabat Gap": f"{talabat_gap.sum():,.0f} ({(talabat_gap.sum() / total_tal_target_all * 100):.2f}%)" if total_tal_target_all else "0",
                    "Top KA Salesman": f"{total_sales_top.idxmax()}: {total_sales_top.max():,.0f}" if not total_sales_top.empty else "—: 0",
                    "Top Talabat Salesman": f"{talabat_sales_top.idxmax()}: {talabat_sales_top.max():,.0f}" if not talabat_sales_top.empty else "—: 0",
                    "Days Finished (working)": f"{days_finish}",
                    "Per Day KA Target": f"{per_day_ka_target:,.0f}",
                    "Current Sales Per Day": f"{current_sales_per_day:,.0f}",
                    "Forecasted Month-End KA Sales": f"{forecast_month_end_ka:,.0f}"
                }

                tabs = st.tabs(["📈 KPIs", "📋 Tables", "📊 Charts", "💾 Downloads", "🔍 Insights"])

                with tabs[0]:
                    st.subheader("🏆 Key Metrics")
                    r1c1, r1c2, r1c3, r1c4 = st.columns(4)
                    r1c1.metric(
                        "Total KA Sales",
                        f"{total_sales.sum():,.0f}",
                        f"{((total_sales.sum() / total_ka_target_all) * 100):.2f}%" if total_ka_target_all else "0.00%",
                        delta_color="normal"
                    )
                    r1c2.metric(
                        "Total Talabat Sales",
                        f"{talabat_sales.sum():,.0f}",
                        f"{((talabat_sales.sum() / total_tal_target_all) * 100):.2f}%" if total_tal_target_all else "0.00%",
                        delta_color="normal"
                    )
                    r1c3.metric(
                        "Total KA Gap",
                        f"{ka_gap.sum():,.0f}",
                        f"{(ka_gap.sum() / total_ka_target_all * 100):.2f}%" if total_ka_target_all else "0.00%",
                        delta_color="inverse"
                    )
                    r1c4.metric(
                        "Total Talabat Gap",
                        f"{talabat_gap.sum():,.0f}",
                        f"{(talabat_gap.sum() / total_tal_target_all * 100):.2f}%" if total_tal_target_all else "0.00%",
                        delta_color="inverse"
                    )

                    r2c1, r2c2 = st.columns(2)
                    r2c1.metric("👑 Top KA Salesman", total_sales_top.idxmax() if not total_sales_top.empty else "—", f"{total_sales_top.max():,.0f}" if not total_sales_top.empty else "0")
                    r2c2.metric("👑 Top Talabat Salesman", talabat_sales_top.idxmax() if not talabat_sales_top.empty else "—", f"{talabat_sales_top.max():,.0f}" if not talabat_sales_top.empty else "0")

                    r3c1, r3c2, r3c3, r3c4 = st.columns(4)
                    r3c1.metric("📅 Days Finished (working)", days_finish)
                    r3c2.metric("🎯 Per Day KA Target", f"{per_day_ka_target:,.0f}")
                    r3c3.metric("📈 Current Sales Per Day", f"{current_sales_per_day:,.0f}")
                    r3c4.metric("🔮 Forecasted Month-End KA Sales", f"{forecast_month_end_ka:,.0f}")

                with tabs[1]:
                    st.subheader("📋 Sales & Targets Summary")
                    report_df = pd.DataFrame({
                        "KA Target": ka_targets,
                        "KA Sales": total_sales,
                        "KA Remaining": ka_gap,
                        "KA % Achieved": np.where(ka_targets != 0, (total_sales / ka_targets * 100).round(2), 0),
                        "Talabat Target": talabat_targets,
                        "Talabat Sales": talabat_sales,
                        "Talabat Remaining": talabat_gap,
                        "Talabat % Achieved": np.where(talabat_targets != 0, (talabat_sales / talabat_targets * 100).round(2), 0)
                    })
                    st.dataframe(
                        report_df.style
                            .applymap(color_positive_negative, subset=["KA % Achieved", "Talabat % Achieved"])
                            .highlight_max(subset=["KA % Achieved", "Talabat % Achieved"], color="#FFD700")
                            .format("{:,.0f}"),
                        use_container_width=True
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
                    display_df = billing_wide.rename(columns={"ZFR": "Sales Group", "YKF2": "HHT"})
                    display_df["Sales Total"] = billing_wide.sum(axis=1)
                    display_df["Return"] = billing_wide["YKRE"] + billing_wide["ZRE"]
                    display_df["Return %"] = np.where(display_df["Sales Total"] != 0, (display_df["Return"] / display_df["Sales Total"] * 100).round(2), 0)
                    display_df["Cancel Total"] = billing_wide[["YKS1", "YKS2", "ZCAN"]].sum(axis=1)
                    ordered_cols = ["Sales Group", "HHT", "Sales Total", "YKS1", "YKS2", "ZCAN", "Cancel Total", "YKRE", "ZRE", "Return", "Return %"]
                    display_df = display_df.reindex(columns=ordered_cols, fill_value=0)
                    total_row = pd.DataFrame(display_df.sum(numeric_only=True)).T
                    total_row.index = ["Total"]
                    total_row["Return %"] = round((total_row["Return"] / total_row["Sales Total"] * 100), 2) if total_row["Sales Total"].iloc[0] != 0 else 0
                    billing_df = pd.concat([display_df, total_row])

                    col_to_color = {
                        **{c: "background-color: #CCFFE6" for c in ["Sales Group", "HHT", "Sales Total"]},
                        **{c: "background-color: #FFE4CC" for c in ["YKS1", "YKS2", "ZCAN", "Cancel Total"]},
                        **{c: "background-color: #FFF2CC" for c in ["YKRE", "ZRE", "Return", "Return %"]}
                    }

                    def highlight_columns(s):
                        return [col_to_color.get(c, "") for c in s.index]

                    def highlight_total_row(row):
                        return ['background-color: #BFDBFE; color: #1E3A8A; font-weight: bold' if row.name == "Total" else '' for _ in row]

                    st.dataframe(
                        billing_df.style
                            .apply(highlight_columns, axis=1)
                            .apply(highlight_total_row, axis=1)
                            .format({
                                "Sales Group": "{:,.0f}",
                                "HHT": "{:,.0f}",
                                "Sales Total": "{:,.0f}",
                                "YKS1": "{:,.0f}",
                                "YKS2": "{:,.0f}",
                                "ZCAN": "{:,.0f}",
                                "Cancel Total": "{:,.0f}",
                                "YKRE": "{:,.0f}",
                                "ZRE": "{:,.0f}",
                                "Return": "{:,.0f}",
                                "Return %": "{:.2f}%"
                            }),
                        use_container_width=True
                    )

                    st.subheader("🏬 Sales by PY Name 1")
                    py_df = df_filtered.groupby("PY Name 1")["Net Value"].sum().sort_values(ascending=False)
                    st.dataframe(
                        py_df.to_frame().style.background_gradient(cmap="Blues").format("{:,.0f}"),
                        use_container_width=True
                    )

                with tabs[2]:
                    st.subheader("📊 Sales Trend with Forecast")
                    df_time = df_filtered.groupby(pd.Grouper(key="Billing Date", freq="D"))["Net Value"].sum().reset_index()
                    df_time.rename(columns={"Billing Date": "Date", "Net Value": "Sales"}, inplace=True)
                    
                    # Generate forecast data points
                    if len(df_time) > 1:
                        model = LinearRegression()
                        X = np.arange(len(df_time)).reshape(-1, 1)
                        y = df_time["Sales"].values
                        model.fit(X, y)
                        df_time["Forecast"] = model.predict(X)
                    else:
                        df_time["Forecast"] = df_time["Sales"]
                    
                    fig_trend = go.Figure()
                    
                    # Add the actual sales data
                    fig_trend.add_trace(go.Scatter(
                        x=df_time["Date"],
                        y=df_time["Sales"],
                        mode='lines+markers',
                        name='Actual Sales',
                        line=dict(color='#1E3A8A', width=3)
                    ))
                    
                    # Add the forecast line
                    fig_trend.add_trace(go.Scatter(
                        x=df_time["Date"],
                        y=df_time["Forecast"],
                        mode='lines',
                        name='Linear Forecast',
                        line=dict(color='#3B82F6', width=2, dash='dash')
                    ))
                    
                    fig_trend.update_layout(
                        title="Daily Sales Trend & Forecast",
                        xaxis_title="Date",
                        yaxis_title="Net Value",
                        font=dict(family="Roboto", size=12),
                        plot_bgcolor="#F3F4F6",
                        paper_bgcolor="#F3F4F6",
                        hovermode="x unified"
                    )
                    st.plotly_chart(fig_trend, use_container_width=True)
                    
                    # Sales Breakdown by PY Name
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
                    fig_py.update_layout(
                        font=dict(family="Roboto", size=12),
                        showlegend=True
                    )
                    st.plotly_chart(fig_py, use_container_width=True)
                    
                    figs_dict = {}
                    figs_dict["Daily Sales Trend"] = fig_trend
                    figs_dict["Sales by PY Name"] = fig_py
                    
                with tabs[3]:
                    st.subheader("📦 Consolidated Downloads")
                    @st.cache_data
                    def convert_dfs_to_excel(dfs, sheet_names):
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                            for df, sheet_name in zip(dfs, sheet_names):
                                # Ensure the index is a column for export
                                df.to_excel(writer, sheet_name=sheet_name, index=True)
                        return output.getvalue()
                    
                    excel_data_combined = convert_dfs_to_excel(
                        [report_df, billing_df, py_df.to_frame()],
                        ["Sales_Targets_Summary", "Billing_Types", "Sales_by_PY_Name"]
                    )
                    st.download_button(
                        "💾 Download Consolidated Excel Report",
                        data=excel_data_combined,
                        file_name="Sales_Report_Consolidated.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                    
                    @st.cache_data
                    def convert_df_to_csv(df):
                        return df.to_csv(index=False).encode('utf-8')
                    
                    st.download_button(
                        "📥 Download Filtered Raw Data (CSV)",
                        data=convert_df_to_csv(df_filtered),
                        file_name="filtered_sales_data.csv",
                        mime="text/csv",
                    )
                    
                    st.subheader("💾 Download Reports")
                    pptx_data = create_pptx(report_df, billing_df, py_df, figs_dict, kpi_data)
                    st.download_button("📄 Download PPTX Report", data=pptx_data, file_name="sales_report.pptx")
                    
                with tabs[4]:
                    st.subheader("🔍 Sales Insights")
                    # Sales Trend Forecast
                    df_time = df_filtered.groupby("Billing Date")["Net Value"].sum().reset_index()
                    if len(df_time) > 1:
                        model = LinearRegression()
                        X = np.arange(len(df_time)).reshape(-1, 1)
                        y = df_time["Net Value"].values
                        model.fit(X, y)
                        df_time["Forecast"] = model.predict(X)
                        fig_trend = px.line(
                            df_time,
                            x="Billing Date",
                            y=["Net Value", "Forecast"],
                            title="Sales Trend with Forecast",
                            color_discrete_sequence=["#1E3A8A", "#EF4444"]
                        )
                        fig_trend.update_layout(
                            font=dict(family="Roboto", size=12),
                            plot_bgcolor="#F3F4F6",
                            paper_bgcolor="#F3F4F6",
                            xaxis_title="Billing Date",
                            yaxis_title="Net Value"
                        )
                        st.plotly_chart(fig_trend, use_container_width=True)

# --- YTD Comparison Page ---
elif choice == "Year to Date Comparison":
    if "ytd_df" in st.session_state and not st.session_state["ytd_df"].empty:
        ytd_df = st.session_state["ytd_df"]

        # Convert 'Billing Date' to datetime for correct filtering
        ytd_df['Billing Date'] = pd.to_datetime(ytd_df['Billing Date'])

        st.title("📅 Year to Date Comparison")

        st.subheader("📊 Choose Dimension")
        dimension = st.selectbox(
            "Choose dimension",
            ["PY Name 1", "Driver Name EN", "SP Name1"],
            index=0
        )

        st.subheader("📆 Select Two Periods")
        # The code is already structured to display in two columns
        col1, col2 = st.columns(2)

        with col1:
            st.write("Period 1")
            period1_range = st.date_input("Select Date Range", value=(ytd_df["Billing Date"].min(), ytd_df["Billing Date"].max()), key="ytd_p1_range")
        
        with col2:
            st.write("Period 2")
            period2_range = st.date_input("Select Date Range", value=(ytd_df["Billing Date"].min(), ytd_df["Billing Date"].max()), key="ytd_p2_range")

        # Ensure dates are selected before proceeding
        if period1_range and period2_range:
            period1_start, period1_end = period1_range
            period2_start, period2_end = period2_range

            # Filter the DataFrame for Period 1
            df_p1 = ytd_df[
                (ytd_df["Billing Date"] >= pd.to_datetime(period1_start)) & 
                (ytd_df["Billing Date"] <= pd.to_datetime(period1_end))
            ]
            
            # Filter the DataFrame for Period 2
            df_p2 = ytd_df[
                (ytd_df["Billing Date"] >= pd.to_datetime(period2_start)) & 
                (ytd_df["Billing Date"] <= pd.to_datetime(period2_end))
            ]

            # Calculate total net value for each dimension
            summary_p1 = df_p1.groupby(dimension)["Net Value"].sum().reset_index().rename(columns={"Net Value": "First Period Value"})
            summary_p2 = df_p2.groupby(dimension)["Net Value"].sum().reset_index().rename(columns={"Net Value": "2nd Period Value"})

            # Merge the summaries on the chosen dimension
            ytd_comparison = pd.merge(summary_p1, summary_p2, on=dimension, how="outer").fillna(0)

            # Calculate the difference
            ytd_comparison["Difference"] = ytd_comparison["2nd Period Value"] - ytd_comparison["First Period Value"]

            # Rename the dimension column to "Name" for the final display
            ytd_comparison.rename(columns={dimension: "Name"}, inplace=True)

            st.subheader(f"📋 Comparison by {dimension}")
            st.dataframe(ytd_comparison.style.format({
                "First Period Value": "{:,.2f}",
                "2nd Period Value": "{:,.2f}",
                "Difference": "{:,.2f}"
            }))
    else:
        st.warning("⚠️ Please ensure the 'MTD.XLSX - YTD.csv' sheet is present in your uploaded file.")


# --- Custom Analysis Page ---
elif choice == "Custom Analysis":
    st.title("🔍 Custom Analysis")
    if "data_loaded" not in st.session_state:
        st.warning("⚠️ Please upload the Excel file in the sidebar (one-time).")
    else:
        sales_df = st.session_state["sales_df"]
        st.subheader("💡 Explore your data by any column.")

        available_cols = list(sales_df.columns)
        group_col = st.selectbox("Group by (e.g., 'PY Name 1')", available_cols)
        value_col = st.selectbox("Value to analyze (e.g., 'Net Value')", available_cols)
        
        if st.button("Generate Analysis"):
            if group_col and value_col:
                try:
                    analysis_df = sales_df.groupby(group_col)[value_col].sum().reset_index()
                    st.subheader(f"Analysis of {value_col} by {group_col}")
                    st.dataframe(analysis_df.style.format({value_col: "{:,.2f}"}), use_container_width=True)

                    fig = px.bar(
                        analysis_df.sort_values(by=value_col, ascending=False),
                        x=group_col,
                        y=value_col,
                        title=f"Total {value_col} by {group_col}",
                        color=group_col,
                        color_discrete_sequence=px.colors.qualitative.Plotly
                    )
                    st.plotly_chart(fig, use_container_width=True)
                except Exception as e:
                    st.error(f"An error occurred: {e}")
            else:
                st.warning("Please select both a grouping column and a value column.")

# --- SP/PY Target Allocation Page ---
elif choice == "SP/PY Target Allocation":
    st.title("🎯 SP/PY Target Allocation")

    if "data_loaded" not in st.session_state:
        st.warning("⚠️ Please upload the Excel file from the sidebar first.")
        st.stop()

    # ✅ Reuse uploaded data instead of asking again
    sales_df = st.session_state["sales_df"]
    ytd_df   = st.session_state["ytd_df"]
    target_df = st.session_state.get("target_df", pd.DataFrame())

    # --- User Inputs & Target Selection ---
    st.subheader("Configuration")
    allocation_type = st.radio("Select Target Allocation Type", ["By Branch (SP Name1)", "Customer (PY Name 1)"])
    group_col = "SP Name1" if allocation_type == "By Branch (SP Name1)" else "PY Name 1"
    
    target_option = st.radio("Select Target Input Option", ["Manual", "Auto (from 'Target' sheet)"])
    
    total_target = 0
    if target_option == "Manual":
        total_target = st.number_input(
            "Enter the Total Target to be Allocated for this Month",
            min_value=0,
            value=1000000,
            step=1000
        )
    else: # Auto from sheet
        if target_df.empty or "KA Target" not in target_df.columns:
            st.error("❌ 'Target' sheet or 'KA Target' column not found. Please upload a file with this sheet for 'Auto' mode.")
            st.stop()
        total_target = target_df["KA Target"].sum()
        st.info(f"Using Total Target from 'Target' sheet: KD {total_target:,.0f}")

    if total_target <= 0:
        st.warning("Please ensure the total target is greater than 0.")
        st.stop()

    # --- Data Period Selection ---
    st.subheader("Historical Data Period")
    today = pd.Timestamp.today().normalize()
    data_period_option = st.radio("Select Historical Data Period", ["Last 6 Months", "Manual Days"], index=1)
    
    if data_period_option == "Last 6 Months":
        lookback_period = pd.DateOffset(months=6)
        days_label = "6 Months"
        months_count = 6
        end_date_selected = today
        start_date_selected = today - lookback_period
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

    # --- Data Processing ---
    historical_df = ytd_df[
        (ytd_df["Billing Date"] >= pd.Timestamp(start_date_selected)) &
        (ytd_df["Billing Date"] <= pd.Timestamp(end_date_selected))
    ].copy()

    if historical_df.empty:
        st.warning(f"⚠️ No sales data available in 'YTD' for {days_label}.")
        st.stop()

    historical_sales = historical_df.groupby(group_col)["Net Value"].sum()
    total_historical_sales_value = historical_sales.sum()
    
    current_month_sales_df = sales_df[
        (sales_df["Billing Date"].dt.month == today.month) &
        (sales_df["Billing Date"].dt.year == today.year)
    ].copy()
    current_month_sales = current_month_sales_df.groupby(group_col)["Net Value"].sum()
    total_current_month_sales = current_month_sales.sum()

    target_balance = total_target - total_current_month_sales

    # --- KPI Metrics ---
    if total_target > 0:
        average_historical_sales = total_historical_sales_value / months_count
        st.subheader("🎯 Target Analysis")
        col1, col2, col3 = st.columns(3)
        col4, col5 = st.columns(2)
        
        with col1:
            st.metric("Historical Sales Total", f"KD {total_historical_sales_value:,.0f}")
        with col2:
            st.metric("Allocated Target Total", f"KD {total_target:,.0f}")
        with col3:
            if average_historical_sales > 0:
                percentage_increase_needed = ((total_target - average_historical_sales) / average_historical_sales) * 100
                delta_value = total_target - average_historical_sales
                st.metric("Increase Needed vs Avg Sales",
                          f"{percentage_increase_needed:.2f}%",
                          delta=f"KD {delta_value:,.0f}")
            else:
                st.metric("Increase Needed vs Avg Sales", "N/A", delta="Historical = 0")
        
        st.markdown("---")
        with col4:
            st.metric("Current Month Sales", f"KD {total_current_month_sales:,.0f}")
        with col5:
            st.metric("Target Balance", f"KD {target_balance:,.0f}")

    # --- Allocation Table ---
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
    allocation_table_with_total = pd.concat([allocation_table, total_row])

    def color_target_balance(val):
        if isinstance(val, (int, float)):
            color = 'red' if val > 0 else 'green'
            return f'color: {color}'
        return ''

    st.subheader(f"📊 Auto-Allocated Targets Based on {days_label}")
    st.dataframe(
        allocation_table_with_total.astype(int).style
        .format("{:,.0f}")
        .applymap(color_target_balance, subset=['Target Balance'])
    )

    # --- Download Excel ---
    @st.cache_data
    def convert_df_to_excel(df):
        output = io.BytesIO()
        df = df.astype(int)
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, sheet_name="Allocated_Targets")
        return output.getvalue()

    excel_data = convert_df_to_excel(allocation_table)
    st.download_button(
        "💾 Download Target Allocation Table",
        data=excel_data,
        file_name=f"target_allocation_{allocation_type.replace(' ', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
