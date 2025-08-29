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

        # KPIs Slide
        slide_layout = prs.slide_layouts[5]
        slide = prs.slides.add_slide(slide_layout)
        slide.shapes.title.text = "📈 Key Performance Indicators"
        slide.shapes.title.text_frame.paragraphs[0].font.size = Pt(24)
        slide.shapes.title.text_frame.paragraphs[0].font.name = 'Roboto'
        slide.shapes.title.text_frame.paragraphs[0].font.color.rgb = RGBColor(30, 58, 138)

        # Mimic four-column layout
        left_cols = [Inches(0.5), Inches(3.5), Inches(6.5), Inches(9.5)]
        top_rows = [Inches(1.5), Inches(2.5), Inches(3.5)]
        metrics = [
            ("Total KA Sales", kpi_data["Total KA Sales"]),
            ("Total Talabat Sales", kpi_data["Total Talabat Sales"]),
            ("Total KA Gap", kpi_data["Total KA Gap"]),
            ("Total Talabat Gap", kpi_data["Total Talabat Gap"]),
            ("Top KA Salesman", kpi_data["Top KA Salesman"]),
            ("Top Talabat Salesman", kpi_data["Top Talabat Salesman"]),
            ("Days Finished (working)", kpi_data["Days Finished (working)"]),
            ("Per Day KA Target", kpi_data["Per Day KA Target"]),
            ("Current Sales Per Day", kpi_data["Current Sales Per Day"]),
            ("Forecasted Month-End KA Sales", kpi_data["Forecasted Month-End KA Sales"])
        ]
        for i, (label, value) in enumerate(metrics):
            row = i // 4
            col = i % 4
            textbox = slide.shapes.add_textbox(left_cols[col], top_rows[row], Inches(2.5), Inches(0.8))
            tf = textbox.text_frame
            p = tf.add_paragraph()
            p.text = f"{label}\n{value}"
            p.font.size = Pt(12)
            p.font.name = 'Roboto'
            p.font.color.rgb = RGBColor(55, 65, 81)  # #374151
            p.font.bold = True if "Top" in label else False
            tf.paragraphs[0].alignment = PP_ALIGN.CENTER

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
        add_table_slide(py_df.reset_index(), "🏬 Sales by PY Name 1")

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

# --- Sidebar Menu ---
st.sidebar.title("🧭 Menu")
menu = ["Home", "Sales Tracking", "YTD", "Custom Analysis", "SP/PY Target Allocation"]
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

            Use the sidebar to navigate.
            """,
            unsafe_allow_html=True
        )

# --- Sales Tracking Page ---
elif choice == "Sales Tracking":
    st.title("📊 Sales Tracking Dashboard")
    uploaded_file = st.file_uploader("📂 Upload Excel File", type=["xlsx"], key="sales_upload")

    if uploaded_file:
        sales_df, target_df, ytd_df = load_data(uploaded_file)
        if not sales_df.empty and not target_df.empty:
            with st.container():
                st.sidebar.subheader("🔍 Filters")
                salesmen = st.sidebar.multiselect(
                    "👥 Select Salesmen",
                    options=sorted(sales_df["Driver Name EN"].dropna().unique()),
                    default=sorted(sales_df["Driver Name EN"].dropna().unique()),
                )
                billing_types = st.sidebar.multiselect(
                    "📋 Select Billing Types",
                    options=sorted(sales_df["Billing Type"].dropna().unique()),
                    default=sorted(sales_df["Billing Type"].dropna().unique()),
                )
                py_filter = st.sidebar.multiselect(
                    "🏬 Select PY Name",
                    options=sorted(sales_df["PY Name 1"].dropna().unique()),
                    default=sorted(sales_df["PY Name 1"].dropna().unique()),
                )
                sp_filter = st.sidebar.multiselect(
                    "🏷️ Select SP Name1",
                    options=sorted(sales_df["SP Name1"].dropna().unique()),
                    default=sorted(sales_df["SP Name1"].dropna().unique()),
                )

                preset = st.sidebar.radio("📅 Quick Date Presets", ["Custom Range", "Last 7 Days", "This Month", "YTD"])
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

                            excel_stream1 = io.BytesIO()
                            with pd.ExcelWriter(excel_stream1, engine="xlsxwriter") as writer:
                                report_df.to_excel(writer, sheet_name="Sales_Targets_Summary")
                            excel_stream1.seek(0)
                            st.download_button(
                                "💾 Download Excel - Sales & Targets",
                                data=excel_stream1,
                                file_name="Sales_Targets_Summary.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            )

                            excel_stream2 = io.BytesIO()
                            with pd.ExcelWriter(excel_stream2, engine="xlsxwriter") as writer:
                                billing_df.to_excel(writer, sheet_name="Billing_Types")
                            excel_stream2.seek(0)
                            st.download_button(
                                "💾 Download Excel - Billing Types",
                                data=excel_stream2,
                                file_name="Billing_Types.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            )

                            st.subheader("🏬 Sales by PY Name 1")
                            py_df = df_filtered.groupby("PY Name 1")["Net Value"].sum().sort_values(ascending=False)
                            st.dataframe(
                                py_df.to_frame().style.background_gradient(cmap="Blues").format("{:,.0f}"),
                                use_container_width=True
                            )

                            excel_stream3 = io.BytesIO()
                            with pd.ExcelWriter(excel_stream3, engine="xlsxwriter") as writer:
                                py_df.to_frame().to_excel(writer, sheet_name="Sales_by_PY_Name")
                            excel_stream3.seek(0)
                            st.download_button(
                                "💾 Download Excel - PY Name Sales",
                                data=excel_stream3,
                                file_name="Sales_by_PY_Name.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            )

                        with tabs[2]:
                            st.subheader("📊 Sales Visualizations")
                            figs_dict = {}
                            report_df_for_chart = report_df.reset_index().rename(columns={'index': 'Driver Name EN'}) if report_df.index.name else report_df
                            fig_ka = px.bar(
                                report_df_for_chart,
                                x="Driver Name EN",
                                y="KA Sales",
                                title="KA Sales by Salesman",
                                text="KA Sales",
                                color_discrete_sequence=["#1E3A8A"]
                            )
                            fig_ka.update_traces(texttemplate='%{text:,.0f}', textposition='auto')
                            fig_ka.update_layout(
                                font=dict(family="Roboto", size=12),
                                plot_bgcolor="#F3F4F6",
                                paper_bgcolor="#F3F4F6",
                                showlegend=False,
                                xaxis_title="Salesman",
                                yaxis_title="KA Sales",
                                xaxis_tickangle=45
                            )
                            fig_talabat = px.bar(
                                report_df_for_chart,
                                x="Driver Name EN",
                                y="Talabat Sales",
                                title="Talabat Sales by Salesman",
                                text="Talabat Sales",
                                color_discrete_sequence=["#3B82F6"]
                            )
                            fig_talabat.update_traces(texttemplate='%{text:,.0f}', textposition='auto')
                            fig_talabat.update_layout(
                                font=dict(family="Roboto", size=12),
                                plot_bgcolor="#F3F4F6",
                                paper_bgcolor="#F3F4F6",
                                showlegend=False,
                                xaxis_title="Salesman",
                                yaxis_title="Talabat Sales",
                                xaxis_tickangle=45
                            )
                            st.plotly_chart(fig_ka, use_container_width=True)
                            st.plotly_chart(fig_talabat, use_container_width=True)
                            figs_dict["KA Sales by Salesman"] = fig_ka
                            figs_dict["Talabat Sales by Salesman"] = fig_talabat

                        with tabs[3]:
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
                                    xaxis_title="Date",
                                    yaxis_title="Net Value",
                                    legend_title="Data"
                                )
                                st.plotly_chart(fig_trend, use_container_width=True)
                                figs_dict["Sales Trend with Forecast"] = fig_trend
                            else:
                                st.info("ℹ️ Not enough data for sales trend forecast.")

                            # KA Target Daily Trend
                            st.subheader("🎯 KA Target Daily Trend")
                            if not df_time.empty:
                                daily_ka_target = per_day_ka_target
                                df_time["Daily KA Target"] = daily_ka_target
                                fig_ka_trend = px.line(
                                    df_time,
                                    x="Billing Date",
                                    y=["Net Value", "Daily KA Target"],
                                    title="KA Sales vs Daily Target",
                                    color_discrete_sequence=["#1E3A8A", "#EF4444"]
                                )
                                fig_ka_trend.update_layout(
                                    font=dict(family="Roboto", size=12),
                                    plot_bgcolor="#F3F4F6",
                                    paper_bgcolor="#F3F4F6",
                                    xaxis_title="Date",
                                    yaxis_title="Net Value",
                                    legend_title="Data"
                                )
                                st.plotly_chart(fig_ka_trend, use_container_width=True)
                                figs_dict["KA Sales vs Daily Target"] = fig_ka_trend
                            else:
                                st.info("ℹ️ Not enough data for KA target daily trend.")

# --- YTD Page ---
elif choice == "YTD":
    st.title("📅 Year-to-Date (YTD) Comparison")
    uploaded_file = st.file_uploader("📂 Upload Excel for YTD", type=["xlsx"], key="ytd_upload")

    if uploaded_file:
        sales_df, target_df, ytd_df = load_data(uploaded_file)
        if not sales_df.empty:
            with st.container():
                dimension = st.selectbox("📊 Compare By", ["Driver Name EN", "PY Name 1", "SP Name1"])

                st.subheader("📆 Select Two Periods")
                col1, col2 = st.columns(2)
                with col1:
                    period1 = st.date_input("Period 1 Start-End", [sales_df["Billing Date"].min(), sales_df["Billing Date"].max()])
                with col2:
                    period2 = st.date_input("Period 2 Start-End", [sales_df["Billing Date"].min(), sales_df["Billing Date"].max()])

                if period1[0] > period1[1] or period2[0] > period2[1]:
                    st.error("❌ Start date must be before end date.")
                else:
                    period1_label = f"{pd.to_datetime(period1[0]).strftime('%d-%b-%Y')} to {pd.to_datetime(period1[1]).strftime('%d-%b-%Y')}"
                    period2_label = f"{pd.to_datetime(period2[0]).strftime('%d-%b-%Y')} to {pd.to_datetime(period2[1]).strftime('%d-%b-%Y')}"

                    df1 = sales_df[(sales_df["Billing Date"] >= pd.to_datetime(period1[0])) & (sales_df["Billing Date"] <= pd.to_datetime(period1[1]))]
                    df2 = sales_df[(sales_df["Billing Date"] >= pd.to_datetime(period2[0])) & (sales_df["Billing Date"] <= pd.to_datetime(period2[1]))]

                    agg1 = df1.groupby(dimension)["Net Value"].sum()
                    agg2 = df2.groupby(dimension)["Net Value"].sum()
                    all_index = agg1.index.union(agg2.index)
                    agg1 = agg1.reindex(all_index, fill_value=0)
                    agg2 = agg2.reindex(all_index, fill_value=0)

                    comparison_df = pd.DataFrame({period1_label: agg1, period2_label: agg2})
                    comparison_df["Difference"] = comparison_df[period2_label] - comparison_df[period1_label]
                    comparison_df["Comparison %"] = np.where(comparison_df[period1_label] != 0, (comparison_df["Difference"] / comparison_df[period1_label] * 100).round(2), 0)
                    comparison_df = comparison_df.sort_values(by=period2_label, ascending=False)

                    def highlight_date_columns(row):
                        return [
                            "background-color: #CCFFE6; font-weight: bold" if col in [period1_label, period2_label]
                            else color_positive_negative(row[col]) for col in row.index
                        ]

                    st.subheader(f"📋 YTD Comparison by {dimension}")
                    st.dataframe(
                        comparison_df.style.format({
                            period1_label: "{:,.0f}",
                            period2_label: "{:,.0f}",
                            "Difference": "{:,.0f}",
                            "Comparison %": "{:.2f}%"
                        }).apply(highlight_date_columns, axis=1),
                        use_container_width=True
                    )

                    excel_stream_ytd = io.BytesIO()
                    with pd.ExcelWriter(excel_stream_ytd, engine="xlsxwriter") as writer:
                        comparison_df.to_excel(writer, sheet_name="YTD Comparison")
                    excel_stream_ytd.seek(0)
                    st.download_button(
                        "💾 Download Excel - YTD Comparison",
                        data=excel_stream_ytd,
                        file_name="YTD_Comparison.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

                    st.subheader("📊 YTD Comparison Chart")
                    fig = go.Figure()
                    fig.add_trace(go.Bar(
                        x=comparison_df.index,
                        y=comparison_df[period1_label],
                        name=period1_label,
                        marker_color="#1E3A8A"
                    ))
                    fig.add_trace(go.Bar(
                        x=comparison_df.index,
                        y=comparison_df[period2_label],
                        name=period2_label,
                        marker_color="#3B82F6"
                    ))
                    fig.add_trace(go.Scatter(
                        x=comparison_df.index,
                        y=comparison_df["Comparison %"],
                        name="% Difference",
                        mode="lines+markers",
                        yaxis="y2",
                        line=dict(color="#EF4444")
                    ))
                    fig.update_layout(
                        barmode="group",
                        title=f"YTD Comparison by {dimension}",
                        xaxis_title=dimension,
                        yaxis=dict(title="Net Value"),
                        yaxis2=dict(title="% Difference", overlaying="y", side="right", showgrid=False),
                        font=dict(family="Roboto", size=12),
                        plot_bgcolor="#F3F4F6",
                        paper_bgcolor="#F3F4F6",
                        xaxis_tickangle=45
                    )
                    st.plotly_chart(fig, use_container_width=True)

# --- Custom Analysis ---
elif choice == "Custom Analysis":
    st.title("🔍 Custom Analysis")
    uploaded_file = st.file_uploader("📂 Upload Excel for Analysis", type=["xlsx"], key="custom_upload")

    if uploaded_file:
        sales_df, target_df, ytd_df = load_data(uploaded_file)
        if not sales_df.empty:
            with st.container():
                with st.expander("👀 View Available Columns"):
                    col_info = pd.DataFrame({"Column": sales_df.columns, "Dtype": [str(sales_df[c].dtype) for c in sales_df.columns]})
                    st.dataframe(col_info, use_container_width=True)

                if "Billing Date" in sales_df.columns:
                    sales_df["Billing Month"] = sales_df["Billing Date"].dt.to_period("M").astype(str)
                    sales_df["Billing Year"] = sales_df["Billing Date"].dt.year.astype(str)

                st.subheader("1️⃣ Choose Grouping Columns")
                categorical_cols = [
                    c for c in sales_df.columns
                    if (sales_df[c].dtype == "object" or str(sales_df[c].dtype) == "category" or sales_df[c].dtype == "bool")
                ]
                for c in ["Billing Month", "Billing Year"]:
                    if c in sales_df.columns and c not in categorical_cols:
                        categorical_cols.append(c)
                if "Net Value" in categorical_cols:
                    categorical_cols.remove("Net Value")

                default_dims = ["Driver Name EN"] if "Driver Name EN" in categorical_cols else (categorical_cols[:1] if categorical_cols else [])
                dims = st.multiselect("📋 Group by (1–3 columns)", options=categorical_cols, default=default_dims, max_selections=3)

                if len(dims) == 0:
                    st.warning("⚠️ Please select at least one column to group by.")
                else:
                    st.subheader("2️⃣ Apply Filters")
                    filt_cols = ["Driver Name EN", "Billing Type", "PY Name 1", "SP Name1"]
                    filter_widgets = {}
                    fl_cols = st.columns(len(filt_cols))
                    for i, colname in enumerate(filt_cols):
                        if colname in sales_df.columns:
                            options = sorted(sales_df[colname].dropna().unique().tolist())
                            filter_widgets[colname] = fl_cols[i].multiselect(f"🔍 Filter: {colname}", options=options, default=options)

                    st.subheader("3️⃣ Select Periods")
                    left, right = st.columns(2)
                    min_d, max_d = sales_df["Billing Date"].min(), sales_df["Billing Date"].max()
                    with left:
                        period1 = st.date_input("📅 Period 1", [min_d, max_d], key="custom_p1")
                    with right:
                        period2 = st.date_input("📅 Period 2", [min_d, max_d], key="custom_p2")

                    if period1[0] > period1[1] or period2[0] > period2[1]:
                        st.error("❌ Start date must be before end date.")
                    else:
                        p1_label = f"{pd.to_datetime(period1[0]).strftime('%d-%b-%Y')} to {pd.to_datetime(period1[1]).strftime('%d-%b-%Y')}"
                        p2_label = f"{pd.to_datetime(period2[0]).strftime('%d-%b-%Y')} to {pd.to_datetime(period2[1]).strftime('%d-%b-%Y')}"

                        df = sales_df.copy()
                        for colname, selected in filter_widgets.items():
                            if colname in df.columns and len(selected) != len(df[colname].dropna().unique()):
                                df = df[df[colname].isin(selected)]

                        df_p1 = df[(df["Billing Date"] >= pd.to_datetime(period1[0])) & (df["Billing Date"] <= pd.to_datetime(period1[1]))]
                        df_p2 = df[(df["Billing Date"] >= pd.to_datetime(period2[0])) & (df["Billing Date"] <= pd.to_datetime(period2[1]))]

                        agg1 = df_p1.groupby(dims)["Net Value"].sum().rename(p1_label)
                        agg2 = df_p2.groupby(dims)["Net Value"].sum().rename(p2_label)
                        full_index = agg1.index.union(agg2.index)
                        agg1 = agg1.reindex(full_index, fill_value=0)
                        agg2 = agg2.reindex(full_index, fill_value=0)

                        comparison_dyn = pd.concat([agg1, agg2], axis=1)
                        comparison_dyn["Difference"] = comparison_dyn[p2_label] - comparison_dyn[p1_label]
                        comparison_dyn["Comparison %"] = np.where(comparison_dyn[p1_label] != 0, (comparison_dyn["Difference"] / comparison_dyn[p1_label] * 100).round(2), 0)
                        comparison_dyn = comparison_dyn.sort_values(by=p2_label, ascending=False)

                        st.subheader("4️⃣ Results")
                        st.dataframe(
                            comparison_dyn.style.format({
                                p1_label: "{:,.0f}",
                                p2_label: "{:,.0f}",
                                "Difference": "{:,.0f}",
                                "Comparison %": "{:.2f}%"
                            }).applymap(color_positive_negative, subset=["Difference", "Comparison %"]),
                            use_container_width=True
                        )

                        excel_stream_dyn = io.BytesIO()
                        with pd.ExcelWriter(excel_stream_dyn, engine="xlsxwriter") as writer:
                            comparison_dyn.to_excel(writer, sheet_name="Custom_Comparison")
                            comparison_dyn.reset_index().to_excel(writer, sheet_name="Custom_Comparison_Flat", index=False)
                        excel_stream_dyn.seek(0)
                        st.download_button(
                            "💾 Download Excel - Custom Comparison",
                            data=excel_stream_dyn,
                            file_name="Custom_Comparison.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )

                        st.subheader("5️⃣ Visualizations")
                        flat_df = comparison_dyn.reset_index()
                        if len(dims) == 1:
                            xcol = dims[0]
                        else:
                            xcol = "__label__"
                            flat_df[xcol] = flat_df[dims].astype(str).agg(" | ".join, axis=1)

                        fig = go.Figure()
                        fig.add_trace(go.Bar(
                            x=flat_df[xcol],
                            y=flat_df[p1_label],
                            name=p1_label,
                            marker_color="#1E3A8A"
                        ))
                        fig.add_trace(go.Bar(
                            x=flat_df[xcol],
                            y=flat_df[p2_label],
                            name=p2_label,
                            marker_color="#3B82F6"
                        ))
                        fig.update_layout(
                            barmode="group",
                            xaxis_title=" x ".join(dims),
                            yaxis_title="Net Value",
                            title="Custom Comparison",
                            font=dict(family="Roboto", size=12),
                            plot_bgcolor="#F3F4F6",
                            paper_bgcolor="#F3F4F6",
                            xaxis_tickangle=45
                        )
                        st.plotly_chart(fig, use_container_width=True)

                        fig2 = go.Figure()
                        fig2.add_trace(go.Scatter(
                            x=flat_df[xcol],
                            y=flat_df["Comparison %"],
                            mode="lines+markers",
                            name="Comparison %",
                            line=dict(color="#EF4444")
                        ))
                        fig2.update_layout(
                            xaxis_title=" x ".join(dims),
                            yaxis_title="% Difference",
                            title="Custom Comparison - % Difference",
                            font=dict(family="Roboto", size=12),
                            plot_bgcolor="#F3F4F6",
                            paper_bgcolor="#F3F4F6",
                            xaxis_tickangle=45
                        )
                        st.plotly_chart(fig2, use_container_width=True)

# --- SP/PY Target Allocation Page ---
elif choice == "SP/PY Target Allocation":
    st.title("🎯 SP/PY Target Allocation")
    uploaded_file = st.file_uploader("📂 Upload Excel File", type=["xlsx"], key="allocation_upload")

    if uploaded_file:
        try:
            xls = pd.ExcelFile(uploaded_file)
            
            required_sheets = {"YTD", "sales data"}
            if not required_sheets.issubset(xls.sheet_names):
                st.error(f"❌ The uploaded file must contain the following sheets: {', '.join(required_sheets)}.")
                st.stop()
            
            ytd_df = pd.read_excel(xls, sheet_name="YTD")
            sales_df = pd.read_excel(xls, sheet_name="sales data")

            # Validate required columns
            required_cols_ytd = ["Billing Date", "SP Name1", "PY Name 1", "Net Value"]
            if not all(col in ytd_df.columns for col in required_cols_ytd):
                st.error(f"❌ Missing required columns in 'YTD' sheet: {set(required_cols_ytd) - set(ytd_df.columns)}")
                st.stop()

            required_cols_sales = ["Billing Date", "SP Name1", "PY Name 1", "Net Value"]
            if not all(col in sales_df.columns for col in required_cols_sales):
                st.error(f"❌ Missing required columns in 'sales data' sheet: {set(required_cols_sales) - set(sales_df.columns)}")
                st.stop()

            # Ensure date columns are in datetime format
            ytd_df["Billing Date"] = pd.to_datetime(ytd_df["Billing Date"], errors='coerce')
            sales_df["Billing Date"] = pd.to_datetime(sales_df["Billing Date"], errors='coerce')

        except Exception as e:
            st.error(f"❌ Error loading or processing Excel file: {e}")
            st.stop()

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
            if "Target" not in xls.sheet_names:
                st.error("❌ 'Target' sheet not found. Please upload a file with this sheet for 'Auto' mode.")
                st.stop()
            try:
                ka_target_df = pd.read_excel(xls, sheet_name="Target")
                if "KA Target" not in ka_target_df.columns:
                    st.error("❌ 'KA Target' column not found in 'Target' sheet.")
                    st.stop()
                total_target = ka_target_df["KA Target"].sum()
                st.info(f"Using Total Target from 'Target' sheet: KD {total_target:,.0f}")
            except Exception as e:
                st.error(f"❌ Error reading 'Target' sheet: {e}")
                st.stop()

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
            # New and corrected code for manual date selection
            start_date_manual = today - pd.DateOffset(days=180) # Default to 180 days ago
            selected_dates = st.date_input("Select date range", value=(start_date_manual, today))
            
            if len(selected_dates) == 2:
                start_date_selected = selected_dates[0]
                end_date_selected = selected_dates[1]
                
                lookback_period = end_date_selected - start_date_selected
                days_label = f"From {start_date_selected.strftime('%Y-%m-%d')} to {end_date_selected.strftime('%Y-%m-%d')}"
                months_count = lookback_period.days / 30  # Approximation for average monthly sales calculation
            else:
                st.warning("Please select both a start and an end date.")
                st.stop()

        # --- Data Processing ---
        # The start_date and end_date are now determined by the date input
        # No need to recalculate them here
        historical_df = ytd_df[
            (ytd_df["Billing Date"] >= pd.Timestamp(start_date_selected)) &
            (ytd_df["Billing Date"] <= pd.Timestamp(end_date_selected))
        ].copy()

        if historical_df.empty:
            st.warning(f"⚠️ No sales data available in the 'YTD' sheet for the last {days_label}.")
            st.stop()

        historical_sales = historical_df.groupby(group_col)["Net Value"].sum()
        total_historical_sales_value = historical_sales.sum()
        
        # Calculate current month sales total for the new metric
        current_month_sales_df = sales_df[
            (sales_df["Billing Date"].dt.month == today.month) &
            (sales_df["Billing Date"].dt.year == today.year)
        ].copy()
        current_month_sales = current_month_sales_df.groupby(group_col)["Net Value"].sum()
        total_current_month_sales = current_month_sales.sum()

        # Calculate target balance total for the new metric
        target_balance = total_target - total_current_month_sales
        
        # Calculate and display the new metrics
        if total_target > 0:
            average_historical_sales = total_historical_sales_value / months_count
            
            st.subheader("🎯 Target Analysis")
            
            col1, col2, col3 = st.columns(3)
            col4, col5 = st.columns(2)
            
            with col1:
                st.metric(
                    label=f"Historical Sales Total ({days_label})",
                    value=f"KD {total_historical_sales_value:,.0f}"
                )
            
            with col2:
                st.metric(
                    label="Allocated Target Total",
                    value=f"KD {total_target:,.0f}"
                )
                
            with col3:
                # New calculation to show percentage increase needed from historical sales
                if average_historical_sales > 0:
                    percentage_increase_needed = ((total_target - average_historical_sales) / average_historical_sales) * 100
                    delta_value = total_target - average_historical_sales
                    st.metric(
                        label="Percentage Increase from Avg Sales to Target",
                        value=f"{percentage_increase_needed:.2f}%",
                        delta=f"KD {delta_value:,.0f} more needed"
                    )
                else:
                    st.metric(
                        label="Percentage Increase from Avg Sales to Target",
                        value="N/A",
                        delta="Historical sales are zero"
                    )
            
            st.markdown("---")
            
            with col4:
                st.metric(
                    label="Current Month Sales Total",
                    value=f"KD {total_current_month_sales:,.0f}"
                )
                
            with col5:
                st.metric(
                    label="Total Target Balance",
                    value=f"KD {target_balance:,.0f}",
                    delta=f"{'KD' if target_balance >= 0 else ''} {target_balance:,.0f}"
                )

        # --- Create the Target Allocation Table ---
        current_month_sales = current_month_sales_df.groupby(group_col)["Net Value"].sum()
        
        allocation_table = pd.DataFrame(index=historical_sales.index.union(current_month_sales.index).unique())
        allocation_table.index.name = "Name"

        allocation_table[f"Last {days_label} Total Sales"] = historical_sales.reindex(allocation_table.index, fill_value=0)
        allocation_table[f"Last {days_label} Average Sales"] = (allocation_table[f"Last {days_label} Total Sales"] / months_count).fillna(0)
        
        if total_historical_sales_value > 0:
            allocation_table["This Month Auto-Allocated Target"] = (allocation_table[f"Last {days_label} Total Sales"] / total_historical_sales_value * total_target).fillna(0)
        else:
            allocation_table["This Month Auto-Allocated Target"] = 0
            st.warning("Total historical sales for the selected period is zero. Cannot perform performance-based allocation. Targets are set to 0.")

        allocation_table["Current Month Sales"] = current_month_sales.reindex(allocation_table.index, fill_value=0)
        
        # Rename the column
        allocation_table["Target Balance"] = allocation_table["This Month Auto-Allocated Target"] - allocation_table["Current Month Sales"]
        
        allocation_table = allocation_table.fillna(0)

        # Add a total row
        total_row = allocation_table.sum().to_frame().T
        total_row.index = ["Total"]
        allocation_table_with_total = pd.concat([allocation_table, total_row])

        # Function to apply color styling
        def color_target_balance(val):
            # Only apply color to numbers, not the 'Total' label
            if isinstance(val, (int, float)):
                color = 'red' if val > 0 else 'green'
                return f'color: {color}'
            return ''

        # Apply bold styling to all cells in the dataframe
        def bold_all(val):
            return 'font-weight: bold'

        # Display the table with integer formatting and styling
        st.subheader(f"📊 Auto-Allocated Targets Based on Last {days_label} Performance")
        st.dataframe(
            allocation_table_with_total.astype(int).style
            .format("{:,.0f}")
            .applymap(color_target_balance, subset=['Target Balance'])
            .applymap(bold_all),
            use_container_width=True
        )

        # --- Download Button ---
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
