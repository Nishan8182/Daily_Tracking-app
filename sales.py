import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import io

# --- Page config ---
st.set_page_config(page_title=" Welcome to Haneef Data", layout="wide")

# --- Cache function ---
@st.cache_data
def load_data(file):
    try:
        sales_df = pd.read_excel(file, sheet_name="sales data")
        target_df = pd.read_excel(file, sheet_name="Target")
        return sales_df, target_df
    except Exception as e:
        st.error(f"Error loading Excel file: {e}")
        return None, None

# --- PPTX helper ---
def create_pptx(report_df, billing_type_df, py_name_df, fig_sales, fig_ka, fig_talabat, fig_daily, fig_daily_ka):
    prs = Presentation()

    # Title slide
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "Sales & Targets Report"
    slide.placeholders[1].text = "Generated from Sales Data"

    # Helper to add table slide
    def add_table_slide(df, title, cmap=None):
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = title
        rows, cols = df.shape
        table = slide.shapes.add_table(rows+1, cols, Inches(0.5), Inches(1.5), Inches(9), Inches(4)).table
        # Header
        for j, col in enumerate(df.columns):
            table.cell(0, j).text = str(col)
            table.cell(0, j).text_frame.paragraphs[0].font.bold = True
        # Data
        for i, row in enumerate(df.itertuples(index=False), start=1):
            for j, val in enumerate(row):
                table.cell(i, j).text = f"{int(val):,}" if isinstance(val, (int, float)) else str(val)

    # Helper to add chart slide
    def add_chart_slide(fig, title):
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = title
        img_stream = io.BytesIO()
        fig.savefig(img_stream, format='png', bbox_inches='tight')
        img_stream.seek(0)
        slide.shapes.add_picture(img_stream, Inches(0.5), Inches(1.5), width=Inches(8))

    # Add tables
    add_table_slide(report_df, "Sales & Targets Summary")
    add_table_slide(billing_type_df.reset_index(), "Sales by Billing Type per Salesman")
    add_table_slide(py_name_df.reset_index(), "Sales by PY Name 1")

    # Add charts
    add_chart_slide(fig_sales, "Sales & Targets by Salesman")
    add_chart_slide(fig_ka, "KA Target vs Sales")
    add_chart_slide(fig_talabat, "Talabat Target vs Sales")
    add_chart_slide(fig_daily, "Daily Sales Trend - All Salesmen")
    add_chart_slide(fig_daily_ka, "Daily KA Sales Trend")

    pptx_stream = io.BytesIO()
    prs.save(pptx_stream)
    pptx_stream.seek(0)
    return pptx_stream

# --- Sidebar ---
st.sidebar.title("Menu")
menu = ["Home", "Sales Tracking"]
choice = st.sidebar.selectbox("Menu", menu)

# --- Home ---
if choice == "Home":
    st.title("🏠 Welcome to Sales Tracking Dashboard")
    st.markdown("### Use the sidebar to navigate to Sales Tracking.")

# --- Sales Tracking ---
elif choice == "Sales Tracking":
    st.title("📊 Sales Tracking Dashboard")

    uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])
    if uploaded_file:
        sales_df, target_df = load_data(uploaded_file)
        if sales_df is not None and target_df is not None:

            # --- Process Sales & Targets ---
            total_by_salesman = sales_df.groupby("Driver Name EN")["Net Value"].sum()
            talabat_df = sales_df[sales_df["PY Name 1"] == "STORES SERVICES KUWAIT CO."]
            talabat_sales_by_salesman = talabat_df.groupby("Driver Name EN")["Net Value"].sum()

            targets_series = target_df.set_index("Driver Name EN")["KA Target"]
            talabat_targets_series = target_df.set_index("Driver Name EN")["Talabat Target"]

            all_salesmen = total_by_salesman.index.union(talabat_sales_by_salesman.index).union(targets_series.index).union(talabat_targets_series.index)
            total_by_salesman = total_by_salesman.reindex(all_salesmen, fill_value=0)
            talabat_sales_by_salesman = talabat_sales_by_salesman.reindex(all_salesmen, fill_value=0)
            targets_series = targets_series.reindex(all_salesmen, fill_value=0)
            talabat_targets_series = talabat_targets_series.reindex(all_salesmen, fill_value=0)

            gap_ka = (targets_series - total_by_salesman).clip(lower=0)
            gap_talabat = (talabat_targets_series - talabat_sales_by_salesman).clip(lower=0)

            report_df = pd.DataFrame({
                "KA Total Sales": total_by_salesman,
                "KA Remaining": gap_ka,
                "Talabat Sales": talabat_sales_by_salesman,
                "Talabat Remaining": gap_talabat,
                "KA Target": targets_series,
                "Talabat Target": talabat_targets_series
            })

            st.subheader("Sales & Targets Table")
            st.dataframe(report_df.style.format("{:,.0f}").background_gradient(cmap="Blues"), use_container_width=True)

            # Billing Type Table
            billing_type_by_salesman = sales_df.groupby(["Driver Name EN", "Billing Type"])["Net Value"].sum().unstack(fill_value=0)
            billing_type_by_salesman["Total"] = billing_type_by_salesman.sum(axis=1)
            st.subheader("Sales by Billing Type per Salesman")
            st.dataframe(billing_type_by_salesman.style.format("{:,.0f}").background_gradient(cmap="Blues"), use_container_width=True)

            # PY Name 1 Table
            py_name_df = sales_df.groupby("PY Name 1")["Net Value"].sum()
            st.subheader("Sales by PY Name 1")
            st.dataframe(py_name_df.to_frame().style.background_gradient(cmap="Greens").format("{:,.0f}"), use_container_width=True)

            # --- Charts ---
            # Sales & Targets by Salesman
            fig_sales, ax = plt.subplots(figsize=(12, 6))
            bar_width = 0.25
            y_pos = np.arange(len(all_salesmen))
            pos_total = y_pos - bar_width/2
            pos_talabat = y_pos + bar_width/2
            ax.barh(pos_total, total_by_salesman, height=bar_width, color='skyblue', label='KA Sales')
            ax.barh(pos_total, gap_ka, left=total_by_salesman, height=bar_width, color='lightgray', label='KA Gap')
            ax.barh(pos_talabat, talabat_sales_by_salesman, height=bar_width, color='orange', label='Talabat Sales')
            ax.barh(pos_talabat, gap_talabat, left=talabat_sales_by_salesman, height=bar_width, color='lightgreen', label='Talabat Gap')
            ax.set_yticks(y_pos)
            ax.set_yticklabels(all_salesmen)
            ax.invert_yaxis()
            for i, v in enumerate(total_by_salesman):
                ax.text(v+gap_ka[i]+500, pos_total[i], f"{int(v):,}", va='center')
            for i, v in enumerate(talabat_sales_by_salesman):
                ax.text(v+gap_talabat[i]+500, pos_talabat[i], f"{int(v):,}", va='center')
            ax.set_title("Sales & Targets by Salesman")
            ax.legend()
            st.pyplot(fig_sales)

            # KA Pie
            fig_ka, ax2 = plt.subplots()
            ax2.pie([total_by_salesman.sum(), gap_ka.sum()], labels=[f"Sales {int(total_by_salesman.sum()):,}", f"Gap {int(gap_ka.sum()):,}"], colors=['skyblue', 'lightgray'])
            ax2.set_title("KA Target vs Sales")
            st.pyplot(fig_ka)

            # Talabat Pie
            fig_talabat, ax3 = plt.subplots()
            ax3.pie([talabat_sales_by_salesman.sum(), gap_talabat.sum()], labels=[f"Sales {int(talabat_sales_by_salesman.sum()):,}", f"Gap {int(gap_talabat.sum()):,}"], colors=['orange', 'lightgreen'])
            ax3.set_title("Talabat Target vs Sales")
            st.pyplot(fig_talabat)

            # Daily Sales Trend - All Salesmen
            daily_sales = sales_df.groupby(["Billing Date", "Driver Name EN"])["Net Value"].sum().reset_index()
            fig_daily, ax4 = plt.subplots(figsize=(12, 6))
            for s in all_salesmen:
                sub_df = daily_sales[daily_sales["Driver Name EN"] == s]
                ax4.plot(sub_df["Billing Date"], sub_df["Net Value"], marker='o', label=s)
            ax4.set_title("Daily Sales Trend - All Salesmen")
            plt.xticks(rotation=45)
            ax4.legend()
            st.pyplot(fig_daily)

            # Daily KA Sales Trend
            daily_ka_sales = sales_df.groupby("Billing Date")["Net Value"].sum()
            fig_daily_ka, ax5 = plt.subplots()
            ax5.plot(daily_ka_sales.index, daily_ka_sales.values, marker='o', color='skyblue')
            ax5.set_title("Daily KA Sales Trend")
            plt.xticks(rotation=45)
            st.pyplot(fig_daily_ka)

            # --- Download PPTX ---
            if st.button("Download Report as PPTX"):
                pptx_data = create_pptx(report_df, billing_type_by_salesman, py_name_df, fig_sales, fig_ka, fig_talabat, fig_daily, fig_daily_ka)
                st.download_button("Download PPTX", data=pptx_data, file_name="sales_report.pptx",
                                   mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")

    else:
        st.info("Please upload your Excel file with sheets 'sales data' and 'Target'.")
