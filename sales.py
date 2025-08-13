import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from pptx import Presentation
from pptx.util import Inches
import io
import os

st.set_page_config(page_title="Sales Tracking Dashboard", layout="wide")

# --- Helper function to load Excel ---
def load_data(uploaded_file):
    try:
        xls = pd.ExcelFile(uploaded_file)
        sales_df = pd.read_excel(xls, sheet_name="sales data")
        target_df = pd.read_excel(xls, sheet_name="Target")
        return sales_df, target_df
    except Exception as e:
        st.error(f"Error loading Excel file: {e}")
        return None, None

# --- PowerPoint export functions ---
def add_table_slide(prs, df, title, cmap="Blues"):
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = title
    rows, cols = df.shape
    table = slide.shapes.add_table(rows+1, cols, Inches(0.5), Inches(1.5), Inches(9), Inches(4)).table

    # Header
    for j, col_name in enumerate(df.columns):
        table.cell(0, j).text = str(col_name)

    # Body
    for i, row in enumerate(df.itertuples(index=False), start=1):
        for j, value in enumerate(row):
            if isinstance(value, (int, float)):
                table.cell(i, j).text = f"{int(value):,}"
            else:
                table.cell(i, j).text = str(value)

def add_chart_slide(prs, fig, title):
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = title
    img_stream = io.BytesIO()
    fig.savefig(img_stream, format='png', bbox_inches='tight')
    img_stream.seek(0)
    slide.shapes.add_picture(img_stream, Inches(0.5), Inches(1.5), width=Inches(8))

def create_pptx(report_df, billing_type_df, py_name_df, fig_salesman, fig_ka, fig_talabat, fig_trend, fig_ka_trend):
    prs = Presentation()
    # Title slide
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "Sales & Targets Report"
    slide.placeholders[1].text = "Generated from Sales Data"

    # Sales & Targets Table
    add_table_slide(prs, report_df.reset_index(), "Sales & Targets Summary")

    # Billing Type Table
    add_table_slide(prs, billing_type_df.reset_index(), "Sales by Billing Type per Salesman")

    # PY Name 1 Table
    add_table_slide(prs, py_name_df.reset_index(), "Sales by PY Name 1")

    # Charts
    add_chart_slide(prs, fig_salesman, "Sales and Targets by Salesman")
    add_chart_slide(prs, fig_ka, "KA Target vs Sales")
    add_chart_slide(prs, fig_talabat, "Talabat Target vs Sales")
    add_chart_slide(prs, fig_trend, "Daily Sales Trend - All Salesmen")
    add_chart_slide(prs, fig_ka_trend, "Daily KA Sales Trend")

    pptx_stream = io.BytesIO()
    prs.save(pptx_stream)
    pptx_stream.seek(0)
    return pptx_stream

# --- MAIN APP ---
st.title("📊 Sales Tracking Dashboard")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    sales_df, target_df = load_data(uploaded_file)

    if sales_df is not None and target_df is not None:
        # All Salesmen list
        all_salesmen = sales_df["Driver Name EN"].unique()
        # KA & Talabat sales
        ka_sales = sales_df.groupby("Driver Name EN")["Net Value"].sum()
        talabat_df = sales_df[sales_df["PY Name 1"]=="STORES SERVICES KUWAIT CO."]
        talabat_sales = talabat_df.groupby("Driver Name EN")["Net Value"].sum()

        # Targets
        ka_targets = target_df.set_index("Driver Name EN")["KA Target"]
        talabat_targets = target_df.set_index("Driver Name EN")["Talabat Target"]

        # Combine all salesmen
        all_index = ka_sales.index.union(talabat_sales.index).union(ka_targets.index).union(talabat_targets.index)
        ka_sales = ka_sales.reindex(all_index, fill_value=0)
        talabat_sales = talabat_sales.reindex(all_index, fill_value=0)
        ka_targets = ka_targets.reindex(all_index, fill_value=0)
        talabat_targets = talabat_targets.reindex(all_index, fill_value=0)

        # Remaining gaps
        gap_ka = (ka_targets - ka_sales).clip(lower=0)
        gap_talabat = (talabat_targets - talabat_sales).clip(lower=0)

        # Report Table
        report_df = pd.DataFrame({
            "KA Total Sales": ka_sales,
            "KA Target": ka_targets,
            "Remaining to KA Target": gap_ka,
            "Talabat Sales": talabat_sales,
            "Talabat Target": talabat_targets,
            "Remaining to Talabat Target": gap_talabat
        })

        st.subheader("Sales & Targets Summary")
        st.dataframe(report_df.style.format("{:,.0f}").background_gradient(cmap="Blues"), use_container_width=True)

        # Billing Type
        billing_type_df = sales_df.groupby(["Driver Name EN","Billing Type"])["Net Value"].sum().unstack(fill_value=0)
        billing_type_df["Total"] = billing_type_df.sum(axis=1)
        st.subheader("Sales by Billing Type per Salesman")
        st.dataframe(billing_type_df.style.format("{:,.0f}").background_gradient(cmap="Blues"), use_container_width=True)

        # PY Name 1
        py_name_df = sales_df.groupby("PY Name 1")["Net Value"].sum().reset_index()
        st.subheader("Sales by PY Name 1")
        st.dataframe(py_name_df.style.format("{:,.0f}").background_gradient(subset=["Net Value"], cmap="Greens"), use_container_width=True)

        # Sales & Targets by Salesman Chart
        fig_salesman, ax = plt.subplots(figsize=(12,6))
        y_pos = np.arange(len(all_index))
        bar_width = 0.35
        ax.bar(y_pos - bar_width/2, ka_sales, bar_width, label="KA Sales", color="skyblue")
        ax.bar(y_pos - bar_width/2, gap_ka, bar_width, bottom=ka_sales, color="lightgray", label="KA Gap")
        ax.bar(y_pos + bar_width/2, talabat_sales, bar_width, label="Talabat Sales", color="orange")
        ax.bar(y_pos + bar_width/2, gap_talabat, bar_width, bottom=talabat_sales, color="lightgreen", label="Talabat Gap")
        ax.set_xticks(y_pos)
        ax.set_xticklabels(all_index, rotation=45)
        ax.set_ylabel("Value")
        for i, v in enumerate(ka_sales):
            ax.text(i - bar_width/2, v + 0.02*v, f"{int(v):,}", ha='center', va='bottom')
        for i, v in enumerate(talabat_sales):
            ax.text(i + bar_width/2, v + 0.02*v, f"{int(v):,}", ha='center', va='bottom')
        ax.legend()
        st.pyplot(fig_salesman)

        # KA Pie Chart
        fig_ka, ax2 = plt.subplots()
        ax2.pie([ka_sales.sum(), gap_ka.sum()], labels=["Sales","Gap"], autopct=lambda p: f'{int(round(p/100*ka_targets.sum())):,}')
        ax2.set_title("KA Target vs Sales")
        st.pyplot(fig_ka)

        # Talabat Pie Chart
        fig_talabat, ax3 = plt.subplots()
        ax3.pie([talabat_sales.sum(), gap_talabat.sum()], labels=["Sales","Gap"], autopct=lambda p: f'{int(round(p/100*talabat_targets.sum())):,}')
        ax3.set_title("Talabat Target vs Sales")
        st.pyplot(fig_talabat)

        # Daily Sales Trend
        daily_trend = sales_df.groupby(["Billing Date","Driver Name EN"])["Net Value"].sum().reset_index()
        fig_trend, ax4 = plt.subplots(figsize=(12,6))
        for salesman in all_index:
            sub = daily_trend[daily_trend["Driver Name EN"]==salesman]
            ax4.plot(sub["Billing Date"], sub["Net Value"], marker='o', label=salesman)
        ax4.set_title("Daily Sales Trend - All Salesmen")
        ax4.set_ylabel("Value")
        ax4.legend()
        plt.xticks(rotation=45)
        st.pyplot(fig_trend)

        # Daily KA Trend
        daily_ka = sales_df.groupby("Billing Date")["Net Value"].sum().reset_index()
        fig_ka_trend, ax5 = plt.subplots(figsize=(12,6))
        ax5.plot(daily_ka["Billing Date"], daily_ka["Net Value"], marker='o', color='skyblue')
        ax5.set_title("Daily KA Sales Trend")
        ax5.set_ylabel("Value")
        plt.xticks(rotation=45)
        st.pyplot(fig_ka_trend)

        # Download PPTX
        if st.button("Download PowerPoint Report"):
            pptx_data = create_pptx(report_df, billing_type_df, py_name_df, fig_salesman, fig_ka, fig_talabat, fig_trend, fig_ka_trend)
            st.download_button(
                label="Download PPTX",
                data=pptx_data,
                file_name="sales_report.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
else:
    st.info("Please upload your Excel file with sheets 'sales data' and 'Target'.")
