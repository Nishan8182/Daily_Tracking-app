import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from fpdf import FPDF
import tempfile
import os
from datetime import datetime
from matplotlib import colors

st.set_page_config(page_title="Sales & Targets Dashboard", layout="wide")
st.title("📊 Sales and Targets by Salesman - KA and Talabat")

@st.cache_data
def load_excel(uploaded_file):
    sales_df = pd.read_excel(uploaded_file, sheet_name="sales data")
    target_df = pd.read_excel(uploaded_file, sheet_name="Target")
    return sales_df, target_df

# Format numbers into 3-digit style (1.2K, 1.5M)
def format_k(val):
    if val >= 1_000_000:
        return f"{val/1_000_000:.1f}M"
    elif val >= 1_000:
        return f"{val/1_000:.1f}K"
    else:
        return f"{val:.0f}"

def add_fig_to_pdf(pdf, fig, title):
    pdf.add_page()
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 10, title, 0, 1, 'C')
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmpfile:
        fig.savefig(tmpfile.name, bbox_inches='tight')
        tmpfile_path = tmpfile.name
    pdf.image(tmpfile_path, x=10, y=25, w=pdf.w - 20)
    os.remove(tmpfile_path)

def add_colored_table(pdf, df, title, cmap='YlOrRd', max_bar_width=50):
    pdf.add_page()
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 10, title, 0, 1, 'C')
    pdf.set_font("Arial", size=10)
    numeric_cols = df.select_dtypes(include=np.number).columns.tolist()
    norm = colors.Normalize(vmin=df[numeric_cols].min().min(), vmax=df[numeric_cols].max().max())
    sm = plt.cm.ScalarMappable(cmap=cmap, norm=norm)
    sm.set_array([])
    col_names = df.columns.tolist()
    pdf.set_font("Arial", 'B', 10)
    for col in col_names:
        pdf.cell(max_bar_width+10, 6, col, border=1, ln=0, align='C')
    pdf.ln()
    pdf.set_font("Arial", '', 10)
    for i, row in df.iterrows():
        for col in col_names:
            val = row[col]
            if col in numeric_cols:
                rgba = sm.to_rgba(val)
                hex_color = (int(rgba[0]*255), int(rgba[1]*255), int(rgba[2]*255))
                pdf.set_fill_color(*hex_color)
                pdf.cell(max_bar_width, 6, '', border=1, fill=True)
                pdf.set_xy(pdf.get_x() - max_bar_width, pdf.get_y())
                pdf.set_text_color(0,0,0)
                pdf.cell(max_bar_width, 6, f"{int(val):,}", border=1, align='C')
            else:
                pdf.set_fill_color(255,255,255)
                pdf.cell(max_bar_width, 6, str(val), border=1, align='C')
        pdf.ln()

def create_pdf(report_df, fig, fig2, billing_type_df, daily_sales_all, fig4, py_name_df, fig_ka_trend):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, f"Sales & Targets Report - {datetime.now().strftime('%Y-%m-%d')}", 0, 1, 'C')
    add_colored_table(pdf, report_df, "Sales & Targets Table", cmap='YlOrRd')
    add_colored_table(pdf, billing_type_df, "Sales by Billing Type per Salesman", cmap='Blues')
    add_colored_table(pdf, py_name_df, "Sales by PY Name 1", cmap='Greens')
    add_fig_to_pdf(pdf, fig, "Sales and Targets by Salesman - KA and Talabat")
    add_fig_to_pdf(pdf, fig2, "KA & Talabat Targets vs Sales (Values with Gap)")
    add_fig_to_pdf(pdf, fig4, "Daily Sales Trend - All Salesmen")
    add_fig_to_pdf(pdf, fig_ka_trend, "Daily KA Sales Trend")
    return pdf.output(dest='S').encode('latin1')

uploaded_file = st.file_uploader("📂 Upload your Excel file", type=["xlsx"])

if uploaded_file:
    sales_df, target_df = load_excel(uploaded_file)

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
        "KA Target Total": targets_series,
        "Remaining to KA Target": gap_ka,
        "Talabat Sales": talabat_sales_by_salesman,
        "Talabat Target Total": talabat_targets_series,
        "Remaining to Talabat Target": gap_talabat
    }).sort_values("KA Total Sales", ascending=False)

    st.subheader("📋 Sales & Targets Table")
    st.dataframe(report_df.style.format({
        "KA Total Sales":"{:,.0f}",
        "KA Target Total":"{:,.0f}",
        "Remaining to KA Target":"{:,.0f}",
        "Talabat Sales":"{:,.0f}",
        "Talabat Target Total":"{:,.0f}",
        "Remaining to Talabat Target":"{:,.0f}"
    }).background_gradient(subset=report_df.select_dtypes(np.number).columns, cmap="YlOrRd"), use_container_width=True)

    billing_type_by_salesman = sales_df.groupby(["Driver Name EN", "Billing Type"])["Net Value"].sum().unstack(fill_value=0)
    billing_type_by_salesman["Total"] = billing_type_by_salesman.sum(axis=1)
    st.subheader("🧾 Sales by Billing Type per Salesman")
    st.dataframe(billing_type_by_salesman.style.format("{:,.0f}").background_gradient(subset=billing_type_by_salesman.select_dtypes(np.number).columns, cmap="Blues"), use_container_width=True)

    py_name_df = sales_df.groupby("PY Name 1")["Net Value"].sum().reset_index().sort_values("Net Value", ascending=False)
    st.subheader("🏷 Sales by PY Name 1")
    st.dataframe(py_name_df.style.format({"Net Value":"{:,.0f}"}).background_gradient(subset=["Net Value"], cmap="Greens"), use_container_width=True)

    # Bar chart with formatted 3-digit values
    fig, ax = plt.subplots(figsize=(12,8))
    bar_width = 0.25
    y_pos = np.arange(len(all_salesmen))
    pos_total = y_pos - bar_width*1.5
    pos_talabat = y_pos + bar_width*0.5

    ax.barh(pos_total, total_by_salesman, height=bar_width, color='skyblue', label='KA Sales')
    ax.barh(pos_total, gap_ka, left=total_by_salesman, height=bar_width, color='lightgray', label='KA Gap')
    ax.barh(pos_talabat, talabat_sales_by_salesman, height=bar_width, color='orange', label='Talabat Sales')
    ax.barh(pos_talabat, gap_talabat, left=talabat_sales_by_salesman, height=bar_width, color='lightgreen', label='Talabat Gap')

    for i, (ka, ka_gap, tal, tal_gap) in enumerate(zip(total_by_salesman, gap_ka, talabat_sales_by_salesman, gap_talabat)):
        ax.text(ka/2, pos_total[i], format_k(ka), va='center', ha='center', color='black', fontweight='bold')
        ax.text(ka+ka_gap/2, pos_total[i], format_k(ka_gap), va='center', ha='center', color='black')
        ax.text(tal/2, pos_talabat[i], format_k(tal), va='center', ha='center', color='black', fontweight='bold')
        ax.text(tal+tal_gap/2, pos_talabat[i], format_k(tal_gap), va='center', ha='center', color='black')

    ax.set_yticks(y_pos)
    ax.set_yticklabels(all_salesmen)
    ax.invert_yaxis()
    ax.set_title("Sales and Targets by Salesman")
    ax.legend()
    plt.tight_layout()
    st.pyplot(fig)

    # Pie chart with values
    fig2, ax2 = plt.subplots(figsize=(8,4))
    ax2.pie([total_by_salesman.sum(), gap_ka.sum()], 
            labels=[f"KA Sales ({total_by_salesman.sum():,.0f})", f"KA Gap ({gap_ka.sum():,.0f})"],
            colors=['skyblue','lightgray'], startangle=90)
    ax2.pie([talabat_sales_by_salesman.sum(), gap_talabat.sum()], 
            labels=[f"Talabat Sales ({talabat_sales_by_salesman.sum():,.0f})", f"Talabat Gap ({gap_talabat.sum():,.0f})"],
            colors=['orange','lightgreen'], startangle=90, radius=0.7)
    ax2.set_title("KA & Talabat Targets vs Sales (Values with Gap)")
    st.pyplot(fig2)

    # Daily trend all salesmen
    st.subheader("📅 Daily Sales Trend - All Salesmen")
    daily_sales_all = sales_df.groupby(["Billing Date","Driver Name EN"])["Net Value"].sum().reset_index()
    fig4, ax4 = plt.subplots(figsize=(12,6))
    for salesman in daily_sales_all["Driver Name EN"].unique():
        sub_df = daily_sales_all[daily_sales_all["Driver Name EN"] == salesman]
        ax4.plot(sub_df["Billing Date"], sub_df["Net Value"], marker='o', label=salesman)
    ax4.set_title("Daily Sales Trend - All Salesmen")
    plt.xticks(rotation=45)
    ax4.legend()
    st.pyplot(fig4)

    # Daily KA trend
    st.subheader("📈 Daily KA Sales Trend")
    daily_ka_sales = sales_df.groupby("Billing Date")["Net Value"].sum().reset_index()
    fig_ka_trend, ax_ka = plt.subplots(figsize=(12,6))
    ax_ka.plot(daily_ka_sales["Billing Date"], daily_ka_sales["Net Value"], marker='o', color='skyblue')
    ax_ka.set_title("Daily KA Sales Trend")
    plt.xticks(rotation=45)
    st.pyplot(fig_ka_trend)

    # Download PDF
    if st.button("📄 Download Report as PDF"):
        pdf_data = create_pdf(report_df, fig, fig2, billing_type_by_salesman, daily_sales_all, fig4, py_name_df, fig_ka_trend)
        st.download_button(label="💾 Save PDF", data=pdf_data, file_name="sales_report.pdf", mime="application/pdf")

else:
    st.info("Please upload your Excel file with sheets 'sales data' and 'Target'.")
