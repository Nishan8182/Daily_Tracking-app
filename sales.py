import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from fpdf import FPDF
import tempfile
import os

st.title("Sales and Targets by Salesman - KA and Talabat")

def add_fig_to_pdf(pdf, fig, title):
    pdf.add_page()
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 10, title, 0, 1, 'C')

    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmpfile:
        fig.savefig(tmpfile.name, bbox_inches='tight')
        tmpfile_path = tmpfile.name

    pdf.image(tmpfile_path, x=10, y=25, w=pdf.w - 20)
    os.remove(tmpfile_path)

def create_pdf(report_df, fig, fig2, fig3, billing_type_df, daily_sales_grouped, selected_salesman, fig4):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)

    # Sales & Targets Table as text
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, "Sales & Targets Report", 0, 1, 'C')
    pdf.set_font("Arial", size=10)
    table_str = report_df.to_string(float_format="{:,.0f}".format)
    for line in table_str.split('\n'):
        pdf.cell(0, 6, line, 0, 1)

    # Billing Type table as text
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, "Sales by Billing Type per Salesman", 0, 1, 'C')
    pdf.set_font("Arial", size=10)
    billing_type_str = billing_type_df.round(0).astype(int).to_string()
    for line in billing_type_str.split('\n'):
        pdf.cell(0, 6, line, 0, 1)

    # Add charts
    add_fig_to_pdf(pdf, fig, "Sales and Targets by Salesman - KA and Talabat")
    add_fig_to_pdf(pdf, fig2, "Summary: KA Target vs Total Sales")
    add_fig_to_pdf(pdf, fig3, "Summary: Talabat Target vs Total Sales")

    # Daily sales data table
    pdf.add_page()
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 10, f"Daily Sales for {selected_salesman}", 0, 1, 'C')
    pdf.set_font("Arial", size=10)
    daily_sales_str = daily_sales_grouped.to_string(index=False, float_format="{:,.0f}".format)
    for line in daily_sales_str.split('\n'):
        pdf.cell(0, 6, line, 0, 1)

    # Daily sales chart
    add_fig_to_pdf(pdf, fig4, f"Daily Sales Trend - {selected_salesman}")

    # Output PDF to bytes buffer
    pdf_output = pdf.output(dest='S').encode('latin1')
    return pdf_output


uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    sales_df = pd.read_excel(uploaded_file, sheet_name="sales data")
    target_df = pd.read_excel(uploaded_file, sheet_name="Target")

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

    gap_ka = targets_series - total_by_salesman
    gap_ka[gap_ka < 0] = 0

    gap_talabat = talabat_targets_series - talabat_sales_by_salesman
    gap_talabat[gap_talabat < 0] = 0

    report_df = pd.DataFrame({
        "KA Total Sales": total_by_salesman,
        "Remaining to KA Target": gap_ka,
        "Talabat Sales": talabat_sales_by_salesman,
        "Remaining to Talabat Target": gap_talabat
    })

    st.subheader("Sales & Targets Table")
    st.dataframe(report_df.style.format("{:,.0f}"), use_container_width=True)

    # Billing Type summary by salesman with Total
    billing_type_by_salesman = sales_df.groupby(["Driver Name EN", "Billing Type"])["Net Value"].sum().unstack(fill_value=0)
    billing_type_by_salesman["Total"] = billing_type_by_salesman.sum(axis=1)

    st.subheader("Sales by Billing Type per Salesman")
    st.dataframe(billing_type_by_salesman.style.format("{:,.0f}"), use_container_width=True)

    bar_width = 0.25
    y_pos = np.arange(len(all_salesmen))

    fig, ax = plt.subplots(figsize=(12, 8))
    pos_total = y_pos - bar_width*1.5
    pos_talabat = y_pos + bar_width*0.5

    ax.barh(pos_total, total_by_salesman, height=bar_width, color='skyblue', label='KA Total Sales')
    ax.barh(pos_total, gap_ka, left=total_by_salesman, height=bar_width, color='lightgray', label='Remaining to KA Target')

    ax.barh(pos_talabat, talabat_sales_by_salesman, height=bar_width, color='orange', label='Talabat Sales')
    ax.barh(pos_talabat, gap_talabat, left=talabat_sales_by_salesman, height=bar_width, color='lightgreen', label='Remaining to Talabat Target')

    ax.set_yticks(y_pos)
    ax.set_yticklabels(all_salesmen)
    ax.invert_yaxis()
    ax.set_xlabel("Net Value")
    ax.set_title("Sales and Targets by Salesman - KA and Talabat")

    for i in range(len(all_salesmen)):
        ax.text(total_by_salesman[i]/2, pos_total[i], f'{total_by_salesman[i]:,.0f}', va='center', ha='center', color='black')
        if gap_ka[i] > 0:
            ax.text(total_by_salesman[i] + gap_ka[i]/2, pos_total[i], f'{gap_ka[i]:,.0f}', va='center', ha='center', color='black')
        if talabat_sales_by_salesman[i] > 0:
            ax.text(talabat_sales_by_salesman[i]/2, pos_talabat[i], f'{talabat_sales_by_salesman[i]:,.0f}', va='center', ha='center', color='black')
        if gap_talabat[i] > 0:
            ax.text(talabat_sales_by_salesman[i] + gap_talabat[i]/2, pos_talabat[i], f'{gap_talabat[i]:,.0f}', va='center', ha='center', color='black')

    ax.legend(loc='lower right')
    plt.tight_layout()
    st.pyplot(fig)

    st.subheader("Summary: KA Target vs Total Sales")
    sum_ka_target = targets_series.sum()
    sum_total_sales = total_by_salesman.sum()
    sum_gap = max(sum_ka_target - sum_total_sales, 0)

    fig2, ax2 = plt.subplots(figsize=(6, 3))
    bars = ax2.bar(['Total Sales', 'Remaining Gap'], [sum_total_sales, sum_gap], color=['skyblue', 'lightgray'])
    ax2.set_title("Total KA Target vs Total Sales")
    ax2.set_ylabel("Net Value")
    for bar in bars:
        height = bar.get_height()
        ax2.text(bar.get_x() + bar.get_width()/2, height/2, f'{int(height):,}', ha='center', va='center', color='black')
    st.pyplot(fig2)

    st.subheader("Summary: Talabat Target vs Total Sales")
    sum_talabat_target = talabat_targets_series.sum()
    sum_talabat_sales = talabat_sales_by_salesman.sum()
    sum_talabat_gap = max(sum_talabat_target - sum_talabat_sales, 0)

    fig3, ax3 = plt.subplots(figsize=(6, 3))
    bars2 = ax3.bar(['Total Sales', 'Remaining Gap'], [sum_talabat_sales, sum_talabat_gap], color=['orange', 'lightgreen'])
    ax3.set_title("Total Talabat Target vs Total Sales")
    ax3.set_ylabel("Net Value")
    for bar in bars2:
        height = bar.get_height()
        ax3.text(bar.get_x() + bar.get_width()/2, height/2, f'{int(height):,}', ha='center', va='center', color='black')
    st.pyplot(fig3)

    st.subheader("Daily Sales by Salesman")
    salesmen_list = sales_df["Driver Name EN"].unique()
    selected_salesman = st.selectbox("Select Salesman to see daily sales", options=salesmen_list)

    if selected_salesman:
        daily_sales = sales_df[sales_df["Driver Name EN"] == selected_salesman]
        daily_sales_grouped = daily_sales.groupby("Billing Date")["Net Value"].sum().reset_index()

        st.write(f"Daily Sales for {selected_salesman}")
        st.dataframe(daily_sales_grouped.style.format({"Net Value": "{:,.0f}"}), use_container_width=True)

        fig4, ax4 = plt.subplots(figsize=(10, 4))
        ax4.plot(daily_sales_grouped["Billing Date"], daily_sales_grouped["Net Value"], marker='o', linestyle='-')
        ax4.set_title(f"Daily Sales Trend - {selected_salesman}")
        ax4.set_xlabel("Date")
        ax4.set_ylabel("Net Value")
        plt.xticks(rotation=45)
        plt.tight_layout()
        st.pyplot(fig4)

        if st.button("Download Report as PDF"):
            pdf_data = create_pdf(report_df, fig, fig2, fig3, billing_type_by_salesman, daily_sales_grouped, selected_salesman, fig4)
            st.download_button(
                label="Download PDF",
                data=pdf_data,
                file_name="sales_report.pdf",
                mime="application/pdf"
            )

else:
    st.info("Please upload your Excel file with sheets 'sales data' and 'Target'.")
