import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from pptx import Presentation
from pptx.util import Inches
import io

# --- Page Config ---
st.set_page_config(page_title="📊 Haneef Sales Dashboard", layout="wide")

# --- Cache data loading ---
@st.cache_data
def load_data(file):
    try:
        sales_df = pd.read_excel(file, sheet_name="sales data")
        target_df = pd.read_excel(file, sheet_name="Target")
        sales_df['Billing Date'] = pd.to_datetime(sales_df['Billing Date'])
        return sales_df, target_df
    except Exception as e:
        st.error(f"Error loading Excel file: {e}")
        return None, None

# --- PPTX Export ---
def create_pptx(report_df, billing_df, py_df, figs_dict):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "Sales & Targets Report"
    slide.placeholders[1].text = "Generated from Sales Data"

    def add_table_slide(df, title):
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
                table.cell(i, j).text = f"{val:,}" if isinstance(val, (int,float,np.integer)) else str(val)

    def add_chart_slide(fig, title):
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = title
        img_stream = io.BytesIO()
        try:
            # Export Plotly figure to PNG
            fig.write_image(img_stream, format='png')  # Requires kaleido
            img_stream.seek(0)
            slide.shapes.add_picture(img_stream, Inches(0.5), Inches(1.5), width=Inches(8))
        except Exception:
            slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(8), Inches(4)).text_frame.text = \
                "Chart cannot be embedded. Install kaleido."

    # --- Add tables ---
    add_table_slide(report_df.reset_index(), "Sales & Targets Summary")
    add_table_slide(billing_df.reset_index(), "Sales by Billing Type per Salesman")
    add_table_slide(py_df.reset_index(), "Sales by PY Name 1")

    # --- Add charts ---
    for key, fig in figs_dict.items():
        add_chart_slide(fig, key)

    pptx_stream = io.BytesIO()
    prs.save(pptx_stream)
    pptx_stream.seek(0)
    return pptx_stream

# --- Sidebar ---
st.sidebar.title("Menu")
menu = ["Home", "Sales Tracking"]
choice = st.sidebar.selectbox("Navigate", menu)

# --- Home ---
if choice == "Home":
    st.title("🏠 Welcome to Haneef Sales Dashboard")
    st.markdown("Use the sidebar to navigate to Sales Tracking and view reports.")

# --- Sales Tracking ---
elif choice == "Sales Tracking":
    st.title("📊 Sales Tracking Dashboard")
    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

    if uploaded_file:
        sales_df, target_df = load_data(uploaded_file)
        if sales_df is not None and target_df is not None:

            # --- Filters ---
            st.sidebar.subheader("Filters")
            salesmen = st.sidebar.multiselect("Select Salesmen", options=sales_df['Driver Name EN'].unique(), default=sales_df['Driver Name EN'].unique())
            billing_types = st.sidebar.multiselect("Select Billing Types", options=sales_df['Billing Type'].unique(), default=sales_df['Billing Type'].unique())
            date_range = st.sidebar.date_input("Select Date Range", [sales_df['Billing Date'].min(), sales_df['Billing Date'].max()])

            df_filtered = sales_df[
                (sales_df['Driver Name EN'].isin(salesmen)) &
                (sales_df['Billing Type'].isin(billing_types)) &
                (sales_df['Billing Date'] >= pd.to_datetime(date_range[0])) &
                (sales_df['Billing Date'] <= pd.to_datetime(date_range[1]))
            ]

            # --- Calculations ---
            total_sales = df_filtered.groupby("Driver Name EN")["Net Value"].sum()
            talabat_df = df_filtered[df_filtered["PY Name 1"] == "STORES SERVICES KUWAIT CO."]
            talabat_sales = talabat_df.groupby("Driver Name EN")["Net Value"].sum()

            ka_targets = target_df.set_index("Driver Name EN")["KA Target"]
            talabat_targets = target_df.set_index("Driver Name EN")["Talabat Target"]

            all_salesmen = total_sales.index.union(talabat_sales.index).union(ka_targets.index).union(talabat_targets.index)
            total_sales = total_sales.reindex(all_salesmen, fill_value=0).astype(int)
            talabat_sales = talabat_sales.reindex(all_salesmen, fill_value=0).astype(int)
            ka_targets = ka_targets.reindex(all_salesmen, fill_value=0).astype(int)
            talabat_targets = talabat_targets.reindex(all_salesmen, fill_value=0).astype(int)

            ka_gap = (ka_targets - total_sales).clip(lower=0).astype(int)
            talabat_gap = (talabat_targets - talabat_sales).clip(lower=0).astype(int)

            # KPI Table
            report_df = pd.DataFrame({
                "KA Total Sales": total_sales,
                "KA Remaining": ka_gap,
                "KA % Achieved": ((total_sales / ka_targets)*100).round(2),
                "Talabat Sales": talabat_sales,
                "Talabat Remaining": talabat_gap,
                "Talabat % Achieved": ((talabat_sales / talabat_targets)*100).round(2),
                "KA Target": ka_targets,
                "Talabat Target": talabat_targets
            })

            st.subheader("Sales & Targets Summary")
            st.dataframe(
                report_df.style
                    .background_gradient(subset=["KA % Achieved","Talabat % Achieved"], cmap="Greens")
                    .format("{:,.0f}"),
                use_container_width=True
            )

            # Billing Type Table
            billing_df = df_filtered.groupby(["Driver Name EN","Billing Type"])["Net Value"].sum().unstack(fill_value=0)
            billing_df["Total"] = billing_df.sum(axis=1)
            st.subheader("Sales by Billing Type per Salesman")
            st.dataframe(billing_df.style.background_gradient(cmap="Blues").format("{:,.0f}"), use_container_width=True)

            # PY Name Table (sorted top values)
            py_df = df_filtered.groupby("PY Name 1")["Net Value"].sum().sort_values(ascending=False)
            st.subheader("Sales by PY Name 1")
            st.dataframe(py_df.to_frame().style.background_gradient(cmap="Greens").format("{:,.0f}"), use_container_width=True)

            # --- Charts ---
            figs = {}

            # Horizontal stacked bar: Sales & Targets by Salesman
            fig_sales = go.Figure()
            fig_sales.add_trace(go.Bar(
                y=all_salesmen,
                x=total_sales.loc[all_salesmen],
                name="KA Sales",
                orientation='h',
                marker_color='skyblue',
                text=[f"{v:,}" for v in total_sales.loc[all_salesmen]],
                textposition='outside'
            ))
            fig_sales.add_trace(go.Bar(
                y=all_salesmen,
                x=ka_gap.loc[all_salesmen],
                name="KA Gap",
                orientation='h',
                marker_color='lightgray',
                text=[f"{v:,}" for v in ka_gap.loc[all_salesmen]],
                textposition='outside'
            ))
            fig_sales.add_trace(go.Bar(
                y=all_salesmen,
                x=talabat_sales.loc[all_salesmen],
                name="Talabat Sales",
                orientation='h',
                marker_color='orange',
                text=[f"{v:,}" for v in talabat_sales.loc[all_salesmen]],
                textposition='outside'
            ))
            fig_sales.add_trace(go.Bar(
                y=all_salesmen,
                x=talabat_gap.loc[all_salesmen],
                name="Talabat Gap",
                orientation='h',
                marker_color='lightgreen',
                text=[f"{v:,}" for v in talabat_gap.loc[all_salesmen]],
                textposition='outside'
            ))
            fig_sales.update_layout(
                barmode='stack',
                title="Sales & Targets by Salesman",
                yaxis=dict(autorange="reversed"),
                xaxis_title="Value"
            )
            figs["Sales & Targets by Salesman"] = fig_sales

            # Combined KA & Talabat Pie Chart
            fig_pie = px.pie(
                names=["KA Sales","KA Gap","Talabat Sales","Talabat Gap"],
                values=[total_sales.sum(), ka_gap.sum(), talabat_sales.sum(), talabat_gap.sum()],
                color_discrete_map={"KA Sales":"skyblue","KA Gap":"lightgray","Talabat Sales":"orange","Talabat Gap":"lightgreen"},
                title="KA & Talabat Combined Target vs Sales"
            )
            figs["Combined Pie Chart"] = fig_pie

            # Daily Sales Trend
            daily_df = df_filtered.groupby(["Billing Date", "Driver Name EN"])["Net Value"].sum().reset_index()
            fig_daily = go.Figure()
            for s in all_salesmen:
                sub = daily_df[daily_df["Driver Name EN"]==s]
                if not sub.empty:
                    fig_daily.add_trace(go.Scatter(x=sub["Billing Date"], y=sub["Net Value"], mode='lines+markers', name=s))
            fig_daily.update_layout(title="Daily Sales Trend - All Salesmen")
            figs["Daily Sales Trend"] = fig_daily

            # Daily KA Trend
            daily_ka = df_filtered.groupby("Billing Date")["Net Value"].sum().reset_index()
            fig_daily_ka = px.line(daily_ka, x="Billing Date", y="Net Value", title="Daily KA Sales Trend", markers=True)
            figs["Daily KA Sales Trend"] = fig_daily_ka

            # Show charts
            st.plotly_chart(fig_sales, use_container_width=True)
            st.plotly_chart(fig_pie, use_container_width=True)
            st.plotly_chart(fig_daily, use_container_width=True)
            st.plotly_chart(fig_daily_ka, use_container_width=True)

            # --- Downloads ---
            st.subheader("Download Reports")
            pptx_data = create_pptx(report_df, billing_df, py_df, figs)
            st.download_button("Download PPTX", data=pptx_data, file_name="sales_report.pptx",
                               mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")

            excel_stream = io.BytesIO()
            with pd.ExcelWriter(excel_stream, engine='xlsxwriter') as writer:
                report_df.to_excel(writer, sheet_name="Sales Summary")
                billing_df.to_excel(writer, sheet_name="Billing Type")
                py_df.to_frame().to_excel(writer, sheet_name="PY Name 1")
            excel_stream.seek(0)
            st.download_button("Download Excel", data=excel_stream, file_name="sales_report.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.info("Please upload your Excel file with sheets 'sales data' and 'Target'.")
