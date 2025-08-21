import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from pptx import Presentation
from pptx.util import Inches
import io
from datetime import datetime, timedelta
from sklearn.linear_model import LinearRegression

# --- Page Config ---
st.set_page_config(page_title="📊 Haneef Sales Dashboard", layout="wide", page_icon="📈")

# --- Cache Data Loading ---
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
        for j, col in enumerate(df.columns):
            table.cell(0, j).text = str(col)
            table.cell(0, j).text_frame.paragraphs[0].font.bold = True
        for i, row in enumerate(df.itertuples(index=False), start=1):
            for j, val in enumerate(row):
                table.cell(i, j).text = f"{val:,}" if isinstance(val, (int,float,np.integer)) else str(val)

    def add_chart_slide(fig, title):
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = title
        img_stream = io.BytesIO()
        try:
            fig.write_image(img_stream, format='png')
            img_stream.seek(0)
            slide.shapes.add_picture(img_stream, Inches(0.5), Inches(1.5), width=Inches(8))
        except Exception:
            slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(8), Inches(4)).text_frame.text = "Chart cannot be embedded. Install kaleido."

    add_table_slide(report_df.reset_index(), "Sales & Targets Summary")
    add_table_slide(billing_df.reset_index(), "Sales by Billing Type per Salesman")
    add_table_slide(py_df.reset_index(), "Sales by PY Name 1")

    for key, fig in figs_dict.items():
        add_chart_slide(fig, key)

    pptx_stream = io.BytesIO()
    prs.save(pptx_stream)
    pptx_stream.seek(0)
    return pptx_stream

# --- Sidebar Menu ---
st.sidebar.title("Menu")
menu = ["Home", "Sales Tracking"]
choice = st.sidebar.selectbox("Navigate", menu)

# --- Home Page ---
if choice == "Home":
    st.title("🏠 Welcome to Haneef Sales Dashboard")
    st.markdown("""
        **Features:**
        - View sales & targets by salesman, PY Name, and SP Name
        - Interactive charts for trends & gaps
        - Download reports in PPTX & Excel
        Use the sidebar to navigate to Sales Tracking.
    """)

# --- Sales Tracking Page ---
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
            py_filter = st.sidebar.multiselect("Select PY Name", options=sales_df['PY Name 1'].unique(), default=sales_df['PY Name 1'].unique())
            sp_filter = st.sidebar.multiselect("Select SP Name1", options=sales_df['SP Name1'].unique(), default=sales_df['SP Name1'].unique())

            # Quick Date Presets
            preset = st.sidebar.radio("Quick Date Presets", ["Custom Range", "Last 7 Days", "This Month", "YTD"])
            today = pd.Timestamp.today()
            if preset == "Last 7 Days":
                date_range = [today - pd.Timedelta(days=7), today]
            elif preset == "This Month":
                date_range = [today.replace(day=1), today]
            elif preset == "YTD":
                date_range = [today.replace(month=1, day=1), today]
            else:
                date_range = st.sidebar.date_input("Select Date Range", [sales_df['Billing Date'].min(), sales_df['Billing Date'].max()])

            top_n = st.sidebar.slider("Show Top N Salesmen", min_value=1, max_value=len(sales_df['Driver Name EN'].unique()), value=5)

            # --- Filter Data ---
            df_filtered = sales_df[
                (sales_df['Driver Name EN'].isin(salesmen)) &
                (sales_df['Billing Type'].isin(billing_types)) &
                (sales_df['Billing Date'] >= pd.to_datetime(date_range[0])) &
                (sales_df['Billing Date'] <= pd.to_datetime(date_range[1])) &
                (sales_df['PY Name 1'].isin(py_filter)) &
                (sales_df['SP Name1'].isin(sp_filter))
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
            top_salesmen = total_sales.sort_values(ascending=False).head(top_n).index
            total_sales = total_sales.loc[top_salesmen]
            talabat_sales = talabat_sales.loc[top_salesmen]
            ka_gap = ka_gap.loc[top_salesmen]
            talabat_gap = talabat_gap.loc[top_salesmen]

            # --- Tabs ---
            tabs = st.tabs(["KPIs", "Summary Tables", "Charts", "Download Reports", "Advanced Insights"])

            # --- KPIs ---
            with tabs[0]:
                st.subheader("🏆 Key Metrics")
                col1, col2, col3, col4 = st.columns(4)
                col1.metric("Total KA Sales", f"{total_sales.sum():,}", f"{((total_sales.sum()/ka_targets.sum())*100):.2f}%")
                col2.metric("Total Talabat Sales", f"{talabat_sales.sum():,}", f"{((talabat_sales.sum()/talabat_targets.sum())*100):.2f}%")
                col3.metric("Total KA Gap", f"{ka_gap.sum():,}", f"{(ka_gap.sum()/ka_targets.sum()*100):.2f}%")
                col4.metric("Total Talabat Gap", f"{talabat_gap.sum():,}", f"{(talabat_gap.sum()/talabat_targets.sum()*100):.2f}%")
                col5, col6 = st.columns(2)
                col5.metric("Top KA Salesman", total_sales.idxmax(), f"{total_sales.max():,}")
                col6.metric("Top Talabat Salesman", talabat_sales.idxmax(), f"{talabat_sales.max():,}")

            # --- Summary Tables ---
            with tabs[1]:
                report_df = pd.DataFrame({
                    "KA Target": ka_targets.loc[top_salesmen],
                    "KA Sales": total_sales,
                    "KA Remaining": ka_gap,
                    "KA % Achieved": ((total_sales / ka_targets.loc[top_salesmen])*100).round(2),
                    "Talabat Target": talabat_targets.loc[top_salesmen],
                    "Talabat Sales": talabat_sales,
                    "Talabat Remaining": talabat_gap,
                    "Talabat % Achieved": ((talabat_sales / talabat_targets.loc[top_salesmen])*100).round(2)
                })
                st.subheader("Sales & Targets Summary")
                st.dataframe(
                    report_df.style.background_gradient(subset=["KA % Achieved","Talabat % Achieved"], cmap="Greens")
                             .highlight_max(subset=["KA % Achieved","Talabat % Achieved"], color="gold")
                             .format("{:,.0f}"), use_container_width=True
                )

                billing_df = df_filtered.groupby(["Driver Name EN","Billing Type"])["Net Value"].sum().unstack(fill_value=0)
                billing_df["Total"] = billing_df.sum(axis=1)

                required_cols = ["ZFR", "YKF2", "YKRE", "YKS1", "YKS2", "ZCAN", "ZRE"]
                for col in required_cols:
                    if col not in billing_df.columns:
                        billing_df[col] = 0

                billing_df["Return"] = billing_df["YKRE"] + billing_df["ZRE"]
                billing_df["Return %"] = (billing_df["Return"] / billing_df["Total"] * 100).round(2)
                billing_df.rename(columns={"YKF2":"HHT", "ZFR":"PRESALES"}, inplace=True)
                billing_df = billing_df[["Total", "PRESALES", "HHT", "YKRE", "YKS1", "YKS2", "ZCAN", "ZRE", "Return", "Return %"]]

                total_row = pd.DataFrame([{
                    "Total": billing_df["Total"].sum(),
                    "PRESALES": 0,
                    "HHT": 0,
                    "YKRE": 0,
                    "YKS1": 0,
                    "YKS2": 0,
                    "ZCAN": 0,
                    "ZRE": 0,
                    "Return": billing_df["Return"].sum(),
                    "Return %": (billing_df["Return"].sum() / billing_df["Total"].sum() * 100).round(2)
                }], index=["Total"])
                billing_df = pd.concat([billing_df, total_row])

                def highlight_total_row(row):
                    return ['background-color: #add8e6; color: #00008B; font-weight: bold' if row.name == "Total" else '' for _ in row]

                st.subheader("Sales by Billing Type per Salesman")
                st.dataframe(
                    billing_df.style.background_gradient(cmap="Blues", subset=billing_df.columns[:-2])
                              .format({
                                  "Total": "{:,.0f}",
                                  "PRESALES": "{:,.0f}",
                                  "HHT": "{:,.0f}",
                                  "YKRE": "{:,.0f}",
                                  "YKS1": "{:,.0f}",
                                  "YKS2": "{:,.0f}",
                                  "ZCAN": "{:,.0f}",
                                  "ZRE": "{:,.0f}",
                                  "Return": "{:,.0f}",
                                  "Return %": "{:.2f}%"
                              })
                              .apply(highlight_total_row, axis=1),
                    use_container_width=True
                )

                py_df = df_filtered.groupby("PY Name 1")["Net Value"].sum().sort_values(ascending=False)
                st.subheader("Sales by PY Name 1")
                st.dataframe(py_df.to_frame().style.background_gradient(cmap="Greens").format("{:,.0f}"), use_container_width=True)

            # --- Charts ---
            with tabs[2]:
                figs = {}
                fig_sales = go.Figure()
                fig_sales.add_trace(go.Bar(y=total_sales.index, x=total_sales, name="KA Sales", orientation='h',
                                           marker_color='skyblue', text=[f"{v:,}" for v in total_sales], textposition='outside'))
                fig_sales.add_trace(go.Bar(y=total_sales.index, x=ka_gap, name="KA Gap", orientation='h',
                                           marker_color='lightgray', text=[f"{v:,}" for v in ka_gap], textposition='outside'))
                fig_sales.add_trace(go.Bar(y=talabat_sales.index, x=talabat_sales, name="Talabat Sales", orientation='h',
                                           marker_color='orange', text=[f"{v:,}" for v in talabat_sales], textposition='outside'))
                fig_sales.add_trace(go.Bar(y=talabat_sales.index, x=talabat_gap, name="Talabat Gap", orientation='h',
                                           marker_color='lightgreen', text=[f"{v:,}" for v in talabat_gap], textposition='outside'))
                fig_sales.update_layout(barmode='stack', title="Sales & Targets by Salesman", yaxis=dict(autorange="reversed"), xaxis_title="Value")
                figs["Sales & Targets by Salesman"] = fig_sales
                st.plotly_chart(fig_sales, use_container_width=True)

                fig_pie = px.pie(names=["KA Sales","Talabat Sales"], values=[total_sales.sum(), talabat_sales.sum()],
                                 color_discrete_map={"KA Sales":"skyblue","Talabat Sales":"orange"},
                                 title=f"Total Sales: KA {total_sales.sum():,} | Talabat {talabat_sales.sum():,}")
                fig_pie.update_traces(textinfo='percent+label+value', hovertemplate='%{label}: %{value:,} (%{percent})', textfont_size=14)
                figs["Combined Pie Chart"] = fig_pie
                st.plotly_chart(fig_pie, use_container_width=True)

                daily_df = df_filtered.groupby(["Billing Date", "Driver Name EN"])["Net Value"].sum().reset_index()
                fig_daily = go.Figure()
                for s in total_sales.index:
                    sub = daily_df[daily_df["Driver Name EN"]==s]
                    if not sub.empty:
                        fig_daily.add_trace(go.Scatter(x=sub["Billing Date"], y=sub["Net Value"], mode='lines+markers', name=s))
                        sub["7d_avg"] = sub["Net Value"].rolling(7, 1).mean()
                        fig_daily.add_trace(go.Scatter(x=sub["Billing Date"], y=sub["7d_avg"], mode='lines', line=dict(dash='dot'), name=f"{s} 7-Day Avg"))
                fig_daily.update_layout(title="Daily Sales Trend - Top Salesmen")
                figs["Daily Sales Trend"] = fig_daily
                st.plotly_chart(fig_daily, use_container_width=True)

            # --- Download Reports ---
            with tabs[3]:
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

            # --- Advanced Insights ---
            with tabs[4]:
                st.subheader("🚀 Advanced Insights & Analytics")
                
                st.markdown("**KA Gap Progress**")
                for s in total_sales.index:
                    st.progress(int((total_sales[s]/ka_targets[s])*100))
                    st.write(f"{s}: {total_sales[s]:,}/{ka_targets[s]:,}")
                st.markdown("**Talabat Gap Progress**")
                for s in talabat_sales.index:
                    st.progress(int((talabat_sales[s]/talabat_targets[s])*100))
                    st.write(f"{s}: {talabat_sales[s]:,}/{talabat_targets[s]:,}")

                st.markdown("**Top 3 Performers per Week**")
                weekly_df = df_filtered.groupby([pd.Grouper(key="Billing Date", freq="W"), "Driver Name EN"])["Net Value"].sum().reset_index()
                top_weekly = weekly_df.groupby("Billing Date").apply(lambda x: x.nlargest(3, "Net Value"))
                st.dataframe(top_weekly)

                st.markdown("**Forecast Next Month's KA & Talabat Sales**")
                forecast_data = []
                for s in total_sales.index:
                    sub = daily_df[daily_df["Driver Name EN"]==s]
                    if len(sub) >= 2:
                        lr = LinearRegression()
                        X = np.arange(len(sub)).reshape(-1,1)
                        y = sub["Net Value"].values
                        lr.fit(X, y)
                        pred = lr.predict(np.array([[len(sub)+30]]))[0]
                        forecast_data.append({"Driver Name EN": s, "Forecast": int(pred)})
                forecast_df = pd.DataFrame(forecast_data)
                st.dataframe(forecast_df)

                fig_forecast = go.Figure()
                for s in total_sales.index:
                    sub = daily_df[daily_df["Driver Name EN"]==s]
                    if not sub.empty:
                        fig_forecast.add_trace(go.Scatter(x=sub["Billing Date"], y=sub["Net Value"], mode='lines+markers', name=f"{s} Actual"))
                        fig_forecast.add_trace(go.Scatter(x=[sub["Billing Date"].max()+pd.Timedelta(days=30)],
                                                          y=[forecast_df.loc[forecast_df["Driver Name EN"]==s, "Forecast"].values[0]], 
                                                          mode='markers', marker=dict(symbol='diamond', size=12), name=f"{s} Forecast"))
                fig_forecast.update_layout(title="KA & Talabat Sales Forecast Next Month", xaxis_title="Date", yaxis_title="Sales")
                st.plotly_chart(fig_forecast, use_container_width=True)

        else:
            st.info("Please upload an Excel file with sheets 'sales data' and 'Target'.")
