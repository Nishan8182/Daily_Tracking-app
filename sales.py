import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from pptx import Presentation
from pptx.util import Inches
import io
from datetime import datetime
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

# --- Positive/Negative Coloring ---
def color_positive_negative(val):
    if val > 0:
        color = 'green'
    elif val < 0:
        color = 'red'
    else:
        color = ''
    return f'color: {color}; font-weight: bold'

# --- Sidebar Menu ---
st.sidebar.title("Menu")
menu = ["Home", "Sales Tracking", "YTD", "Custom Analysis"]  # <— NEW PAGE ADDED
choice = st.sidebar.selectbox("Navigate", menu)

# --- Home Page ---
if choice == "Home":
    st.title("🏠 Welcome to Haneef Sales Dashboard")
    st.markdown("""
        **Features:**
        - View sales & targets by salesman, PY Name, and SP Name
        - Interactive charts for trends & gaps
        - Download reports in PPTX & Excel
        - Compare sales across two custom periods
        - NEW: Custom Analysis (pick any column(s) to compare)
        Use the sidebar to navigate to Sales Tracking, YTD, or Custom Analysis.
    """)

# --- Sales Tracking Page ---
elif choice == "Sales Tracking":
    st.title("📊 Sales Tracking Dashboard")
    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"], key="sales_upload")

    if uploaded_file:
        sales_df, target_df = load_data(uploaded_file)
        if sales_df is not None and target_df is not None:

            # --- Filters ---
            st.sidebar.subheader("Filters")
            salesmen = st.sidebar.multiselect("Select Salesmen", options=sales_df['Driver Name EN'].unique(), default=sales_df['Driver Name EN'].unique())
            billing_types = st.sidebar.multiselect("Select Billing Types", options=sales_df['Billing Type'].unique(), default=sales_df['Billing Type'].unique())
            py_filter = st.sidebar.multiselect("Select PY Name", options=sales_df['PY Name 1'].unique(), default=sales_df['PY Name 1'].unique())
            sp_filter = st.sidebar.multiselect("Select SP Name1", options=sales_df['SP Name1'].unique(), default=sales_df['SP Name1'].unique())

            # --- Quick Date Presets ---
            preset = st.sidebar.radio("Quick Date Presets", ["Custom Range", "Last 7 Days", "This Month", "YTD"])
            today = pd.Timestamp.today()
            if preset == "Last 7 Days":
                date_range = [today - pd.Timedelta(days=7), today]
            elif preset == "This Month":
                month_start = today.replace(day=1)
                month_end = (month_start + pd.offsets.MonthEnd(0))
                date_range = [month_start, month_end]
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

            # --- Days Finish & Per Day KA Target ---
            billing_start = df_filtered['Billing Date'].min()
            billing_end = df_filtered['Billing Date'].max()
            all_days = pd.date_range(billing_start, billing_end)
            days_finish = sum(1 for d in all_days if d.weekday() != 4)  # Exclude Fridays
            total_ka_target = target_df['KA Target'].sum()
            current_month_start = today.replace(day=1)
            current_month_end = (current_month_start + pd.offsets.MonthEnd(0))
            current_month_days = pd.date_range(current_month_start, current_month_end)
            working_days_current_month = sum(1 for d in current_month_days if d.weekday() != 4)
            per_day_ka_target = total_ka_target / working_days_current_month if working_days_current_month > 0 else 0

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
            total_sales_top = total_sales.loc[top_salesmen]
            talabat_sales_top = talabat_sales.loc[top_salesmen]
            ka_gap_top = ka_gap.loc[top_salesmen]
            talabat_gap_top = talabat_gap.loc[top_salesmen]

            # --- Tabs ---
            tabs = st.tabs(["KPIs", "Summary Tables", "Charts", "Download Reports", "Advanced Insights"])

            # --- KPIs ---
            with tabs[0]:
                st.subheader("🏆 Key Metrics")
                r1c1, r1c2, r1c3, r1c4 = st.columns(4)
                r1c1.metric("Total KA Sales", f"{total_sales.sum():,}", f"{((total_sales.sum()/ka_targets.sum())*100):.2f}%")
                r1c2.metric("Total Talabat Sales", f"{talabat_sales.sum():,}", f"{((talabat_sales.sum()/talabat_targets.sum())*100):.2f}%")
                r1c3.metric("Total KA Gap", f"{ka_gap.sum():,}", f"{(ka_gap.sum()/ka_targets.sum()*100):.2f}%")
                r1c4.metric("Total Talabat Gap", f"{talabat_gap.sum():,}", f"{(talabat_gap.sum()/talabat_targets.sum()*100):.2f}%")
                r2c1, r2c2 = st.columns(2)
                r2c1.metric("Top KA Salesman", total_sales_top.idxmax(), f"{total_sales_top.max():,}")
                r2c2.metric("Top Talabat Salesman", talabat_sales_top.idxmax(), f"{talabat_sales_top.max():,}")
                r3c1, r3c2, r3c3, r3c4 = st.columns(4)
                r3c1.metric("Day's Finish", days_finish)
                r3c2.metric("Per Day KA Target", f"{per_day_ka_target:,.0f}")
                current_sales_per_day = total_sales.sum() / days_finish if days_finish > 0 else 0
                remaining_per_day_target = per_day_ka_target - current_sales_per_day
                r3c3.metric("Current Sales Per Day", f"{current_sales_per_day:,.0f}")
                r3c4.metric("Remaining KA per Day", f"{remaining_per_day_target:,.0f}")

            # --- Summary Tables ---
            with tabs[1]:
                st.subheader("Sales & Targets Summary")
                report_df = pd.DataFrame({
                    "KA Target": ka_targets,
                    "KA Sales": total_sales,
                    "KA Remaining": ka_gap,
                    "KA % Achieved": ((total_sales / ka_targets)*100).round(2),
                    "Talabat Target": talabat_targets,
                    "Talabat Sales": talabat_sales,
                    "Talabat Remaining": talabat_gap,
                    "Talabat % Achieved": ((talabat_sales / talabat_targets)*100).round(2)
                })
                st.dataframe(report_df.style.applymap(color_positive_negative, subset=["KA % Achieved","Talabat % Achieved"])
                                             .highlight_max(subset=["KA % Achieved","Talabat % Achieved"], color="gold")
                                             .format("{:,.0f}"), use_container_width=True)

                # Excel Download
                excel_stream1 = io.BytesIO()
                with pd.ExcelWriter(excel_stream1, engine='xlsxwriter') as writer:
                    report_df.to_excel(writer, sheet_name="Sales_Targets_Summary")
                excel_stream1.seek(0)
                st.download_button("Download Excel - Sales & Targets Summary", data=excel_stream1, file_name="Sales_Targets_Summary.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

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
                    "PRESALES": 0, "HHT": 0, "YKRE": 0, "YKS1": 0, "YKS2": 0, "ZCAN": 0, "ZRE": 0,
                    "Return": billing_df["Return"].sum(),
                    "Return %": (billing_df["Return"].sum() / billing_df["Total"].sum() * 100).round(2)
                }], index=["Total"])
                billing_df = pd.concat([billing_df, total_row])
                def highlight_total_row(row):
                    return ['background-color: #add8e6; color: #00008B; font-weight: bold' if row.name == "Total" else '' for _ in row]
                st.subheader("Sales by Billing Type per Salesman")
                st.dataframe(billing_df.style.background_gradient(cmap="Blues", subset=billing_df.columns[:-2])
                              .format({
                                  "Total": "{:,.0f}","PRESALES": "{:,.0f}","HHT": "{:,.0f}",
                                  "YKRE": "{:,.0f}","YKS1": "{:,.0f}","YKS2": "{:,.0f}",
                                  "ZCAN": "{:,.0f}","ZRE": "{:,.0f}","Return": "{:,.0f}","Return %": "{:.2f}%"
                              }).apply(highlight_total_row, axis=1)
                              .applymap(color_positive_negative, subset=["Return %"]), use_container_width=True)
                
                excel_stream2 = io.BytesIO()
                with pd.ExcelWriter(excel_stream2, engine='xlsxwriter') as writer:
                    billing_df.to_excel(writer, sheet_name="Billing_Types")
                excel_stream2.seek(0)
                st.download_button("Download Excel - Billing Types", data=excel_stream2, file_name="Billing_Types.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

                py_df = df_filtered.groupby("PY Name 1")["Net Value"].sum().sort_values(ascending=False)
                st.subheader("Sales by PY Name 1")
                st.dataframe(py_df.to_frame().style.background_gradient(cmap="Greens").format("{:,.0f}"), use_container_width=True)

                excel_stream3 = io.BytesIO()
                with pd.ExcelWriter(excel_stream3, engine='xlsxwriter') as writer:
                    py_df.to_frame().to_excel(writer, sheet_name="Sales_by_PY_Name")
                excel_stream3.seek(0)
                st.download_button("Download Excel - Sales by PY Name", data=excel_stream3, file_name="Sales_by_PY_Name.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            # --- Charts Tab ---
            with tabs[2]:
                st.subheader("Sales Charts")
                figs_dict = {}
                fig_ka = px.bar(report_df, x=report_df.index, y="KA Sales", title="KA Sales by Salesman", text="KA Sales")
                fig_talabat = px.bar(report_df, x=report_df.index, y="Talabat Sales", title="Talabat Sales by Salesman", text="Talabat Sales")
                st.plotly_chart(fig_ka, use_container_width=True)
                st.plotly_chart(fig_talabat, use_container_width=True)
                figs_dict["KA Sales by Salesman"] = fig_ka
                figs_dict["Talabat Sales by Salesman"] = fig_talabat

            # --- Download Reports Tab ---
            with tabs[3]:
                st.subheader("Download PPTX Report")
                pptx_data = create_pptx(report_df, billing_df, py_df, figs_dict)
                st.download_button("Download PPTX Report", data=pptx_data, file_name="sales_report.pptx")

            # --- Advanced Insights Tab ---
            with tabs[4]:
                st.subheader("Sales Trend Forecast")
                df_time = df_filtered.groupby("Billing Date")["Net Value"].sum().reset_index()
                if len(df_time) > 1:
                    model = LinearRegression()
                    model.fit(np.arange(len(df_time)).reshape(-1,1), df_time["Net Value"])
                    df_time["Forecast"] = model.predict(np.arange(len(df_time)).reshape(-1,1))
                    fig_trend = px.line(df_time, x="Billing Date", y=["Net Value", "Forecast"], title="Sales Trend with Forecast")
                    st.plotly_chart(fig_trend, use_container_width=True)
                else:
                    st.info("Not enough data to generate trend forecast.")

# --- YTD Page ---
elif choice == "YTD":
    st.title("📅 Year-to-Date (YTD) Comparison")
    uploaded_file = st.file_uploader("Upload Excel for YTD Comparison", type=["xlsx"], key="ytd_upload")

    if uploaded_file:
        sales_df, _ = load_data(uploaded_file)
        if sales_df is not None:
            dimension = st.selectbox("Compare By", ["Driver Name EN", "PY Name 1", "SP Name1"])

            st.subheader("Select Two Periods to Compare")
            col1, col2 = st.columns(2)
            with col1:
                period1 = st.date_input("Period 1 Start-End", [sales_df['Billing Date'].min(), sales_df['Billing Date'].max()])
            with col2:
                period2 = st.date_input("Period 2 Start-End", [sales_df['Billing Date'].min(), sales_df['Billing Date'].max()])

            period1_label = f"{period1[0].strftime('%d-%b-%Y')} to {period1[1].strftime('%d-%b-%Y')}"
            period2_label = f"{period2[0].strftime('%d-%b-%Y')} to {period2[1].strftime('%d-%b-%Y')}"

            df1 = sales_df[(sales_df['Billing Date'] >= pd.to_datetime(period1[0])) & (sales_df['Billing Date'] <= pd.to_datetime(period1[1]))]
            df2 = sales_df[(sales_df['Billing Date'] >= pd.to_datetime(period2[0])) & (sales_df['Billing Date'] <= pd.to_datetime(period2[1]))]

            agg1 = df1.groupby(dimension)["Net Value"].sum()
            agg2 = df2.groupby(dimension)["Net Value"].sum()
            all_index = agg1.index.union(agg2.index)
            agg1 = agg1.reindex(all_index, fill_value=0)
            agg2 = agg2.reindex(all_index, fill_value=0)

            comparison_df = pd.DataFrame({
                period1_label: agg1,
                period2_label: agg2
            })
            comparison_df["Difference"] = comparison_df[period2_label] - comparison_df[period1_label]
            comparison_df["Comparison %"] = np.where(comparison_df[period1_label] != 0,
                                                     (comparison_df["Difference"] / comparison_df[period1_label] * 100).round(2),
                                                     0)
            comparison_df = comparison_df.sort_values(by=period2_label, ascending=False)

            def highlight_date_columns(row):
                return [
                    'background-color: #d1e7dd; font-weight: bold' if col in [period1_label, period2_label] 
                    else '' for col in row.index
                ]

            st.subheader(f"YTD Comparison by {dimension}")
            st.dataframe(
                comparison_df.style.format({
                    period1_label: "{:,.0f}",
                    period2_label: "{:,.0f}",
                    "Difference": "{:,.0f}",
                    "Comparison %": "{:.2f}%"
                }).applymap(color_positive_negative, subset=["Difference","Comparison %"])
                  .apply(highlight_date_columns, axis=1),
                use_container_width=True
            )

            # Excel download for YTD
            excel_stream_ytd = io.BytesIO()
            with pd.ExcelWriter(excel_stream_ytd, engine='xlsxwriter') as writer:
                comparison_df.to_excel(writer, sheet_name="YTD Comparison")
            excel_stream_ytd.seek(0)
            st.download_button("Download Excel - YTD Comparison", data=excel_stream_ytd, file_name="YTD_Comparison.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            st.subheader("YTD Comparison Chart")
            fig = go.Figure()
            fig.add_trace(go.Bar(x=comparison_df.index, y=comparison_df[period1_label], name=period1_label, marker_color='skyblue'))
            fig.add_trace(go.Bar(x=comparison_df.index, y=comparison_df[period2_label], name=period2_label, marker_color='orange'))
            fig.add_trace(go.Scatter(x=comparison_df.index, y=comparison_df["Comparison %"], name="% Difference", mode='lines+markers', yaxis='y2', line=dict(color='green', width=3)))
            fig.update_layout(
                barmode='group',
                title=f"YTD Comparison by {dimension}",
                xaxis_title=dimension,
                yaxis=dict(title="Net Value"),
                yaxis2=dict(title="% Difference", overlaying='y', side='right', showgrid=False)
            )
            st.plotly_chart(fig, use_container_width=True)

# --- Custom Analysis (NEW PAGE) ---
elif choice == "Custom Analysis":
    st.title("🔍 Custom Analysis (Pick Any Columns)")

    uploaded_file = st.file_uploader("Upload Excel for Custom Analysis", type=["xlsx"], key="custom_upload")

    if uploaded_file:
        sales_df, _ = load_data(uploaded_file)
        if sales_df is not None:
            # Show available columns
            with st.expander("👀 Show available columns in your data"):
                col_info = pd.DataFrame({
                    "Column": sales_df.columns,
                    "Dtype": [str(sales_df[c].dtype) for c in sales_df.columns]
                })
                st.dataframe(col_info, use_container_width=True)

            # Derived time columns (added without touching original data)
            if "Billing Date" in sales_df.columns:
                sales_df["Billing Month"] = sales_df["Billing Date"].dt.to_period("M").astype(str)
                sales_df["Billing Year"] = sales_df["Billing Date"].dt.year.astype(str)

            # Dimension pickers
            st.subheader("1) Choose columns to group/compare")
            # Treat object/category/bool as categorical; also include derived month/year
            categorical_cols = [c for c in sales_df.columns if (sales_df[c].dtype == 'object' or str(sales_df[c].dtype) == 'category' or sales_df[c].dtype == 'bool')]
            for c in ["Billing Month", "Billing Year"]:
                if c in sales_df.columns and c not in categorical_cols:
                    categorical_cols.append(c)
            # Remove the measure column if it's object accidentally
            if "Net Value" in categorical_cols:
                categorical_cols.remove("Net Value")

            dims = st.multiselect("Group by (1–3 columns)", options=categorical_cols, default=["Driver Name EN"] if "Driver Name EN" in categorical_cols else categorical_cols[:1], max_selections=3)
            if len(dims) == 0:
                st.warning("Please select at least one column to group by.")
            else:
                # Optional filters similar to Sales Tracking (do not change original pages)
                st.subheader("2) Optional filters")
                filt_cols = ["Driver Name EN", "Billing Type", "PY Name 1", "SP Name1"]
                filter_widgets = {}
                fl_cols = st.columns(len(filt_cols))
                for i, colname in enumerate(filt_cols):
                    if colname in sales_df.columns:
                        options = sorted(sales_df[colname].dropna().unique().tolist())
                        default_vals = options
                        filter_widgets[colname] = fl_cols[i].multiselect(f"Filter: {colname}", options=options, default=default_vals)

                # Date ranges (two periods)
                st.subheader("3) Select two periods to compare")
                left, right = st.columns(2)
                min_d, max_d = sales_df["Billing Date"].min(), sales_df["Billing Date"].max()
                with left:
                    period1 = st.date_input("Period 1", [min_d, max_d], key="custom_p1")
                with right:
                    period2 = st.date_input("Period 2", [min_d, max_d], key="custom_p2")

                p1_label = f"{period1[0].strftime('%d-%b-%Y')} to {period1[1].strftime('%d-%b-%Y')}"
                p2_label = f"{period2[0].strftime('%d-%b-%Y')} to {period2[1].strftime('%d-%b-%Y')}"

                # Apply filters
                df = sales_df.copy()
                for colname, selected in filter_widgets.items():
                    if colname in df.columns and len(selected) != len(df[colname].dropna().unique()):
                        df = df[df[colname].isin(selected)]

                df_p1 = df[(df["Billing Date"] >= pd.to_datetime(period1[0])) & (df["Billing Date"] <= pd.to_datetime(period1[1]))]
                df_p2 = df[(df["Billing Date"] >= pd.to_datetime(period2[0])) & (df["Billing Date"] <= pd.to_datetime(period2[1]))]

                # Aggregate by chosen dimensions
                agg1 = df_p1.groupby(dims)["Net Value"].sum().rename(p1_label)
                agg2 = df_p2.groupby(dims)["Net Value"].sum().rename(p2_label)

                # Merge to full view
                full_index = agg1.index.union(agg2.index)
                agg1 = agg1.reindex(full_index, fill_value=0)
                agg2 = agg2.reindex(full_index, fill_value=0)
                comparison_dyn = pd.concat([agg1, agg2], axis=1)
                comparison_dyn["Difference"] = comparison_dyn[p2_label] - comparison_dyn[p1_label]
                comparison_dyn["Comparison %"] = np.where(comparison_dyn[p1_label] != 0,
                                                          (comparison_dyn["Difference"] / comparison_dyn[p1_label] * 100).round(2),
                                                          0)
                comparison_dyn = comparison_dyn.sort_values(by=p2_label, ascending=False)

                st.subheader("4) Result table")
                st.dataframe(
                    comparison_dyn.style.format({
                        p1_label: "{:,.0f}",
                        p2_label: "{:,.0f}",
                        "Difference": "{:,.0f}",
                        "Comparison %": "{:.2f}%"
                    }).applymap(color_positive_negative, subset=["Difference","Comparison %"]),
                    use_container_width=True
                )

                # Excel download
                excel_stream_dyn = io.BytesIO()
                with pd.ExcelWriter(excel_stream_dyn, engine='xlsxwriter') as writer:
                    # Write as a regular sheet (multi-index preserved)
                    comparison_dyn.to_excel(writer, sheet_name="Custom_Comparison")
                    # Also write a flattened version for easier Excel consumption
                    flat = comparison_dyn.reset_index()
                    flat.to_excel(writer, sheet_name="Custom_Comparison_Flat", index=False)
                excel_stream_dyn.seek(0)
                st.download_button("Download Excel - Custom Comparison", data=excel_stream_dyn, file_name="Custom_Comparison.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

                # Chart
                st.subheader("5) Visualization")
                # For charting, create a flat df and choose a display dimension
                flat_df = comparison_dyn.reset_index()
                chart_dim = None
                if len(dims) == 1:
                    chart_dim = dims[0]
                else:
                    chart_dim = st.selectbox("Choose a dimension for the x-axis", dims)

                # Build a readable x label if multiple dims
                if len(dims) >= 2:
                    # Combine dims into a single label for plotting
                    flat_df["__label__"] = flat_df[dims].astype(str).agg(" | ".join, axis=1)
                    xcol = "__label__"
                else:
                    xcol = chart_dim

                fig = go.Figure()
                fig.add_trace(go.Bar(x=flat_df[xcol], y=flat_df[p1_label], name=p1_label))
                fig.add_trace(go.Bar(x=flat_df[xcol], y=flat_df[p2_label], name=p2_label))
                fig.update_layout(barmode='group', xaxis_title=" x ".join(dims), yaxis_title="Net Value", title="Custom Comparison")
                st.plotly_chart(fig, use_container_width=True)

                # Optional line for % difference (secondary y)
                fig2 = go.Figure()
                fig2.add_trace(go.Scatter(x=flat_df[xcol], y=flat_df["Comparison %"], mode='lines+markers', name="Comparison %"))
                fig2.update_layout(xaxis_title=" x ".join(dims), yaxis_title="% Difference", title="Custom Comparison - % Difference")
                st.plotly_chart(fig2, use_container_width=True)
