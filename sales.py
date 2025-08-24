import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from pptx import Presentation
from pptx.util import Inches
import io
from sklearn.linear_model import LinearRegression

# ===============================
# Page Config
# ===============================
st.set_page_config(page_title="📊 Haneef Sales Dashboard", layout="wide", page_icon="📈")

# ===============================
# Helpers
# ===============================
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

def color_positive_negative(val):
    if pd.isna(val):
        return ""
    try:
        v = float(val)
    except Exception:
        return ""
    if v > 0:
        return 'color: green; font-weight: bold'
    if v < 0:
        return 'color: red; font-weight: bold'
    return ''

def working_days_between(start_date: pd.Timestamp, end_date: pd.Timestamp) -> int:
    """Count days excluding Fridays (weekday() == 4) inclusive of endpoints."""
    if pd.isna(start_date) or pd.isna(end_date):
        return 0
    all_days = pd.date_range(start_date, end_date, freq='D')
    return sum(1 for d in all_days if d.weekday() != 4)

def working_days_in_month(anchor_date: pd.Timestamp) -> int:
    """Total working days in the month of anchor_date, excluding Fridays."""
    month_start = anchor_date.replace(day=1)
    month_end = month_start + pd.offsets.MonthEnd(0)
    days = pd.date_range(month_start, month_end, freq='D')
    return sum(1 for d in days if d.weekday() != 4)

def create_pptx(report_df, billing_df, py_df, figs_dict):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "Sales & Targets Report"
    slide.placeholders[1].text = "Generated from Sales Data"

    def add_table_slide(df, title):
        slide2 = prs.slides.add_slide(prs.slide_layouts[5])
        slide2.shapes.title.text = title
        rows, cols = df.shape
        table = slide2.shapes.add_table(rows + 1, cols, Inches(0.5), Inches(1.5), Inches(9), Inches(4)).table
        for j, col in enumerate(df.columns):
            table.cell(0, j).text = str(col)
            table.cell(0, j).text_frame.paragraphs[0].font.bold = True
        for i, row in enumerate(df.itertuples(index=False), start=1):
            for j, val in enumerate(row):
                table.cell(i, j).text = f"{val:,}" if isinstance(val, (int, float, np.integer, np.floating)) else str(val)

    def add_chart_slide(fig, title):
        slide3 = prs.slides.add_slide(prs.slide_layouts[5])
        slide3.shapes.title.text = title
        img_stream = io.BytesIO()
        try:
            # Requires kaleido installed to export plotly figures
            fig.write_image(img_stream, format='png')
            img_stream.seek(0)
            slide3.shapes.add_picture(img_stream, Inches(0.5), Inches(1.5), width=Inches(8))
        except Exception:
            slide3.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(8), Inches(4)).text_frame.text = \
                "Chart cannot be embedded. Install 'kaleido' to export plotly figures."

    add_table_slide(report_df.reset_index(), "Sales & Targets Summary")
    add_table_slide(billing_df.reset_index(), "Sales by Billing Type per Salesman")
    add_table_slide(py_df.reset_index(), "Sales by PY Name 1")

    for key, fig in figs_dict.items():
        add_chart_slide(fig, key)

    pptx_stream = io.BytesIO()
    prs.save(pptx_stream)
    pptx_stream.seek(0)
    return pptx_stream

# ===============================
# Sidebar Navigation
# ===============================
st.sidebar.title("Menu")
choice = st.sidebar.selectbox("Navigate", ["Home", "Sales Tracking", "YTD", "Custom Analysis"])

# ===============================
# Home
# ===============================
if choice == "Home":
    st.title("🏠 Welcome to Haneef Sales Dashboard")
    st.markdown("""
**Features**
- View sales & targets by Salesman, PY Name, and SP Name  
- Interactive charts for trends & gaps  
- Download reports in PPTX & Excel  
- YTD comparison across two periods  
- Custom Analysis: pick any columns to compare  

Use the sidebar to navigate to **Sales Tracking**, **YTD**, or **Custom Analysis**.
    """)

# ===============================
# Sales Tracking
# ===============================
elif choice == "Sales Tracking":
    st.title("📊 Sales Tracking Dashboard")
    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"], key="sales_upload")

    if uploaded_file:
        sales_df, target_df = load_data(uploaded_file)
        if (sales_df is None) or (target_df is None):
            st.stop()

        # --------- Filters in Sidebar ----------
        st.sidebar.subheader("Filters")
        salesmen = st.sidebar.multiselect(
            "Select Salesmen",
            options=sorted(sales_df['Driver Name EN'].dropna().unique()),
            default=sorted(sales_df['Driver Name EN'].dropna().unique())
        )
        billing_types = st.sidebar.multiselect(
            "Select Billing Types",
            options=sorted(sales_df['Billing Type'].dropna().unique()),
            default=sorted(sales_df['Billing Type'].dropna().unique())
        )
        py_filter = st.sidebar.multiselect(
            "Select PY Name",
            options=sorted(sales_df['PY Name 1'].dropna().unique()),
            default=sorted(sales_df['PY Name 1'].dropna().unique())
        )
        sp_filter = st.sidebar.multiselect(
            "Select SP Name1",
            options=sorted(sales_df['SP Name1'].dropna().unique()),
            default=sorted(sales_df['SP Name1'].dropna().unique())
        )

        preset = st.sidebar.radio("Quick Date Presets", ["Custom Range", "Last 7 Days", "This Month", "YTD"])
        today = pd.Timestamp.today().normalize()
        if preset == "Last 7 Days":
            date_range = [today - pd.Timedelta(days=7), today]
        elif preset == "This Month":
            month_start = today.replace(day=1)
            month_end = (month_start + pd.offsets.MonthEnd(0))
            date_range = [month_start, month_end]
        elif preset == "YTD":
            date_range = [today.replace(month=1, day=1), today]
        else:
            date_range = st.sidebar.date_input(
                "Select Date Range",
                [sales_df['Billing Date'].min(), sales_df['Billing Date'].max()]
            )
            if not isinstance(date_range, (list, tuple)) or len(date_range) != 2:
                st.warning("Please select a valid date range.")
                st.stop()
            date_range = [pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])]
        top_n = st.sidebar.slider(
            "Show Top N Salesmen",
            min_value=1,
            max_value=max(1, len(sales_df['Driver Name EN'].dropna().unique())),
            value=min(5, max(1, len(sales_df['Driver Name EN'].dropna().unique())))
        )

        # --------- Filter Data ----------
        df_filtered = sales_df[
            (sales_df['Driver Name EN'].isin(salesmen)) &
            (sales_df['Billing Type'].isin(billing_types)) &
            (sales_df['Billing Date'] >= pd.to_datetime(date_range[0])) &
            (sales_df['Billing Date'] <= pd.to_datetime(date_range[1])) &
            (sales_df['PY Name 1'].isin(py_filter)) &
            (sales_df['SP Name1'].isin(sp_filter))
        ].copy()

        if df_filtered.empty:
            st.info("No data for the selected filters/date range.")
            st.stop()

        # --------- Working Days & Sales Per Day ----------
        billing_start = df_filtered['Billing Date'].min()
        billing_end = df_filtered['Billing Date'].max()
        days_finish = working_days_between(billing_start, billing_end)  # exclude Fridays

        total_sales_series = df_filtered.groupby("Driver Name EN", dropna=True)["Net Value"].sum()
        total_sales_sum = float(total_sales_series.sum())
        current_sales_per_day = total_sales_sum / days_finish if days_finish > 0 else 0.0

        # --------- Month working days & forecast ----------
        total_working_days_month = working_days_in_month(today)  # exclude Fridays
        forecast_month_end_ka = current_sales_per_day * total_working_days_month  # per your spec

        # --------- Targets & Gaps ----------
        ka_targets_series = target_df.set_index("Driver Name EN")["KA Target"]
        talabat_targets_series = target_df.set_index("Driver Name EN")["Talabat Target"]
        talabat_df = df_filtered[df_filtered["PY Name 1"] == "STORES SERVICES KUWAIT CO."]
        talabat_sales_series = talabat_df.groupby("Driver Name EN")["Net Value"].sum()

        all_salesmen_idx = total_sales_series.index.union(talabat_sales_series.index).union(
            ka_targets_series.index).union(talabat_targets_series.index)

        total_sales = total_sales_series.reindex(all_salesmen_idx, fill_value=0).astype(int)
        talabat_sales = talabat_sales_series.reindex(all_salesmen_idx, fill_value=0).astype(int)
        ka_targets = ka_targets_series.reindex(all_salesmen_idx, fill_value=0).astype(int)
        talabat_targets = talabat_targets_series.reindex(all_salesmen_idx, fill_value=0).astype(int)

        ka_gap = (ka_targets - total_sales).clip(lower=0).astype(int)
        talabat_gap = (talabat_targets - talabat_sales).clip(lower=0).astype(int)

        top_salesmen_idx = total_sales.sort_values(ascending=False).head(top_n).index
        total_sales_top = total_sales.loc[top_salesmen_idx]
        talabat_sales_top = talabat_sales.loc[top_salesmen_idx]

        # --------- Tabs ----------
        tabs = st.tabs(["KPIs", "Summary Tables", "Charts", "Download Reports", "Advanced Insights"])

        # ---------- KPIs ----------
        with tabs[0]:
            st.subheader("🏆 Key Metrics")
            r1c1, r1c2, r1c3, r1c4 = st.columns(4)
            ka_target_sum = float(ka_targets.sum())
            tal_target_sum = float(talabat_targets.sum())

            ka_pct = (total_sales_sum / ka_target_sum * 100) if ka_target_sum > 0 else 0.0
            tal_pct = (float(talabat_sales.sum()) / tal_target_sum * 100) if tal_target_sum > 0 else 0.0

            r1c1.metric("Total KA Sales", f"{total_sales_sum:,.0f}", f"{ka_pct:.2f}%")
            r1c2.metric("Total Talabat Sales", f"{float(talabat_sales.sum()):,.0f}", f"{tal_pct:.2f}%")
            r1c3.metric("Total KA Gap", f"{float(ka_gap.sum()):,.0f}")
            r1c4.metric("Total Talabat Gap", f"{float(talabat_gap.sum()):,.0f}")

            r2c1, r2c2 = st.columns(2)
            if len(total_sales_top) > 0:
                r2c1.metric("Top KA Salesman", total_sales_top.idxmax(), f"{float(total_sales_top.max()):,.0f}")
            else:
                r2c1.metric("Top KA Salesman", "-", "0")
            if len(talabat_sales_top) > 0:
                r2c2.metric("Top Talabat Salesman", talabat_sales_top.idxmax(), f"{float(talabat_sales_top.max()):,.0f}")
            else:
                r2c2.metric("Top Talabat Salesman", "-", "0")

            r3c1, r3c2, r3c3 = st.columns(3)
            r3c1.metric("Working Days Finished", days_finish)
            r3c2.metric("Current Sales Per Day", f"{current_sales_per_day:,.0f}")
            r3c3.metric("Forecasted Month-End KA Sales", f"{forecast_month_end_ka:,.0f}")

        # ---------- Summary Tables ----------
        with tabs[1]:
            st.subheader("Sales & Targets Summary")

            # percentage achieved defensively
            ka_pct_ach = np.where(ka_targets > 0, (total_sales / ka_targets * 100).round(2), 0.0)
            tal_pct_ach = np.where(talabat_targets > 0, (talabat_sales / talabat_targets * 100).round(2), 0.0)

            report_df = pd.DataFrame({
                "KA Target": ka_targets,
                "KA Sales": total_sales,
                "KA Remaining": ka_gap,
                "KA % Achieved": ka_pct_ach,
                "Talabat Target": talabat_targets,
                "Talabat Sales": talabat_sales,
                "Talabat Remaining": talabat_gap,
                "Talabat % Achieved": tal_pct_ach
            })

            st.dataframe(
                report_df.style
                    .applymap(color_positive_negative, subset=["KA % Achieved", "Talabat % Achieved"])
                    .highlight_max(subset=["KA % Achieved", "Talabat % Achieved"], color="#ffef9a")
                    .format({
                        "KA Target": "{:,.0f}",
                        "KA Sales": "{:,.0f}",
                        "KA Remaining": "{:,.0f}",
                        "KA % Achieved": "{:.2f}",
                        "Talabat Target": "{:,.0f}",
                        "Talabat Sales": "{:,.0f}",
                        "Talabat Remaining": "{:,.0f}",
                        "Talabat % Achieved": "{:.2f}",
                    }),
                use_container_width=True
            )

            # Excel download for report_df
            excel_stream1 = io.BytesIO()
            with pd.ExcelWriter(excel_stream1, engine='xlsxwriter') as writer:
                report_df.to_excel(writer, sheet_name="Sales_Targets_Summary")
            excel_stream1.seek(0)
            st.download_button(
                "Download Excel - Sales & Targets Summary",
                data=excel_stream1,
                file_name="Sales_Targets_Summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # Billing Type Table
            billing_df = df_filtered.groupby(["Driver Name EN", "Billing Type"])["Net Value"].sum().unstack(fill_value=0)
            billing_df["Total"] = billing_df.sum(axis=1)

            required_cols = ["ZFR", "YKF2", "YKRE", "YKS1", "YKS2", "ZCAN", "ZRE"]
            for col in required_cols:
                if col not in billing_df.columns:
                    billing_df[col] = 0

            billing_df["Return"] = billing_df.get("YKRE", 0) + billing_df.get("ZRE", 0)
            billing_df["Return %"] = np.where(
                billing_df["Total"] != 0,
                (billing_df["Return"] / billing_df["Total"] * 100).round(2),
                0.0
            )
            billing_df.rename(columns={"YKF2": "HHT", "ZFR": "PRESALES"}, inplace=True)
            # Reorder if exists
            ordered_cols = [c for c in ["Total", "PRESALES", "HHT", "YKRE", "YKS1", "YKS2", "ZCAN", "ZRE", "Return", "Return %"] if c in billing_df.columns]
            billing_df = billing_df[ordered_cols]

            total_row = pd.DataFrame([{
                "Total": billing_df["Total"].sum(),
                "PRESALES": billing_df["PRESALES"].sum() if "PRESALES" in billing_df else 0,
                "HHT": billing_df["HHT"].sum() if "HHT" in billing_df else 0,
                "YKRE": billing_df["YKRE"].sum() if "YKRE" in billing_df else 0,
                "YKS1": billing_df["YKS1"].sum() if "YKS1" in billing_df else 0,
                "YKS2": billing_df["YKS2"].sum() if "YKS2" in billing_df else 0,
                "ZCAN": billing_df["ZCAN"].sum() if "ZCAN" in billing_df else 0,
                "ZRE": billing_df["ZRE"].sum() if "ZRE" in billing_df else 0,
                "Return": billing_df["Return"].sum(),
                "Return %": (billing_df["Return"].sum() / billing_df["Total"].sum() * 100).round(2) if billing_df["Total"].sum() != 0 else 0.0
            }], index=["Total"])

            billing_df_display = pd.concat([billing_df, total_row])

            def highlight_total_row(row):
                return ['background-color: #add8e6; color: #00008B; font-weight: bold' if row.name == "Total" else '' for _ in row]

            st.subheader("Sales by Billing Type per Salesman")
            st.dataframe(
                billing_df_display.style
                    .background_gradient(cmap="Blues", subset=[c for c in billing_df_display.columns if c not in ["Return %"]])
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
                    .apply(highlight_total_row, axis=1)
                    .applymap(color_positive_negative, subset=["Return %"]),
                use_container_width=True
            )

            # Excel download for billing_df
            excel_stream2 = io.BytesIO()
            with pd.ExcelWriter(excel_stream2, engine='xlsxwriter') as writer:
                billing_df_display.to_excel(writer, sheet_name="Billing_Types")
            excel_stream2.seek(0)
            st.download_button(
                "Download Excel - Billing Types",
                data=excel_stream2,
                file_name="Billing_Types.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # Sales by PY Name 1
            py_df = df_filtered.groupby("PY Name 1")["Net Value"].sum().sort_values(ascending=False)
            st.subheader("Sales by PY Name 1")
            st.dataframe(py_df.to_frame().style.background_gradient(cmap="Greens").format("{:,.0f}"), use_container_width=True)

            excel_stream3 = io.BytesIO()
            with pd.ExcelWriter(excel_stream3, engine='xlsxwriter') as writer:
                py_df.to_frame().to_excel(writer, sheet_name="Sales_by_PY_Name")
            excel_stream3.seek(0)
            st.download_button(
                "Download Excel - Sales by PY Name",
                data=excel_stream3,
                file_name="Sales_by_PY_Name.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # ---------- Charts ----------
        with tabs[2]:
            st.subheader("Sales Charts")
            figs_dict = {}

            # Use the same report_df from Summary Tables; recompute here for safety
            ka_pct_ach = np.where(ka_targets > 0, (total_sales / ka_targets * 100).round(2), 0.0)
            tal_pct_ach = np.where(talabat_targets > 0, (talabat_sales / talabat_targets * 100).round(2), 0.0)
            report_df_for_chart = pd.DataFrame({
                "KA Target": ka_targets,
                "KA Sales": total_sales,
                "KA Remaining": ka_gap,
                "KA % Achieved": ka_pct_ach,
                "Talabat Target": talabat_targets,
                "Talabat Sales": talabat_sales,
                "Talabat Remaining": talabat_gap,
                "Talabat % Achieved": tal_pct_ach
            })

            fig_ka = px.bar(
                report_df_for_chart.reset_index(),
                x="Driver Name EN", y="KA Sales",
                title="KA Sales by Salesman", text="KA Sales"
            )
            fig_talabat = px.bar(
                report_df_for_chart.reset_index(),
                x="Driver Name EN", y="Talabat Sales",
                title="Talabat Sales by Salesman", text="Talabat Sales"
            )
            st.plotly_chart(fig_ka, use_container_width=True)
            st.plotly_chart(fig_talabat, use_container_width=True)

            figs_dict["KA Sales by Salesman"] = fig_ka
            figs_dict["Talabat Sales by Salesman"] = fig_talabat

        # ---------- Download Reports ----------
        with tabs[3]:
            st.subheader("Download PPTX Report")
            # Reuse dataframes computed above
            # For PPT, rebuild the same report_df and billing_df to pass
            ka_pct_ach = np.where(ka_targets > 0, (total_sales / ka_targets * 100).round(2), 0.0)
            tal_pct_ach = np.where(talabat_targets > 0, (talabat_sales / talabat_targets * 100).round(2), 0.0)
            report_df_dl = pd.DataFrame({
                "KA Target": ka_targets,
                "KA Sales": total_sales,
                "KA Remaining": ka_gap,
                "KA % Achieved": ka_pct_ach,
                "Talabat Target": talabat_targets,
                "Talabat Sales": talabat_sales,
                "Talabat Remaining": talabat_gap,
                "Talabat % Achieved": tal_pct_ach
            })

            billing_df = df_filtered.groupby(["Driver Name EN", "Billing Type"])["Net Value"].sum().unstack(fill_value=0)
            billing_df["Total"] = billing_df.sum(axis=1)
            required_cols = ["ZFR", "YKF2", "YKRE", "YKS1", "YKS2", "ZCAN", "ZRE"]
            for col in required_cols:
                if col not in billing_df.columns:
                    billing_df[col] = 0
            billing_df["Return"] = billing_df.get("YKRE", 0) + billing_df.get("ZRE", 0)
            billing_df["Return %"] = np.where(billing_df["Total"] != 0, (billing_df["Return"] / billing_df["Total"] * 100).round(2), 0.0)
            billing_df.rename(columns={"YKF2": "HHT", "ZFR": "PRESALES"}, inplace=True)
            ordered_cols = [c for c in ["Total", "PRESALES", "HHT", "YKRE", "YKS1", "YKS2", "ZCAN", "ZRE", "Return", "Return %"] if c in billing_df.columns]
            billing_df = billing_df[ordered_cols]

            py_df = df_filtered.groupby("PY Name 1")["Net Value"].sum().sort_values(ascending=False)

            # Simple figs for PPT
            figs_dict = {}
            figs_dict["KA Sales by Salesman"] = px.bar(report_df_dl.reset_index(), x="Driver Name EN", y="KA Sales", title="KA Sales by Salesman")
            figs_dict["Talabat Sales by Salesman"] = px.bar(report_df_dl.reset_index(), x="Driver Name EN", y="Talabat Sales", title="Talabat Sales by Salesman")

            pptx_data = create_pptx(report_df_dl, billing_df, py_df, figs_dict)
            st.download_button("Download PPTX Report", data=pptx_data, file_name="sales_report.pptx")

        # ---------- Advanced Insights ----------
        with tabs[4]:
            st.subheader("Sales Trend Forecast (Simple Linear Regression)")
            df_time = df_filtered.groupby("Billing Date")["Net Value"].sum().reset_index().sort_values("Billing Date")
            if len(df_time) > 1:
                X = np.arange(len(df_time)).reshape(-1, 1)
                y = df_time["Net Value"].values
                model = LinearRegression()
                model.fit(X, y)
                df_time["Forecast"] = model.predict(X)

                fig_trend = px.line(df_time, x="Billing Date", y=["Net Value", "Forecast"], title="Sales Trend with Forecast")
                st.plotly_chart(fig_trend, use_container_width=True)
            else:
                st.info("Not enough data to generate trend forecast.")

# ===============================
# YTD
# ===============================
elif choice == "YTD":
    st.title("📅 Year-to-Date (YTD) Comparison")
    uploaded_file = st.file_uploader("Upload Excel for YTD Comparison", type=["xlsx"], key="ytd_upload")

    if uploaded_file:
        sales_df, _ = load_data(uploaded_file)
        if sales_df is None:
            st.stop()

        dimension = st.selectbox("Compare By", ["Driver Name EN", "PY Name 1", "SP Name1"])

        st.subheader("Select Two Periods to Compare")
        c1, c2 = st.columns(2)
        with c1:
            period1 = st.date_input("Period 1 Start-End", [sales_df['Billing Date'].min(), sales_df['Billing Date'].max()], key="ytd_p1")
        with c2:
            period2 = st.date_input("Period 2 Start-End", [sales_df['Billing Date'].min(), sales_df['Billing Date'].max()], key="ytd_p2")

        if len(period1) != 2 or len(period2) != 2:
            st.warning("Please select valid date ranges.")
            st.stop()

        period1 = [pd.to_datetime(period1[0]), pd.to_datetime(period1[1])]
        period2 = [pd.to_datetime(period2[0]), pd.to_datetime(period2[1])]

        period1_label = f"{period1[0].strftime('%d-%b-%Y')} to {period1[1].strftime('%d-%b-%Y')}"
        period2_label = f"{period2[0].strftime('%d-%b-%Y')} to {period2[1].strftime('%d-%b-%Y')}"

        df1 = sales_df[(sales_df['Billing Date'] >= period1[0]) & (sales_df['Billing Date'] <= period1[1])]
        df2 = sales_df[(sales_df['Billing Date'] >= period2[0]) & (sales_df['Billing Date'] <= period2[1])]

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
        comparison_df["Comparison %"] = np.where(
            comparison_df[period1_label] != 0,
            (comparison_df["Difference"] / comparison_df[period1_label] * 100).round(2),
            0.0
        )
        comparison_df = comparison_df.sort_values(by=period2_label, ascending=False)

        def highlight_date_columns(row):
            return [
                'background-color: #d1e7dd; font-weight: bold' if col in [period1_label, period2_label] else ''
                for col in row.index
            ]

        st.subheader(f"YTD Comparison by {dimension}")
        st.dataframe(
            comparison_df.style.format({
                period1_label: "{:,.0f}",
                period2_label: "{:,.0f}",
                "Difference": "{:,.0f}",
                "Comparison %": "{:.2f}%"
            }).applymap(color_positive_negative, subset=["Difference", "Comparison %"])
              .apply(highlight_date_columns, axis=1),
            use_container_width=True
        )

        excel_stream_ytd = io.BytesIO()
        with pd.ExcelWriter(excel_stream_ytd, engine='xlsxwriter') as writer:
            comparison_df.to_excel(writer, sheet_name="YTD Comparison")
        excel_stream_ytd.seek(0)
        st.download_button(
            "Download Excel - YTD Comparison",
            data=excel_stream_ytd,
            file_name="YTD_Comparison.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.subheader("YTD Comparison Chart")
        fig = go.Figure()
        fig.add_trace(go.Bar(x=comparison_df.index, y=comparison_df[period1_label], name=period1_label))
        fig.add_trace(go.Bar(x=comparison_df.index, y=comparison_df[period2_label], name=period2_label))
        fig.add_trace(go.Scatter(x=comparison_df.index, y=comparison_df["Comparison %"], name="% Difference",
                                 mode='lines+markers', yaxis='y2'))
        fig.update_layout(
            barmode='group',
            title=f"YTD Comparison by {dimension}",
            xaxis_title=dimension,
            yaxis=dict(title="Net Value"),
            yaxis2=dict(title="% Difference", overlaying='y', side='right', showgrid=False)
        )
        st.plotly_chart(fig, use_container_width=True)

# ===============================
# Custom Analysis
# ===============================
elif choice == "Custom Analysis":
    st.title("🔍 Custom Analysis (Pick Any Columns)")
    uploaded_file = st.file_uploader("Upload Excel for Custom Analysis", type=["xlsx"], key="custom_upload")

    if uploaded_file:
        sales_df, _ = load_data(uploaded_file)
        if sales_df is None:
            st.stop()

        with st.expander("👀 Show available columns in your data"):
            col_info = pd.DataFrame({
                "Column": sales_df.columns,
                "Dtype": [str(sales_df[c].dtype) for c in sales_df.columns]
            })
            st.dataframe(col_info, use_container_width=True)

        # Derived time columns
        if "Billing Date" in sales_df.columns:
            sales_df["Billing Month"] = sales_df["Billing Date"].dt.to_period("M").astype(str)
            sales_df["Billing Year"] = sales_df["Billing Date"].dt.year.astype(str)

        st.subheader("1) Choose columns to group/compare")
        categorical_cols = [
            c for c in sales_df.columns
            if (sales_df[c].dtype == 'object' or str(sales_df[c].dtype) == 'category' or sales_df[c].dtype == 'bool')
        ]
        for c in ["Billing Month", "Billing Year"]:
            if c in sales_df.columns and c not in categorical_cols:
                categorical_cols.append(c)
        if "Net Value" in categorical_cols:
            categorical_cols.remove("Net Value")

        default_dims = ["Driver Name EN"] if "Driver Name EN" in categorical_cols else (categorical_cols[:1] if len(categorical_cols) else [])
        dims = st.multiselect("Group by (1–3 columns)", options=categorical_cols, default=default_dims, max_selections=3)

        if len(dims) == 0:
            st.warning("Please select at least one column to group by.")
            st.stop()

        # Optional filters
        st.subheader("2) Optional filters")
        filt_cols = ["Driver Name EN", "Billing Type", "PY Name 1", "SP Name1"]
        filter_widgets = {}
        fl_cols = st.columns(len(filt_cols))
        for i, colname in enumerate(filt_cols):
            if colname in sales_df.columns:
                options = sorted(sales_df[colname].dropna().unique().tolist())
                default_vals = options
                filter_widgets[colname] = fl_cols[i].multiselect(f"Filter: {colname}", options=options, default=default_vals)

        # Date ranges
        st.subheader("3) Select two periods to compare")
        left, right = st.columns(2)
        min_d, max_d = sales_df["Billing Date"].min(), sales_df["Billing Date"].max()
        with left:
            period1 = st.date_input("Period 1", [min_d, max_d], key="custom_p1")
        with right:
            period2 = st.date_input("Period 2", [min_d, max_d], key="custom_p2")

        if len(period1) != 2 or len(period2) != 2:
            st.warning("Please select valid date ranges.")
            st.stop()

        p1 = [pd.to_datetime(period1[0]), pd.to_datetime(period1[1])]
        p2 = [pd.to_datetime(period2[0]), pd.to_datetime(period2[1])]
        p1_label = f"{p1[0].strftime('%d-%b-%Y')} to {p1[1].strftime('%d-%b-%Y')}"
        p2_label = f"{p2[0].strftime('%d-%b-%Y')} to {p2[1].strftime('%d-%b-%Y')}"

        # Apply filters
        df = sales_df.copy()
        for colname, selected in filter_widgets.items():
            if colname in df.columns and len(selected) != len(df[colname].dropna().unique()):
                df = df[df[colname].isin(selected)]

        df_p1 = df[(df["Billing Date"] >= p1[0]) & (df["Billing Date"] <= p1[1])]
        df_p2 = df[(df["Billing Date"] >= p2[0]) & (df["Billing Date"] <= p2[1])]

        agg1 = df_p1.groupby(dims)["Net Value"].sum().rename(p1_label)
        agg2 = df_p2.groupby(dims)["Net Value"].sum().rename(p2_label)

        full_index = agg1.index.union(agg2.index)
        agg1 = agg1.reindex(full_index, fill_value=0)
        agg2 = agg2.reindex(full_index, fill_value=0)
        comparison_dyn = pd.concat([agg1, agg2], axis=1)
        comparison_dyn["Difference"] = comparison_dyn[p2_label] - comparison_dyn[p1_label]
        comparison_dyn["Comparison %"] = np.where(
            comparison_dyn[p1_label] != 0,
            (comparison_dyn["Difference"] / comparison_dyn[p1_label] * 100).round(2),
            0.0
        )
        comparison_dyn = comparison_dyn.sort_values(by=p2_label, ascending=False)

        st.subheader("4) Result table")
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
        with pd.ExcelWriter(excel_stream_dyn, engine='xlsxwriter') as writer:
            comparison_dyn.to_excel(writer, sheet_name="Custom_Comparison")
            flat = comparison_dyn.reset_index()
            flat.to_excel(writer, sheet_name="Custom_Comparison_Flat", index=False)
        excel_stream_dyn.seek(0)
        st.download_button(
            "Download Excel - Custom Comparison",
            data=excel_stream_dyn,
            file_name="Custom_Comparison.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Visualization
        st.subheader("5) Visualization")
        flat_df = comparison_dyn.reset_index()
        if len(dims) == 1:
            xcol = dims[0]
        else:
            xcol = st.selectbox("Choose a dimension for the x-axis", dims)
            # Combine multiple dims to a single label if user picks one that's shared
            flat_df["__label__"] = flat_df[dims].astype(str).agg(" | ".join, axis=1)
            if xcol not in dims:
                xcol = "__label__"

        fig = go.Figure()
        fig.add_trace(go.Bar(x=flat_df[xcol], y=flat_df[p1_label], name=p1_label))
        fig.add_trace(go.Bar(x=flat_df[xcol], y=flat_df[p2_label], name=p2_label))
        fig.update_layout(barmode='group', xaxis_title=" x ".join(dims), yaxis_title="Net Value", title="Custom Comparison")
        st.plotly_chart(fig, use_container_width=True)

        fig2 = go.Figure()
        fig2.add_trace(go.Scatter(x=flat_df[xcol], y=flat_df["Comparison %"], mode='lines+markers', name="Comparison %"))
        fig2.update_layout(xaxis_title=" x ".join(dims), yaxis_title="% Difference", title="Custom Comparison - % Difference")
        st.plotly_chart(fig2, use_container_width=True)
