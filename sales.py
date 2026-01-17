import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from plotly.subplots import make_subplots
import plotly.graph_objects as go
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import io
from sklearn.linear_model import LinearRegression
from statsmodels.tsa.holtwinters import ExponentialSmoothing
import os
from datetime import datetime
from prophet import Prophet
import io
import base64
import streamlit_authenticator as stauth
import hashlib
from difflib import SequenceMatcher
from fuzzywuzzy import fuzz
from io import BytesIO

# --- Language Selector ---
st.sidebar.header("Language / اللغة")
language = st.sidebar.selectbox("Choose / اختر", ["English", "العربية"])
lang = "en" if language == "English" else "ar"

if lang == "ar":
    st.markdown("""
    <style>
    .stApp {
        direction: rtl;
        text-align: right;
    }
    .stButton > button {
        float: right;
    }
    .dataframe th, .dataframe td {
        text-align: right !important;
        line-height: normal !important;
    }
    </style>
    """, unsafe_allow_html=True)

texts = {
    "en": {
        "page_title": "📊 Haneef Data Dashboard",
        "layout": "wide",
        "page_icon": "📈",
        "welcome": "Welcome {0} 👋",
        "logout": "Logout",
        "incorrect_login": "❌ Username/password is incorrect",
        "no_login": "⚠️ Please enter your username and password",
        "dark_mode": "🌙 Dark Mode",
        "upload_header": "📂 Upload Excel (one-time)",
        "upload_tooltip": "Upload an Excel file with sheets: sales data, Target, sales channels, and optionally YTD.",
        "clear_data": "🔁 Clear data",
        "file_loaded": "✅ File loaded — now use the menu to go to any page.",
        "menu_title": "🧭 Menu",
        "navigate": "Navigate",
        "home": "Home",
        "sales_tracking": "Sales Tracking",
        "ytd_comparison": "Year to Date Comparison",
        "custom_analysis": "Custom Analysis",
        "target_allocation": "SP/PY Target Allocation",
        "ai_insights": "AI Insights",
        "customer_insights": "Customer Insights",
        "customer_insights_title": "Customer Insights Dashboard",
        "material_forecast": "Material Forecast",
        "material_forecast_title": "📈 Material Forecast",
        "rfm_analysis_sub": "RFM Analysis",
        "rfm_table_sub": "RFM Table",
        "rfm_chart_sub": "RFM Visualization",
        "rfm_no_data": "No data available for RFM analysis.",
        "rfm_download": "Download RFM Report",
        "rfm_cohort_sub": "RFM Cohort Analysis",
        "rfm_cohort_info": "Analyzes how RFM scores evolve over time for customer acquisition cohorts (grouped by first purchase month).",
        "rfm_cohort_table_sub": "Cohort Summary Table",
        "rfm_cohort_insights_sub": "Key Insights",
        "rfm_cohort_download": "Download Cohort Report (Excel)",
        "rfm_cohort_no_data": "Insufficient data for cohort analysis.",
        "product_availability_checker": "Product Availability Checker",
        "home_title": "🏠 Haneef Data Dashboard",
        "home_welcome": "**Welcome to your Sales Analytics Hub!**\n- 📈 Track sales & targets by salesman, By Customer Name, By Branch Name\n- 📊 Visualize trends with interactive charts (now with advanced forecasting)\n- 💾 Download reports in PPTX & Excel\n- 📅 Compare sales across custom periods\n- 🎯 Allocate SP/PY targets based on recent performance\nUse the sidebar to navigate and upload data once.",
        "data_loaded_msg": "Data is loaded — choose a page from the menu.",
        "upload_prompt": "Please upload your Excel file in the sidebar to start.",
        "sales_tracking_title": "📊 MTD Tracking",
        "no_data_warning": "⚠️ Please upload the Excel file in the sidebar (one-time).",
        "filters_header": "🔍 Filters (Sales Tracking)",
        "filters_tooltip": "Filter data by salesmen, billing types, PY, SP, and date range.",
        "select_salesmen": "👥 Select Salesmen",
        "select_billing_types": "📋 Select Billing Types",
        "select_py": "🏬 Select PY Name",
        "select_sp": "🏷️ Select SP Name1",
        "date_presets": "📅 Quick Date Presets",
        "date_presets_options": ["Custom Range", "Last 7 Days", "This Month", "YTD"],
        "select_date_range": "📆 Select Date Range",
        "date_error": "❌ Start date must be before end date.",
        "top_n_salesmen": "🏆 Show Top N Salesmen",
        "no_match_warning": "⚠️ No data matches the selected filters.",
        "kpis_tab": "📈 KPIs",
        "tables_tab": "📋 Tables",
        "charts_tab": "📊 Charts",
        "downloads_tab": "💾 Downloads",
        "key_metrics_sub": "🏆 Key Metrics",
        "total_ka_sales": "Total KA Sales",
        "of_ka_target": "{0:.0f}% of KA Target Achieved",
        "ka_other_ecom": "KA & Other E-com",
        "of_ka_target_pct": "{0:.0f}% of KA Target",
        "talabat_sales": "Talabat Sales",
        "of_talabat_target": "{0:.0f}% of Talabat Target Achieved",
        "target_overview_sub": "🎯 Target Overview",
        "ka_target": "KA Target",
        "talabat_target": "Talabat Target",
        "ka_gap": "KA Gap",
        "talabat_gap": "Talabat Gap",
        "channel_sales_sub": "📊 Channel Sales",
        "retail_sales": "Retail Sales",
        "of_total_ka": "{0:.0f}% of Total KA Sales",
        "ecom_sales": "E-com Sales",
        "performance_metrics_sub": "📈 Performance Metrics",
        "days_finished": "Days Finished (working)",
        "current_sales_per_day": "Current Sales Per Day",
        "forecast_month_end": "Forecasted Month-End KA Sales",
        "sales_targets_summary_sub": "📋 Sales & Targets Summary-Value",
        "download_sales_targets": "⬇️ Download Sales & Targets Summary (Excel)",
        "sales_by_billing_sub": "📊 Sales By Billing Type-Value",
        "download_billing": "⬇️ Download Billing Type Table (Excel)",
        "sales_by_py_sub": "🏬 Sales Summary By Customer-Value",
        "download_py": "⬇️ Download PY Name Table (Excel)",
        "daily_sales_trend_sub": "📊 Daily Sales Trend And Forecast",
        "daily_sales_title": "Daily Sales Trend And Forecast",
        "not_enough_data": "Not enough data to perform a time-series forecast.",
        "market_vs_ecom_sub": "📊 Market vs E-com Sales",
        "market_vs_ecom_title": "Market vs E-com Sales Distribution",
        "daily_ka_target_sub": "📊 Daily KA Target vs Actual Sales",
        "daily_ka_title": "Daily KA Target vs Actual Sales",
        "salesman_ka_sub": "📊 Salesman KA Target vs Actual",
        "salesman_ka_title": "KA Target vs Actual Sales by Salesman",
        "top10_py_sub": "📊 Top 10 Customers",
        "top10_py_title": "Top 10 Customer 1 by Sales",
        "download_reports_sub": "💾 Download Reports",
        "generate_pptx": "📑 Generate PPTX Report",
        "download_pptx": "⬇️ Download PPTX Report",
        "ytd_title": "📊 Year to Date Comparison",
        "ytd_filters_header": "🔍 Filters (YTD)",
        "ytd_filters_tooltip": "Filter data by salesmen, billing types, PY, SP, and date range.",
        "ytd_select_group": "Group By",
        "ytd_select_value": "Value Column",
        "ytd_period1": "Select Period 1",
        "ytd_period2": "Select Period 2",
        "ytd_no_data": "⚠️ No data matches the filters.",
        "ytd_comparison_sub": "📋 YTD Comparison Table",
        "ytd_download": "⬇️ Download YTD Comparison (Excel)",
        "ytd_chart_title": "YTD Comparison Chart",
        "custom_title": "📊 Custom Analysis",
        "custom_upload": "Upload Additional Sheet (optional)",
        "custom_extra_loaded": "✅ Extra sheet loaded.",
        "custom_select_sheet": "📑 Select Sheet for Analysis",
        "custom_sheet_empty": "⚠️ The sheet '{0}' is empty or not available in your file.",
        "custom_explore": "💡 Explore your data by multiple columns & compare two periods.",
        "custom_group_cols": "Group by columns",
        "custom_value_col": "Value to analyze",
        "custom_periods_sub": "📆 Select Two Periods",
        "custom_period1": "Period 1",
        "custom_period2": "Period 2",
        "custom_select_p1": "Select Period 1",
        "custom_select_p2": "Select Period 2",
        "custom_select_prompt": "👉 Please select at least one group column, one value column, and valid date ranges.",
        "custom_comparison_sub": "📋 Comparison of {0} by {1}",
        "custom_download": "⬇️ Download Comparison (Excel)",
        "target_alloc_title": "🎯 SP/PY Target Allocation",
        "target_config_sub": "Configuration",
        "target_alloc_tooltip": "Allocate targets by branch or customer based on historical sales.",
        "target_alloc_type": "Select Target Allocation Type",
        "target_alloc_options": ["By Branch", "By Customer"],
        "target_input_option": "Select Target Input Option",
        "target_input_options": ["Manual", "Auto (from 'Target' sheet)"],
        "target_enter_total": "Enter the Total Target to be Allocated for this Month (KD)",
        "target_auto_info": "Using Total Target from 'Target' sheet: KD {0:,.0f}",
        "target_zero_warning": "Please ensure the total target is greater than 0.",
        "target_hist_sub": "Historical Data Period",
        "target_hist_option": "Select Historical Data Period",
        "target_hist_options": ["Last 6 Months", "Manual Days"],
        "target_select_range": "Select date range",
        "target_date_warning": "Please select both a start and an end date.",
        "target_no_hist": "⚠️ No sales data available in 'YTD' for {0}.",
        "target_analysis_sub": "🎯 Target Analysis",
        "hist_sales_total": "Historical Sales Total",
        "alloc_target_total": "Allocated Target Total",
        "increase_needed": "Increase Needed vs Avg Sales",
        "current_month_sales": "Current Month Sales",
        "target_balance": "Target Balance",
        "alloc_targets_sub": "📊 Auto-Allocated Targets Based on {0}",
        "download_alloc": "💾 Download Target Allocation Table",
        "ai_title": "🤖 AI Insights",
        "ai_scope_sub": "Scope and filters",
        "ai_select_period": "Select period for insights",
        "ai_top_n": "Top-N salesmen spotlight",
        "ai_no_data": "No data in the selected period. Try expanding the date range.",
        "ai_exec_sub": "📜 Executive summary",
        "ai_prescript_sub": "🛠️ Prescriptive Recommendations",
        "ai_visual_sub": "📊 AI-Generated Visuals",
        "ai_section_sub": "🧭 Section insights",
        "ai_download_sub": "📥 Download executive summary",
        "ai_download_button": "💾 Download AI executive summary (TXT)",
        "ai_ask_sub": "💬 Ask a question about your data",
        "ai_ask_prompt": "Type a question (e.g., 'Which salesman is growing fastest?', 'Where are returns highest?', 'Correlation between targets and sales?')",
        "admin_tools": "Admin Tools",
        "view_logs": "View Audit Logs",
        "audit_title": "📋 Audit Logs",
        "download_logs": "⬇️ Download Audit Logs (Excel)",
        "sheet_missing": "❌ Excel file must contain sheets: {0}. Missing: {1}",
        "cols_missing": "❌ Missing required columns: {0}",
        "load_error": "❌ Error loading Excel file: {0}",
        "pptx_title": "Sales & Targets Report",
        "pptx_generated": "Generated on {0}",
        "pptx_kpi_title": "📈 Key Performance Indicators",
        "pptx_summary_title": "📋 Sales & Targets Summary",
        "pptx_billing_title": "📊 Sales by Billing Type per Salesman",
        "pptx_py_title": "🏬 Sales by PY Name 1",
        "pptx_embed_error": "Chart cannot be embedded: {0}. Install kaleido if missing.",
        "generating_pptx": "⏳ Generating PPTX report...",
        "loading_data": "⏳ Loading Excel data...",
    },
    "ar": {
        "page_title": "📊 لوحة تحكم بيانات حنيف",
        "layout": "wide",
        "page_icon": "📈",
        "welcome": "مرحبا {0} 👋",
        "logout": "تسجيل الخروج",
        "incorrect_login": "❌ اسم المستخدم/كلمة المرور غير صحيحة",
        "no_login": "⚠️ يرجى إدخال اسم المستخدم وكلمة المرور",
        "dark_mode": "🌙 الوضع الداكن",
        "upload_header": "📂 تحميل إكسل (مرة واحدة)",
        "upload_tooltip": "تحميل ملف إكسل يحتوي على أوراق: بيانات المبيعات، الهدف، قنوات المبيعات، واختيارياً YTD.",
        "clear_data": "🔁 مسح البيانات",
        "file_loaded": "✅ تم تحميل الملف — استخدم القائمة الآن للذهاب إلى أي صفحة.",
        "menu_title": "🧭 القائمة",
        "navigate": "التنقل",
        "home": "الرئيسية",
        "sales_tracking": "تتبع المبيعات",
        "ytd_comparison": "مقارنة من بداية السنة",
        "custom_analysis": "تحليل مخصص",
        "target_allocation": "تخصيص أهداف SP/PY",
        "ai_insights": "رؤى الذكاء الاصطناعي",
        "customer_insights": "رؤى العملاء",
        "customer_insights_title": "لوحة رؤى العملاء",
        "material_forecast": "توقعات المواد",
        "material_forecast_title": "📈 توقعات المواد",
        "rfm_analysis_sub": "تحليل RFM",
        "rfm_table_sub": "جدول RFM",
        "rfm_chart_sub": "تصور RFM",
        "rfm_no_data": "لا توجد بيانات متاحة لتحليل RFM.",
        "rfm_download": "تنزيل تقرير RFM",
        "rfm_cohort_sub": "تحليل الأتراب RFM",
        "rfm_cohort_info": "يحلل كيفية تطور درجات RFM بمرور الوقت لأتراب اكتساب العملاء (مجمعة حسب شهر الشراء الأول).",
        "rfm_cohort_table_sub": "جدول ملخص الأتراب",
        "rfm_cohort_insights_sub": "الرؤى الرئيسية",
        "rfm_cohort_download": "تنزيل تقرير الأتراب (إكسل)",
        "rfm_cohort_no_data": "بيانات غير كافية لتحليل الأتراب.",
        "product_availability_checker": "مدقق توفر المنتج",
        "home_title": "🏠 لوحة تحكم بيانات حنيف",
        "home_welcome": "**مرحبا بك في مركز تحليلات المبيعات!**\n- 📈 تتبع المبيعات والأهداف حسب البائع، اسم العميل، اسم الفرع\n- 📊 تصور الاتجاهات مع رسوم بيانية تفاعلية (الآن مع تنبؤ متقدم)\n- 💾 تنزيل التقارير في PPTX و Excel\n- 📅 مقارنة المبيعات عبر فترات مخصصة\n- 🎯 تخصيص أهداف SP/PY بناءً على الأداء الأخير\nاستخدم الشريط الجانبي للتنقل وتحميل البيانات مرة واحدة.",
        "data_loaded_msg": "تم تحميل البيانات — اختر صفحة من القائمة.",
        "upload_prompt": "يرجى تحميل ملف الإكسل في الشريط الجانبي للبدء.",
        "sales_tracking_title": "📊 تتبع MTD",
        "no_data_warning": "⚠️ يرجى تحميل ملف الإكسل في الشريط الجانبي (مرة واحدة).",
        "filters_header": "🔍 فلاتر (تتبع المبيعات)",
        "filters_tooltip": "تصفية البيانات حسب البائعين، أنواع الفواتير، PY، SP، ونطاق التاريخ.",
        "select_salesmen": "👥 اختر البائعين",
        "select_billing_types": "📋 اختر أنواع الفواتير",
        "select_py": "🏬 اختر اسم PY",
        "select_sp": "🏷️ اختر اسم SP1",
        "date_presets": "📅 إعدادات تاريخ سريعة",
        "date_presets_options": ["نطاق مخصص", "آخر 7 أيام", "هذا الشهر", "من بداية السنة"],
        "select_date_range": "📆 اختر نطاق التاريخ",
        "date_error": "❌ تاريخ البداية يجب أن يكون قبل تاريخ النهاية.",
        "top_n_salesmen": "🏆 عرض أفضل N بائعين",
        "no_match_warning": "⚠️ لا توجد بيانات تطابق الفلاتر المختارة.",
        "kpis_tab": "📈 المؤشرات الرئيسية",
        "tables_tab": "📋 الجداول",
        "charts_tab": "📊 الرسوم البيانية",
        "downloads_tab": "💾 التنزيلات",
        "key_metrics_sub": "🏆 المقاييس الرئيسية",
        "total_ka_sales": "إجمالي مبيعات KA",
        "of_ka_target": "{0:.0f}% من هدف KA المحقق",
        "ka_other_ecom": "KA و E-com أخرى",
        "of_ka_target_pct": "{0:.0f}% من هدف KA",
        "talabat_sales": "مبيعات طلبات",
        "of_talabat_target": "{0:.0f}% من هدف طلبات المحقق",
        "target_overview_sub": "🎯 نظرة عامة على الأهداف",
        "ka_target": "هدف KA",
        "talabat_target": "هدف طلبات",
        "ka_gap": "فجوة KA",
        "talabat_gap": "فجوة طلبات",
        "channel_sales_sub": "📊 مبيعات القنوات",
        "retail_sales": "مبيعات التجزئة",
        "of_total_ka": "{0:.0f}% من إجمالي مبيعات KA",
        "ecom_sales": "مبيعات E-com",
        "performance_metrics_sub": "📈 مقاييس الأداء",
        "days_finished": "الأيام المكتملة (العملية)",
        "current_sales_per_day": "المبيعات الحالية لكل يوم",
        "forecast_month_end": "مبيعات KA المتوقعة في نهاية الشهر",
        "sales_targets_summary_sub": "📋 ملخص المبيعات والأهداف",
        "download_sales_targets": "⬇️ تنزيل ملخص المبيعات والأهداف (إكسل)",
        "sales_by_billing_sub": "📊 المبيعات حسب نوع الفاتورة لكل بائع",
        "download_billing": "⬇️ تنزيل جدول أنواع الفواتير (إكسل)",
        "sales_by_py_sub": "🏬 المبيعات حسب اسم PY 1",
        "download_py": "⬇️ تنزيل جدول اسم PY (إكسل)",
        "daily_sales_trend_sub": "📊 اتجاه المبيعات اليومي مع تنبؤ Prophet",
        "daily_sales_title": "اتجاه المبيعات اليومي، تنبؤ Prophet والشذوذات",
        "not_enough_data": "لا توجد بيانات كافية لإجراء تنبؤ زمني.",
        "market_vs_ecom_sub": "📊 السوق مقابل مبيعات E-com",
        "market_vs_ecom_title": "توزيع مبيعات السوق مقابل E-com",
        "daily_ka_target_sub": "📊 هدف KA اليومي مقابل المبيعات الفعلية",
        "daily_ka_title": "هدف KA اليومي مقابل المبيعات الفعلية",
        "salesman_ka_sub": "📊 هدف KA للبائع مقابل الفعلي",
        "salesman_ka_title": "هدف KA مقابل المبيعات الفعلية حسب البائع",
        "top10_py_sub": "📊 أفضل 10 أسماء PY 1 حسب المبيعات",
        "top10_py_title": "أفضل 10 أسماء PY 1 حسب المبيعات",
        "download_reports_sub": "💾 تنزيل التقارير",
        "generate_pptx": "📑 توليد تقرير PPTX",
        "download_pptx": "⬇️ تنزيل تقرير PPTX",
        "ytd_title": "📊 مقارنة من بداية السنة",
        "ytd_filters_header": "🔍 فلاتر (YTD)",
        "ytd_filters_tooltip": "تصفية البيانات حسب البائعين، أنواع الفواتير، PY، SP، ونطاق التاريخ.",
        "ytd_select_group": "التجميع حسب",
        "ytd_select_value": "عمود القيمة",
        "ytd_period1": "اختر الفترة 1",
        "ytd_period2": "اختر الفترة 2",
        "ytd_no_data": "⚠️ لا توجد بيانات تطابق الفلاتر.",
        "ytd_comparison_sub": "📋 جدول مقارنة YTD",
        "ytd_download": "⬇️ تنزيل مقارنة YTD (إكسل)",
        "ytd_chart_title": "رسم بياني لمقارنة YTD",
        "custom_title": "📊 تحليل مخصص",
        "custom_upload": "تحميل ورقة إضافية (اختياري)",
        "custom_extra_loaded": "✅ تم تحميل الورقة الإضافية.",
        "custom_select_sheet": "📑 اختر الورقة للتحليل",
        "custom_sheet_empty": "⚠️ الورقة '{0}' فارغة أو غير متوفرة في ملفك.",
        "custom_explore": "💡 استكشف بياناتك حسب أعمدة متعددة وقارن فترتين.",
        "custom_group_cols": "التجميع حسب الأعمدة",
        "custom_value_col": "القيمة للتحليل",
        "custom_periods_sub": "📆 اختر فترتين",
        "custom_period1": "الفترة 1",
        "custom_period2": "الفترة 2",
        "custom_select_p1": "اختر الفترة 1",
        "custom_select_p2": "اختر الفترة 2",
        "custom_select_prompt": "👉 يرجى اختيار عمود تجميع واحد على الأقل، عمود قيمة واحد، ونطاقات تاريخ صالحة.",
        "custom_comparison_sub": "📋 مقارنة {0} حسب {1}",
        "custom_download": "⬇️ تنزيل المقارنة (إكسل)",
        "target_alloc_title": "🎯 تخصيص أهداف SP/PY",
        "target_config_sub": "التكوين",
        "target_alloc_tooltip": "تخصيص الأهداف حسب الفرع أو العميل بناءً على المبيعات التاريخية.",
        "target_alloc_type": "اختر نوع تخصيص الهدف",
        "target_alloc_options": ["حسب الفرع", "حسب العميل"],
        "target_input_option": "اختر خيار إدخال الهدف",
        "target_input_options": ["يدوي", "تلقائي (من ورقة 'Target')"],
        "target_enter_total": "أدخل إجمالي الهدف للتخصيص لهذا الشهر (د.ك)",
        "target_auto_info": "استخدام إجمالي الهدف من ورقة 'Target': د.ك {0:,.0f}",
        "target_zero_warning": "يرجى التأكد من أن إجمالي الهدف أكبر من 0.",
        "target_hist_sub": "فترة البيانات التاريخية",
        "target_hist_option": "اختر فترة البيانات التاريخية",
        "target_hist_options": ["آخر 6 أشهر", "أيام يدوية"],
        "target_select_range": "اختر نطاق التاريخ",
        "target_date_warning": "يرجى اختيار تاريخ بداية ونهاية.",
        "target_no_hist": "⚠️ لا توجد بيانات مبيعات متوفرة في 'YTD' لـ {0}.",
        "target_analysis_sub": "🎯 تحليل الهدف",
        "hist_sales_total": "إجمالي المبيعات التاريخية",
        "alloc_target_total": "إجمالي الهدف المخصص",
        "increase_needed": "الزيادة المطلوبة مقابل متوسط المبيعات",
        "current_month_sales": "مبيعات الشهر الحالي",
        "target_balance": "رصيد الهدف",
        "alloc_targets_sub": "📊 الأهداف المخصصة تلقائياً بناءً على {0}",
        "download_alloc": "💾 تنزيل جدول تخصيص الهدف",
        "ai_title": "🤖 رؤى الذكاء الاصطناعي",
        "ai_scope_sub": "النطاق والفلاتر",
        "ai_select_period": "اختر الفترة للرؤى",
        "ai_top_n": "أبرز أفضل N بائعين",
        "ai_no_data": "لا توجد بيانات في الفترة المختارة. جرب توسيع نطاق التاريخ.",
        "ai_exec_sub": "📜 ملخص تنفيذي",
        "ai_prescript_sub": "🛠️ توصيات إرشادية",
        "ai_visual_sub": "📊 رسوم بيانية مولدة بالذكاء الاصطناعي",
        "ai_section_sub": "🧭 رؤى حسب القسم",
        "ai_download_sub": "📥 تنزيل الملخص التنفيذي",
        "ai_download_button": "💾 تنزيل ملخص تنفيذي AI (TXT)",
        "ai_ask_sub": "💬 اسأل سؤالاً عن بياناتك",
        "ai_ask_prompt": "اكتب سؤالاً (مثل 'أي بائع ينمو أسرع؟'، 'أين المرتجعات أعلى؟'، 'الارتباط بين الأهداف والمبيعات؟')",
        "admin_tools": "أدوات الإدارة",
        "view_logs": "عرض سجلات التدقيق",
        "audit_title": "📋 سجلات التدقيق",
        "download_logs": "⬇️ تنزيل سجلات التدقيق (إكسل)",
        "sheet_missing": "❌ ملف الإكسل يجب أن يحتوي على أوراق: {0}. المفقودة: {1}",
        "cols_missing": "❌ أعمدة مطلوبة مفقودة: {0}",
        "load_error": "❌ خطأ في تحميل ملف الإكسل: {0}",
        "pptx_title": "تقرير المبيعات والأهداف",
        "pptx_generated": "تم توليده في {0}",
        "pptx_kpi_title": "📈 مؤشرات الأداء الرئيسية",
        "pptx_summary_title": "📋 ملخص المبيعات والأهداف",
        "pptx_billing_title": "📊 المبيعات حسب نوع الفاتورة لكل بائع",
        "pptx_py_title": "🏬 المبيعات حسب اسم PY 1",
        "pptx_embed_error": "لا يمكن تضمين الرسم البياني: {0}. قم بتثبيت kaleido إذا كان مفقوداً.",
        "generating_pptx": "⏳ جاري توليد تقرير PPTX...",
        "loading_data": "⏳ جاري تحميل بيانات الإكسل...",
    }
}

# --- Page Config ---
st.set_page_config(page_title=texts[lang]["page_title"], layout=texts[lang]["layout"], page_icon=texts[lang]["page_icon"])

# --- Streamlit Authenticator (v0.4.2) ---
# Hash the plain-text passwords
hashed_passwords = [
    stauth.Hasher.hash("admin123"),
    stauth.Hasher.hash("salesman1"),
    stauth.Hasher.hash("salesman2")
]

# Define credentials
credentials = {
    "usernames": {
        "admin": {
            "email": "admin@example.com",
            "name": "Admin User",
            "password": hashed_passwords[0],
            "role": "admin"
        },
        "salesman1": {
            "email": "sales1@example.com",
            "name": "Salesman One",
            "password": hashed_passwords[1],
            "role": "salesman",
            "salesman_name": "Salesman One"
        },
        "salesman2": {
            "email": "sales2@example.com",
            "name": "Salesman Two",
            "password": hashed_passwords[2],
            "role": "salesman",
            "salesman_name": "Salesman Two"
        }
    }
}

# Initialize authenticator
authenticator = stauth.Authenticate(
    credentials,
    cookie_name="sales_app",
    key="auth_key",
    cookie_expiry_days=30
)

# Initialize variables
user_role = None
salesman_name = None
username = None

# --- Login ---
authenticator.login(location="main")

if st.session_state.get("authentication_status"):
    st.success(texts[lang]["welcome"].format(st.session_state["name"]))
    authenticator.logout(texts[lang]["logout"], location="sidebar")
    username = st.session_state["username"]
    user_role = credentials["usernames"][username]["role"]
    salesman_name = credentials["usernames"][username].get("salesman_name", None)

elif st.session_state.get("authentication_status") is False:
    st.error(texts[lang]["incorrect_login"])
    st.stop()

elif st.session_state.get("authentication_status") is None:
    st.warning(texts[lang]["no_login"])
    st.stop()

# --- Custom CSS for Visual Enhancements ---
st.markdown(
    """
    <style>
    /* General layout and typography */
    .main {
        background-color: #F8FAFC;
        padding: 30px;
        border-radius: 12px;
        box-shadow: 0 10px 25px rgba(0, 0, 0, 0.05);
        transition: background-color 0.3s;
    }
    h1, h2, h3 {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        color: #0F172A;
        margin-bottom: 15px;
        font-weight: 700;
    }

    /* Buttons */
    .stButton>button {
        background: linear-gradient(135deg, #1565C0, #1E88E5);
        color: white;
        border: 2px solid #0EA5E9;
        border-radius: 12px;
        padding: 12px 24px;
        font-weight: 700;
        letter-spacing: 0.3px;
        transition: all 0.2s ease;
        box-shadow: 0 6px 18px rgba(2,132,199,0.25);
    }
    .stButton>button:hover {
        transform: translateY(-1px);
        box-shadow: 0 10px 20px rgba(2,132,199,0.35);
    }

    /* Dataframe (tables) */
    .dataframe {
        border-radius: 12px !important;
        overflow: hidden !important;
        box-shadow: 0 4px 16px rgba(0, 0, 0, 0.08);
        border: 1px solid #E2E8F0;
    }
    .dataframe th {
        background: #1E3A8A !important; /* Dark blue for headers */
        color: #FFFFFF !important;
        font-weight: 800 !important;
        padding: 12px !important;
        text-transform: uppercase;
        letter-spacing: 0.4px;
        border: 1px solid #E5E7EB !important;
        line-height: normal !important;
    }
    .dataframe td {
        background-color: #FFFFFF;
        border: 1px solid #E5E7EB !important;
        padding: 10px !important;
        font-weight: 600;
        color: #0F172A;
        vertical-align: middle !important;
        line-height: normal !important;
    }

    /* Sidebar */
    .css-1d391kg {
        background-color: #E2E8F0;
        border-radius: 12px;
        padding: 20px;
        box-shadow: 0 4px 10px rgba(0, 0, 0, 0.05);
    }

    /* Metric card styling with pretty border */
    div[data-testid="stMetric"] {
        border: 2px solid #38BDF8;
        border-radius: 14px;
        padding: 16px 14px;
        background: linear-gradient(180deg, #FFFFFF, #F0F9FF);
        box-shadow: 0 10px 20px rgba(56,189,248,0.25);
        white-space: nowrap;
        overflow: visible;
    }
    div[data-testid="stMetric"] > div {
        color: #0F172A !important;
        font-size: 16px;
    }

    /* Dark mode */
    .dark-mode .main { background-color: #1F2937; }
    .dark-mode h1, .dark-mode h2, .dark-mode h3 { color: #F3F4F6; }
    .dark-mode         line-height: normal !important;
.dataframe td { background-color: #111827; color: #F3F4F6; }
    .dark-mode         line-height: normal !important;
.dataframe th { background: #1E3A8A !important; } /* Dark blue headers in dark mode */
    .dark-mode div[data-testid="stMetric"] { background: linear-gradient(180deg,#111827,#0B1220); border-color:#60A5FA; box-shadow: 0 10px 20px rgba(59,130,246,0.25); }

    /* Tooltip */
    .tooltip { position: relative; display: inline-block; cursor: pointer; }
    .tooltip .tooltiptext {
        visibility: hidden; width: 220px; background-color: #0F172A; color: #fff; text-align: center;
        border-radius: 8px; padding: 8px; position: absolute; z-index: 1; bottom: 125%; left: 50%; margin-left: -110px;
        opacity: 0; transition: opacity 0.3s;
    }
    .tooltip:hover .tooltiptext { visibility: visible; opacity: 1; }

    .progress-bar {
        background-color: #e0e0e0;
        border-radius: 10px;
        margin-top: 5px;
    }
    .progress-bar-fill {
        background-color: #4CAF50;
        height: 15px;
        border-radius: 10px;
        text-align: right;
        padding-right: 5px;
        color: white;
        font-weight: bold;
        transition: width 0.5s ease-in-out;
    }

    /* Green caption styling for specific percentage captions */
    .green-caption {
        color: #15803D !important;
        font-weight: 600;
        font-size: 14px;
    }
    .dark-mode .green-caption {
        color: #6EE7B7 !important; /* Lighter green for dark mode */
    }
    </style>
    """,
    unsafe_allow_html=True
)

# --- Dark Mode Toggle ---
if "dark_mode" not in st.session_state:
    st.session_state.dark_mode = False

st.sidebar.checkbox(
    texts[lang]["dark_mode"],
    value=st.session_state.dark_mode,
    key="dark_mode_toggle",
    on_change=lambda: setattr(st.session_state, "dark_mode", not st.session_state.dark_mode)
)

if st.session_state.dark_mode:
    st.markdown('<script>document.body.classList.add("dark-mode");</script>', unsafe_allow_html=True)
else:
    st.markdown('<script>document.body.classList.remove("dark-mode");</script>', unsafe_allow_html=True)

# --- Cache Data Loading ---
@st.cache_data
def load_data(file):
    with st.spinner(texts[lang]["loading_data"]):
        try:
            xls = pd.ExcelFile(file)
            required_sheets = ["sales data", "Target", "sales channels"]
            missing = [s for s in required_sheets if s not in xls.sheet_names]
            if missing:
                st.error(texts[lang]["sheet_missing"].format(', '.join(required_sheets), ', '.join(missing)))
                return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

            sales_df = pd.read_excel(xls, sheet_name="sales data")
            target_df = pd.read_excel(xls, sheet_name="Target")
            channels_df = pd.read_excel(xls, sheet_name="sales channels")
            ytd_df = pd.read_excel(xls, sheet_name="YTD") if "YTD" in xls.sheet_names else pd.DataFrame()

            required_cols = ["Billing Date", "Driver Name EN", "Net Value", "Billing Type", "PY Name 1", "SP Name1"]
            if not all(col in sales_df.columns for col in required_cols):
                st.error(texts[lang]["cols_missing"].format(set(required_cols) - set(sales_df.columns)))
                return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

            def normalize_series(s):
                try:
                    return s.astype(str).str.strip().str.lower().replace({'nan': ''})
                except Exception:
                    return s

            sales_df["Billing Date"] = pd.to_datetime(sales_df["Billing Date"], errors='coerce')
            if "PY Name 1" in sales_df.columns:
                sales_df["_py_name_norm"] = normalize_series(sales_df["PY Name 1"])
            else:
                sales_df["_py_name_norm"] = ""

            if "PY Name 1" in channels_df.columns:
                channels_df["_py_name_norm"] = normalize_series(channels_df["PY Name 1"])
            else:
                channels_df["_py_name_norm"] = ""
            if "Channels" in channels_df.columns:
                channels_df["_channels_norm"] = normalize_series(channels_df["Channels"])
            else:
                channels_df["_channels_norm"] = ""

            if not ytd_df.empty and "Billing Date" in ytd_df.columns:
                ytd_df["Billing Date"] = pd.to_datetime(ytd_df["Billing Date"], errors='coerce')
            if not ytd_df.empty and "PY Name 1" in ytd_df.columns:
                ytd_df["_py_name_norm"] = normalize_series(ytd_df["PY Name 1"])

            # Hash sensitive columns for privacy (e.g., Driver Name EN)
            def hash_column(col):
                return col.apply(lambda x: hashlib.sha256(str(x).encode()).hexdigest() if pd.notnull(x) else x)

            sales_df["Driver Name EN Hashed"] = hash_column(sales_df["Driver Name EN"])
            if "Driver Name EN" in ytd_df.columns:
                ytd_df["Driver Name EN Hashed"] = hash_column(ytd_df["Driver Name EN"])
            if "Driver Name EN" in target_df.columns:
                target_df["Driver Name EN Hashed"] = hash_column(target_df["Driver Name EN"])

            return sales_df, target_df, ytd_df, channels_df
        except Exception as e:
            st.error(texts[lang]["load_error"].format(e))
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

# --- Helpers: Downloads ---
@st.cache_data
def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Sheet1", index: bool = False) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=index)
    return output.getvalue()

@st.cache_data
def to_multi_sheet_excel_bytes(dfs, sheet_names) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for df, sn in zip(dfs, sheet_names):
            df.to_excel(writer, sheet_name=sn, index=True)
    return output.getvalue()

# --- PPTX Export ---
def create_pptx(report_df, billing_df, py_table, figs_dict, kpi_data, talabat_tables=None):
    with st.spinner(texts[lang]["generating_pptx"]):
        prs = Presentation()
        slide_layout = prs.slide_layouts[0]
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        title.text = texts[lang]["pptx_title"]
        title.text_frame.paragraphs[0].font.size = Pt(32)
        title.text_frame.paragraphs[0].font.name = 'Roboto'
        title.text_frame.paragraphs[0].font.color.rgb = RGBColor(30, 58, 138)
        try:
            subtitle = slide.placeholders[1]
            subtitle.text = texts[lang]["pptx_generated"].format(datetime.now().strftime('%Y-%m-%d'))
            subtitle.text_frame.paragraphs[0].font.size = Pt(18)
            subtitle.text_frame.paragraphs[0].font.name = 'Roboto'
            subtitle.text_frame.paragraphs[0].font.color.rgb = RGBColor(55, 65, 81)
        except Exception:
            pass

        slide_layout = prs.slide_layouts[5]
        slide = prs.slides.add_slide(slide_layout)
        slide.shapes.title.text = texts[lang]["pptx_kpi_title"]
        slide.shapes.title.text_frame.paragraphs[0].font.size = Pt(24)
        slide.shapes.title.text_frame.paragraphs[0].font.name = 'Roboto'
        slide.shapes.title.text_frame.paragraphs[0].font.color.rgb = RGBColor(30, 58, 138)

        rows = 4
        cols = 3
        left = Inches(1)
        top = Inches(1.5)
        width = Inches(8)
        height = Inches(4)
        table = slide.shapes.add_table(rows, cols, left, top, width, height).table
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
                    cell.fill.fore_color.rgb = RGBColor(243, 244, 246)

        def add_table_slide(df, title):
            slide_layout = prs.slide_layouts[5]
            slide = prs.slides.add_slide(slide_layout)
            slide.shapes.title.text = title
            slide.shapes.title.text_frame.paragraphs[0].font.size = Pt(24)
            slide.shapes.title.text_frame.paragraphs[0].font.name = 'Roboto'
            slide.shapes.title.text_frame.paragraphs[0].font.color.rgb = RGBColor(30, 58, 138)
            rows, cols = df.shape
            table = slide.shapes.add_table(rows + 1, cols, Inches(0.5), Inches(1.5), Inches(9), Inches(4)).table
            for j, col in enumerate(df.columns):
                cell = table.cell(0, j)
                cell.text = str(col)
                cell.text_frame.paragraphs[0].font.size = Pt(14)
                cell.text_frame.paragraphs[0].font.name = 'Roboto'
                cell.text_frame.paragraphs[0].font.bold = True
                cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(30, 58, 138)
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
                    texts[lang]["pptx_embed_error"].format(e)
                )

        add_table_slide(report_df.reset_index(), texts[lang]["pptx_summary_title"])
        add_table_slide(billing_df.reset_index(), texts[lang]["pptx_billing_title"])
        add_table_slide(py_table.reset_index(), texts[lang]["pptx_py_title"])

        # --- Talabat tables (optional) ---
        if isinstance(talabat_tables, dict):
            tb = talabat_tables.get("billing_split")
            if tb is not None and hasattr(tb, "empty") and (not tb.empty):
                add_table_slide(tb, "🛵 Talabat – Billing Split")
            tc = talabat_tables.get("customers")
            if tc is not None and hasattr(tc, "empty") and (not tc.empty):
                add_table_slide(tc, "🛵 Talabat – Customer Summary")
        for key, fig in figs_dict.items():
            add_chart_slide(fig, key)
        pptx_stream = io.BytesIO()
        prs.save(pptx_stream)
        pptx_stream.seek(0)
        return pptx_stream

# --- Table Rendering Helpers (consistent headers & full-row highlights) ---
def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    # Remove axis names that can create blank header cells
    try:
        df.rename_axis(None, axis=0, inplace=True)
        df.rename_axis(None, axis=1, inplace=True)
    except Exception:
        pass
    # Ensure all column names are strings (avoid missing/blank headers)
    try:
        df.columns = [("" if c is None else str(c)) for c in df.columns]
    except Exception:
        pass
    return df

def render_table(df, *, formats: dict | None = None, total_row_match=None, hide_index: bool = True):
    '''
    df can be a DataFrame or a pandas Styler.
    formats: dict like {"Sales":"{:,.0f}", "%":"{:.0f}%"}
    total_row_match: function(row)->bool, highlights full row (e.g. lambda r: r.get("Salesman Name")=="Total")
    '''
    try:
        from pandas.io.formats.style import Styler as _Styler
        is_styler = isinstance(df, _Styler)
    except Exception:
        is_styler = False

    if (not is_styler) and isinstance(df, pd.DataFrame):
        df = clean_columns(df)

        sty = df.style.set_table_styles([
            {'selector': 'th', 'props': [
                ('background', '#1E3A8A'),
                ('color', '#FFFFFF'),
                ('font-weight', '800'),
                ('border', '1px solid #E5E7EB'),
                ('text-align', 'center')
            ]}
        ])

        if total_row_match:
            def _hl(row):
                try:
                    if total_row_match(row):
                        return ['background-color: #BFDBFE; color: #1E3A8A; font-weight: 900' for _ in row]
                except Exception:
                    pass
                return ['' for _ in row]
            sty = sty.apply(_hl, axis=1)

        if formats:
            sty = sty.format(formats)

        st.dataframe(sty, use_container_width=True, hide_index=hide_index)
        return

    # If it's already a Styler, just display (best effort)
    st.dataframe(df, use_container_width=True, hide_index=hide_index)

# --- Positive/Negative Coloring ---
def color_positive_negative(val):
    try:
        v = float(val)
        color = "#15803D" if v > 0 else "#B91C1C" if v < 0 else ""
        return f"color: {color}; font-weight: bold"
    except:
        return ""

def create_progress_bar_html(percentage):
    safe_pct = max(0, min(100, percentage))
    fill_color = "#4CAF50" if safe_pct >= 100 else "#2196F3"
    html = f"""
    <div style="background-color: #f0f0f0; border-radius: 5px; height: 10px; margin-top: 5px;">
        <div style="background-color: {fill_color}; height: 100%; width: {safe_pct}%; border-radius: 5px; text-align: right; font-size: 8px; color: white;">
        </div>
    </div>
    """
    return html

# --- SINGLE SIDEBAR UPLOADER ---
st.sidebar.header(texts[lang]["upload_header"])
st.sidebar.markdown(f'<div class="tooltip">ℹ️<span class="tooltiptext">{texts[lang]["upload_tooltip"]}</span></div>', unsafe_allow_html=True)
uploaded = st.sidebar.file_uploader("", type=["xlsx"], key="single_upload")
if st.sidebar.button(texts[lang]["clear_data"]):
    for k in ["sales_df", "target_df", "ytd_df", "channels_df", "data_loaded"]:
        if k in st.session_state:
            del st.session_state[k]
    st.experimental_rerun()

if uploaded is not None and "data_loaded" not in st.session_state:
    sales_df, target_df, ytd_df, channels_df = load_data(uploaded)
    st.session_state["sales_df"] = sales_df
    st.session_state["target_df"] = target_df
    st.session_state["ytd_df"] = ytd_df
    st.session_state["channels_df"] = channels_df
    st.session_state["data_loaded"] = True
    st.session_state["audit_log"] = []  # Initialize audit log
    st.success(texts[lang]["file_loaded"])

    # Log upload action
    st.session_state["audit_log"].append({
        "user": username,
        "action": "upload_file",
        "details": uploaded.name,
        "timestamp": datetime.now().isoformat()
    })

# --- Sidebar Menu with multilingual support ---
st.sidebar.title(texts[lang]["menu_title"])

menu = [
    texts[lang]["home"],
    texts[lang]["sales_tracking"],
    texts[lang]["ytd_comparison"],
    texts[lang]["custom_analysis"],
    texts[lang]["target_allocation"],
    texts[lang]["ai_insights"],
    texts[lang]["customer_insights"],
    texts[lang]["material_forecast"],
    "🧭 Management Command Center"
]

choice = st.sidebar.selectbox(texts[lang]["navigate"], menu)
# Role-based filtering for data
if "data_loaded" in st.session_state:
    sales_df = st.session_state["sales_df"].copy()
    target_df = st.session_state["target_df"].copy()
    ytd_df = st.session_state["ytd_df"].copy()
    channels_df = st.session_state["channels_df"].copy()

    if user_role == "salesman" and salesman_name:
        sales_df = sales_df[sales_df["Driver Name EN"] == salesman_name]
        ytd_df = ytd_df[ytd_df.get("Driver Name EN", pd.Series()) == salesman_name]
        target_df = target_df[target_df.get("Driver Name EN", pd.Series()) == salesman_name]

# --- Home Page ---
if choice == texts[lang]["home"]:
    st.title(texts[lang]["home_title"])
    with st.container():
        st.markdown(
            texts[lang]["home_welcome"],
            unsafe_allow_html=True
        )
    if "data_loaded" in st.session_state: st.success(texts[lang]["data_loaded_msg"])
    else: st.info(texts[lang]["upload_prompt"])


# --- Sales Tracking Page ---
elif choice == texts[lang]["sales_tracking"]:
    st.title(texts[lang]["sales_tracking_title"])
    if "data_loaded" not in st.session_state:
        st.warning(texts[lang]["no_data_warning"])
    else:
        # Filters
        st.sidebar.subheader(texts[lang]["filters_header"])
        st.sidebar.markdown(f'<div class="tooltip">ℹ️<span class="tooltiptext">{texts[lang]["filters_tooltip"]}</span></div>', unsafe_allow_html=True)
        salesmen = st.sidebar.multiselect(
            texts[lang]["select_salesmen"],
            options=sorted(sales_df["Driver Name EN"].dropna().unique()),
            default=sorted(sales_df["Driver Name EN"].dropna().unique()),
            key="st_salesmen"
        )
        billing_types = st.sidebar.multiselect(
            texts[lang]["select_billing_types"],
            options=sorted(sales_df["Billing Type"].dropna().unique()),
            default=sorted(sales_df["Billing Type"].dropna().unique()),
            key="st_billing_types"
        )
        py_filter = st.sidebar.multiselect(
            texts[lang]["select_py"],
            options=sorted(sales_df["PY Name 1"].dropna().unique()),
            default=sorted(sales_df["PY Name 1"].dropna().unique()),
            key="st_py_filter"
        )
        sp_filter = st.sidebar.multiselect(
            texts[lang]["select_sp"],
            options=sorted(sales_df["SP Name1"].dropna().unique()),
            default=sorted(sales_df["SP Name1"].dropna().unique()),
            key="st_sp_filter"
        )

        preset = st.sidebar.radio(texts[lang]["date_presets"], texts[lang]["date_presets_options"], key="st_preset")
        today = pd.Timestamp.today().normalize()
        if preset == texts[lang]["date_presets_options"][1]:  # Last 7 Days
            date_range = [today - pd.Timedelta(days=7), today]
        elif preset == texts[lang]["date_presets_options"][2]:  # This Month
            month_start = today.replace(day=1)
            month_end = month_start + pd.offsets.MonthEnd(0)
            date_range = [month_start, month_end]
        elif preset == texts[lang]["date_presets_options"][3]:  # YTD
            date_range = [today.replace(month=1, day=1), today]
        else:
            date_range = st.sidebar.date_input(
                texts[lang]["select_date_range"],
                [sales_df["Billing Date"].min(), sales_df["Billing Date"].max()],
                key="st_date_range"
            )
            if isinstance(date_range, tuple) and len(date_range) == 2:
                date_range = [pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])]
            else:
                date_range = [sales_df["Billing Date"].min(), sales_df["Billing Date"].max()]

        if date_range[0] > date_range[1]:
            st.error(texts[lang]["date_error"])
        else:
            top_n = st.sidebar.slider(
                texts[lang]["top_n_salesmen"],
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
            ].copy()

            if df_filtered.empty:
                st.warning(texts[lang]["no_match_warning"])
            else:
                billing_start = df_filtered["Billing Date"].min().normalize()
                billing_end = df_filtered["Billing Date"].max().normalize()
                all_days = pd.date_range(billing_start, billing_end, freq="D")
                days_finish = int(sum(1 for d in all_days if d.weekday() != 4))

                current_month_start = today.replace(day=1)
                current_month_end = current_month_start + pd.offsets.MonthEnd(0)
                month_days = pd.date_range(current_month_start, current_month_end, freq="D")
                working_days_current_month = int(sum(1 for d in month_days if d.weekday() != 4))

                # --- Base aggregates ---
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

                # --- Talabat detailed tables (for Tables / Downloads / PPTX) ---
                TALABAT_PY = "STORES SERVICES KUWAIT CO."
                talabat_df_detail = df_filtered[df_filtered["PY Name 1"] == TALABAT_PY].copy()

                def _talabat_group_row(row):
                    bt = str(row.get("Billing Type", "")).strip().upper()
                    try:
                        nv = float(row.get("Net Value", 0) or 0)
                    except Exception:
                        nv = 0.0
                    if bt == "ZFR":
                        return "ZFR"
                    if bt == "HHT":
                        return "HHT"
                    # Returns: negative value OR common return codes / patterns
                    if (nv < 0) or ("RE" in bt) or bt in {"YKF2", "YKRE", "ZRE", "ZCR", "CR"}:
                        return "Returns"
                    return "Other"

                if not talabat_df_detail.empty:
                    talabat_df_detail["Talabat Billing Group"] = talabat_df_detail.apply(_talabat_group_row, axis=1)
                else:
                    talabat_df_detail["Talabat Billing Group"] = pd.Series(dtype=str)

                talabat_billing_split = (
                    talabat_df_detail
                    .pivot_table(
                        index="Driver Name EN",
                        columns="Talabat Billing Group",
                        values="Net Value",
                        aggfunc="sum",
                        fill_value=0
                    )
                    .reindex(all_salesmen_idx, fill_value=0)
                )
                for _c in ["ZFR", "HHT", "Returns", "Other"]:
                    if _c not in talabat_billing_split.columns:
                        talabat_billing_split[_c] = 0.0
                talabat_billing_split = talabat_billing_split[["ZFR", "HHT", "Returns", "Other"]]
                talabat_billing_split["Total"] = talabat_billing_split.sum(axis=1)
                talabat_billing_split = talabat_billing_split.reset_index().rename(columns={"Driver Name EN": "Salesman"})

                def _pick_customer_col(df):
                    for _col in ["Customer Name", "Branch Name", "PY Name 1"]:
                        if _col in df.columns and df[_col].notna().any():
                            return _col
                    # Fallback (if you don't have outlet/branch columns)
                    return "SP Name1" if "SP Name1" in df.columns else "PY Name 1"

                _cust_col = _pick_customer_col(talabat_df_detail)
                if (not talabat_df_detail.empty) and (_cust_col in talabat_df_detail.columns):
                    talabat_customer_table = (
                        talabat_df_detail
                        .groupby(_cust_col, dropna=False)
                        .agg(
                            **{
                                "Talabat Sales": ("Net Value", "sum"),
                                "Orders": ("Net Value", "size"),
                            }
                        )
                        .sort_values("Talabat Sales", ascending=False)
                        .reset_index()
                        .rename(columns={_cust_col: "Customer"})
                    )
                else:
                    talabat_customer_table = pd.DataFrame(columns=["Customer", "Talabat Sales", "Orders"])

                if (not talabat_df_detail.empty) and ("Billing Date" in talabat_df_detail.columns):
                    talabat_df_detail["Talabat Date"] = pd.to_datetime(talabat_df_detail["Billing Date"], errors="coerce").dt.date
                    talabat_daily_trend = (
                        talabat_df_detail
                        .groupby(["Talabat Date", "Talabat Billing Group"])["Net Value"].sum()
                        .reset_index()
                        .pivot_table(
                            index="Talabat Date",
                            columns="Talabat Billing Group",
                            values="Net Value",
                            aggfunc="sum",
                            fill_value=0
                        )
                        .sort_index()
                    )
                    for _c in ["ZFR", "HHT", "Returns", "Other"]:
                        if _c not in talabat_daily_trend.columns:
                            talabat_daily_trend[_c] = 0.0
                    talabat_daily_trend = talabat_daily_trend[["ZFR", "HHT", "Returns", "Other"]]
                    talabat_daily_trend["Total"] = talabat_daily_trend.sum(axis=1)
                    talabat_daily_trend = talabat_daily_trend.reset_index().rename(columns={"Talabat Date": "Date"})
                else:
                    talabat_daily_trend = pd.DataFrame(columns=["Date", "ZFR", "HHT", "Returns", "Other", "Total"])

                total_ka_target_all = float(ka_targets.sum())
                total_tal_target_all = float(talabat_targets.sum())
                per_day_ka_target = (total_ka_target_all / working_days_current_month) if working_days_current_month > 0 else 0
                current_sales_per_day = (total_sales.sum() / days_finish) if days_finish > 0 else 0
                forecast_month_end_ka = current_sales_per_day * working_days_current_month

                # --- Channels mapping: Market (Retail) vs E-com ---
                df_py_sales = df_filtered.groupby("_py_name_norm")["Net Value"].sum().reset_index()
                df_channels_merged = df_py_sales.merge(
                    channels_df[["_py_name_norm", "Channels"]],
                    on="_py_name_norm",
                    how="left"
                )
                df_channels_merged["Channels"] = df_channels_merged["Channels"].str.strip().str.lower().fillna("uncategorized")
                channel_sales = df_channels_merged.groupby("Channels")["Net Value"].sum()
                total_retail_sales = float(channel_sales.get("market", 0.0) + channel_sales.get("uncategorized", 0.0))
                total_ecom_sales = float(channel_sales.get("e-com", 0.0))
                total_channel_sales = total_retail_sales + total_ecom_sales
                retail_sales_pct = (total_retail_sales / total_channel_sales * 100) if total_channel_sales > 0 else 0
                ecom_sales_pct = (total_ecom_sales / total_channel_sales * 100) if total_channel_sales > 0 else 0

                # --- KA & Other E-com Calculation ---
                ka_other_ecom_sales = total_sales.sum() - talabat_sales.sum()
                ka_other_ecom_pct = (ka_other_ecom_sales / total_ka_target_all * 100) if total_ka_target_all > 0 else 0

                # --- KPI Data for PPTX ---
                kpi_data = {
                    texts[lang]["ka_target"]: f"KD {total_ka_target_all:,.0f}",
                    texts[lang]["talabat_target"]: f"KD {total_tal_target_all:,.0f}",
                    texts[lang]["ka_gap"]: f"KD {(total_ka_target_all - total_sales.sum()):,.0f}",
                    "Total Talabat Gap": f"KD {talabat_gap.sum():,.0f}",
                    texts[lang]["total_ka_sales"]: f"KD {total_sales.sum():,.0f} ({((total_sales.sum() / total_ka_target_all) * 100):.0f}%)" if total_ka_target_all else f"KD {total_sales.sum():,.0f} (0%)",
                    "Total Talabat Sales": f"KD {talabat_sales.sum():,.0f} ({((talabat_sales.sum() / total_tal_target_all) * 100):.0f}%)" if total_tal_target_all else f"KD {talabat_sales.sum():,.0f} (0%)",
                    texts[lang]["ka_other_ecom"]: f"KD {ka_other_ecom_sales:,.0f} ({ka_other_ecom_pct:.0f}%)",
                    texts[lang]["retail_sales"]: f"KD {total_retail_sales:,.0f} ({retail_sales_pct:.0f}%)",
                    texts[lang]["ecom_sales"]: f"KD {total_ecom_sales:,.0f} ({ecom_sales_pct:.0f}%)",
                    texts[lang]["days_finished"]: f"{days_finish}",
                    "Per Day KA Target": f"KD {per_day_ka_target:,.0f}",
                    texts[lang]["current_sales_per_day"]: f"KD {current_sales_per_day:,.0f}",
                    texts[lang]["forecast_month_end"]: f"KD {forecast_month_end_ka:,.0f}"
                }

                tabs = st.tabs([texts[lang]["kpis_tab"], texts[lang]["tables_tab"], texts[lang]["charts_tab"], texts[lang]["downloads_tab"]])

                # --- KPIs with progress bars ---
                with tabs[0]:
                    st.subheader(texts[lang]["key_metrics_sub"])
                    r1c1 = st.columns(1)[0]
                    with r1c1:
                        st.metric(texts[lang]["total_ka_sales"], f"KD {total_sales.sum():,.0f}")
                        progress_pct_ka = (total_sales.sum() / total_ka_target_all * 100) if total_ka_target_all > 0 else 0
                        st.markdown(create_progress_bar_html(progress_pct_ka), unsafe_allow_html=True)
                        st.markdown(f'<div class="green-caption">{texts[lang]["of_ka_target"].format(progress_pct_ka)}</div>', unsafe_allow_html=True)

                    r2c1, r2c2 = st.columns(2)
                    with r2c1:
                        st.metric(texts[lang]["ka_other_ecom"], f"KD {ka_other_ecom_sales:,.0f}")
                        st.markdown(create_progress_bar_html(ka_other_ecom_pct), unsafe_allow_html=True)
                        st.markdown(f'<div class="green-caption">{texts[lang]["of_ka_target_pct"].format(ka_other_ecom_pct)}</div>', unsafe_allow_html=True)
                    with r2c2:
                        st.metric(texts[lang]["talabat_sales"], f"KD {talabat_sales.sum():,.0f}")
                        progress_pct_talabat = (talabat_sales.sum() / total_tal_target_all * 100) if total_tal_target_all > 0 else 0
                        st.markdown(create_progress_bar_html(progress_pct_talabat), unsafe_allow_html=True)
                        st.markdown(f'<div class="green-caption">{texts[lang]["of_talabat_target"].format(progress_pct_talabat)}</div>', unsafe_allow_html=True)

                    st.subheader(texts[lang]["target_overview_sub"])
                    r3c1, r3c2, r3c3, r3c4 = st.columns(4)
                    r3c1.metric(texts[lang]["ka_target"], f"KD {total_ka_target_all:,.0f}")
                    r3c2.metric(texts[lang]["talabat_target"], f"KD {total_tal_target_all:,.0f}")
                    r3c3.metric(texts[lang]["ka_gap"], f"KD {(total_ka_target_all - total_sales.sum()):,.0f}")
                    r3c4.metric(texts[lang]["talabat_gap"], f"KD {talabat_gap.sum():,.0f}")

                    st.subheader(texts[lang]["channel_sales_sub"])
                    r4c1, r4c2 = st.columns(2)
                    with r4c1:
                        st.metric(texts[lang]["retail_sales"], f"KD {total_retail_sales:,.0f}")
                        retail_contribution_pct = (total_retail_sales / total_sales.sum() * 100) if total_sales.sum() > 0 else 0
                        st.caption(texts[lang]["of_total_ka"].format(retail_contribution_pct))
                    with r4c2:
                        st.metric(texts[lang]["ecom_sales"], f"KD {total_ecom_sales:,.0f}")
                        ecom_contribution_pct = (total_ecom_sales / total_sales.sum() * 100) if total_sales.sum() > 0 else 0
                        st.caption(texts[lang]["of_total_ka"].format(ecom_contribution_pct))

                    st.subheader(texts[lang]["performance_metrics_sub"])
                    r5c1, r5c2, r5c3 = st.columns(3)
                    r5c1.metric(texts[lang]["days_finished"], days_finish)
                    r5c2.metric(texts[lang]["current_sales_per_day"], f"KD {current_sales_per_day:,.0f}")
                    r5c3.metric(texts[lang]["forecast_month_end"], f"KD {forecast_month_end_ka:,.0f}")

                    # --- TABLES ---
                    with tabs[1]:

                        st.subheader(texts[lang]["sales_targets_summary_sub"])

                        # ================= TOTAL SECTION (KA) =================
                        total_target = ka_targets          # index: Driver Name EN
                        total_sales = total_sales          # same index

                        total_balance = (total_target - total_sales).clip(lower=0)
                        total_percent = np.where(
                            total_target != 0,
                            (total_sales / total_target * 100).round(0),
                            0
                        )

                                                # ================= MERGE SALES + CHANNELS (BY CUSTOMER) =================
                        # Sales sheet: Driver Name EN + PY Name 1 + Net Value
                        df_sales = df_filtered[["Driver Name EN", "PY Name 1", "Net Value"]].copy()

                        # --- Normalize PY names for safer merge (avoids mismatches due to spaces/case) ---
                        df_sales["_py_name_norm"] = (
                            df_sales["PY Name 1"].astype(str).str.upper().str.strip()
                        )

                        ch_tmp = channels_df.copy()
                        if "_py_name_norm" not in ch_tmp.columns:
                            ch_tmp["_py_name_norm"] = ch_tmp["PY Name 1"].astype(str).str.upper().str.strip()

                        df_sales = df_sales.merge(
                            ch_tmp[["_py_name_norm", "Channels"]],
                            on="_py_name_norm",
                            how="left"
                        )

                        df_sales["Channels"] = (
                            df_sales["Channels"]
                            .astype(str)
                            .str.lower()
                            .str.strip()
                        )

                        # Treat empty / missing channel as Market by default
                        df_sales.loc[df_sales["Channels"].isin(["", "nan", "none"]), "Channels"] = "market"

                        # ================= E-COM SALES (BY SALESMAN) =================
                        # Accept multiple variants: e-com, ecom, ecommerce, online, talabat
                        ecom_mask = df_sales["Channels"].str.contains(r"e-?com|ecommerce|online|talabat", regex=True, na=False)

                        ecom_sales_by_sm = (
                            df_sales[ecom_mask]
                            .groupby("Driver Name EN")["Net Value"].sum()
                            .reindex(total_target.index, fill_value=0)
                        )

                        # ================= MARKET SALES (BY SALESMAN) =================
                        market_sales_by_sm = (
                            df_sales[~ecom_mask]
                            .groupby("Driver Name EN")["Net Value"].sum()
                            .reindex(total_target.index, fill_value=0)
                        )

                        # ================= E-COM TARGETS =================
                        # Be robust to slightly different column names in Target sheet
                        ecom_target = pd.Series(0.0, index=total_target.index)
                        if (not target_df.empty) and ("Driver Name EN" in target_df.columns):
                            _col_map = {str(c).strip().lower(): c for c in target_df.columns}
                            _candidates = [
                                "e-com target", "ecom target", "e-commerce target", "e commerce target",
                                "e-com target kd", "ecom target kd", "ecommerce target", "ecomtarget"
                            ]
                            _found = None
                            for _k in _candidates:
                                if _k in _col_map:
                                    _found = _col_map[_k]
                                    break
                            if _found is not None:
                                ecom_target = (
                                    target_df.set_index("Driver Name EN")[_found]
                                    .apply(lambda x: pd.to_numeric(x, errors="coerce"))
                                    .fillna(0)
                                    .reindex(total_target.index, fill_value=0)
                                )

                        # ============== E-COM KPIs =====================
                        ecom_balance = (ecom_target - ecom_sales_by_sm).clip(lower=0)
                        ecom_percent = np.where(
                            ecom_target != 0,
                            (ecom_sales_by_sm / ecom_target * 100).round(0),
                            0
                        )

                        # ============== MARKET KPIs =====================
                        market_target = (total_target - ecom_target).clip(lower=0)
                        market_balance = (market_target - market_sales_by_sm).clip(lower=0)
                        market_percent = np.where(
                            market_target != 0,
                            (market_sales_by_sm / market_target * 100).round(0),
                            0
                        )

                        # ================= FINAL TABLE (ROW LEVEL) ==================
                        report_df = pd.DataFrame({
                            "Salesman Name": total_target.index,

                            "Total Target": total_target.values,
                            "Total Sales": total_sales.values,
                            "Total Balance": total_balance.values,
                            "Total % Achieved": total_percent,

                            "Market Target": market_target.values,
                            "Market Sales": market_sales_by_sm.values,
                            "Market Balance": market_balance.values,
                            "Market % Achieved": market_percent,

                            "E-Com Target": ecom_target.values,
                            "E-Com Sales": ecom_sales_by_sm.values,
                            "E-Com Balance": ecom_balance.values,
                            "E-Com % Achieved": ecom_percent
                        })

                        # ================= TOTAL ROW (RECALC FROM TOTALS) ==================
                        total_row = report_df.sum(numeric_only=True).to_frame().T
                        total_row.index = ["Total"]

                        tt = total_row["Total Target"].iloc[0]
                        ts = total_row["Total Sales"].iloc[0]
                        mt = total_row["Market Target"].iloc[0]
                        ms = total_row["Market Sales"].iloc[0]
                        et = total_row["E-Com Target"].iloc[0]
                        es = total_row["E-Com Sales"].iloc[0]

                        # Balances from GRAND totals
                        total_row["Total Balance"] = max(tt - ts, 0)
                        total_row["Market Balance"] = max(mt - ms, 0)
                        total_row["E-Com Balance"] = max(et - es, 0)

                        # % Achieved from GRAND totals
                        total_row["Total % Achieved"] = round(ts / tt * 100, 0) if tt != 0 else 0
                        total_row["Market % Achieved"] = round(ms / mt * 100, 0) if mt != 0 else 0
                        total_row["E-Com % Achieved"] = round(es / et * 100, 0) if et != 0 else 0

                        total_row["Salesman Name"] = "Total"
                        total_row = total_row[report_df.columns]

                        report_df_with_total = pd.concat([report_df, total_row], ignore_index=True)

                        # ================= SORTING (BEST FIRST, TOTAL LAST) =================
                        data_part = report_df_with_total[report_df_with_total["Salesman Name"] != "Total"].copy()
                        total_part = report_df_with_total[report_df_with_total["Salesman Name"] == "Total"].copy()

                        data_part = data_part.sort_values("Total % Achieved", ascending=False)
                        report_df_sorted = pd.concat([data_part, total_part], ignore_index=True)

                        # ================= STYLING HELPERS =================
                        def zebra_row_style(row):
                            """Zebra rows + highlight Total row."""
                            if row["Salesman Name"] == "Total":
                                return ['background-color: #BFDBFE; color: #1E3A8A; font-weight:900'
                                        for _ in row]
                            return [
                                'background-color: #F9FAFB' if row.name % 2 == 0 else ''
                                for _ in row
                            ]

                        def kpi_color(val):
                            """Green / amber / red for % KPIs."""
                            if pd.isna(val):
                                return ''
                            try:
                                v = float(val)
                            except Exception:
                                return ''
                            if v >= 100:
                                return 'color: #166534; font-weight:700'   # dark green
                            elif v >= 80:
                                return 'color: #92400E; font-weight:600'   # amber
                            else:
                                return 'color: #991B1B; font-weight:700'   # red

                        percent_cols = [
                            "Total % Achieved",
                            "Market % Achieved",
                            "E-Com % Achieved",
                        ]

                        # ================= FINAL STYLED TABLE =================
                        styled_report = (
                            report_df_sorted.style
                            .set_table_styles([
                                {
                                    'selector': 'th',
                                    'props': [
                                        ('background', '#BFDBFE'),
                                        ('color', '#1E3A8A'),
                                        ('font-weight', '900'),
                                        ('height', '40px'),
                                        ('line-height', '40px'),
                                        ('border', '1px solid #E5E7EB')
                                    ]
                                }
                            ])
                            .apply(zebra_row_style, axis=1)
                            .format("{:,.0f}", subset=[
                                "Total Target","Total Sales","Total Balance",
                                "Market Target","Market Sales","Market Balance",
                                "E-Com Target","E-Com Sales","E-Com Balance"
                            ])
                            .format("{:.0f}%", subset=percent_cols)
                            .applymap(kpi_color, subset=percent_cols)
                        )

                        st.dataframe(styled_report, use_container_width=True, hide_index=True)
                        # --- Sales by Billing Type per Salesman ---
                        st.subheader(texts[lang]["sales_by_billing_sub"])
                        billing_wide = df_filtered.pivot_table(
                            index="Driver Name EN",
                            columns="Billing Type",
                            values="Net Value",
                            aggfunc="sum",
                            fill_value=0
                        )

                        required_cols_raw = ["ZFR", "YKF2", "YKRE", "YKS1", "YKS2", "ZCAN", "ZRE"]
                        billing_wide = billing_wide.reindex(columns=required_cols_raw, fill_value=0)
                        display_df = billing_wide.rename(columns={"ZFR": "Presales", "YKF2": "HHT"})
                        display_df["Sales Total"] = billing_wide.sum(axis=1)
                        display_df["Return"] = billing_wide["YKRE"] + billing_wide["ZRE"]
                        display_df["Return %"] = np.where(display_df["Sales Total"] != 0,
                                                        (display_df["Return"] / display_df["Sales Total"] * 100).round(0), 0)
                        display_df["Cancel Total"] = billing_wide[["YKS1", "YKS2", "ZCAN"]].sum(axis=1)

                        ordered_cols = ["Presales", "HHT", "Sales Total", "YKS1", "YKS2", "ZCAN",
                                        "Cancel Total", "YKRE", "ZRE", "Return", "Return %"]
                        display_df = display_df.reindex(columns=ordered_cols, fill_value=0)

                        total_row = pd.DataFrame(display_df.sum(numeric_only=True)).T
                        total_row.index = ["Total"]
                        total_row["Return %"] = round((total_row["Return"] / total_row["Sales Total"] * 100), 0) if total_row["Sales Total"].iloc[0] != 0 else 0
                        billing_df = pd.concat([display_df, total_row])
                        billing_df.index.name = "Salesman"
                        # --- Styling + Display (fixed Total row color & header) ---
                        billing_df_show = billing_df.reset_index()  # brings 'Salesman' as a real column
                        render_table(
                            billing_df_show,
                            hide_index=True,
                            total_row_match=lambda r: str(r.get("Salesman", "")).strip() == "Total",
                            formats={
                                "Presales": "{:,.0f}", "HHT": "{:,.0f}", "Sales Total": "{:,.0f}",
                                "YKS1": "{:,.0f}", "YKS2": "{:,.0f}", "ZCAN": "{:,.0f}", "Cancel Total": "{:,.0f}",
                                "YKRE": "{:,.0f}", "ZRE": "{:,.0f}", "Return": "{:,.0f}", "Return %": "{:.0f}%"
                            }
                        )

                        if st.download_button(
                            texts[lang]["download_billing"],
                            data=to_excel_bytes(billing_df.reset_index(), sheet_name="Billing_Types", index=False),
                            file_name=f"Billing_Types_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        ):
                            st.session_state["audit_log"].append({
                                "user": username,
                                "action": "download",
                                "details": "Billing Type Table Excel",
                                "timestamp": datetime.now()
                            })


                        # --- Sales Summary By Customer – Value ---
                        st.subheader("📌 Sales Summary By Customer – Value")

                        # Sales grouped by Customer
                        py_table = df_filtered.groupby("PY Name 1")["Net Value"].sum().sort_values(ascending=False).to_frame(name="Sales")

                        # Returns grouped by Customer (Billing Type = YKRE or ZRE)
                        returns_df = df_filtered[df_filtered["Billing Type"].isin(["YKRE", "ZRE"])] \
                            .groupby("PY Name 1")["Net Value"].sum()

                        # Add returns into table
                        py_table["Returns"] = py_table.index.map(returns_df).fillna(0.0)

                        # Calculate return %
                        py_table["Return %"] = np.where(
                            py_table["Sales"] > 0,
                            (py_table["Returns"] / py_table["Sales"] * 100).round(1),
                            0
                        )

                        # Contribution %
                        py_table["Contribution %"] = np.where(
                            py_table["Sales"] > 0,
                            (py_table["Sales"] / py_table["Sales"].sum() * 100).round(1),
                            0
                        )

                        # Total Row
                        total_row = pd.DataFrame({
                            "Sales": [py_table["Sales"].sum()],
                            "Returns": [py_table["Returns"].sum()],
                            "Return %": [(py_table["Returns"].sum() / py_table["Sales"].sum() * 100).round(1) if py_table["Sales"].sum() > 0 else 0],
                            "Contribution %": [100.0]
                        }, index=["Total"])

                        py_table_with_total = pd.concat([py_table, total_row])

                        py_table_with_total.index.name = "Customer Name"

                        # Styling function
                        # --- Styling + Display (fixed Total row color & header) ---
                        py_show = py_table_with_total.reset_index()  # brings 'Customer Name' as a real column
                        # Ensure the first column is clearly named
                        if py_show.columns[0] == "":
                            py_show = py_show.rename(columns={py_show.columns[0]: "Customer Name"})
                        render_table(
                            py_show,
                            hide_index=True,
                            total_row_match=lambda r: str(r.get("Customer Name", "")).strip() == "Total",
                            formats={
                                "Sales": "{:,.0f}",
                                "Returns": "{:,.0f}",
                                "Return %": "{:.1f}%",
                                "Contribution %": "{:.1f}%"
                            }
                        )

                        # Download Button
                        if st.download_button(
                            "⬇️ Download Customer Summary (Excel)",
                            data=to_excel_bytes(py_table_with_total.reset_index(), sheet_name="Sales_by_Customer", index=False),
                            file_name=f"Sales_by_Customer_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        ):
                            st.session_state["audit_log"].append({
                                "user": username,
                                "action": "download",
                                "details": "Customer Summary Excel",
                                "timestamp": datetime.now()
                            })


                        # --- Return by SP Name1 ---
                        st.subheader("🔄 Return Summary By Branch-Value")
                        sp_billing = df_filtered.pivot_table(
                            index="SP Name1",
                            columns="Billing Type",
                            values="Net Value",
                            aggfunc="sum",
                            fill_value=0
                        )
                        sp_billing = sp_billing.reindex(columns=required_cols_raw, fill_value=0)
                        sp_billing["Sales Total"] = sp_billing.sum(axis=1)
                        sp_billing["Return"] = sp_billing["YKRE"] + sp_billing["ZRE"]
                        sp_billing["Cancel Total"] = sp_billing[["YKS1", "YKS2", "ZCAN"]].sum(axis=1)
                        sp_billing = sp_billing.rename(columns={"ZFR": "Presales", "YKF2": "HHT"})
                        sp_billing["Return %"] = np.where(sp_billing["Sales Total"] != 0,
                                                        (sp_billing["Return"] / sp_billing["Sales Total"] * 100).round(0), 0)
                        sp_billing = sp_billing.reindex(columns=ordered_cols, fill_value=0).astype(int)

                        total_row = pd.DataFrame(sp_billing.sum(numeric_only=True)).T
                        total_row.index = ["Total"]
                        total_row["Return %"] = round((total_row["Return"]/total_row["Sales Total"]*100), 0) if total_row["Sales Total"].iloc[0]!=0 else 0
                        sp_billing = pd.concat([sp_billing, total_row])
                        # --- Styling + Display (fixed missing header & Total row color) ---
                        sp_billing_show = sp_billing.reset_index().rename(columns={"SP Name1": "Branch Name"})
                        render_table(
                            sp_billing_show,
                            hide_index=True,
                            total_row_match=lambda r: str(r.get("Branch Name", "")).strip() == "Total",
                            formats={
                                "Presales": "{:,.0f}", "HHT": "{:,.0f}", "Sales Total": "{:,.0f}",
                                "YKS1": "{:,.0f}", "YKS2": "{:,.0f}", "ZCAN": "{:,.0f}", "Cancel Total": "{:,.0f}",
                                "YKRE": "{:,.0f}", "ZRE": "{:,.0f}", "Return": "{:,.0f}", "Return %": "{:.0f}%"
                            }
                        )


                        # --- Return by Material Description ---
                        st.subheader("🔄 Return Summary By Product")
                        if "Material Description" in df_filtered.columns:
                            material_billing = df_filtered.pivot_table(
                                index="Material Description",
                                columns="Billing Type",
                                values="Net Value",
                                aggfunc="sum",
                                fill_value=0
                            )

                            material_cols_raw = ["ZFR", "YKF2", "YKRE", "YKS1", "YKS2", "ZCAN", "ZRE"]
                            material_billing = material_billing.reindex(columns=material_cols_raw, fill_value=0)
                            material_billing["Sales Total"] = material_billing.sum(axis=1)
                            material_billing["Return"] = material_billing["YKRE"] + material_billing["ZRE"]
                            material_billing["Cancel Total"] = material_billing[["YKS1", "YKS2", "ZCAN"]].sum(axis=1)
                            material_billing = material_billing.rename(columns={"ZFR": "Presales", "YKF2": "HHT"})
                            material_billing["Return %"] = np.where(material_billing["Sales Total"] != 0,
                                                                    (material_billing["Return"] / material_billing["Sales Total"] * 100).round(0), 0)

                            ordered_cols_material = ["Presales", "HHT", "Sales Total", "YKS1", "YKS2", "ZCAN",
                                                    "Cancel Total", "YKRE", "ZRE", "Return", "Return %"]
                            material_billing = material_billing.reindex(columns=ordered_cols_material, fill_value=0)

                            total_row = pd.DataFrame(material_billing.sum(numeric_only=True)).T
                            total_row.index = ["Total"]
                            total_row["Return %"] = round((total_row["Return"]/total_row["Sales Total"]*100), 0) if total_row["Sales Total"].iloc[0] != 0 else 0
                            material_billing = pd.concat([material_billing, total_row])

                            def highlight_total_row_material(row):
                                return ['background-color: #BFDBFE; color: #1E3A8A; font-weight: 900' if row.name == "Total" else '' for _ in row]

                            styled_material = (
                                material_billing.style
                                .set_table_styles([
                                    {'selector': 'th', 'props': [('background', '#1E3A8A'), ('color', 'white'),
                                                                ('font-weight', '800'), ('height', '40px'),
                                                                ('line-height', '40px'), ('border', '1px solid #E5E7EB')] }
                                ])
                                .apply(highlight_total_row_material, axis=1)
                                .format({
                                    "Presales": "{:,.0f}", "HHT": "{:,.0f}", "Sales Total": "{:,.0f}",
                                    "YKS1": "{:,.0f}", "YKS2": "{:,.0f}", "ZCAN": "{:,.0f}", "Cancel Total": "{:,.0f}",
                                    "YKRE": "{:,.0f}", "ZRE": "{:,.0f}", "Return": "{:,.0f}", "Return %": "{:.0f}%"
                                })
                            )
                            st.dataframe(styled_material, use_container_width=True, hide_index=False)

                            if st.download_button(
                                texts[lang].get("download_material", "Download Return by Material Description"),
                                data=to_excel_bytes(material_billing.reset_index(), sheet_name="Return_by_Material", index=False),
                                file_name=f"Return_by_Material_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            ):
                                st.session_state["audit_log"].append({
                                    "user": username,
                                    "action": "download",
                                    "details": "Return by Material Description Excel",
                                    "timestamp": datetime.now()
                                })
                        else:
                            st.info("No 'Material Description' column found in data — skipping Material Description table.")


                        # --- Return by SP Name1 + Material Description ---
                        st.subheader("🔄 Return Summary By Branch & Product ")
                        required_cols = {"SP Name1", "Material Description", "Billing Type", "Net Value"}
                        if required_cols.issubset(df_filtered.columns):
                            # Pivot table
                            sp_mat_table = pd.pivot_table(
                                df_filtered,
                                index=["SP Name1", "Material Description"],
                                columns="Billing Type",
                                values="Net Value",
                                aggfunc="sum",
                                fill_value=0
                            )

                            # Ensure all billing columns exist
                            billing_cols = ["ZFR", "YKF2", "YKRE", "YKS1", "YKS2", "ZCAN", "ZRE"]
                            for col in billing_cols:
                                if col not in sp_mat_table.columns:
                                    sp_mat_table[col] = 0

                            # Rename for display
                            sp_mat_table = sp_mat_table.rename(columns={"ZFR": "Presales", "YKF2": "HHT"})
                            
                            # Calculate totals
                            sp_mat_table["Sales Total"] = sp_mat_table.sum(axis=1, numeric_only=True)
                            sp_mat_table["Return"] = sp_mat_table["YKRE"] + sp_mat_table["ZRE"]
                            sp_mat_table["Cancel Total"] = sp_mat_table[["YKS1", "YKS2", "ZCAN"]].sum(axis=1)
                            sp_mat_table["Return %"] = np.where(sp_mat_table["Sales Total"] != 0,
                                                                (sp_mat_table["Return"] / sp_mat_table["Sales Total"] * 100).round(0), 0)

                            # Reorder columns
                            ordered_cols = ["Presales", "HHT", "Sales Total", "YKS1", "YKS2", "ZCAN",
                                            "Cancel Total", "YKRE", "ZRE", "Return", "Return %"]
                            sp_mat_table = sp_mat_table.reindex(columns=ordered_cols, fill_value=0)

                            # Add total row
                            total_row = pd.DataFrame(sp_mat_table.sum(numeric_only=True)).T
                            total_row.index = [("Total", "")]
                            total_row["Return %"] = round((total_row["Return"] / total_row["Sales Total"] * 100), 0) if total_row["Sales Total"].iloc[0]!=0 else 0
                            sp_mat_table = pd.concat([sp_mat_table, total_row])

                            # Highlighting function
                            def highlight_sp_mat(row):
                                styles = []
                                for col in row.index:
                                    if row.name == ("Total", ""):
                                        styles.append('background-color: #BFDBFE; color: #1E3A8A; font-weight: 900')
                                    elif col == "Return" and row[col] > 0:
                                        styles.append('background-color: #FECACA; color: #991B1B; font-weight: 700')  # highlight returns
                                    elif col == "Cancel Total" and row[col] > 0:
                                        styles.append('background-color: #FDE68A; color: #92400E; font-weight: 700')  # highlight cancels
                                    elif col == "Sales Total" and row[col] > 0:
                                        styles.append('background-color: #D1FAE5; color: #065F46; font-weight: 700')  # highlight sales
                                    else:
                                        styles.append('')
                                return styles

                            # Style the table
                            styled_sp_mat = (
                                sp_mat_table.style
                                .set_table_styles([
                                    {'selector': 'th', 'props': [('background', '#1E3A8A'), ('color', 'white'),
                                                                ('font-weight', '800'), ('height', '40px'),
                                                                ('line-height', '40px'), ('border', '1px solid #E5E7EB')] }
                                ])
                                .apply(highlight_sp_mat, axis=1)
                                .format({
                                    "Presales": "{:,.0f}", "HHT": "{:,.0f}", "Sales Total": "{:,.0f}",
                                    "YKS1": "{:,.0f}", "YKS2": "{:,.0f}", "ZCAN": "{:,.0f}", "Cancel Total": "{:,.0f}",
                                    "YKRE": "{:,.0f}", "ZRE": "{:,.0f}", "Return": "{:,.0f}", "Return %": "{:.0f}%"
                                })
                            )

                            st.dataframe(styled_sp_mat, use_container_width=True, hide_index=True)

                            # Download button
                            if st.download_button(
                                texts[lang].get("download_sp_material", "Download Return by SP+Material"),
                                data=to_excel_bytes(sp_mat_table.reset_index(), sheet_name="Return_by_SP_Material", index=False),
                                file_name=f"Return_by_SP_Material_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            ):
                                st.session_state["audit_log"].append({
                                    "user": username,
                                    "action": "download",
                                    "details": "Return by SP+Material Excel",
                                    "timestamp": datetime.now()
                                })
                        else:
                            st.info("Required columns are missing in your data for SP+Material table.")
                            
                            


                        # ------------------------------------------------
                        # 🛵 Talabat – MTD details (Billing split / Customers / Daily trend + Excel)
                        # Place this inside: with tabs[1]:
                        # ------------------------------------------------

                        TALABAT_PY = "STORES SERVICES KUWAIT CO."

                        talabat_only_df = df_filtered[df_filtered["PY Name 1"] == TALABAT_PY].copy()

                        # Defaults (so display won't break)
                        talabat_billing_split = pd.DataFrame()
                        talabat_customer_table = pd.DataFrame()
                        talabat_daily_trend = pd.DataFrame()

                        if not talabat_only_df.empty:

                            # ---- Billing Type Group Mapping (FIXED) ----
                            SALES_ZFR = {"ZFR"}
                            SALES_HHT = {"YKF2"}              # HHT in your app = YKF2
                            RETURNS_CODES = {"YKRE", "ZRE"}   # ONLY these are returns

                            def _bt_group(bt):
                                bt = str(bt).strip().upper()
                                if bt in SALES_ZFR:
                                    return "ZFR"
                                if bt in SALES_HHT:
                                    return "HHT"
                                if bt in RETURNS_CODES:
                                    return "Returns"
                                return "Other"

                            talabat_only_df["_bt_group"] = talabat_only_df["Billing Type"].apply(_bt_group)

                            # =========================
                            # 1) Billing split by Salesman
                            # =========================
                            talabat_billing_split = (
                                talabat_only_df
                                .groupby(["Driver Name EN", "_bt_group"])["Net Value"]
                                .sum()
                                .unstack(fill_value=0)
                            )

                            for c in ["ZFR", "HHT", "Returns", "Other"]:
                                if c not in talabat_billing_split.columns:
                                    talabat_billing_split[c] = 0

                            talabat_billing_split = talabat_billing_split[["ZFR", "HHT", "Returns", "Other"]]
                            talabat_billing_split["Total"] = talabat_billing_split.sum(axis=1)
                            talabat_billing_split = talabat_billing_split.reset_index().rename(columns={"Driver Name EN": "Salesman"})

                            # Total row
                            talabat_billing_split = pd.concat([
                                talabat_billing_split,
                                pd.DataFrame([{
                                    "Salesman": "Total",
                                    "ZFR": talabat_billing_split["ZFR"].sum(),
                                    "HHT": talabat_billing_split["HHT"].sum(),
                                    "Returns": talabat_billing_split["Returns"].sum(),
                                    "Other": talabat_billing_split["Other"].sum(),
                                    "Total": talabat_billing_split["Total"].sum(),
                                }])
                            ], ignore_index=True)

                            # =========================
                            # 2) Customer / Outlet summary
                            # =========================
                            candidate_customer_cols = [
                                "Customer", "Customer Name", "Outlet", "Outlet Name", "Branch Name",
                                "Ship-to Name", "Sold-to Name", "SP Name1", "PY Name 1"
                            ]
                            customer_col = next((c for c in candidate_customer_cols if c in talabat_only_df.columns), None)

                            candidate_order_cols = ["Billing Document", "Invoice", "Invoice No", "Sales Document", "Document No"]
                            order_col = next((c for c in candidate_order_cols if c in talabat_only_df.columns), None)

                            if customer_col:
                                if order_col:
                                    orders_series = talabat_only_df.groupby(customer_col)[order_col].nunique()
                                else:
                                    orders_series = talabat_only_df.groupby(customer_col).size()

                                talabat_customer_table = (
                                    talabat_only_df.groupby(customer_col)["Net Value"].sum()
                                    .to_frame("Talabat Sales")
                                    .join(orders_series.to_frame("Orders"))
                                    .reset_index()
                                    .rename(columns={customer_col: "Customer"})
                                    .sort_values("Talabat Sales", ascending=False)
                                )

                                talabat_customer_table = pd.concat([
                                    talabat_customer_table,
                                    pd.DataFrame([{
                                        "Customer": "Total",
                                        "Talabat Sales": talabat_customer_table["Talabat Sales"].sum(),
                                        "Orders": talabat_customer_table["Orders"].sum()
                                    }])
                                ], ignore_index=True)

                            # =========================
                            # 3) Daily trend (MTD)
                            # =========================
                            if "Billing Date" in talabat_only_df.columns:
                                talabat_only_df["Billing Date"] = pd.to_datetime(talabat_only_df["Billing Date"], errors="coerce")

                                daily = (
                                    talabat_only_df
                                    .dropna(subset=["Billing Date"])
                                    .groupby([talabat_only_df["Billing Date"].dt.date, "_bt_group"])["Net Value"]
                                    .sum()
                                    .unstack(fill_value=0)
                                )

                                for c in ["ZFR", "HHT", "Returns", "Other"]:
                                    if c not in daily.columns:
                                        daily[c] = 0

                                daily = daily[["ZFR", "HHT", "Returns", "Other"]]
                                daily["Total"] = daily.sum(axis=1)
                                talabat_daily_trend = daily.reset_index().rename(columns={"Billing Date": "Date"})
                                talabat_daily_trend.columns = ["Date"] + [c for c in talabat_daily_trend.columns if c != "Date"]


                        # ------------------------------------------------
                        # 🛵 Talabat – DISPLAY (your same style)
                        # ------------------------------------------------
                        st.markdown("---")
                        st.subheader("🛵 Talabat – Billing Type Split (ZFR / HHT / Returns)")

                        if isinstance(talabat_billing_split, pd.DataFrame) and (not talabat_billing_split.empty):
                            st.dataframe(
                                talabat_billing_split.style.format({
                                    "ZFR": "{:,.0f}",
                                    "HHT": "{:,.0f}",
                                    "Returns": "{:,.0f}",
                                    "Other": "{:,.0f}",
                                    "Total": "{:,.0f}",
                                }),
                                use_container_width=True,
                                hide_index=True
                            )
                        else:
                            st.info("No Talabat data found in the selected filters.")

                        st.subheader("🛵 Talabat – Customer / Outlet Summary")
                        if isinstance(talabat_customer_table, pd.DataFrame) and (not talabat_customer_table.empty):
                            st.dataframe(
                                talabat_customer_table.style.format({
                                    "Talabat Sales": "{:,.0f}",
                                    "Orders": "{:,.0f}",
                                }),
                                use_container_width=True,
                                hide_index=True
                            )
                        else:
                            st.info("No Talabat customer-level data available (missing customer/outlet columns or no rows).")

                        st.subheader("🛵 Talabat – Daily Trend (MTD)")
                        if isinstance(talabat_daily_trend, pd.DataFrame) and (not talabat_daily_trend.empty):
                            st.dataframe(
                                talabat_daily_trend.style.format({
                                    "ZFR": "{:,.0f}",
                                    "HHT": "{:,.0f}",
                                    "Returns": "{:,.0f}",
                                    "Other": "{:,.0f}",
                                    "Total": "{:,.0f}",
                                }),
                                use_container_width=True,
                                hide_index=True
                            )

                            # Simple line chart (Total)
                            try:
                                fig_tal_daily = px.line(
                                    talabat_daily_trend,
                                    x="Date",
                                    y="Total",
                                    markers=True,
                                    title="Talabat Daily Sales (Total)"
                                )
                                st.plotly_chart(fig_tal_daily, use_container_width=True)
                            except Exception:
                                pass
                        else:
                            st.info("No Talabat daily trend data available.")


                        # ------------------------------------------------
                        # ⬇️ Talabat Excel Download (3 sheets)
                        # Uses your helper: to_multi_sheet_excel_bytes(dfs, sheet_names)
                        # ------------------------------------------------
                        if (isinstance(talabat_billing_split, pd.DataFrame) and not talabat_billing_split.empty) or \
                        (isinstance(talabat_customer_table, pd.DataFrame) and not talabat_customer_table.empty) or \
                        (isinstance(talabat_daily_trend, pd.DataFrame) and not talabat_daily_trend.empty):

                            dfs = [
                                talabat_billing_split.set_index("Salesman") if (not talabat_billing_split.empty and "Salesman" in talabat_billing_split.columns) else pd.DataFrame(),
                                talabat_customer_table.set_index("Customer") if (not talabat_customer_table.empty and "Customer" in talabat_customer_table.columns) else pd.DataFrame(),
                                talabat_daily_trend.set_index("Date") if (not talabat_daily_trend.empty and "Date" in talabat_daily_trend.columns) else pd.DataFrame(),
                            ]
                            sheet_names = ["Talabat_Billing_Split", "Talabat_Customers", "Talabat_Daily_Trend"]

                            tal_excel = to_multi_sheet_excel_bytes(dfs, sheet_names)

                            st.download_button(
                                "⬇️ Download Talabat Details (Excel)",
                                data=tal_excel,
                                file_name=f"Talabat_MTD_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )


                        # Download Talabat Excel (3 sheets)
                        talabat_xlsx_bytes = to_multi_sheet_excel_bytes(
                            dfs=[
                                talabat_billing_split,
                                talabat_customer_table,
                                talabat_daily_trend,
                            ],
                            sheet_names=[
                                "Talabat_Billing_Split",
                                "Talabat_Customers",
                                "Talabat_Daily_Trend",
                            ]
                        )

                        if st.download_button(
                            "⬇️ Download Talabat Details (Excel)",
                            data=talabat_xlsx_bytes,
                            file_name=f"Talabat_MTD_Details_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        ):
                            st.session_state["audit_log"].append({
                                "user": username,
                                "action": "download",
                                "details": "Talabat MTD Details (Excel)",
                                "timestamp": datetime.now()
                            })

                    # ================================
                    # 📊 CHARTS TAB (GM Premium Visuals)
                    # ================================
                    with tabs[2]:

                        # 🔹 Section Title
                        st.subheader("📈 Sales Trends & Channel Performance – Manager Dashboard")

                        # ----------------------------------
                        # 1️⃣ Prepare Channel Mapping for Charts
                        # ----------------------------------
                        df_channel_temp = df_filtered.groupby("PY Name 1")["Net Value"].sum().reset_index()

                        df_ch_merge = df_channel_temp.merge(
                            channels_df[["PY Name 1", "Channels"]],
                            on="PY Name 1",
                            how="left"
                        )

                        df_ch_merge["Channels"] = (
                            df_ch_merge["Channels"]
                            .astype(str).str.strip().str.lower()
                            .replace({"": "market", "nan": "market"})
                        )

                        total_ecom_sales = float(df_ch_merge[df_ch_merge["Channels"] == "e-com"]["Net Value"].sum())
                        total_market_sales = float(df_ch_merge[df_ch_merge["Channels"] != "e-com"]["Net Value"].sum())
                        total_channel_sales = total_market_sales + total_ecom_sales

                        # ----------------------------------
                        # 2️⃣ Daily Sales Trend + Forecast + Anomaly
                        # ----------------------------------
                        st.markdown("### 📌 Daily Sales Trend + Forecast")

                        df_time = df_filtered.groupby(pd.Grouper(key="Billing Date", freq="D"))["Net Value"].sum().reset_index()
                        df_time.rename(columns={"Billing Date": "ds", "Net Value": "y"}, inplace=True)

                        if len(df_time) > 2:
                            m = Prophet()
                            m.fit(df_time)
                            future = m.make_future_dataframe(periods=30)
                            forecast = m.predict(future)

                            # Anomaly detection
                            df_time["median"] = df_time["y"].rolling(7).median()
                            df_time["mdev"] = abs(df_time["y"] - df_time["median"]).rolling(7).median()
                            df_time["anomaly"] = np.where(abs(df_time["y"] - df_time["median"]) > 2*df_time["mdev"], df_time["y"], np.nan)

                            fig_trend = go.Figure()
                            fig_trend.add_trace(go.Scatter(
                                x=df_time["ds"], y=df_time["y"],
                                name="Actual Sales",
                                mode="lines+markers",
                                line=dict(color="#1E3A8A", width=3)
                            ))
                            fig_trend.add_trace(go.Scatter(
                                x=forecast["ds"], y=forecast["yhat"],
                                name="Forecast",
                                line=dict(color="#22C55E", width=2, dash="dash")
                            ))
                            fig_trend.add_trace(go.Scatter(
                                x=df_time["ds"], y=df_time["anomaly"],
                                name="Anomaly",
                                mode="markers",
                                marker=dict(color="red", size=12, symbol="x")
                            ))

                            fig_trend.update_layout(
                                xaxis_title="Date",
                                yaxis_title="Net Value (KD)",
                                hovermode="x unified",
                                template="plotly_white"
                            )
                            st.plotly_chart(fig_trend, use_container_width=True)
                        else:
                            st.info("Not enough data to generate trend.")

                        # ----------------------------------
                        # 🛒 Market vs E-com Performance (Sales + Share %)
                        # ----------------------------------
                        st.markdown("### 🛒 Market vs E-com Performance (Sales + Share %)")

                        # Total Sales KPI Display Above Chart
                        st.metric(
                            label="Total Channel Sales",
                            value=f"KD {total_channel_sales:,.0f}"
                        )

                        fig_market = make_subplots(
                            rows=1, cols=2,
                            specs=[[{"type": "bar"}, {"type": "pie"}]],
                            column_widths=[0.55, 0.45],
                            horizontal_spacing=0.08
                        )

                        # Bar chart (value view)
                        fig_market.add_trace(
                            go.Bar(
                                x=["Market", "E-com", "TOTAL"],
                                y=[total_market_sales, total_ecom_sales, total_channel_sales],
                                text=[f"KD {total_market_sales:,.0f}", 
                                    f"KD {total_ecom_sales:,.0f}", 
                                    f"KD {total_channel_sales:,.0f}"],
                                textposition="outside",
                                marker=dict(
                                    color=["#0EA5E9", "#A78BFA", "#22C55E"],
                                    line=dict(color="black", width=1)
                                ),
                                name="KD Value"
                            ),
                            row=1, col=1
                        )

                        # Pie chart (% share view)
                        fig_market.add_trace(
                            go.Pie(
                                labels=["Market", "E-com"],
                                values=[total_market_sales, total_ecom_sales],
                                hole=0.55,
                                textinfo="percent+label",
                                marker=dict(colors=["#0EA5E9", "#A78BFA"]),
                                name="Share %"
                            ),
                            row=1, col=2
                        )

                        # Center label in donut
                        fig_market.update_layout(
                            annotations=[
                                dict(
                                    text=f"KD<br>{total_channel_sales:,.0f}",
                                    x=0.86,
                                    y=0.5,
                                    showarrow=False,
                                    font=dict(size=15, color="black")
                                )
                            ],
                            template="plotly_white",
                            showlegend=False
                        )

                        fig_market.update_layout(
                            title="Channel Value + % Contribution",
                            xaxis_title="Channel",
                            yaxis_title="KD Value",
                        )

                        st.plotly_chart(fig_market, use_container_width=True)

                        # ----------------------------------
                        # 4️⃣ Daily KA Target vs Actual
                        # ----------------------------------
                        st.markdown("### 🎯 Daily KA Target vs Actual Sales")

                        df_daily = df_filtered.groupby(pd.Grouper(key="Billing Date", freq="D"))["Net Value"].sum().reset_index()
                        df_daily.rename(columns={"Billing Date": "Date", "Net Value": "Sales"}, inplace=True)
                        df_daily["Daily KA Target"] = per_day_ka_target

                        fig_target = go.Figure()
                        fig_target.add_trace(go.Scatter(
                            x=df_daily["Date"], y=df_daily["Sales"],
                            name="Sales", mode="lines+markers",
                            line=dict(color="#16A34A", width=3)
                        ))
                        fig_target.add_trace(go.Scatter(
                            x=df_daily["Date"], y=df_daily["Daily KA Target"],
                            name="Target", mode="lines",
                            line=dict(color="#FACC15", width=2, dash="dot")
                        ))
                        fig_target.update_layout(
                            xaxis_title="Date",
                            yaxis_title="Net Value (KD)",
                            hovermode="x unified",
                            template="plotly_white"
                        )
                        st.plotly_chart(fig_target, use_container_width=True)
                        st.markdown("---")
                        st.subheader("💪 Salesman KA Target vs Actual")

                        # Detect Salesman column dynamically
                        salesman_col = None
                        for c in ["Driver Name EN", "Salesman", "Sales Rep", "Salesperson"]:
                            if c in df_filtered.columns:
                                salesman_col = c
                                break

                        if salesman_col is None:
                            st.error("⚠️ Salesman column not found!")
                        else:
                            # Sales by Salesman (Filtered period)
                            sales_by_sm = df_filtered.groupby(salesman_col)["Net Value"].sum()

                            # KA Target aligned with Salesman list
                            target_df_aligned = target_df.set_index(salesman_col)
                            ka_targets_full = sales_by_sm.reindex(target_df_aligned.index, fill_value=0)
                            
                            # Summary Table
                            salesman_data = pd.DataFrame({
                                "Salesman": sales_by_sm.index,
                                "KA Sales": sales_by_sm.values,
                                "KA Target": target_df_aligned["KA Target"].values
                            })

                            # Order by performance
                            salesman_data.sort_values("KA Sales", ascending=False, inplace=True)

                            fig_salesman_new = go.Figure()

                            # Actual Sales Bar
                            fig_salesman_new.add_trace(go.Bar(
                                x=salesman_data["Salesman"],
                                y=salesman_data["KA Sales"],
                                name="Actual Sales",
                                marker=dict(
                                    color=[
                                        "#10B981" if s >= t else "#EF4444"
                                        for s, t in zip(salesman_data["KA Sales"], salesman_data["KA Target"])
                                    ],
                                    line=dict(color="black", width=1)
                                ),
                                text=[f"KD {v:,.0f}" for v in salesman_data["KA Sales"]],
                                textposition="outside"
                            ))

                            # Target Line
                            fig_salesman_new.add_trace(go.Scatter(
                                x=salesman_data["Salesman"],
                                y=salesman_data["KA Target"],
                                name="Target",
                                mode="lines+markers",
                                line=dict(color="#1E3A8A", width=2)
                            ))

                            fig_salesman_new.update_layout(
                                title="Salesman KA Target vs Actual",
                                xaxis_title="Salesman",
                                yaxis_title="Net Value (KD)",
                                hovermode="x unified",
                                template="plotly_white"
                            )

                            st.plotly_chart(fig_salesman_new, use_container_width=True)


                        # ----------------------------------
                        # 5️⃣ Top 10 Customers Chart
                        # ----------------------------------
                        st.markdown("### 🏆 Top 10 Customers by Sales")

                        top10 = df_filtered.groupby("PY Name 1")["Net Value"].sum().sort_values(ascending=False).head(10)

                        fig_top10 = go.Figure(go.Bar(
                            x=top10.index,
                            y=top10.values,
                            text=[f"KD {v:,.0f}" for v in top10.values],
                            textposition="outside",
                            marker=dict(color="#1D4ED8")
                        ))

                        fig_top10.update_layout(
                            xaxis_title="Customer",
                            yaxis_title="Net Value (KD)",
                            template="plotly_white"
                        )
                        st.plotly_chart(fig_top10, use_container_width=True)

                    # --- DOWNLOADS ---
                    with tabs[3]:
                        st.subheader(texts[lang]["download_reports_sub"])
                        col1, col2 = st.columns(2)

                        # --------------------------
                        # COL 1: PPTX GENERATION
                        # --------------------------
                        with col1:
                            if st.button(texts[lang]["generate_pptx"]):

                                # ✅ FIX: Build figs_dict safely (avoid NameError if any fig not created)
                                figs_dict = {}

                                if "fig_trend_new" in globals() and fig_trend_new is not None:
                                    figs_dict["Daily Sales Trend"] = fig_trend_new

                                if "fig_channel_new" in globals() and fig_channel_new is not None:
                                    figs_dict["Market vs E-com Sales"] = fig_channel_new

                                if "fig_target_new" in globals() and fig_target_new is not None:
                                    figs_dict["Daily KA Target vs Actual"] = fig_target_new

                                if "fig_salesman_new" in globals() and fig_salesman_new is not None:
                                    figs_dict["Salesman KA Target vs Actual"] = fig_salesman_new

                                if "fig_top10_new" in globals() and fig_top10_new is not None:
                                    figs_dict["Top 10 Customers by Sales"] = fig_top10_new

                                talabat_ppt_tables = {
                                    "billing_split": talabat_billing_split if "talabat_billing_split" in locals() else pd.DataFrame(),
                                    "customers": talabat_customer_table.head(25) if ("talabat_customer_table" in locals() and isinstance(talabat_customer_table, pd.DataFrame)) else pd.DataFrame(),
                                }

                                pptx_stream = create_pptx(
                                    report_df_with_total,
                                    billing_df,
                                    py_table_with_total,
                                    figs_dict,
                                    kpi_data,
                                    talabat_tables=talabat_ppt_tables
                                )

                                st.download_button(
                                    texts[lang]["download_pptx"],
                                    pptx_stream,
                                    file_name=f"sales_report_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.pptx",
                                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                                )

                                # Audit safe
                                if "audit_log" in st.session_state:
                                    st.session_state["audit_log"].append({
                                        "user": username,
                                        "action": "download",
                                        "details": "PPTX Report",
                                        "timestamp": datetime.now()
                                    })

                        # --------------------------
                        # COL 2: EXCEL DOWNLOADS
                        # --------------------------
                        with col2:
                            # ==============================
                            # 🎯 KA Target Daily Sheet (Full Month) - DOWNLOAD
                            # ==============================
                            st.markdown("### 🎯 KA Target – Daily Sheet (Full Month)")

                            # Safety checks (avoid crash if variables missing)
                            if "df_filtered" not in locals() or not isinstance(df_filtered, pd.DataFrame) or df_filtered.empty:
                                st.info("No data available to generate KA Target Daily file.")
                            else:
                                # Build month days
                                month_start_ds = current_month_start
                                month_end_ds = current_month_end
                                days_ds = pd.date_range(month_start_ds, month_end_ds, freq="D")

                                # Completed days: up to today (future days stay blank)
                                cutoff = min(today, month_end_ds)

                                # Prepare working df for daily totals
                                df_ds = df_filtered.copy()
                                df_ds["__date"] = pd.to_datetime(df_ds["Billing Date"], errors="coerce").dt.normalize()

                                # --- Sales type masks (adjust if your billing codes differ) ---
                                hht_mask = df_ds["Billing Type"].astype(str).str.upper().isin(["YKF2", "HHT"])
                                presales_mask = df_ds["Billing Type"].astype(str).str.upper().isin(["ZFR", "PRESALES"])
                                tal_mask = df_ds["PY Name 1"].astype(str).str.strip().str.upper() == str(TALABAT_PY).strip().upper()

                                # Group daily sums
                                g_hht = df_ds.loc[hht_mask].groupby(["Driver Name EN", "__date"])["Net Value"].sum()
                                g_pre = df_ds.loc[presales_mask].groupby(["Driver Name EN", "__date"])["Net Value"].sum()
                                g_tal = df_ds.loc[tal_mask].groupby(["Driver Name EN", "__date"])["Net Value"].sum()

                                def _get(g, sm, d):
                                    try:
                                        return float(g.get((sm, d), 0.0))
                                    except Exception:
                                        return 0.0

                                # Day columns: Date 1..Date N (full month)
                                day_cols = [f"Date {i}" for i in range(1, len(days_ds) + 1)]

                                # Salesmen list
                                all_salesmen_month = sorted(df_ds["Driver Name EN"].dropna().unique().tolist())

                                rows = []
                                for sm in all_salesmen_month:
                                    ka_t = float(ka_targets.get(sm, 0.0)) if "ka_targets" in locals() else 0.0

                                    # Achieved KA (exclude Talabat)
                                    sm_total = float(df_ds.loc[df_ds["Driver Name EN"] == sm, "Net Value"].sum())
                                    sm_tal = float(df_ds.loc[(df_ds["Driver Name EN"] == sm) & tal_mask, "Net Value"].sum())
                                    achieved_ka = sm_total - sm_tal

                                    per_day_target = (ka_t / working_days_current_month) if ("working_days_current_month" in locals() and working_days_current_month) else 0.0
                                    current_sales_per_day = (achieved_ka / days_finish) if ("days_finish" in locals() and days_finish) else 0.0
                                    balance_vs_target = ka_t - achieved_ka

                                    # Monthly totals per type
                                    sum_hht = float(df_ds.loc[(df_ds["Driver Name EN"] == sm) & hht_mask, "Net Value"].sum())
                                    sum_pre = float(df_ds.loc[(df_ds["Driver Name EN"] == sm) & presales_mask, "Net Value"].sum())
                                    sum_tot = float(df_ds.loc[(df_ds["Driver Name EN"] == sm) & (hht_mask | presales_mask), "Net Value"].sum())
                                    sum_tal2 = float(df_ds.loc[(df_ds["Driver Name EN"] == sm) & tal_mask, "Net Value"].sum())

                                    def _add_row(label, daily_fn, summary_val):
                                        row = {
                                            "KA Target": "KA Target",
                                            "KA Target Value": round(ka_t, 0),
                                            "Per day Target (Total target / Working days)": round(per_day_target, 0),
                                            "Achieved value": round(achieved_ka, 0),
                                            "Current Sales Per day": round(current_sales_per_day, 0),
                                            "Salesman Name": sm,
                                            "Sales Type": label,
                                            "Sales Summary": round(summary_val, 0),
                                            "Balance": round(balance_vs_target, 0),
                                        }
                                        for idx_d, d in enumerate(days_ds, start=1):
                                            col = f"Date {idx_d}"
                                            row[col] = round(daily_fn(sm, d), 0) if d <= cutoff else ""
                                        rows.append(row)

                                    _add_row("HHT", lambda s, d: _get(g_hht, s, d), sum_hht)
                                    _add_row("PRESALES", lambda s, d: _get(g_pre, s, d), sum_pre)
                                    _add_row("Sales Total", lambda s, d: (_get(g_hht, s, d) + _get(g_pre, s, d)), sum_tot)
                                    _add_row("Talabat Sales", lambda s, d: _get(g_tal, s, d), sum_tal2)

                                ka_daily_df = pd.DataFrame(rows)

                                # Order columns nicely
                                base_cols = [
                                    "KA Target", "KA Target Value",
                                    "Per day Target (Total target / Working days)",
                                    "Achieved value", "Current Sales Per day",
                                    "Salesman Name", "Sales Type", "Sales Summary", "Balance"
                                ]
                                ka_daily_df = ka_daily_df.reindex(columns=base_cols + day_cols)

                                ka_daily_bytes = to_excel_bytes(ka_daily_df, sheet_name="KA_Target_Daily", index=False)

                                if st.download_button(
                                    "⬇️ Download KA Target Daily (Excel)",
                                    data=ka_daily_bytes,
                                    file_name=f"KA_Target_Daily_{datetime.now().strftime('%Y-%m')}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                ):
                                    if "audit_log" in st.session_state:
                                        st.session_state["audit_log"].append({
                                            "user": username,
                                            "action": "download",
                                            "details": "KA Target Daily (Excel)",
                                            "timestamp": datetime.now()
                                        })



elif choice == texts[lang]["sales_tracking"]:
    st.title(texts[lang]["sales_tracking_title"])
    if "data_loaded" not in st.session_state:
        st.warning(texts[lang]["no_data_warning"])
    else:
        # Filters
        st.sidebar.subheader(texts[lang]["filters_header"])
        st.sidebar.markdown(f'<div class="tooltip">ℹ️<span class="tooltiptext">{texts[lang]["filters_tooltip"]}</span></div>', unsafe_allow_html=True)
        salesmen = st.sidebar.multiselect(
            texts[lang]["select_salesmen"],
            options=sorted(sales_df["Driver Name EN"].dropna().unique()),
            default=sorted(sales_df["Driver Name EN"].dropna().unique()),
            key="st_salesmen"
        )
        billing_types = st.sidebar.multiselect(
            texts[lang]["select_billing_types"],
            options=sorted(sales_df["Billing Type"].dropna().unique()),
            default=sorted(sales_df["Billing Type"].dropna().unique()),
            key="st_billing_types"
        )
        py_filter = st.sidebar.multiselect(
            texts[lang]["select_py"],
            options=sorted(sales_df["PY Name 1"].dropna().unique()),
            default=sorted(sales_df["PY Name 1"].dropna().unique()),
            key="st_py_filter"
        )
        sp_filter = st.sidebar.multiselect(
            texts[lang]["select_sp"],
            options=sorted(sales_df["SP Name1"].dropna().unique()),
            default=sorted(sales_df["SP Name1"].dropna().unique()),
            key="st_sp_filter"
        )

        preset = st.sidebar.radio(texts[lang]["date_presets"], texts[lang]["date_presets_options"], key="st_preset")
        today = pd.Timestamp.today().normalize()
        if preset == texts[lang]["date_presets_options"][1]:  # Last 7 Days
            date_range = [today - pd.Timedelta(days=7), today]
        elif preset == texts[lang]["date_presets_options"][2]:  # This Month
            month_start = today.replace(day=1)
            month_end = month_start + pd.offsets.MonthEnd(0)
            date_range = [month_start, month_end]
        elif preset == texts[lang]["date_presets_options"][3]:  # YTD
            date_range = [today.replace(month=1, day=1), today]
        else:
            date_range = st.sidebar.date_input(
                texts[lang]["select_date_range"],
                [sales_df["Billing Date"].min(), sales_df["Billing Date"].max()],
                key="st_date_range"
            )
            if isinstance(date_range, tuple) and len(date_range) == 2:
                date_range = [pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])]
            else:
                date_range = [sales_df["Billing Date"].min(), sales_df["Billing Date"].max()]

        if date_range[0] > date_range[1]:
            st.error(texts[lang]["date_error"])
        else:
            top_n = st.sidebar.slider(
                texts[lang]["top_n_salesmen"],
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
            ].copy()

            if df_filtered.empty:
                st.warning(texts[lang]["no_match_warning"])
            else:
                billing_start = df_filtered["Billing Date"].min().normalize()
                billing_end = df_filtered["Billing Date"].max().normalize()
                all_days = pd.date_range(billing_start, billing_end, freq="D")
                days_finish = int(sum(1 for d in all_days if d.weekday() != 4))

                current_month_start = today.replace(day=1)
                current_month_end = current_month_start + pd.offsets.MonthEnd(0)
                month_days = pd.date_range(current_month_start, current_month_end, freq="D")
                working_days_current_month = int(sum(1 for d in month_days if d.weekday() != 4))

                # --- Base aggregates ---
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
                per_day_ka_target = (total_ka_target_all / working_days_current_month) if working_days_current_month > 0 else 0
                current_sales_per_day = (total_sales.sum() / days_finish) if days_finish > 0 else 0
                forecast_month_end_ka = current_sales_per_day * working_days_current_month

                # --- Channels mapping: Market (Retail) vs E-com ---
                df_py_sales = df_filtered.groupby("_py_name_norm")["Net Value"].sum().reset_index()
                df_channels_merged = df_py_sales.merge(
                    channels_df[["_py_name_norm", "Channels"]],
                    on="_py_name_norm",
                    how="left"
                )
                df_channels_merged["Channels"] = df_channels_merged["Channels"].str.strip().str.lower().fillna("uncategorized")
                channel_sales = df_channels_merged.groupby("Channels")["Net Value"].sum()
                total_retail_sales = float(channel_sales.get("market", 0.0) + channel_sales.get("uncategorized", 0.0))
                total_ecom_sales = float(channel_sales.get("e-com", 0.0))
                total_channel_sales = total_retail_sales + total_ecom_sales
                retail_sales_pct = (total_retail_sales / total_channel_sales * 100) if total_channel_sales > 0 else 0
                ecom_sales_pct = (total_ecom_sales / total_channel_sales * 100) if total_channel_sales > 0 else 0

                # --- KA & Other E-com Calculation ---
                ka_other_ecom_sales = total_sales.sum() - talabat_sales.sum()
                ka_other_ecom_pct = (ka_other_ecom_sales / total_ka_target_all * 100) if total_ka_target_all > 0 else 0

                # --- KPI Data for PPTX ---
                kpi_data = {
                    texts[lang]["ka_target"]: f"KD {total_ka_target_all:,.0f}",
                    texts[lang]["talabat_target"]: f"KD {total_tal_target_all:,.0f}",
                    texts[lang]["ka_gap"]: f"KD {(total_ka_target_all - total_sales.sum()):,.0f}",
                    "Total Talabat Gap": f"KD {talabat_gap.sum():,.0f}",
                    texts[lang]["total_ka_sales"]: f"KD {total_sales.sum():,.0f} ({((total_sales.sum() / total_ka_target_all) * 100):.0f}%)" if total_ka_target_all else f"KD {total_sales.sum():,.0f} (0%)",
                    "Total Talabat Sales": f"KD {talabat_sales.sum():,.0f} ({((talabat_sales.sum() / total_tal_target_all) * 100):.0f}%)" if total_tal_target_all else f"KD {talabat_sales.sum():,.0f} (0%)",
                    texts[lang]["ka_other_ecom"]: f"KD {ka_other_ecom_sales:,.0f} ({ka_other_ecom_pct:.0f}%)",
                    texts[lang]["retail_sales"]: f"KD {total_retail_sales:,.0f} ({retail_sales_pct:.0f}%)",
                    texts[lang]["ecom_sales"]: f"KD {total_ecom_sales:,.0f} ({ecom_sales_pct:.0f}%)",
                    texts[lang]["days_finished"]: f"{days_finish}",
                    "Per Day KA Target": f"KD {per_day_ka_target:,.0f}",
                    texts[lang]["current_sales_per_day"]: f"KD {current_sales_per_day:,.0f}",
                    texts[lang]["forecast_month_end"]: f"KD {forecast_month_end_ka:,.0f}"
                }

                tabs = st.tabs([texts[lang]["kpis_tab"], texts[lang]["tables_tab"], texts[lang]["charts_tab"], texts[lang]["downloads_tab"]])

                # --- KPIs with progress bars ---
                with tabs[0]:
                    st.subheader(texts[lang]["key_metrics_sub"])
                    r1c1 = st.columns(1)[0]
                    with r1c1:
                        st.metric(texts[lang]["total_ka_sales"], f"KD {total_sales.sum():,.0f}")
                        progress_pct_ka = (total_sales.sum() / total_ka_target_all * 100) if total_ka_target_all > 0 else 0
                        st.markdown(create_progress_bar_html(progress_pct_ka), unsafe_allow_html=True)
                        st.markdown(f'<div class="green-caption">{texts[lang]["of_ka_target"].format(progress_pct_ka)}</div>', unsafe_allow_html=True)

                    r2c1, r2c2 = st.columns(2)
                    with r2c1:
                        st.metric(texts[lang]["ka_other_ecom"], f"KD {ka_other_ecom_sales:,.0f}")
                        st.markdown(create_progress_bar_html(ka_other_ecom_pct), unsafe_allow_html=True)
                        st.markdown(f'<div class="green-caption">{texts[lang]["of_ka_target_pct"].format(ka_other_ecom_pct)}</div>', unsafe_allow_html=True)
                    with r2c2:
                        st.metric(texts[lang]["talabat_sales"], f"KD {talabat_sales.sum():,.0f}")
                        progress_pct_talabat = (talabat_sales.sum() / total_tal_target_all * 100) if total_tal_target_all > 0 else 0
                        st.markdown(create_progress_bar_html(progress_pct_talabat), unsafe_allow_html=True)
                        st.markdown(f'<div class="green-caption">{texts[lang]["of_talabat_target"].format(progress_pct_talabat)}</div>', unsafe_allow_html=True)

                    st.subheader(texts[lang]["target_overview_sub"])
                    r3c1, r3c2, r3c3, r3c4 = st.columns(4)
                    r3c1.metric(texts[lang]["ka_target"], f"KD {total_ka_target_all:,.0f}")
                    r3c2.metric(texts[lang]["talabat_target"], f"KD {total_tal_target_all:,.0f}")
                    r3c3.metric(texts[lang]["ka_gap"], f"KD {(total_ka_target_all - total_sales.sum()):,.0f}")
                    r3c4.metric(texts[lang]["talabat_gap"], f"KD {talabat_gap.sum():,.0f}")

                    st.subheader(texts[lang]["channel_sales_sub"])
                    r4c1, r4c2 = st.columns(2)
                    with r4c1:
                        st.metric(texts[lang]["retail_sales"], f"KD {total_retail_sales:,.0f}")
                        retail_contribution_pct = (total_retail_sales / total_sales.sum() * 100) if total_sales.sum() > 0 else 0
                        st.caption(texts[lang]["of_total_ka"].format(retail_contribution_pct))
                    with r4c2:
                        st.metric(texts[lang]["ecom_sales"], f"KD {total_ecom_sales:,.0f}")
                        ecom_contribution_pct = (total_ecom_sales / total_sales.sum() * 100) if total_sales.sum() > 0 else 0
                        st.caption(texts[lang]["of_total_ka"].format(ecom_contribution_pct))

                    st.subheader(texts[lang]["performance_metrics_sub"])
                    r5c1, r5c2, r5c3 = st.columns(3)
                    r5c1.metric(texts[lang]["days_finished"], days_finish)
                    r5c2.metric(texts[lang]["current_sales_per_day"], f"KD {current_sales_per_day:,.0f}")
                    r5c3.metric(texts[lang]["forecast_month_end"], f"KD {forecast_month_end_ka:,.0f}")

                    # --- TABLES ---
                    with tabs[1]:
                        # --- Sales & Targets Summary ---
                        st.subheader(texts[lang]["sales_targets_summary_sub"])
                        report_df = pd.DataFrame({
                            "Salesman": ka_targets.index,
                            "KA Target": ka_targets.values,
                            "KA Sales": total_sales.values,
                            "KA Remaining": ka_gap.values,
                            "KA % Achieved": np.where(ka_targets.values != 0, (total_sales.values / ka_targets.values * 100).round(0), 0),
                            "Talabat Target": talabat_targets.values,
                            "Talabat Sales": talabat_sales.values,
                            "Talabat Remaining": talabat_gap.values,
                            "Talabat % Achieved": np.where(talabat_targets.values != 0, (talabat_sales.values / talabat_targets.values * 100).round(0), 0)
                        })

                        ttotal_row = report_df.sum(numeric_only=True).to_frame().T
                        total_row.index = ["Total"]

                        total_row["Target % Achieved"] = round(
                            total_row["Sales"]/total_row["Target"]*100,0
                        ) if total_row["Target"].iloc[0] != 0 else 0

                        total_row["E-Com % Achieved"] = round(
                            total_row["E-Com Sales"]/total_row["E-Com Target"]*100,0
                        ) if total_row["E-Com Target"].iloc[0] != 0 else 0

                        total_row = total_row.reset_index(drop=True)
                        total_row["Salesman Name"] = "Total"

                        total_row = total_row[report_df.columns]
                        report_df_with_total = pd.concat([report_df, total_row], ignore_index=True)

                        report_df_with_total = pd.concat([report_df, total_row], ignore_index=True)

                        def highlight_total_row(row):
                            return ['background-color: #BFDBFE; color: #1E3A8A; font-weight: 900' if row['Salesman'] == "Total" else '' for _ in row]

                        styled_report = (
                            report_df_with_total.style
                            .set_table_styles([
                                {'selector': 'th', 'props': [('background', '#1E3A8A'), ('color', 'white'),
                                                            ('font-weight', '800'), ('height', '40px'),
                                                            ('line-height', '40px'), ('border', '1px solid #E5E7EB')]}
                            ])
                            .apply(highlight_total_row, axis=1)
                            .format("{:,.0f}", subset=["KA Target","KA Sales","KA Remaining",
                                                    "Talabat Target","Talabat Sales","Talabat Remaining"])
                            .format("{:.0f}%", subset=["KA % Achieved","Talabat % Achieved"])
                        )
                        st.dataframe(styled_report, use_container_width=True, hide_index=True)
                        if st.download_button(
                            texts[lang]["download_sales_targets"],
                            data=to_excel_bytes(report_df_with_total, sheet_name="Sales_Targets_Summary", index=False),
                            file_name=f"Sales_Targets_Summary_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        ):
                            st.session_state["audit_log"].append({
                                "user": username,
                                "action": "download",
                                "details": "Sales & Targets Summary Excel",
                                "timestamp": datetime.now()
                            })

                        # --- Sales by Billing Type per Salesman ---
                        st.subheader(texts[lang]["sales_by_billing_sub"])
                        billing_wide = df_filtered.pivot_table(
                            index="Driver Name EN",
                            columns="Billing Type",
                            values="Net Value",
                            aggfunc="sum",
                            fill_value=0
                        )

                        required_cols_raw = ["ZFR", "YKF2", "YKRE", "YKS1", "YKS2", "ZCAN", "ZRE"]
                        billing_wide = billing_wide.reindex(columns=required_cols_raw, fill_value=0)
                        display_df = billing_wide.rename(columns={"ZFR": "Presales", "YKF2": "HHT"})
                        display_df["Sales Total"] = billing_wide.sum(axis=1)
                        display_df["Return"] = billing_wide["YKRE"] + billing_wide["ZRE"]
                        display_df["Return %"] = np.where(display_df["Sales Total"] != 0,
                                                        (display_df["Return"] / display_df["Sales Total"] * 100).round(0), 0)
                        display_df["Cancel Total"] = billing_wide[["YKS1", "YKS2", "ZCAN"]].sum(axis=1)
                        ordered_cols = ["Presales", "HHT", "Sales Total", "YKS1", "YKS2", "ZCAN",
                                        "Cancel Total", "YKRE", "ZRE", "Return", "Return %"]
                        display_df = display_df.reindex(columns=ordered_cols, fill_value=0)

                        total_row = pd.DataFrame(display_df.sum(numeric_only=True)).T
                        total_row.index = ["Total"]
                        total_row["Return %"] = round((total_row["Return"] / total_row["Sales Total"] * 100), 0) if total_row["Sales Total"].iloc[0] != 0 else 0
                        billing_df = pd.concat([display_df, total_row])
                        billing_df.index.name = "Salesman"

                        def highlight_total_row_billing(row):
                            return ['background-color: #BFDBFE; color: #1E3A8A; font-weight: 900' if row.name == "Total" else '' for _ in row]

                        styled_billing = (
                            billing_df.style
                            .set_table_styles([
                                {'selector': 'th', 'props': [('background', '#1E3A8A'), ('color', 'white'),
                                                            ('font-weight', '800'), ('height', '40px'),
                                                            ('line-height', '40px'), ('border', '1px solid #E5E7EB')]}
                            ])
                            .apply(highlight_total_row_billing, axis=1)
                            .format({
                                "Presales": "{:,.0f}", "HHT": "{:,.0f}", "Sales Total": "{:,.0f}",
                                "YKS1": "{:,.0f}", "YKS2": "{:,.0f}", "ZCAN": "{:,.0f}", "Cancel Total": "{:,.0f}",
                                "YKRE": "{:,.0f}", "ZRE": "{:,.0f}", "Return": "{:,.0f}", "Return %": "{:.0f}%"
                            })
                        )
                        st.dataframe(styled_billing, use_container_width=True, hide_index=False)
                        if st.download_button(
                            texts[lang]["download_billing"],
                            data=to_excel_bytes(billing_df.reset_index(), sheet_name="Billing_Types", index=False),
                            file_name=f"Billing_Types_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        ):
                            st.session_state["audit_log"].append({
                                "user": username,
                                "action": "download",
                                "details": "Billing Type Table Excel",
                                "timestamp": datetime.now()
                            })

                        # --- Sales by PY Name 1 ---
                        st.subheader(texts[lang]["sales_by_py_sub"])
                        py_table = df_filtered.groupby("PY Name 1")["Net Value"].sum().sort_values(ascending=False).to_frame(name="Sales")
                        py_table["Contribution %"] = np.where(py_table["Sales"] != 0,
                                                            (py_table["Sales"]/py_table["Sales"].sum()*100).round(0), 0)

                        total_row = py_table.sum(numeric_only=True).to_frame().T
                        total_row.index = ["Total"]
                        py_table_with_total = pd.concat([py_table, total_row])
                        py_table_with_total.index.name = "PY Name 1"

                        def highlight_total_row_py(row):
                            return ['background-color: #BFDBFE; color: #1E3A8A; font-weight: 900' if row.name == "Total" else '' for _ in row]

                        styled_py = (
                            py_table_with_total.style
                            .set_table_styles([
                                {'selector': 'th', 'props': [('background', '#1E3A8A'), ('color', 'white'),
                                                            ('font-weight', '800'), ('height', '40px'),
                                                            ('line-height', '40px'), ('border', '1px solid #E5E7EB')]}
                            ])
                            .apply(highlight_total_row_py, axis=1)
                            .format("{:,.0f}", subset=["Sales"])
                            .format("{:.0f}%", subset=["Contribution %"])
                        )
                        st.dataframe(styled_py, use_container_width=True, hide_index=False)
                        if st.download_button(
                            texts[lang]["download_py"],
                            data=to_excel_bytes(py_table_with_total.reset_index(), sheet_name="Sales_by_PY_Name", index=False),
                            file_name=f"Sales_by_PY_Name_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        ):
                            st.session_state["audit_log"].append({
                                "user": username,
                                "action": "download",
                                "details": "PY Name Table Excel",
                                "timestamp": datetime.now()
                            })

                        # --- Return by PY Name 1 ---
                        st.subheader("🔄 Return by PY Name 1")
                        py_billing = df_filtered.pivot_table(
                            index="PY Name 1",
                            columns="Billing Type",
                            values="Net Value",
                            aggfunc="sum",
                            fill_value=0
                        )

                        py_billing = py_billing.reindex(columns=required_cols_raw, fill_value=0)
                        py_billing["Sales Total"] = py_billing.sum(axis=1)
                        py_billing["Return"] = py_billing["YKRE"] + py_billing["ZRE"]
                        py_billing["Cancel Total"] = py_billing[["YKS1", "YKS2", "ZCAN"]].sum(axis=1)

                        # Reorder columns like Sales by Billing table
                        ordered_cols = ["Presales", "HHT", "Sales Total", "YKS1", "YKS2", "ZCAN",
                                        "Cancel Total", "YKRE", "ZRE", "Return", "Return %"]
                        # Map the original column names to the ordered display names
                        py_billing = py_billing.rename(columns={"ZFR": "Presales", "YKF2": "HHT"})
                        py_billing["Return %"] = np.where(py_billing["Sales Total"] != 0,
                                                        (py_billing["Return"] / py_billing["Sales Total"] * 100).round(0), 0)
                        py_billing = py_billing.reindex(columns=ordered_cols, fill_value=0).astype(int)

                        # Add total row
                        total_row = pd.DataFrame(py_billing.sum(numeric_only=True)).T
                        total_row.index = ["Total"]
                        total_row["Return %"] = round((total_row["Return"]/total_row["Sales Total"]*100), 0) if total_row["Sales Total"].iloc[0]!=0 else 0
                        py_billing = pd.concat([py_billing, total_row])

                        def highlight_total_row_py_return(row):
                            return ['background-color: #BFDBFE; color: #1E3A8A; font-weight: 900' if row.name == "Total" else '' for _ in row]

                        styled_py_return = (
                            py_billing.style
                            .set_table_styles([
                                {'selector': 'th', 'props': [('background', '#1E3A8A'), ('color', 'white'),
                                                            ('font-weight', '800'), ('height', '40px'),
                                                            ('line-height', '40px'), ('border', '1px solid #E5E7EB')]}
                            ])
                            .apply(highlight_total_row_py_return, axis=1)
                            .format({
                                "Presales": "{:,.0f}", "HHT": "{:,.0f}", "Sales Total": "{:,.0f}",
                                "YKS1": "{:,.0f}", "YKS2": "{:,.0f}", "ZCAN": "{:,.0f}", "Cancel Total": "{:,.0f}",
                                "YKRE": "{:,.0f}", "ZRE": "{:,.0f}", "Return": "{:,.0f}", "Return %": "{:.0f}%"
                            })
                        )
                        st.dataframe(styled_py_return, use_container_width=True, hide_index=False)

                        # --- Return by SP Name1 ---
                        st.subheader("🔄 Return By Branch")
                        sp_billing = df_filtered.pivot_table(
                            index="SP Name1",
                            columns="Billing Type",
                            values="Net Value",
                            aggfunc="sum",
                            fill_value=0
                        )

                        sp_billing = sp_billing.reindex(columns=required_cols_raw, fill_value=0)
                        sp_billing["Sales Total"] = sp_billing.sum(axis=1)
                        sp_billing["Return"] = sp_billing["YKRE"] + sp_billing["ZRE"]
                        sp_billing["Cancel Total"] = sp_billing[["YKS1", "YKS2", "ZCAN"]].sum(axis=1)

                        # Reorder and rename like Sales by Billing table
                        sp_billing = sp_billing.rename(columns={"ZFR": "Presales", "YKF2": "HHT"})
                        sp_billing["Return %"] = np.where(sp_billing["Sales Total"] != 0,
                                                        (sp_billing["Return"] / sp_billing["Sales Total"] * 100).round(0), 0)
                        sp_billing = sp_billing.reindex(columns=ordered_cols, fill_value=0).astype(int)

                        # Add total row
                        total_row = pd.DataFrame(sp_billing.sum(numeric_only=True)).T
                        total_row.index = ["Total"]
                        total_row["Return %"] = round((total_row["Return"]/total_row["Sales Total"]*100), 0) if total_row["Sales Total"].iloc[0]!=0 else 0
                        sp_billing = pd.concat([sp_billing, total_row])

                        def highlight_total_row_sp_return(row):
                            return ['background-color: #BFDBFE; color: #1E3A8A; font-weight: 900' if row.name == "Total" else '' for _ in row]

                        styled_sp_return = (
                            sp_billing.style
                            .set_table_styles([
                                {'selector': 'th', 'props': [('background', '#1E3A8A'), ('color', 'white'),
                                                            ('font-weight', '800'), ('height', '40px'),
                                                            ('line-height', '40px'), ('border', '1px solid #E5E7EB')]}
                            ])
                            .apply(highlight_total_row_sp_return, axis=1)
                            .format({
                                "Presales": "{:,.0f}", "HHT": "{:,.0f}", "Sales Total": "{:,.0f}",
                                "YKS1": "{:,.0f}", "YKS2": "{:,.0f}", "ZCAN": "{:,.0f}", "Cancel Total": "{:,.0f}",
                                "YKRE": "{:,.0f}", "ZRE": "{:,.0f}", "Return": "{:,.0f}", "Return %": "{:.0f}%"
                            })
                        )
                        st.dataframe(styled_sp_return, use_container_width=True, hide_index=False)

                # --- CHARTS (GM Premium Visuals Pack) ---
                with tabs[2]:
                    st.subheader("📈 Daily Sales Trend & Forecast (GM View)")
                    df_time = df_filtered.groupby(pd.Grouper(key="Billing Date", freq="D"))["Net Value"].sum().reset_index()
                    df_time.rename(columns={"Billing Date": "ds", "Net Value": "y"}, inplace=True)

                    if len(df_time) > 2:
                        # Prophet Forecast
                        m = Prophet()
                        m.fit(df_time)
                        future = m.make_future_dataframe(periods=30)
                        forecast = m.predict(future)

                        # Identify Anomalies (Based on Rolling Mean & Std)
                        df_time['y_mean'] = df_time['y'].rolling(window=7).mean()
                        df_time['y_std'] = df_time['y'].rolling(window=7).std()
                        df_time['upper'] = df_time['y_mean'] + 2 * df_time['y_std']
                        df_time['lower'] = df_time['y_mean'] - 2 * df_time['y_std']
                        df_time['anomaly'] = np.where(
                            (df_time['y'] > df_time['upper']) | (df_time['y'] < df_time['lower']),
                            df_time['y'], np.nan
                        )

                        fig_trend = go.Figure()
                        fig_trend.add_trace(go.Scatter(
                            x=df_time["ds"], y=df_time["y"],
                            mode="lines+markers",
                            name="Actual Sales",
                            line=dict(color="#0052CC", width=3)
                        ))
                        fig_trend.add_trace(go.Scatter(
                            x=forecast["ds"], y=forecast["yhat"],
                            mode="lines",
                            name="Sales Forecast",
                            line=dict(color="#22C55E", width=2, dash="dash")
                        ))
                        fig_trend.add_trace(go.Scatter(
                            x=df_time["ds"], y=df_time["anomaly"],
                            mode="markers",
                            name="Anomaly",
                            marker=dict(color="red", size=12, symbol="x")
                        ))

                        fig_trend.update_layout(
                            xaxis_title="Date",
                            yaxis_title="Net Value (KD)",
                            hovermode="x unified",
                            template="plotly_white"
                        )
                        st.plotly_chart(fig_trend, use_container_width=True)
                    else:
                        st.info("Not enough trend data")

                    st.markdown("---")
                    st.subheader("🛒 Market vs E-com Performance (Value & Share)")

                    # Market vs E-com
                    market_ecom_df = pd.DataFrame({
                        "Channel": ["Market", "E-com"],
                        "Sales": [total_retail_sales, total_ecom_sales]
                    })

                    fig_channel = go.Figure()

                    fig_channel.add_trace(go.Bar(
                        x=market_ecom_df["Channel"],
                        y=market_ecom_df["Sales"],
                        name="Sales",
                        text=market_ecom_df["Sales"].apply(lambda x: f"KD {x:,.0f}"),
                        textposition="outside",
                        marker=dict(
                            color=["#0EA5E9", "#8B5CF6"],
                            line=dict(color="black", width=1)
                        )
                    ))

                    fig_channel.add_trace(go.Pie(
                        labels=market_ecom_df["Channel"],
                        values=market_ecom_df["Sales"],
                        hole=0.55,
                        name="Share %",
                        textinfo="percent",
                        domain=dict(x=[0.55, 1.0])
                    ))

                    fig_channel.update_layout(
                        barmode="group",
                        showlegend=False,
                        template="plotly_white"
                    )
                    st.plotly_chart(fig_channel, use_container_width=True)

                    st.markdown("---")
                    st.subheader("🎯 Daily KA Target vs Actual Sales")

                    df_time_target = df_filtered.groupby(pd.Grouper(key="Billing Date", freq="D"))["Net Value"].sum().reset_index()
                    df_time_target.rename(columns={"Billing Date": "Date", "Net Value": "Sales"}, inplace=True)
                    df_time_target["Daily KA Target"] = per_day_ka_target

                    fig_target = go.Figure()
                    fig_target.add_trace(go.Scatter(
                        x=df_time_target["Date"], y=df_time_target["Sales"],
                        name="Actual Sales",
                        mode="lines+markers",
                        line=dict(color="#22C55E", width=3)
                    ))
                    fig_target.add_trace(go.Scatter(
                        x=df_time_target["Date"], y=df_time_target["Daily KA Target"],
                        name="Daily KA Target",
                        mode="lines",
                        line=dict(color="#FACC15", width=2, dash="dot")
                    ))

                    fig_target.update_layout(
                        xaxis_title="Date",
                        yaxis_title="Net Value (KD)",
                        hovermode="x unified",
                        template="plotly_white"
                    )
                    st.plotly_chart(fig_target, use_container_width=True)

                    st.markdown("---")
                    st.subheader("💪 Salesman KA Target vs Actual")

                    target_actual_df = pd.DataFrame({
                        "Salesman": ka_targets.index,
                        "KA Target": ka_targets.values,
                        "KA Sales": sales_by_sm.values
                    }).sort_values("KA Sales", ascending=False)

                    fig_salesman = go.Figure()
                    fig_salesman.add_trace(go.Bar(
                        x=target_actual_df["Salesman"],
                        y=target_actual_df["KA Sales"],
                        name="Sales",
                        text=target_actual_df["KA Sales"].apply(lambda x: f"{x:,.0f}"),
                        textposition="inside",
                        marker=dict(color=[
                            "#22C55E" if s >= t else "#EF4444"
                            for s, t in zip(target_actual_df["KA Sales"], target_actual_df["KA Target"])
                        ])
                    ))

                    fig_salesman.add_trace(go.Scatter(
                        x=target_actual_df["Salesman"],
                        y=target_actual_df["KA Target"],
                        name="Target",
                        mode="lines+markers",
                        line=dict(color="#1E3A8A", width=2)
                    ))

                    fig_salesman.update_layout(
                        xaxis_title="Salesman",
                        yaxis_title="Net Value (KD)",
                        template="plotly_white"
                    )
                    st.plotly_chart(fig_salesman, use_container_width=True)

                    st.markdown("---")
                    st.subheader("🏆 Top 10 Customers by Sales (KD)")

                    top10 = df_filtered.groupby("PY Name 1")["Net Value"].sum().sort_values(ascending=False).head(10)

                    fig_top10 = go.Figure(go.Bar(
                        x=top10.index,
                        y=top10.values,
                        text=[f"KD {v:,.0f}" for v in top10.values],
                        textposition="outside",
                        marker=dict(color="#2563EB")
                    ))

                    fig_top10.update_layout(
                        xaxis_title="Customer",
                        yaxis_title="Net Value (KD)",
                        template="plotly_white"
                    )
                    st.plotly_chart(fig_top10, use_container_width=True)
                    
                    

# --- YTD Comparison Page ---
elif choice == "Year to Date Comparison":
    if "ytd_df" in st.session_state and not st.session_state["ytd_df"].empty:
        ytd_df = st.session_state["ytd_df"].copy()

        if "Billing Date" not in ytd_df.columns:
            st.error("❌ 'Billing Date' column not found in YTD sheet.")
            st.stop()

        ytd_df["Billing Date"] = pd.to_datetime(ytd_df["Billing Date"], errors="coerce")
        ytd_df = ytd_df.dropna(subset=["Billing Date"])

        if "Net Value" not in ytd_df.columns:
            st.error("❌ 'Net Value' column not found in YTD sheet.")
            st.stop()

        st.title("📅 Year to Date Comparison")
        st.markdown(
            '<div class="tooltip">ℹ️<span class="tooltiptext">Compare sales across two periods by a selected dimension.</span></div>',
            unsafe_allow_html=True
        )

        # --- Select Dimension ---
        st.subheader("📊 Choose Dimension")

        # Map friendly labels to actual column names (added By Material)
        dim_options = {
            "By Customer": "PY Name 1",          # Customer Name column
            "By Salesman": "Driver Name EN",     # Salesman Name
            "By Branch": "SP Name1",             # Branch Name column
            "By Material": "Material Description"  # 👈 NEW: Material Description
        }

        # Filter only dimensions where the column actually exists
        dim_options_available = {
            label: col for label, col in dim_options.items() if col in ytd_df.columns
        }

        if not dim_options_available:
            st.error("❌ None of the expected dimension columns (Customer, Salesman, Branch, Material) are present.")
            st.stop()

        # User sees friendly names; program uses real column names
        dim_label = st.selectbox(
            "Choose dimension",
            list(dim_options_available.keys()),
            index=0
        )
        dimension = dim_options_available[dim_label]

        # --- Select Two Periods ---
        st.subheader("📆 Select Two Periods")
        col1, col2 = st.columns(2)

        min_date = ytd_df["Billing Date"].min().date()
        max_date = ytd_df["Billing Date"].max().date()

        with col1:
            st.write("Period 1")
            period1_range = st.date_input(
                "Select Date Range ( Click Twice Start & End Date)",
                value=(min_date, max_date),
                key="ytd_p1_range"
            )
        with col2:
            st.write("Period 2")
            period2_range = st.date_input(
                "Select Date Range ( Click Twice Start & End Date)",
                value=(min_date, max_date),
                key="ytd_p2_range"
            )

        if period1_range and period2_range and len(period1_range) == 2 and len(period2_range) == 2:
            period1_start, period1_end = period1_range
            period2_start, period2_end = period2_range

            df_p1 = ytd_df[
                (ytd_df["Billing Date"] >= pd.to_datetime(period1_start)) &
                (ytd_df["Billing Date"] <= pd.to_datetime(period1_end))
            ]
            df_p2 = ytd_df[
                (ytd_df["Billing Date"] >= pd.to_datetime(period2_start)) &
                (ytd_df["Billing Date"] <= pd.to_datetime(period2_end))
            ]

            if df_p1.empty or df_p2.empty:
                st.warning("⚠️ One of the selected periods has no data.")
            else:
                # --- YTD Comparison Table ---
                col_p1_name = f"{period1_start.strftime('%Y-%m-%d')} to {period1_end.strftime('%Y-%m-%d')} Sales"
                col_p2_name = f"{period2_start.strftime('%Y-%m-%d')} to {period2_end.strftime('%Y-%m-%d')} Sales"

                summary_p1 = (
                    df_p1.groupby(dimension)["Net Value"]
                    .sum()
                    .reset_index()
                    .rename(columns={"Net Value": col_p1_name})
                )
                summary_p2 = (
                    df_p2.groupby(dimension)["Net Value"]
                    .sum()
                    .reset_index()
                    .rename(columns={"Net Value": col_p2_name})
                )

                ytd_comparison = pd.merge(summary_p1, summary_p2, on=dimension, how="outer")

                # 🔒 Safety: ensure both period columns exist (avoid KeyError)
                if col_p1_name not in ytd_comparison.columns:
                    ytd_comparison[col_p1_name] = 0
                if col_p2_name not in ytd_comparison.columns:
                    ytd_comparison[col_p2_name] = 0

                ytd_comparison = ytd_comparison.fillna(0)

                # Difference = Period2 - Period1
                ytd_comparison["Difference"] = ytd_comparison[col_p2_name] - ytd_comparison[col_p1_name]

                # Rename dimension column to generic "Name"
                ytd_comparison.rename(columns={dimension: "Name"}, inplace=True)

                # Total row
                total_row = {
                    "Name": "Total",
                    col_p1_name: ytd_comparison[col_p1_name].sum(),
                    col_p2_name: ytd_comparison[col_p2_name].sum(),
                    "Difference": ytd_comparison["Difference"].sum()
                }
                ytd_comparison = pd.concat(
                    [ytd_comparison, pd.DataFrame([total_row])],
                    ignore_index=True
                )

                # Use friendly label in the heading
                st.subheader(f"📋 Comparison by {dim_label}")

                styled_ytd = (
                    ytd_comparison.style
                    .set_table_styles([
                        {
                            'selector': 'th',
                            'props': [
                                ('background', '#1E3A8A'),
                                ('color', 'white'),
                                ('font-weight', '800'),
                                ('height', '40px'),
                                ('line-height', '40px'),
                                ('border', '1px solid #E5E7EB')
                            ]
                        }
                    ])
                    .apply(
                        lambda x: [
                            'background-color: #BFDBFE; color: #1E3A8A; font-weight: 900'
                            if x.name == len(x) - 1 else ''  # last row = Total
                            for _ in x
                        ],
                        axis=1
                    )
                    .format(
                        {
                            col_p1_name: "{:,.0f}",
                            col_p2_name: "{:,.0f}",
                            'Difference': "{:,.0f}"
                        }
                    )
                )
                st.dataframe(styled_ytd, use_container_width=True, hide_index=False)

                st.download_button(
                    "⬇️ Download YTD Comparison (Excel)",
                    data=to_excel_bytes(ytd_comparison, sheet_name="YTD_Comparison", index=False),
                    file_name=f"YTD_Comparison_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                # --- NEW: Sales by Month (YTD) ---
                st.subheader("📈 Sales by Month (YTD)")
                monthly = ytd_df.copy()
                monthly["YearMonth"] = monthly["Billing Date"].dt.to_period("M").astype(str)

                monthly_sales = (
                    monthly.groupby("YearMonth")["Net Value"]
                    .sum()
                    .reset_index()
                    .sort_values("YearMonth")
                )

                if not monthly_sales.empty:
                    fig_month = px.bar(
                        monthly_sales,
                        x="YearMonth",
                        y="Net Value",
                        title="Sales by Month (YTD)",
                        labels={"YearMonth": "Year-Month", "Net Value": "Net Value (KD)"}
                    )
                    fig_month.update_layout(height=400)
                    st.plotly_chart(fig_month, use_container_width=True)
                else:
                    st.info("No monthly data available in YTD sheet.")

                # --- Top 10 Customers: Last Year vs Current Year ---
                st.subheader("🏆 Top 10 Customers – Last Year vs Current Year")

                # Make sure Billing Date is datetime
                if not np.issubdtype(ytd_df["Billing Date"].dtype, np.datetime64):
                    ytd_df["Billing Date"] = pd.to_datetime(ytd_df["Billing Date"], errors="coerce")

                ytd_df["Year"] = ytd_df["Billing Date"].dt.year
                available_years = sorted(ytd_df["Year"].dropna().unique().tolist())
                if not available_years:
                    st.info("⚠️ No valid years found in YTD data.")
                    st.stop()
                default_year = max(available_years)
                selected_current_year = st.selectbox("Select Current Year:", available_years, index=available_years.index(default_year))
                current_year = int(selected_current_year)
                last_year = current_year - 1

                # Aggregate sales by Customer + Year
                cust_sales = (
                    ytd_df[ytd_df["Year"].isin([last_year, current_year])]
                    .groupby(["PY Name 1", "Year"])["Net Value"]
                    .sum()
                    .reset_index()
                )

                if cust_sales.empty:
                    st.info("⚠️ No customer sales found for last year or current year.")
                else:
                    cust_pivot = cust_sales.pivot(
                        index="PY Name 1",
                        columns="Year",
                        values="Net Value"
                    ).fillna(0)

                    # Use only year columns that actually exist (avoid KeyError)
                    year_cols = [y for y in [last_year, current_year] if y in cust_pivot.columns]

                    if not year_cols:
                        st.info("⚠️ No valid year columns found to build Top 10 customers chart.")
                    else:
                        cust_pivot["Total"] = cust_pivot[year_cols].sum(axis=1)
                        top10_cust = cust_pivot.sort_values("Total", ascending=False).head(10).reset_index()

                        # Melt for plotting
                        top10_melt = top10_cust.melt(
                            id_vars="PY Name 1",
                            value_vars=year_cols,
                            var_name="Year",
                            value_name="Sales",
                        )

                        # Merge back the wide data for status calculation
                        merge_cols = ["PY Name 1"] + year_cols
                        top10_melt = top10_melt.merge(
                            top10_cust[merge_cols],
                            on="PY Name 1",
                            how="left",
                        )

                        def classify_status(row):
                            # Both last_year & current_year exist → Achieved vs Not Achieved
                            if (last_year in year_cols) and (current_year in year_cols) and row["Year"] == current_year:
                                return "Achieved" if row.get(current_year, 0) >= row.get(last_year, 0) else "Not Achieved"
                            # Only current year data exists
                            if row["Year"] == current_year:
                                return "Current"
                            # Only previous year / others
                            return "Previous"

                        top10_melt["Status"] = top10_melt.apply(classify_status, axis=1)

                        color_map = {
                            "Achieved": "green",
                            "Not Achieved": "red",
                            "Previous": "gray",
                            "Current": "blue",
                        }

                        fig_top10 = px.bar(
                            top10_melt,
                            x="PY Name 1",
                            y="Sales",
                            color="Status",
                            color_discrete_map=color_map,
                            barmode="group",
                            text=top10_melt["Sales"].apply(lambda x: f"{x:,.0f}"),
                        )
                        fig_top10.update_traces(
                            textposition="inside",
                            insidetextanchor="middle",
                            textfont=dict(color="white", size=12),
                        )
                        fig_top10.update_layout(
                            title=f"Top 10 Customers: {last_year} vs {current_year}",
                            xaxis_title="Customer",
                            yaxis_title="Sales (KD)",
                            font=dict(family="Roboto", size=12),
                            plot_bgcolor="#F3F4F6",
                            paper_bgcolor="#F3F4F6",
                        )
                        st.plotly_chart(fig_top10, use_container_width=True)

                # --- Return by SP Name1 + Material Description (YTD) ---
                st.subheader("🔄 Return By Branch + Material Description (YTD)")
                required_cols = {"SP Name1", "Material Description", "Billing Type", "Net Value"}
                if required_cols.issubset(ytd_df.columns):
                    sp_mat_ytd = pd.pivot_table(
                        ytd_df,
                        index=["SP Name1", "Material Description"],
                        columns="Billing Type",
                        values="Net Value",
                        aggfunc="sum",
                        fill_value=0
                    )
                    billing_cols = ["ZFR", "YKF2", "YKRE", "YKS1", "YKS2", "ZCAN", "ZRE"]
                    for col in billing_cols:
                        if col not in sp_mat_ytd.columns:
                            sp_mat_ytd[col] = 0
                    sp_mat_ytd = sp_mat_ytd.rename(columns={"ZFR": "Presales", "YKF2": "HHT"})
                    sp_mat_ytd["Sales Total"] = sp_mat_ytd.sum(axis=1, numeric_only=True)
                    sp_mat_ytd["Return"] = sp_mat_ytd["YKRE"] + sp_mat_ytd["ZRE"]
                    sp_mat_ytd["Cancel Total"] = sp_mat_ytd[["YKS1", "YKS2", "ZCAN"]].sum(axis=1)
                    sp_mat_ytd["Return %"] = np.where(
                        sp_mat_ytd["Sales Total"] != 0,
                        (sp_mat_ytd["Return"] / sp_mat_ytd["Sales Total"] * 100).round(0),
                        0
                    )
                    ordered_cols = [
                        "Presales", "HHT", "Sales Total", "YKS1", "YKS2", "ZCAN",
                        "Cancel Total", "YKRE", "ZRE", "Return", "Return %"
                    ]
                    sp_mat_ytd = sp_mat_ytd.reindex(columns=ordered_cols, fill_value=0)
                    total_row = pd.DataFrame(sp_mat_ytd.sum(numeric_only=True)).T
                    total_row.index = [("Total", "")]
                    total_row["Return %"] = (
                        round((total_row["Return"] / total_row["Sales Total"] * 100), 0)
                        if total_row["Sales Total"].iloc[0] != 0 else 0
                    )
                    sp_mat_ytd = pd.concat([sp_mat_ytd, total_row])

                    # Conditional highlights for easy read
                    def highlight_sp_mat(row):
                        styles = []
                        for col in row.index:
                            if row.name == ("Total", ""):
                                styles.append('background-color: #BFDBFE; color: #1E3A8A; font-weight: 900')
                            elif col == "Return" and row[col] > 0:
                                styles.append('background-color: #FECACA; color: #991B1B; font-weight: 700')
                            elif col == "Cancel Total" and row[col] > 0:
                                styles.append('background-color: #FDE68A; color: #92400E; font-weight: 700')
                            elif col == "Sales Total" and row[col] > 0:
                                styles.append('background-color: #D1FAE5; color: #065F46; font-weight: 700')
                            else:
                                styles.append('')
                        return styles

                    styled_sp_mat = (
                        sp_mat_ytd.style
                        .set_table_styles([
                            {
                                'selector': 'th',
                                'props': [
                                    ('background', '#1E3A8A'),
                                    ('color', 'white'),
                                    ('font-weight', '800'),
                                    ('height', '40px'),
                                    ('line-height', '40px'),
                                    ('border', '1px solid #E5E7EB')
                                ]
                            }
                        ])
                        .apply(highlight_sp_mat, axis=1)
                        .format({
                            "Presales": "{:,.0f}", "HHT": "{:,.0f}", "Sales Total": "{:,.0f}",
                            "YKS1": "{:,.0f}", "YKS2": "{:,.0f}", "ZCAN": "{:,.0f}", "Cancel Total": "{:,.0f}",
                            "YKRE": "{:,.0f}", "ZRE": "{:,.0f}", "Return": "{:,.0f}", "Return %": "{:.0f}%"
                        })
                    )
                    st.dataframe(styled_sp_mat, use_container_width=True, hide_index=True)
                    st.download_button(
                        "⬇️ Download Return by SP+Material (YTD)",
                        data=to_excel_bytes(sp_mat_ytd.reset_index(), sheet_name="Return_by_SP_Material_YTD", index=False),
                        file_name=f"Return_by_SP_Material_YTD_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.info("Required columns for SP+Material YTD table are missing.")
    else:
        st.warning("⚠️ Please ensure the 'YTD' sheet is present in your uploaded file.")


# --- Custom Analysis Page ---
elif choice == texts[lang]["custom_analysis"]:
    st.title(texts[lang]["custom_title"])
    if "data_loaded" not in st.session_state:
        st.warning(texts[lang]["no_data_warning"])
    else:
        # ✅ Ensure Extra sheet is loaded into session state
        if "Extra_sheet_df" not in st.session_state:
            try:
                extra_df = pd.read_excel(uploaded, sheet_name="Extra sheet")
            except Exception:
                extra_df = pd.DataFrame()
            st.session_state["Extra_sheet_df"] = extra_df

        # Available sheet options
        sheet_options = {
            "Sales Data": st.session_state.get("sales_df", pd.DataFrame()),
            "YTD": st.session_state.get("ytd_df", pd.DataFrame()),
            "Target": st.session_state.get("target_df", pd.DataFrame()),
            "Sales Channels": st.session_state.get("channels_df", pd.DataFrame()),
            "Extra sheet": st.session_state.get("Extra_sheet_df", pd.DataFrame())
        }

        selected_sheet_name = st.selectbox(texts[lang]["custom_select_sheet"], list(sheet_options.keys()))
        df = sheet_options[selected_sheet_name]

        if df.empty:
            st.warning(texts[lang]["custom_sheet_empty"].format(selected_sheet_name))
        else:
            st.subheader(texts[lang]["custom_explore"])

            available_cols = list(df.columns)
            group_cols = st.multiselect(texts[lang]["custom_group_cols"], available_cols)
            value_col = st.selectbox(texts[lang]["custom_value_col"], available_cols)

            if "Billing Date" in df.columns:
                st.subheader(texts[lang]["custom_periods_sub"])
                col1, col2 = st.columns(2)
                with col1:
                    st.write(texts[lang]["custom_period1"])
                    period1_range = st.date_input(
                        texts[lang]["custom_select_p1"],
                        [df["Billing Date"].min(), df["Billing Date"].max()],
                        key="ca_p1_range"
                    )
                with col2:
                    st.write(texts[lang]["custom_period2"])
                    period2_range = st.date_input(
                        texts[lang]["custom_select_p2"],
                        [df["Billing Date"].min(), df["Billing Date"].max()],
                        key="ca_p2_range"
                    )
            else:
                period1_range = period2_range = None
                st.info("⚠️ No 'Billing Date' column found. Period comparison disabled.")

            if group_cols and value_col and period1_range and period2_range and len(period1_range) == 2 and len(period2_range) == 2:
                # --- Period 1 ---
                p1_start, p1_end = pd.to_datetime(period1_range[0]), pd.to_datetime(period1_range[1])
                df_p1 = df[(df["Billing Date"] >= p1_start) & (df["Billing Date"] <= p1_end)]
                summary_p1 = df_p1.groupby(group_cols)[value_col].sum().reset_index()
                summary_p1.rename(columns={value_col: "Period 1"}, inplace=True)

                # --- Period 2 ---
                p2_start, p2_end = pd.to_datetime(period2_range[0]), pd.to_datetime(period2_range[1])
                df_p2 = df[(df["Billing Date"] >= p2_start) & (df["Billing Date"] <= p2_end)]
                summary_p2 = df_p2.groupby(group_cols)[value_col].sum().reset_index()
                summary_p2.rename(columns={value_col: "Period 2"}, inplace=True)

                # --- Merge & Compare ---
                comparison_df = pd.merge(summary_p1, summary_p2, on=group_cols, how="outer").fillna(0)
                comparison_df["Difference"] = comparison_df["Period 2"] - comparison_df["Period 1"]

                st.subheader(texts[lang]["custom_comparison_sub"].format(value_col, ", ".join(group_cols)))
                styled_custom = (
                    comparison_df.style
                    .set_table_styles([
                        {'selector': 'th', 'props': [('background', '#1E3A8A'), ('color', 'white'),
                                                    ('font-weight', '800'), ('height', '40px'),
                                                    ('line-height', '40px'), ('border', '1px solid #E5E7EB')]}
                    ])
                    .format({
                        "Period 1": "{:,.0f}",
                        "Period 2": "{:,.0f}",
                        "Difference": "{:,.0f}"
                    })
                )
                st.dataframe(styled_custom, use_container_width=True, hide_index=True)

                # --- Plotly Chart Fix ---
                df_plot = comparison_df.sort_values(by="Period 2", ascending=False).copy()

                if len(group_cols) == 1:
                    df_plot["Group"] = df_plot[group_cols[0]].astype(str)
                elif len(group_cols) > 1:
                    df_plot["Group"] = df_plot[group_cols].astype(str).agg(" | ".join, axis=1)
                else:
                    df_plot["Group"] = "All Data"

                df_plot_melted = df_plot.melt(
                    id_vars=["Group"],
                    value_vars=["Period 1", "Period 2"],
                    var_name="Period",
                    value_name="Value"
                )

                fig = px.bar(
                    df_plot_melted,
                    x="Group",
                    y="Value",
                    color="Period",
                    barmode="group",
                    title=f"Comparison of {value_col} by {', '.join(group_cols) if group_cols else 'All'}",
                    color_discrete_sequence=px.colors.qualitative.Set2
                )
                st.plotly_chart(fig, use_container_width=True)

                # --- Download ---
                if st.download_button(
                    texts[lang]["custom_download"],
                    data=to_excel_bytes(comparison_df, sheet_name="Custom_Comparison", index=False),
                    file_name=f"Custom_Comparison_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ):
                    st.session_state["audit_log"].append({
                        "user": username,
                        "action": "download",
                        "details": "Custom Comparison Excel",
                        "timestamp": datetime.now()
                    })
            else:
                st.info(texts[lang]["custom_select_prompt"])

# --- SP/PY Target Allocation Page ---
elif choice == "SP/PY Target Allocation":
    st.title("🎯 By Customer & By Branch Target Allocation")
    if "data_loaded" not in st.session_state:
        st.warning("⚠️ Please upload the Excel file from the sidebar first.")
        st.stop()

    sales_df = st.session_state["sales_df"]
    ytd_df = st.session_state["ytd_df"]
    target_df = st.session_state.get("target_df", pd.DataFrame())

    st.subheader("Configuration")
    st.markdown('<div class="tooltip">ℹ️<span class="tooltiptext">Allocate targets by branch or customer based on historical sales.</span></div>', unsafe_allow_html=True)
        # Target allocation grouping choice
    allocation_type = st.radio(
        "Select Target Allocation Type",
        ["By Branch", "By Customer"]
    )

    # Map choice to actual column
    group_col = "SP Name1" if allocation_type == "By Branch" else "PY Name 1"

    target_option = st.radio("Select Target Input Option", ["Manual", "Auto (from 'Target' sheet)"])

    total_target = 0
    if target_option == "Manual":
        total_target = st.number_input("Enter the Total Target to be Allocated for this Month (KD)", min_value=0, value=1000000, step=1000)
    else:
        if target_df.empty or "KA Target" not in target_df.columns:
            st.error("❌ 'Target' sheet or 'KA Target' column not found. Please upload a file with this sheet for 'Auto' mode.")
            st.stop()
        total_target = target_df["KA Target"].sum()
        st.info(f"Using Total Target from 'Target' sheet: KD {total_target:,.0f}")

    if total_target <= 0:
        st.warning("Please ensure the total target is greater than 0.")
        st.stop()

    st.subheader("Historical Data Period")
    today = pd.Timestamp.today().normalize()
    data_period_option = st.radio("Select Historical Data Period", ["Last 6 Months", "Manual Days"], index=1)

    if data_period_option == "Last 6 Months":
        lookback_period = pd.DateOffset(months=6)
        days_label = "6 Months"; months_count = 6
        end_date_selected = today; start_date_selected = today - lookback_period
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

    historical_df = ytd_df[(ytd_df["Billing Date"] >= pd.Timestamp(start_date_selected)) & (ytd_df["Billing Date"] <= pd.Timestamp(end_date_selected))].copy()
    if historical_df.empty:
        st.warning(f"⚠️ No sales data available in 'YTD' for {days_label}.")
        st.stop()

    historical_sales = historical_df.groupby(group_col)["Net Value"].sum()
    total_historical_sales_value = historical_sales.sum()
    current_month_sales_df = sales_df[(sales_df["Billing Date"].dt.month == today.month) & (sales_df["Billing Date"].dt.year == today.year)].copy()
    current_month_sales = current_month_sales_df.groupby(group_col)["Net Value"].sum()
    total_current_month_sales = current_month_sales.sum()

    target_balance = total_target - total_current_month_sales

    if total_target > 0:
        average_historical_sales = total_historical_sales_value / months_count
        st.subheader("🎯 Target Analysis")
        col1, col2, col3 = st.columns(3)
        col4, col5 = st.columns(2)
        with col1: st.metric("Historical Sales Total", f"KD {total_historical_sales_value:,.0f}")
        with col2: st.metric("Allocated Target Total", f"KD {total_target:,.0f}")
        with col3:
            if average_historical_sales > 0:
                percentage_increase_needed = ((total_target - average_historical_sales) / average_historical_sales) * 100
                delta_value = total_target - average_historical_sales
                st.metric("Increase Needed vs Avg Sales", f"{percentage_increase_needed:.0f}%", delta=f"KD {delta_value:,.0f}")
            else:
                st.metric("Increase Needed vs Avg Sales", "N/A", delta="Historical = 0")
        st.markdown("---")
        with col4: st.metric("Current Month Sales", f"KD {total_current_month_sales:,.0f}")
        with col5: st.metric("Target Balance", f"KD {target_balance:,.0f}")

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
    total_row["Target Balance"] = total_target - total_current_month_sales
    allocation_table_with_total = pd.concat([allocation_table, total_row])

    def color_target_balance(val):
        if isinstance(val, (int, float)):
            color = 'red' if val > 0 else 'green'
            return f'color: {color}'
        return ''

    st.subheader(f"📊 Auto-Allocated Targets Based on {days_label}")

    # Show the first column (Branch/Customer) with a proper header
    name_col = "Branch" if allocation_type == "By Branch" else "Customer"

    allocation_display = allocation_table_with_total.reset_index()
    first_col = allocation_display.columns[0]
    allocation_display = allocation_display.rename(columns={first_col: name_col})

    numeric_cols = [c for c in allocation_display.columns if c != name_col]
    for c in numeric_cols:
        allocation_display[c] = (
            pd.to_numeric(allocation_display[c], errors="coerce")
            .fillna(0)
            .round(0)
            .astype(int)
        )

    styled_allocation = (
        allocation_display.style
        .set_table_styles([
            {'selector': 'th', 'props': [
                ('background', '#1E3A8A'),
                ('color', 'white'),
                ('font-weight', '800'),
                ('height', '40px'),
                ('line-height', '40px'),
                ('border', '1px solid #E5E7EB')
            ]}
        ])
        .apply(
            lambda r: ['background-color: #BFDBFE; color: #1E3A8A; font-weight: 900' if str(r.get(name_col)) == 'Total' else '' for _ in r],
            axis=1
        )
        .format({c: '{:,.0f}' for c in numeric_cols})
    )
    st.dataframe(styled_allocation, use_container_width=True, hide_index=True)

    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    excel_data = to_excel_bytes(allocation_table, sheet_name="Allocated_Targets")
    if st.download_button(
        "💾 Download Target Allocation Table",
        data=excel_data,
        file_name=f"target_allocation_{allocation_type.replace(' ', '_')}_{timestamp}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    ):
        st.session_state["audit_log"].append({
            "user": username,
            "action": "download",
            "details": f"target_allocation_{allocation_type.replace(' ', '_')}_{timestamp}.xlsx",
            "timestamp": datetime.now()
        })

# --- AI Insights Page (New GM Executive View) ---
elif choice == "AI Insights":
    st.title("🤖 AI Insights – GM Executive View")

    if "data_loaded" not in st.session_state:
        st.warning("⚠️ Please upload the Excel file first.")
    else:
        # ------------------------------------------------
        # 1) Base Data
        # ------------------------------------------------
        sales_df = st.session_state["sales_df"].copy()
        target_df = st.session_state["target_df"].copy()
        ytd_df = st.session_state.get("ytd_df", pd.DataFrame()).copy()
        channels_df = st.session_state.get("channels_df", pd.DataFrame()).copy()

        # Helper
        def fmt_kd(x):
            try:
                return f"KD {x:,.0f}"
            except Exception:
                return "KD 0"

        # ------------------------------------------------
        # 2) Scope & Filters
        # ------------------------------------------------
        st.subheader("🎛 Scope & Filters")

        colf1, colf2 = st.columns(2)
        with colf1:
            min_date = pd.to_datetime(sales_df["Billing Date"].min())
            max_date = pd.to_datetime(sales_df["Billing Date"].max())
            date_range_ai = st.date_input(
                "Select analysis period",
                value=(min_date, max_date)
            )

        with colf2:
            top_n_ai = st.slider("Top-N salesmen spotlight (for GM notes)", min_value=3, max_value=15, value=5, step=1)

        if isinstance(date_range_ai, (list, tuple)) and len(date_range_ai) == 2:
            ai_start, ai_end = pd.to_datetime(date_range_ai[0]), pd.to_datetime(date_range_ai[1])
        else:
            ai_start, ai_end = min_date, max_date

        df_ai = sales_df[(sales_df["Billing Date"] >= ai_start) & (sales_df["Billing Date"] <= ai_end)].copy()

        if df_ai.empty:
            st.info("No data in the selected period. Try expanding the date range.")
            st.stop()

        # ------------------------------------------------
        # 3) Core KPIs for GM (Sales, Targets, Channels, Returns)
        # ------------------------------------------------
        st.markdown("---")
        st.subheader("📊 GM Snapshot – Key KPIs")

        total_sales = float(df_ai["Net Value"].sum())

        # Targets by salesman (KA + Talabat)
        if "Driver Name EN" in target_df.columns:
            ka_targets = target_df.set_index("Driver Name EN")["KA Target"] if "KA Target" in target_df.columns else pd.Series(dtype=float)
            tal_targets = target_df.set_index("Driver Name EN")["Talabat Target"] if "Talabat Target" in target_df.columns else pd.Series(dtype=float)
        else:
            ka_targets = pd.Series(dtype=float)
            tal_targets = pd.Series(dtype=float)

        # Sales by salesman (filtered period)
        sales_by_sm = df_ai.groupby("Driver Name EN")["Net Value"].sum()

        # Talabat sales (treated as E-Com – customer name match)
        tal_mask = df_ai["PY Name 1"].astype(str).str.upper().str.contains("STORES SERVICES KUWAIT CO.")
        tal_sales_by_sm = df_ai[tal_mask].groupby("Driver Name EN")["Net Value"].sum() if tal_mask.any() else pd.Series(dtype=float)

        # Align indices
        idx_union = sales_by_sm.index.union(ka_targets.index).union(tal_targets.index).union(tal_sales_by_sm.index)
        sales_by_sm = sales_by_sm.reindex(idx_union, fill_value=0.0)
        ka_targets = ka_targets.reindex(idx_union, fill_value=0.0)
        tal_targets = tal_targets.reindex(idx_union, fill_value=0.0)
        tal_sales_by_sm = tal_sales_by_sm.reindex(idx_union, fill_value=0.0)

        # Only count targets for salesmen who have sales in this period
        active_sm = sales_by_sm[sales_by_sm > 0].index.tolist()
        active_ka_target = float(ka_targets.loc[active_sm].sum()) if len(active_sm) > 0 else 0.0
        active_tal_target = float(tal_targets.loc[active_sm].sum()) if len(active_sm) > 0 else 0.0

        total_tal_sales = float(tal_sales_by_sm.sum())

        ka_ach_pct = (total_sales / active_ka_target * 100) if active_ka_target > 0 else 0
        tal_ach_pct = (total_tal_sales / active_tal_target * 100) if active_tal_target > 0 else 0

        # Channels: Market vs E-com (sales channels sheet)
        total_ecom = 0.0
        total_market = 0.0
        if not channels_df.empty and {"PY Name 1", "Channels"}.issubset(channels_df.columns):
            df_py_sales_ai = df_ai.groupby("PY Name 1")["Net Value"].sum().reset_index()
            df_ch_merge = df_py_sales_ai.merge(
                channels_df[["PY Name 1", "Channels"]],
                on="PY Name 1",
                how="left"
            )
            df_ch_merge["Channels"] = df_ch_merge["Channels"].astype(str).str.strip().str.lower().fillna("market")
            ch_sales = df_ch_merge.groupby("Channels")["Net Value"].sum()
            total_ecom = float(ch_sales.get("e-com", 0.0))
            total_market = float(ch_sales.sum() - total_ecom)

        total_channels = total_market + total_ecom
        ecom_pct = (total_ecom / total_channels * 100) if total_channels > 0 else 0
        market_pct = (total_market / total_channels * 100) if total_channels > 0 else 0

        # Returns (YKRE + ZRE)
        if "Billing Type" in df_ai.columns:
            returns_df = df_ai[df_ai["Billing Type"].isin(["YKRE", "ZRE"])]
            total_returns = float(returns_df["Net Value"].sum())
        else:
            total_returns = 0.0
        returns_pct = (total_returns / total_sales * 100) if total_sales > 0 else 0

        # KPI cards
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Total Sales", fmt_kd(total_sales))
        k2.metric("KA Achievement %", f"{ka_ach_pct:.0f}%" if active_ka_target > 0 else "N/A")
        k3.metric("Talabat Achievement %", f"{tal_ach_pct:.0f}%" if active_tal_target > 0 else "N/A")
        k4.metric("Returns % of Sales", f"{returns_pct:.1f}%" if total_sales > 0 else "N/A")

        c1, c2 = st.columns(2)
        if total_channels > 0:
            c1.metric("Market Sales", f"{fmt_kd(total_market)} ({market_pct:.0f}%)")
            c2.metric("E-com Sales", f"{fmt_kd(total_ecom)} ({ecom_pct:.0f}%)")
        else:
            c1.metric("Market Sales", fmt_kd(total_market))
            c2.metric("E-com Sales", fmt_kd(total_ecom))

        # ------------------------------------------------
        # 4) YTD Growth Pulse (simple, from YTD sheet)
        # ------------------------------------------------
        st.markdown("---")
        st.subheader("📈 YTD Growth Pulse")

        def ytd_pulse_simple(df_ytd: pd.DataFrame):
            if df_ytd.empty:
                return None

            df_ytd = df_ytd.copy()
            df_ytd["Billing Date"] = pd.to_datetime(df_ytd["Billing Date"])

            effective_today = df_ytd["Billing Date"].max()
            p2_end = effective_today
            p2_start = effective_today - pd.Timedelta(days=30)
            p1_end = p2_start
            p1_start = p1_end - pd.Timedelta(days=30)

            df_p1 = df_ytd[(df_ytd["Billing Date"] >= p1_start) & (df_ytd["Billing Date"] < p1_end)]
            df_p2 = df_ytd[(df_ytd["Billing Date"] >= p2_start) & (df_ytd["Billing Date"] <= p2_end)]

            total_p1 = float(df_p1["Net Value"].sum())
            total_p2 = float(df_p2["Net Value"].sum())
            diff = total_p2 - total_p1
            pct = (diff / total_p1 * 100) if total_p1 else None

            current_year = effective_today.year
            prev_year = current_year - 1
            ytd_curr = float(df_ytd[df_ytd["Billing Date"].dt.year == current_year]["Net Value"].sum())
            ytd_prev = float(df_ytd[df_ytd["Billing Date"].dt.year == prev_year]["Net Value"].sum())
            yoy_diff = ytd_curr - ytd_prev
            yoy_pct = (yoy_diff / ytd_prev * 100) if ytd_prev else None

            return dict(
                p1_start=p1_start.date(),
                p1_end=p1_end.date(),
                p2_start=p2_start.date(),
                p2_end=p2_end.date(),
                total_p1=total_p1,
                total_p2=total_p2,
                diff=diff,
                pct=pct,
                ytd_curr=ytd_curr,
                ytd_prev=ytd_prev,
                yoy_diff=yoy_diff,
                yoy_pct=yoy_pct,
                current_year=current_year,
                prev_year=prev_year
            )

        ytd_info = ytd_pulse_simple(ytd_df) if not ytd_df.empty else None

        if ytd_info:
            col_y1, col_y2, col_y3 = st.columns(3)
            col_y1.metric(
                "Last 30d Sales",
                fmt_kd(ytd_info["total_p2"]),
                delta=f"{(ytd_info['pct'] or 0):.1f}% vs prev 30d" if ytd_info["pct"] is not None else None
            )
            col_y2.metric(
                f"YTD {ytd_info['current_year']}",
                fmt_kd(ytd_info["ytd_curr"]),
                delta=f"{(ytd_info['yoy_pct'] or 0):.1f}% vs {ytd_info['prev_year']}" if ytd_info["yoy_pct"] is not None else None
            )
            col_y3.metric(
                "YTD Δ Value",
                fmt_kd(ytd_info["yoy_diff"]),
                delta=None
            )
        else:
            st.info("YTD data not available or not enough history for pulse.")

        # ------------------------------------------------
        # 5) Top Salesmen & Customers (GM view, filtered period)
        # ------------------------------------------------
        st.markdown("---")
        st.subheader("🏆 Drivers & 🔻 Risks")

        # Full sorted lists (GM can scroll)
        top_sm = sales_by_sm.sort_values(ascending=False)
        cust_sales = df_ai.groupby("PY Name 1")["Net Value"].sum().sort_values(ascending=False)

        col_t1, col_t2 = st.columns(2)

        # --- Top Salesmen
        with col_t1:
            st.markdown("**Salesmen – Sales & KA Achievement (Filtered Period)**")
            sm_df = pd.DataFrame({
                "Salesman": top_sm.index,
                "Sales": top_sm.values
            })
            if not ka_targets.empty:
                sm_df["KA Target"] = sm_df["Salesman"].map(ka_targets).fillna(0.0)
                sm_df["KA % Achieved"] = np.where(
                    sm_df["KA Target"] > 0,
                    (sm_df["Sales"] / sm_df["KA Target"] * 100).round(0),
                    0
                )

            styled_sm = sm_df.style.set_table_styles([{
                "selector": "th",
                "props": [
                    ("background", "#1E3A8A"),
                    ("color", "white"),
                    ("font-weight", "800"),
                    ("height", "40px"),
                    ("line-height", "40px"),
                    ("border", "1px solid #E5E7EB"),
                    ("text-align", "center")
                ],
            }])

            num_cols_sm = [c for c in ["Sales", "KA Target", "KA % Achieved"] if c in sm_df.columns]
            if "Sales" in num_cols_sm or "KA Target" in num_cols_sm:
                fmt_map = {}
                if "Sales" in num_cols_sm:
                    fmt_map["Sales"] = "{:,.0f}".format
                if "KA Target" in num_cols_sm:
                    fmt_map["KA Target"] = "{:,.0f}".format
                styled_sm = styled_sm.format(fmt_map)
            if "KA % Achieved" in sm_df.columns:
                styled_sm = styled_sm.format("{:.0f}%", subset=["KA % Achieved"])

            st.dataframe(styled_sm, use_container_width=True, hide_index=True, height=320)

        # --- Top Customers
        with col_t2:
            st.markdown("**Customers – Sales & Returns (Filtered Period)**")
            cust_df = pd.DataFrame({
                "Customer Name": cust_sales.index,
                "Sales": cust_sales.values
            })
            if "Billing Type" in df_ai.columns:
                ret_by_cust = df_ai[df_ai["Billing Type"].isin(["YKRE", "ZRE"])] \
                    .groupby("PY Name 1")["Net Value"].sum()
                cust_df["Returns"] = cust_df["Customer Name"].map(ret_by_cust).fillna(0.0)
                cust_df["Return %"] = np.where(
                    cust_df["Sales"] > 0,
                    (cust_df["Returns"] / cust_df["Sales"] * 100).round(1),
                    0
                )

            styled_cust = cust_df.style.set_table_styles([{
                "selector": "th",
                "props": [
                    ("background", "#1E3A8A"),
                    ("color", "white"),
                    ("font-weight", "800"),
                    ("height", "40px"),
                    ("line-height", "40px"),
                    ("border", "1px solid #E5E7EB"),
                    ("text-align", "center")
                ],
            }])

            num_cols_c = [c for c in ["Sales", "Returns", "Return %"] if c in cust_df.columns]
            if "Sales" in num_cols_c:
                styled_cust = styled_cust.format("{:,.0f}", subset=["Sales"])
            if "Returns" in num_cols_c:
                styled_cust = styled_cust.format("{:,.0f}", subset=["Returns"])
            if "Return %" in num_cols_c:
                styled_cust = styled_cust.format("{:.1f}%", subset=["Return %"])

            st.dataframe(styled_cust, use_container_width=True, hide_index=True, height=320)

        # ------------------------------------------------
        # 6) Risk & Alert Panel for GM
        # ------------------------------------------------
        st.markdown("---")
        st.subheader("🚨 GM Risk Radar & Actions")

        alert_lines = []

        if active_ka_target > 0 and ka_ach_pct < 90:
            alert_lines.append(
                f"KA achievement is only {ka_ach_pct:.0f}% – risk of missing KA target. Push high-potential customers and strong SKUs."
            )
        if active_tal_target > 0 and tal_ach_pct < 90:
            alert_lines.append(
                f"Talabat (E-com) achievement is {tal_ach_pct:.0f}% – review assortment, deals and visibility in e-com."
            )
        if returns_pct > 3:
            alert_lines.append(
                f"Returns are {returns_pct:.1f}% of sales – check top return customers, ageing stock and handling."
            )
        if ytd_info and ytd_info["yoy_pct"] is not None and ytd_info["yoy_pct"] < 0:
            alert_lines.append(
                f"Negative YTD YoY growth ({ytd_info['yoy_pct']:.1f}%) – need recovery plan with promotions and new listings."
            )
        if ytd_info and ytd_info["pct"] is not None and ytd_info["pct"] < 0:
            alert_lines.append(
                f"Last 30 days (YTD sheet) are lower than previous 30 days by {fmt_kd(abs(ytd_info['diff']))} ({ytd_info['pct']:.1f}%)."
            )

        if alert_lines:
            for a in alert_lines:
                st.write("• " + a)
        else:
            st.success("✅ No major red flags detected in this period. Keep current direction and monitor weekly.")

        # ------------------------------------------------
        # 7) Material Heatmap – Material x Salesman (Net Sales)
        # ------------------------------------------------
        st.markdown("---")
        st.subheader("🧩 Material Heatmap – Material x Salesman (Net Sales)")

        if "Driver Name EN" in df_ai.columns and "Material Description" in df_ai.columns:
            mat_pivot = df_ai.pivot_table(
                index="Material Description",
                columns="Driver Name EN",
                values="Net Value",
                aggfunc="sum",
                fill_value=0.0
            )

            # Sort materials by total sales (top SKUs first)
            row_totals = mat_pivot.sum(axis=1).sort_values(ascending=False)
            mat_pivot = mat_pivot.loc[row_totals.index]

            def material_heatmap(df_vals):
                styles = pd.DataFrame("", index=df_vals.index, columns=df_vals.columns)
                row_max = df_vals.max(axis=1).replace(0, np.nan)

                for mat in df_vals.index:          # each material (row)
                    max_val = row_max.loc[mat]
                    for sm in df_vals.columns:     # each salesman (column)
                        v = df_vals.loc[mat, sm]
                        if pd.isna(max_val) or v <= 0:
                            styles.loc[mat, sm] = "background-color: #F9FAFB; color: #9CA3AF"
                        else:
                            ratio = v / max_val
                            if ratio >= 0.7:
                                styles.loc[mat, sm] = "background-color: #D1FAE5; color: #065F46; font-weight: 600"
                            elif ratio >= 0.3:
                                styles.loc[mat, sm] = "background-color: #FEF3C7; color: #92400E; font-weight: 500"
                            else:
                                styles.loc[mat, sm] = "background-color: #FEE2E2; color: #991B1B"
                return styles

            styled_mat = (
                mat_pivot.style
                .set_table_styles([{
                    "selector": "th",
                    "props": [
                        ("background", "#1F2937"),
                        ("color", "white"),
                        ("font-weight", "800"),
                        ("height", "38px"),
                        ("line-height", "38px"),
                        ("border", "1px solid #E5E7EB"),
                        ("text-align", "center")
                    ],
                }])
                .apply(material_heatmap, axis=None)
                .format("{:,.0f}")
            )

            st.caption(
                "Rows = Materials, Columns = Salesmen – "
                "🟢 strong vs best salesman for that SKU, 🟡 medium, 🔴 weak, grey = no sales."
            )
            st.dataframe(styled_mat, use_container_width=True, height=320, hide_index=True)
        else:
            st.info("Material Heatmap not available – 'Driver Name EN' or 'Material Description' missing in Sales sheet.")

        # ------------------------------------------------
        # 8) Sales Trend (Actual vs 7-day Avg)
        # ------------------------------------------------
        st.markdown("---")
        st.subheader("📉 Sales Trend (Selected Period)")

        ts = df_ai.copy()
        ts["Billing Date"] = pd.to_datetime(ts["Billing Date"])
        ts_daily = ts.groupby("Billing Date")["Net Value"].sum().reset_index()
        ts_daily = ts_daily.sort_values("Billing Date")
        ts_daily["7D Avg"] = ts_daily["Net Value"].rolling(window=7, min_periods=1).mean()

        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=ts_daily["Billing Date"],
            y=ts_daily["Net Value"],
            mode="lines+markers",
            name="Daily Sales",
            line=dict(width=2)
        ))
        fig.add_trace(go.Scatter(
            x=ts_daily["Billing Date"],
            y=ts_daily["7D Avg"],
            mode="lines",
            name="7-Day Average",
            line=dict(dash="dash", width=2)
        ))
        fig.update_layout(
            xaxis_title="Date",
            yaxis_title="Net Value (KD)",
            hovermode="x unified"
        )
        st.plotly_chart(fig, use_container_width=True)

        # ------------------------------------------------
        # 9) GM Summary Table + Downloads (Excel + PPT)
        # ------------------------------------------------
        st.markdown("---")
        st.subheader("📥 GM Summary – Downloadable Report")

        gm_rows = [
            {"Metric": "Analysis Period", "Value": f"{ai_start.date()} to {ai_end.date()}"},
            {"Metric": "Total Sales", "Value": fmt_kd(total_sales)},
            {"Metric": "KA Target (Active Salesmen)", "Value": fmt_kd(active_ka_target)},
            {"Metric": "KA Achievement %", "Value": f"{ka_ach_pct:.0f}%" if active_ka_target > 0 else "N/A"},
            {"Metric": "Talabat Sales (E-com)", "Value": fmt_kd(total_tal_sales)},
            {"Metric": "Talabat Target (Active Salesmen)", "Value": fmt_kd(active_tal_target)},
            {"Metric": "Talabat Achievement %", "Value": f"{tal_ach_pct:.0f}%" if active_tal_target > 0 else "N/A"},
            {"Metric": "Market Sales", "Value": fmt_kd(total_market)},
            {"Metric": "E-com Sales", "Value": fmt_kd(total_ecom)},
            {"Metric": "Returns Value", "Value": fmt_kd(total_returns)},
            {"Metric": "Returns % of Sales", "Value": f"{returns_pct:.1f}%" if total_sales > 0 else "N/A"},
        ]

        if ytd_info:
            gm_rows.append({
                "Metric": "Last 30d vs Prev 30d (YTD sheet)",
                "Value": f"{fmt_kd(ytd_info['total_p2'])} vs {fmt_kd(ytd_info['total_p1'])} ({(ytd_info['pct'] or 0):.1f}%)"
            })
            gm_rows.append({
                "Metric": f"YTD {ytd_info['current_year']} vs {ytd_info['prev_year']}",
                "Value": f"{fmt_kd(ytd_info['ytd_curr'])} vs {fmt_kd(ytd_info['ytd_prev'])} ({(ytd_info['yoy_pct'] or 0):.1f}%)"
            })

        gm_df = pd.DataFrame(gm_rows)

        styled_gm = gm_df.style.set_table_styles([{
            "selector": "th",
            "props": [
                ("background", "#1E3A8A"),
                ("color", "white"),
                ("font-weight", "800"),
                ("height", "40px"),
                ("line-height", "40px"),
                ("border", "1px solid #E5E7EB"),
                ("text-align", "center")
            ],
        }])
        st.dataframe(styled_gm, use_container_width=True, hide_index=True)

        # Excel export
        gm_excel_bytes = to_excel_bytes(gm_df, sheet_name="GM_Executive_Summary", index=False)
        st.download_button(
            "⬇️ Download GM Summary (Excel)",
            data=gm_excel_bytes,
            file_name=f"GM_Executive_Summary_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # PPT export (simple 1-slide deck)
        def build_gm_ppt(gm_df, alerts):
            prs = Presentation()
            slide = prs.slides.add_slide(prs.slide_layouts[5])  # blank

            # Title
            title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(9), Inches(1))
            tf = title_box.text_frame
            tf.text = "GM Executive Sales Summary"
            tf.paragraphs[0].font.size = Pt(28)
            tf.paragraphs[0].font.bold = True

            # KPIs
            body_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.4), Inches(9), Inches(3.2))
            bf = body_box.text_frame
            bf.word_wrap = True
            p0 = bf.paragraphs[0]
            p0.text = "Key Metrics:"
            p0.font.bold = True
            p0.font.size = Pt(18)

            for _, row in gm_df.iterrows():
                p = bf.add_paragraph()
                p.text = f"- {row['Metric']}: {row['Value']}"
                p.level = 1
                p.font.size = Pt(14)

            # Alerts
            alert_box = slide.shapes.add_textbox(Inches(0.5), Inches(4.8), Inches(9), Inches(2.2))
            af = alert_box.text_frame
            af.word_wrap = True
            pa0 = af.paragraphs[0]
            pa0.text = "Key Alerts / Actions:"
            pa0.font.bold = True
            pa0.font.size = Pt(18)

            if alerts:
                for a in alerts:
                    pa = af.add_paragraph()
                    pa.text = f"- {a}"
                    pa.level = 1
                    pa.font.size = Pt(14)
            else:
                pa = af.add_paragraph()
                pa.text = "- No critical alerts in this period."
                pa.level = 1
                pa.font.size = Pt(14)

            bio = io.BytesIO()
            prs.save(bio)
            bio.seek(0)
            return bio.getvalue()

        gm_ppt_bytes = build_gm_ppt(gm_df, alert_lines)

        st.download_button(
            "⬇️ Download GM Snapshot (PPTX)",
            data=gm_ppt_bytes,
            file_name=f"GM_Executive_Snapshot_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
            
# ---------------- Customer Insights Page (Full Pro Version, Fixed) ----------------
elif choice == texts[lang]["customer_insights"]:
    st.title(texts[lang]["customer_insights_title"])

    # Pre-checks
    if "data_loaded" not in st.session_state:
        st.warning(texts[lang]["no_data_warning"])
        st.stop()

    df_rfm = st.session_state["sales_df"].copy()
    if df_rfm.empty:
        st.warning(texts[lang]["rfm_no_data"])
        st.stop()

    # Apply salesman filter if applicable
    if user_role == "salesman" and salesman_name:
        df_rfm = df_rfm[df_rfm["Driver Name EN"] == salesman_name]

    # Detect columns robustly
    def find_col(df, candidates):
        for n in candidates:
            if n in df.columns:
                return n
        for c in df.columns:
            lc = c.lower()
            for n in candidates:
                if n.lower() in lc:
                    return c
        return None

    cust_col = find_col(df_rfm, ["SP Name1", "SP Name 1", "SP_Name1", "Customer", "PY Name1", "PY Name 1"])
    date_col = find_col(df_rfm, ["Billing Date", "billing date", "Date"])
    amount_col = find_col(df_rfm, ["Net Value", "NetAmount", "Net Amount", "Amount", "Sales Amount"])
    material_col = find_col(df_rfm, ["Material Description", "Item", "Product", "Material"])

    if None in [cust_col, date_col, amount_col]:
        st.warning(texts[lang]["rfm_no_data"])
        st.stop()

    # Normalize date col
    df_rfm[date_col] = pd.to_datetime(df_rfm[date_col], errors="coerce")
    today = pd.Timestamp.today().normalize()

    # --- Fixed robust RFM aggregation ---
    rfm_group = df_rfm.groupby(cust_col)

    rfm_agg = pd.DataFrame({
        "Customer": rfm_group.apply(lambda g: g.name),
        "Recency": rfm_group[date_col].max().apply(lambda d: (today - d).days),
        "Frequency": rfm_group[date_col].count(),
        "Monetary": rfm_group[amount_col].sum()
    }).reset_index(drop=True)

    rfm_agg = rfm_agg[rfm_agg["Monetary"] > 0].set_index("Customer")

    if rfm_agg.empty:
        st.warning(texts[lang]["rfm_no_data"])
        st.stop()

    # Safe qcut as before
    def safe_qcut(series, q=4, reverse=False):
        s = series.copy().fillna(series.max() + 1)
        unique_vals = pd.unique(s)
        n_unique = len(unique_vals)
        if n_unique == 1:
            return pd.Series([1]*len(s), index=s.index)
        if n_unique < q:
            ranks = s.rank(method='dense', ascending=not reverse)
            return ranks.astype(int)
        labels = list(range(q, 0, -1)) if reverse else list(range(1, q+1))
        try:
            return pd.qcut(s, q=q, labels=labels, duplicates='drop')
        except Exception:
            ranks = s.rank(method='dense', ascending=not reverse)
            return ranks.astype(int)

    rfm_agg["R_Score"] = safe_qcut(rfm_agg["Recency"], q=4, reverse=True).astype(int)
    rfm_agg["F_Score"] = safe_qcut(rfm_agg["Frequency"], q=4).astype(int)
    rfm_agg["M_Score"] = safe_qcut(rfm_agg["Monetary"], q=4).astype(int)
    rfm_agg["RFM_Score"] = rfm_agg["R_Score"].astype(str) + rfm_agg["F_Score"].astype(str) + rfm_agg["M_Score"].astype(str)

    # Segmentation
    def rfm_segment(row):
        if row["RFM_Score"] in ["444", "443", "434", "433"]:
            return "Champions"
        if row["R_Score"] >= 3 and row["F_Score"] >= 3:
            return "Loyal Customers"
        if row["R_Score"] >= 3 and row["M_Score"] >= 3:
            return "Potential Loyalists"
        if row["R_Score"] >= 3:
            return "New Customers"
        if row["R_Score"] <= 2 and row["F_Score"] >= 2 and row["M_Score"] >= 2:
            return "At Risk"
        if row["R_Score"] <= 1 and row["F_Score"] >= 2 and row["M_Score"] >= 2:
            return "Hibernating"
        return "Others"

    rfm_agg["Segment"] = rfm_agg.apply(rfm_segment, axis=1)

    # --- Layout Tabs: RFM, Cohort, Weekly CRM ---
    tab_weekly, tab_360 = st.tabs([
    # texts[lang]["rfm_analysis_sub"], 
    # texts[lang]["rfm_cohort_sub"], 
    "CRM & Weekly Operations",
    "Customer 360°"  # ← NEW TAB
])

    # # ---------------- RFM Tab ----------------
    # with tab_rfm:
    #     st.subheader(texts[lang]["rfm_table_sub"])
    #     display_rfm = rfm_agg.copy()
    #     display_rfm[["Recency","Frequency","Monetary"]] = display_rfm[["Recency","Frequency","Monetary"]].astype(int)
    #     st.dataframe(display_rfm.sort_values("Monetary", ascending=False), use_container_width=True, hide_index=True)

    #     # Download with safe sheet name
    #     ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    #     safe_sheet = "RFM_Analysis"[:31]
    #     if st.download_button(
    #         texts[lang]["rfm_download"],
    #         data=to_excel_bytes(display_rfm.reset_index(), sheet_name=safe_sheet),
    #         file_name=f"rfm_analysis_{ts}.xlsx"
    #     ):
    #         st.session_state["audit_log"].append({"user": username, "action":"download","details":f"rfm_analysis_{ts}.xlsx","timestamp":datetime.now().isoformat()})

    #     # Segment Pie + metrics
    #     st.subheader("RFM Segment Distribution")
    #     seg_counts = display_rfm["Segment"].value_counts().reset_index()
    #     seg_counts.columns = ["Segment", "Count"]
    #     seg_counts["Percentage"] = (seg_counts["Count"]/seg_counts["Count"].sum()*100).round(1)
    #     seg_avg = display_rfm.groupby("Segment")["Monetary"].mean().round(2).reset_index().rename(columns={"Monetary":"Avg Monetary"})
    #     seg_counts = seg_counts.merge(seg_avg, on="Segment", how="left")
    #     fig_seg = px.pie(seg_counts, names="Segment", values="Count", hole=0.35, hover_data=["Percentage","Avg Monetary"], title="RFM Segment Distribution")
    #     fig_seg.update_traces(textinfo='percent+label')
    #     st.plotly_chart(fig_seg, use_container_width=True)

    #     st.subheader("Key Metrics per Segment")
    #     seg_metrics = display_rfm.groupby("Segment").agg(
    #         mean_recency=("Recency","mean"),
    #         mean_frequency=("Frequency","mean"),
    #         mean_monetary=("Monetary","mean"),
    #         count=("R_Score","count")
    #     ).round(2).rename(columns={"count":"Count"})
    #     st.dataframe(seg_metrics, use_container_width=True, hide_index=True)

    #     st.subheader("Prescriptive Actions per Segment")
    #     recs = {
    #         "Champions":"Reward with exclusive offers & loyalty programs.",
    #         "Loyal Customers":"Upsell & referrals.",
    #         "Potential Loyalists":"Nurture with targeted campaigns.",
    #         "New Customers":"Onboard & incentivize repeat purchase.",
    #         "At Risk":"Win-back campaigns and surveys.",
    #         "Hibernating":"Reactivate with promotions.",
    #         "Others":"Investigate further."
    #     }
    #     for s in seg_metrics.index:
    #         st.write(f"- **{s}**: {recs.get(s,'General engagement strategies recommended.')}")

    #     st.subheader(texts[lang]["rfm_chart_sub"])
    #     fig_rfm = px.scatter(display_rfm.reset_index(), x="Recency", y="Monetary", size="Frequency", color="Segment",
    #                          hover_name=display_rfm.reset_index()["Customer"], title="RFM Scatter (Recency vs Monetary; size=Frequency)")
    #     st.plotly_chart(fig_rfm, use_container_width=True)

    # # ---------------- Cohort Tab (Fixed) ----------------
    # with tab_cohort:
    #     st.subheader(texts[lang]["rfm_cohort_sub"])
    #     st.info(texts[lang]["rfm_cohort_info"])

    #     df_cohort = df_rfm[[cust_col, date_col, amount_col]].dropna().copy().rename(columns={cust_col:"Customer", date_col:"Billing Date"})

    #     # Convert periods to strings
    #     df_cohort["Cohort_Month_Str"] = df_cohort.groupby("Customer")["Billing Date"].transform("min").dt.to_period("M").astype(str)
    #     df_cohort["Period_Month_Str"] = df_cohort["Billing Date"].dt.to_period("M").astype(str)
    #     df_cohort["Cohort_Index"] = (pd.to_datetime(df_cohort["Period_Month_Str"]) - pd.to_datetime(df_cohort["Cohort_Month_Str"])).dt.days // 30

    #     cohort_summary = df_cohort.groupby(["Cohort_Month_Str","Cohort_Index"]).agg(
    #         Customer=("Customer","nunique"),
    #         Monetary=(amount_col,"mean")
    #     ).reset_index()

    #     if not cohort_summary.empty:
    #         cohort_pivot = cohort_summary.pivot(index="Cohort_Month_Str", columns="Cohort_Index", values="Monetary").fillna(0)
    #         fig_cohort = px.imshow(cohort_pivot, labels=dict(x="Months after acquisition", y="Cohort", color="Avg Monetary"),
    #                                title="Cohort Monetary Heatmap", text_auto=True, aspect="auto")
    #         st.plotly_chart(fig_cohort, use_container_width=True)

    #         st.subheader(texts[lang]["rfm_cohort_table_sub"])
    #         cohort_table = cohort_summary.pivot(index="Cohort_Month_Str", columns="Cohort_Index", values="Customer").fillna(0).astype(int)
    #         st.dataframe(cohort_table, use_container_width=True, hide_index=True)
    #     else:
    #         st.warning(texts[lang]["rfm_cohort_no_data"])

    # ---------------- CRM & Weekly Operations Tab ----------------
    with tab_weekly:
        st.subheader("CRM Operations — Weekly Tracker, Products & Manager Dashboard")

        # Date selector
        col_left, col_right = st.columns([3,1])
        with col_left:
            auto_date = st.toggle("Use current date automatically", value=True)
            selected_date = datetime.now().date() if auto_date else st.date_input("Select visit date manually", datetime.now().date())
            st.session_state["visit_date"] = selected_date

        with col_right:
            show_manager = st.checkbox("Show Manager KPIs", value=True)
            refresh = st.button("🔄 Refresh")  # triggers rerun

        # Prepare commonly used variables
        ytd_df = st.session_state.get("ytd_df", pd.DataFrame()).copy()
        sales_df = st.session_state.get("sales_df", pd.DataFrame()).copy()

        # Ensure date cols
        if date_col in ytd_df.columns:
            ytd_df[date_col] = pd.to_datetime(ytd_df[date_col], errors="coerce")
        if date_col in sales_df.columns:
            sales_df[date_col] = pd.to_datetime(sales_df[date_col], errors="coerce")

        # ---------------- Weekly Visit Tracker (robust) ----------------
        st.markdown("### Weekly Visit Tracker")
        last_3_months = pd.Timestamp(selected_date) - pd.DateOffset(months=3)
        recent_ytd = ytd_df[ytd_df.get(date_col, pd.Series()) >= last_3_months] if not ytd_df.empty else pd.DataFrame()
        customer_list = pd.Series(recent_ytd.get(cust_col, pd.Series()).dropna().unique()).astype(str) if not recent_ytd.empty else pd.Series(sales_df.get(cust_col, pd.Series()).dropna().unique()).astype(str)

        if customer_list.empty:
            st.info("No customers found in YTD or Sales for the last 3 months.")
        else:
            # last 7-day window
            days_dt = [pd.Timestamp(selected_date) - pd.Timedelta(days=i) for i in range(6, -1, -1)]
            days_str = [d.strftime("%Y-%m-%d") for d in days_dt]
            sales_window_start = days_dt[0]
            sales_window_end = pd.Timestamp(selected_date) + pd.Timedelta(days=1)

            sales_last7 = sales_df[(sales_df[date_col] >= sales_window_start) & (sales_df[date_col] < sales_window_end)].copy()
            sales_last7[cust_col] = sales_last7[cust_col].astype(str)
            sales_last7[amount_col] = pd.to_numeric(sales_last7[amount_col], errors="coerce").fillna(0.0)
            sales_last7["__date_str"] = sales_last7[date_col].dt.strftime("%Y-%m-%d")

            pivot7 = (sales_last7.groupby([cust_col, "__date_str"])[amount_col].sum().reset_index()
                      .pivot(index=cust_col, columns="__date_str", values=amount_col).reindex(columns=days_str, fill_value=0.0).reset_index())
            pivot7 = pivot7.rename(columns={cust_col: "Customer"})
            base = pd.DataFrame({"Customer": customer_list})
            visit_df = base.merge(pivot7, on="Customer", how="left").fillna(0.0)
            visit_df.insert(1, "Visit Date", selected_date)

            existing_days = [c for c in days_str if c in visit_df.columns]
            visit_df["Weekly Total"] = visit_df[existing_days].sum(axis=1) if existing_days else 0

            # 4-week totals
            end_date = pd.Timestamp(selected_date)
            start_date = end_date - pd.Timedelta(weeks=4)
            recent_sales = sales_df[(sales_df[date_col] >= start_date) & (sales_df[date_col] <= end_date)].copy()
            recent_sales[cust_col] = recent_sales[cust_col].astype(str)
            recent_sales[amount_col] = pd.to_numeric(recent_sales[amount_col], errors="coerce").fillna(0.0)
            recent_sales["Week_Number"] = ((recent_sales[date_col] - start_date).dt.days // 7) + 1
            recent_sales.loc[recent_sales["Week_Number"] > 4, "Week_Number"] = 4
            week_totals = (recent_sales.groupby([cust_col, "Week_Number"])[amount_col].sum().unstack(fill_value=0).reset_index().rename(columns={cust_col:"Customer"}))
            week_cols = [c for c in week_totals.columns if c != "Customer"]
            if not week_cols:
                for i in range(1,5):
                    week_totals[i] = 0
                week_cols = [1,2,3,4]
            ordered_week_cols = sorted(week_cols, key=lambda x:int(x))
            week_totals = week_totals[["Customer"] + ordered_week_cols]
            week_headers = []
            for i in range(len(ordered_week_cols)):
                wstart = (start_date + pd.Timedelta(days=i*7)).strftime("%b %d")
                wend = (start_date + pd.Timedelta(days=(i+1)*7-1)).strftime("%b %d")
                week_headers.append(f"Week {i+1} ({wstart}-{wend})")
            new_cols = ["Customer"] + week_headers
            if len(new_cols) != len(week_totals.columns):
                new_cols = ["Customer"] + [f"Week {i+1}" for i in range(len(week_totals.columns)-1)]
            week_totals.columns = new_cols
            visit_df = visit_df.merge(week_totals, on="Customer", how="left").fillna(0.0)

            # Total Sales (3 months)
            total_sales_3m = sales_df[sales_df[date_col] >= last_3_months].groupby(cust_col)[amount_col].sum().reset_index().rename(columns={cust_col:"Customer", amount_col:"Total Sales"})
            visit_df = visit_df.merge(total_sales_3m, on="Customer", how="left").fillna(0.0)

            # Alerts & recommended action
            def compute_alert(row):
                q80 = visit_df["Total Sales"].quantile(0.8) if "Total Sales" in visit_df.columns else 0
                q50 = visit_df["Total Sales"].quantile(0.5) if "Total Sales" in visit_df.columns else 0
                if row.get("Weekly Total",0) == 0:
                    if row.get("Total Sales",0) >= q80:
                        return "🔴 High", "Visit immediately"
                    if row.get("Total Sales",0) >= q50:
                        return "🟠 Medium", "Call / Email"
                    return "🟢 Low", "Monitor"
                return "✅ Visited", "No action"
            visit_df[["Alert Level","Recommended Action"]] = visit_df.apply(lambda r: pd.Series(compute_alert(r)), axis=1)

            # numeric formatting
            numeric_cols = [c for c in existing_days + ["Weekly Total","Total Sales"] + week_headers if c in visit_df.columns]
            if numeric_cols:
                visit_df[numeric_cols] = visit_df[numeric_cols].fillna(0).round(0).astype(int)

            # show table
            with st.expander("Show Weekly Visit Table", expanded=True):
                st.dataframe(visit_df.sort_values(["Total Sales","Weekly Total"], ascending=[False,False]), use_container_width=True, hide_index=True)

            # manager KPIs
            if show_manager:
                st.markdown("### Manager Dashboard")
                col1, col2, col3, col4 = st.columns(4)
                last7_revenue = sales_last7[amount_col].sum() if not sales_last7.empty else 0
                col1.metric("Revenue (Last 7 days)", f"KD {last7_revenue:,.0f}")
                col2.metric("Customers Visited (7d)", int((visit_df["Weekly Total"]>0).sum()))
                col3.metric("High-Value Missed", int((visit_df["Alert Level"]=="🔴 High").sum()))
                col4.metric("Avg Weekly Sales per Customer", f"KD {visit_df['Weekly Total'].mean():,.0f}")

                top_missed = visit_df[visit_df["Alert Level"]=="🔴 High"].sort_values("Total Sales", ascending=False).head(10)
                if not top_missed.empty:
                    st.subheader("Top High-Value Missed Customers")
                    st.bar_chart(top_missed.set_index("Customer")["Total Sales"])

            # Save Visit as record (mini CRM)
            st.markdown("### Visit Planner / Notes")
            planner_customer = st.selectbox("Select Customer to plan visit", options=sorted(visit_df["Customer"].astype(str).unique()))
            note = st.text_area("Note / Follow-up", key=f"note_{planner_customer}")
            next_visit = st.date_input("Next Visit Date", value=(pd.Timestamp(selected_date)+pd.Timedelta(days=7)).date(), key=f"nextvisit_{planner_customer}")
            if st.button("💾 Save Visit Plan"):
                if "visit_plans" not in st.session_state:
                    st.session_state["visit_plans"] = []
                st.session_state["visit_plans"].append({
                    "Customer": planner_customer,
                    "Next Visit": next_visit.isoformat() if hasattr(next_visit, "isoformat") else str(next_visit),
                    "Note": note,
                    "Created By": username,
                    "Created At": datetime.now().isoformat()
                })
                st.success(f"Saved visit plan for {planner_customer}")

            if "visit_plans" in st.session_state and st.session_state["visit_plans"]:
                st.subheader("Saved Visit Plans")
                st.dataframe(pd.DataFrame(st.session_state["visit_plans"]), use_container_width=True, hide_index=True)
                # Download visit plans
                safe_file_name = f"visit_plans_{selected_date}.csv"
                st.download_button("⬇️ Download Visit Plans", data=pd.DataFrame(st.session_state["visit_plans"]).to_csv(index=False).encode('utf-8'), file_name=safe_file_name, mime="text/csv")

        # ---------------- 15-Day Product Analysis ----------------
        st.markdown("### Customer Product Activity (Last 15 Days)")
        product_start_date = pd.Timestamp(selected_date) - pd.Timedelta(days=15)
        prod_sales = sales_df[(sales_df[date_col] >= product_start_date) & (sales_df[date_col] <= pd.Timestamp(selected_date))].copy()

        if prod_sales.empty or material_col is None:
            st.info("No product-sales data available for the last 15 days.")
        else:
            prod_sales[cust_col] = prod_sales[cust_col].astype(str)
            prod_sales[amount_col] = pd.to_numeric(prod_sales[amount_col], errors="coerce").fillna(0.0)
            all_products = sorted(sales_df[material_col].dropna().unique())

            prod_summary = prod_sales.groupby([cust_col, material_col])[amount_col].sum().reset_index().rename(columns={cust_col:"Customer", material_col:"Product", amount_col:"Sales Amount"})
            customers_prod = sorted(prod_summary["Customer"].unique())

            if customers_prod:
                sel_cust = st.selectbox("Select a Customer to inspect products", options=customers_prod)
                sold_by_cust = prod_summary[prod_summary["Customer"]==sel_cust].sort_values("Sales Amount", ascending=False)
                sold_set = set(sold_by_cust["Product"].dropna())
                not_sold = [p for p in all_products if p not in sold_set]
                df_not_sold = pd.DataFrame({"Product": not_sold, "Status":"❌ Not Purchased"})

                with st.expander(f"{sel_cust} - Purchased (Last 15 Days)", expanded=True):
                    if not sold_by_cust.empty:
                        st.dataframe(sold_by_cust, use_container_width=True, hide_index=True)
                    else:
                        st.info("No purchases by this customer in last 15 days.")


                with st.expander(f"{sel_cust} - Not Purchased (Last 15 Days)", expanded=False):
                    if not df_not_sold.empty:
                        st.dataframe(df_not_sold, use_container_width=True, hide_index=True)
                    else:
                        st.info("All products purchased!")


                # Safe sheet name
                safe_sheet_name = (sel_cust+"_Purchased")[:31]
                if st.download_button(f"⬇️ Download {sel_cust} Purchased (15d)",
                                      data=to_excel_bytes(sold_by_cust, sheet_name=safe_sheet_name),
                                      file_name=f"{sel_cust}_purchased_15days_{selected_date}.xlsx"):
                    st.success("Download ready!")

    # ──────────────────────── TAB 4: Customer 360° (FIXED) ────────────────────────
    with tab_360:
        st.markdown("### **Customer 360° – One-Click Profile**")
        
        # === 1. COLUMN DETECTION (Safe & Reusable) ===
        def find_col(df, candidates):
            for c in candidates:
                if c in df.columns:
                    return c
            return None

        cust_col = find_col(df_rfm, ["SP Name1", "SP Name 1", "Customer", "PY Name 1", "Customer Name"])
        date_col = find_col(df_rfm, ["Billing Date", "Date", "Invoice Date"])
        amount_col = find_col(df_rfm, ["Net Value", "Amount", "Sales"])
        driver_col = find_col(df_rfm, ["Driver Name EN", "Salesman", "Driver", "Rep"])
        material_col = find_col(df_rfm, ["Material", "Material Description", "Item", "SKU"])

        if not all([cust_col, date_col, amount_col]):
            st.error("Missing required columns. Check your Excel file.")
            st.stop()

        # === 2. CUSTOMER SELECTOR ===
        all_customers = sorted(df_rfm[cust_col].dropna().unique())
        selected_cust = st.selectbox("**Select Customer**", all_customers, key="cust360_select")

        # === 3. FILTER DATA ===
        cust_sales = df_rfm[df_rfm[cust_col] == selected_cust].copy()
        cust_sales[date_col] = pd.to_datetime(cust_sales[date_col], errors='coerce')
        cust_sales = cust_sales.dropna(subset=[date_col])  # Remove invalid dates
        today = pd.Timestamp.today()

        if cust_sales.empty:
            st.warning(f"No sales data for **{selected_cust}**")
        else:
            # ========================================
            # KPI CARDS – CUSTOMER 360°
            # ========================================
            col1, col2, col3, col4 = st.columns(4)
            
            # --- 1. ENSURE amount_col is numeric ---
            cust_sales[amount_col] = pd.to_numeric(cust_sales[amount_col], errors='coerce').fillna(0)

            # --- 2. NORMALIZE Billing Type ---
            if "Billing Type" in cust_sales.columns:
                cust_sales["Billing Type"] = cust_sales["Billing Type"].astype(str).str.strip().str.upper()

            # --- 3. TOTAL SALES & ORDERS ---
            total_sales = cust_sales[amount_col].sum()
            order_count = len(cust_sales)
            last_visit = cust_sales[date_col].max()
            days_since = (today - last_visit).days if pd.notna(last_visit) else 999

            # --- 4. RETURNS ONLY (YKRE, ZRE) – CANCELLATIONS EXCLUDED ---
            return_codes = ["YKRE", "ZRE"]  # Only Returns
            if "Billing Type" in cust_sales.columns:
                returns_mask = cust_sales["Billing Type"].isin(return_codes)
            else:
                returns_mask = cust_sales[date_col].notna() & False  # no returns if column missing
            returns_df = cust_sales[returns_mask]

            # --- 5. RETURN VALUE = ABSOLUTE SUM (handles negative values) ---
            returns_value = returns_df[amount_col].abs().sum()
            return_rate = (returns_value / total_sales * 100) if total_sales > 0 else 0

            # === KPI CARD 1: Total Sales ===
            col1.metric(
                label="**Total Sales**",
                value=f"KD {total_sales:,.0f}",
                delta=None
            )

            # === KPI CARD 2: Orders ===
            col2.metric(
                label="**Orders**",
                value=f"{order_count:,}",
                delta=None
            )

            # === KPI CARD 3: Last Visit ===
            if pd.notna(last_visit):
                col3.metric(
                    label="**Last Visit**",
                    value=last_visit.strftime("%b %d, %Y"),
                    delta=f"{days_since} days ago" if days_since <= 365 else "Over 1 year",
                    delta_color="inverse" if days_since > 30 else "normal"
                )
            else:
                col3.metric(
                    label="**Last Visit**",
                    value="Never",
                    delta="No data"
                )

            # === KPI CARD 4: Return Rate + Value (COMBINED) ===
            if returns_value > 0:
                col4.metric(
                    label="**Return Rate**",
                    value=f"{return_rate:.2f}%",
                    delta=f"KD {returns_value:,.0f} returned",
                    delta_color="inverse"  # Red = high return
                )
            else:
                col4.metric(
                    label="**Return Rate**",
                    value="0.00%",
                    delta="No returns",
                    delta_color="normal"  # Green = good
                )
            
            # === MINI TABS ===
            mini_tab1, mini_tab2, mini_tab3, mini_tab4, mini_tab5 = st.tabs([
                "Sales Trend", "RFM", "Visits", "Issues", "Actions"
            ])

            # ───────── Mini Tab 1: Sales Trend ─────────
            with mini_tab1:
                daily = cust_sales.groupby(cust_sales[date_col].dt.date)[amount_col].sum().reset_index()
                daily.columns = ["Date", "Sales"]
                fig = px.line(daily, x="Date", y="Sales", title="Sales Trend", markers=True)
                fig.update_layout(height=300)
                st.plotly_chart(fig, use_container_width=True)

            # ───────── Mini Tab 2: RFM ─────────
            with mini_tab2:
                if selected_cust in rfm_agg.index:
                    r = rfm_agg.loc[selected_cust]
                    c1, c2, c3 = st.columns(3)
                    c1.metric("**Recency**", f"{int(r['Recency'])} days")
                    c2.metric("**Frequency**", f"{int(r['Frequency'])}")
                    c3.metric("**Monetary**", f"KD {r['Monetary']:,.0f}")
                    st.success(f"**Segment:** {r['Segment']}")
                else:
                    st.info("RFM not calculated yet")

            # ───────── Mini Tab 3: Visits ─────────
            with mini_tab3:
                if driver_col and driver_col in cust_sales.columns:
                    visits = cust_sales[[date_col, driver_col]].drop_duplicates()
                    visits = visits.sort_values(date_col, ascending=False).head(20)
                    visits["Date"] = visits[date_col].dt.strftime("%Y-%m-%d")
                    visits["Salesman"] = visits[driver_col]
                    st.dataframe(visits[["Date", "Salesman"]], use_container_width=True, hide_index=True)
                else:
                    st.info("No salesman data available")

                        # ───────── Mini Tab 4: Issues (Returns + Material Details) ─────────
                        # ───────── Mini Tab 4: Issues (Returns + Material Details) ─────────
            with mini_tab4:
                issues = cust_sales[returns_mask]

                if not issues.empty:
                    # Main returns summary by invoice/date
                    st.error(f"**Returns: KD {returns_value:,.0f} ({return_rate:.2f}%)**")
                    st.dataframe(
                        issues[[date_col, "Billing Type", amount_col]].rename(
                            columns={date_col: "Billing Date", amount_col: "Return Value"}
                        ),
                        use_container_width=True
                    , hide_index=True)

                    # ---- Return Material Details table ----
                    st.markdown("#### Return Material Details")

                    # Prefer Material Description-like columns first
                    desc_candidates = [
                        "Material Description",
                        "Material Desc",
                        "Material Description EN",
                        "Material Description AR",
                        "MAT Description",
                    ]
                    desc_col = None
                    for c in desc_candidates:
                        if c in issues.columns:
                            desc_col = c
                            break

                    # Fallback: if no description column, use material_col (code)
                    if not desc_col and material_col and material_col in issues.columns:
                        desc_col = material_col

                    if desc_col:
                        issues_mat = issues.copy()
                        issues_mat[amount_col] = (
                            pd.to_numeric(issues_mat[amount_col], errors="coerce")
                            .fillna(0.0)
                            .abs()
                        )

                        # Group by the chosen description column
                        mat_summary = (
                            issues_mat
                            .groupby(desc_col)[amount_col]
                            .sum()
                            .reset_index()
                            .rename(columns={desc_col: "Material Description", amount_col: "Return Value"})
                            .sort_values("Return Value", ascending=False)
                        )

                        # Add Total row at end
                        total_val = mat_summary["Return Value"].sum()
                        total_row = {
                            "Material Description": "Total",
                            "Return Value": total_val
                        }
                        mat_summary = pd.concat(
                            [mat_summary, pd.DataFrame([total_row])],
                            ignore_index=True
                        )

                        st.dataframe(mat_summary, use_container_width=True, hide_index=True)
                    else:
                        st.info("No material description column found for returns.")
                else:
                    st.success("**No Returns – Perfect!**")

            # ───────── Mini Tab 5: Actions ─────────
            with mini_tab5:
                st.markdown("#### **Smart Actions**")
                actions = []
                if days_since > 30:
                    actions.append("**URGENT:** Schedule visit TODAY")
                if return_rate > 10:
                    actions.append("Call about quality issues")
                if total_sales > 5000:
                    actions.append("Offer premium products")
                if order_count > 15:
                    actions.append("Send loyalty reward")

                for a in actions:
                    st.markdown(f"• {a}")

                note_key = f"note_{selected_cust}"
                note = st.text_area("**Add Note**", value=st.session_state.get(note_key, ""), height=80)
                if st.button("**Save Note**", type="primary"):
                    st.session_state[note_key] = note
                    st.success("Note saved!")

                # Download Profile
                profile = pd.DataFrame({
                    "Metric": ["Customer", "Total Sales", "Orders", "Last Visit", "Days Since", "Return Rate %", "Note"],
                    "Value": [
                        selected_cust,
                        total_sales,
                        order_count,
                        last_visit.strftime("%Y-%m-%d") if pd.notna(last_visit) else "N/A",
                        days_since,
                        return_rate,
                        note
                    ]
                })
                st.download_button(
                    "**Download Profile (Excel)**",
                    data=to_excel_bytes(profile),
                    file_name=f"{selected_cust.replace(' ', '_')}_360.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

# --- Material Forecast Page ---
elif choice == texts[lang]["material_forecast"]:
    st.title(texts[lang]["material_forecast_title"])

    if "data_loaded" not in st.session_state:
        st.warning(texts[lang]["no_data_warning"])
        st.stop()

    # Use sales_df directly
    df_sales = st.session_state["sales_df"].copy()

    required_cols = ["Billing Date", "Material", "Material Description"]
    missing_cols = [col for col in required_cols if col not in df_sales.columns]
    if df_sales.empty or missing_cols:
        st.warning(f"⚠️ Sales data is missing required columns: {missing_cols}")
        st.stop()

    # Optional: if you have salesman/user view
    try:
        if "Driver Name EN" in df_sales.columns and "user_role" in st.session_state:
            # Keep existing logic (do not force filter here)
            pass
    except Exception:
        pass

    # Ensure date column is datetime
    df_sales["Billing Date"] = pd.to_datetime(df_sales["Billing Date"], errors="coerce")
    df_sales = df_sales.dropna(subset=["Billing Date"]).copy()

    # Ensure numeric columns
    if "Quantity" in df_sales.columns:
        df_sales["Quantity"] = pd.to_numeric(df_sales["Quantity"], errors="coerce").fillna(0)
    else:
        df_sales["Quantity"] = 0

    if "Net Value" in df_sales.columns:
        df_sales["Net Value"] = pd.to_numeric(df_sales["Net Value"], errors="coerce").fillna(0)

    # Extract Month & Year
    df_sales["Year"] = df_sales["Billing Date"].dt.year.astype(int)
    df_sales["Month"] = df_sales["Billing Date"].dt.month.astype(int)

    # ---------------- Settings ----------------
    with st.expander("⚙️ Forecast Settings", expanded=True):
        metric_choice = st.radio(
            "Forecast Based On",
            options=["Quantity", "Value (Net Value)"],
            horizontal=True,
            index=0
        )

        if metric_choice == "Value (Net Value)" and "Net Value" not in df_sales.columns:
            st.warning("⚠️ 'Net Value' column not found. Switching to Quantity.")
            metric_choice = "Quantity"

        value_col = "Quantity" if metric_choice == "Quantity" else "Net Value"

        # Materials
        all_mats = sorted(df_sales["Material Description"].dropna().astype(str).unique().tolist())
        if not all_mats:
            st.info("No materials found in the data.")
            st.stop()

        # Performance helper: Top-N default when too many
        use_topn = st.checkbox("Use Top-N Materials (recommended for large lists)", value=(len(all_mats) > 60))
        topn = st.slider("Top N Materials", 5, min(200, max(5, len(all_mats))), min(30, len(all_mats))) if use_topn else None

        if use_topn and topn:
            mat_rank = (
                df_sales.groupby("Material Description")[value_col].sum()
                .sort_values(ascending=False)
                .head(topn)
                .index.astype(str)
                .tolist()
            )
            default_mats = mat_rank
        else:
            # User requested: full materials when not selected
            default_mats = all_mats

        selected_mats = st.multiselect(
            "Select Materials (leave as default for all)",
            options=all_mats,
            default=default_mats
        )

        # If user clears selection, fall back to ALL (so nothing becomes empty)
        if not selected_mats:
            selected_mats = all_mats

        exclude_returns = st.checkbox("Exclude Returns (YKRE / ZRE)", value=False)
        exclude_cancels = st.checkbox("Exclude Cancellations (YKS1 / YKS2 / ZCAN)", value=False)

    # Apply optional exclusions
    df_work = df_sales.copy()
    if exclude_returns and "Billing Type" in df_work.columns:
        df_work = df_work[~df_work["Billing Type"].astype(str).str.upper().isin(["YKRE", "ZRE"])].copy()
    if exclude_cancels and "Billing Type" in df_work.columns:
        df_work = df_work[~df_work["Billing Type"].astype(str).str.upper().isin(["YKS1", "YKS2", "ZCAN"])].copy()

    df_work = df_work[df_work["Material Description"].astype(str).isin([str(x) for x in selected_mats])].copy()

    # Tabs for Monthly & Yearly Forecast
    tab_month, tab_year = st.tabs(["Monthly Forecast", "Yearly Forecast"])

    # ---------------- Monthly Forecast ----------------
    with tab_month:
        st.subheader("Monthly Material Forecast")

        years = sorted(df_work["Year"].dropna().unique().tolist())
        if not years:
            st.info("No valid years found after filters.")
            st.stop()

        selected_year = st.selectbox("Select Year:", years, index=len(years)-1)
        df_monthly = df_work[df_work["Year"] == selected_year].copy()

        monthly = (
            df_monthly.groupby(["Month", "Material Description"])[value_col]
            .sum()
            .reset_index()
        )

        # Fill missing months for each material
        all_months = pd.DataFrame({"Month": list(range(1, 13))})
        all_materials = pd.DataFrame({"Material Description": sorted(df_monthly["Material Description"].dropna().astype(str).unique().tolist())})
        full_index = all_materials.merge(all_months, how="cross")

        monthly = full_index.merge(
            monthly, on=["Month", "Material Description"], how="left"
        ).fillna({value_col: 0})

        # Plot
        fig = px.line(
            monthly,
            x="Month",
            y=value_col,
            color="Material Description",
            markers=True,
            title=f"Monthly Trend ({selected_year}) – {metric_choice}"
        )
        st.plotly_chart(fig, use_container_width=True)

        # Pivot
        pivot_table = monthly.pivot(
            index="Material Description", columns="Month", values=value_col
        ).fillna(0)

        st.dataframe(pivot_table, hide_index=True)

        # Download Excel
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        excel_bytes = to_excel_bytes(monthly, sheet_name="Monthly_Forecast")
        if st.download_button(
            texts[lang].get("download_excel", "⬇️ Download Excel"),
            data=excel_bytes,
            file_name=f"monthly_forecast_{selected_year}_{timestamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ):
            st.session_state["audit_log"].append({
                "user": username,
                "action": "download",
                "details": f"monthly_forecast_{selected_year}_{timestamp}.xlsx",
                "timestamp": datetime.now().isoformat()
            })

    # ---------------- Yearly Forecast ----------------
    with tab_year:
        st.subheader("Yearly Material Forecast")

        years = sorted(df_work["Year"].dropna().unique().tolist())
        if not years:
            st.info("No valid years found after filters.")
            st.stop()

        # Let user pick years to compare
        default_years = years[-3:] if len(years) >= 3 else years
        selected_years = st.multiselect("Select Year(s):", options=years, default=default_years)
        if not selected_years:
            selected_years = years

        df_year = df_work[df_work["Year"].isin(selected_years)].copy()

        yearly = (
            df_year.groupby(["Year", "Material Description"])[value_col]
            .sum()
            .reset_index()
        )

        fig = px.bar(
            yearly,
            x="Year",
            y=value_col,
            color="Material Description",
            barmode="group",
            text=value_col,
            title=f"Yearly Trend – {metric_choice}"
        )
        st.plotly_chart(fig, use_container_width=True)

        pivot_table_year = yearly.pivot(
            index="Material Description", columns="Year", values=value_col
        ).fillna(0)

        st.dataframe(pivot_table_year, hide_index=True)

        # Download Excel
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        excel_bytes_year = to_excel_bytes(yearly, sheet_name="Yearly_Forecast")
        if st.download_button(
            texts[lang].get("download_excel", "⬇️ Download Excel"),
            data=excel_bytes_year,
            file_name=f"yearly_forecast_{timestamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ):
            st.session_state["audit_log"].append({
                "user": username,
                "action": "download",
                "details": f"yearly_forecast_{timestamp}.xlsx",
                "timestamp": datetime.now().isoformat()
            })
                      
# --- Management Command Center   ---
elif choice == "🧭 Management Command Center":
    st.title("🧭 Management Command Center")

    # ================= SAFETY CHECK =================
    if sales_df is None or sales_df.empty:
        st.warning("Please load sales data first")
        st.stop()

    df = sales_df.copy()

    # ================= DATE & WORKING DAYS (EXCLUDE FRIDAY ONLY) =================
    today = pd.to_datetime("today").normalize()
    month_start = today.replace(day=1)
    month_end = month_start + pd.offsets.MonthEnd(1)

    # All days in month
    all_days = pd.date_range(month_start, month_end, freq="D")

    # Exclude Friday only (weekday=4)
    working_days = all_days[all_days.weekday != 4]

    total_working_days = len(working_days)

    days_completed = len(working_days[working_days <= today])
    days_completed = max(1, days_completed)  # safety

    # ================= TARGET DATA =================
    if "target_df" in globals() and "KA Target" in target_df.columns:
        ka_target_map = target_df.set_index("Driver Name EN")["KA Target"]
    else:
        ka_target_map = pd.Series(dtype=float)

    # ================= OVERALL KA SALES =================
    total_sales = float(df["Net Value"].sum())
    total_ka_target = float(ka_target_map.sum()) if not ka_target_map.empty else 0.0

    # ================= DAILY-PACE CALCULATION (OVERALL KA) =================
    ka_target_per_day = round(
        total_ka_target / total_working_days, 0
    ) if total_working_days > 0 else 0

    ka_actual_per_day = round(
        total_sales / days_completed, 0
    )

    def pace_status(actual_day, target_day):
        if target_day <= 0:
            return "🟢 GREEN"
        ratio = actual_day / target_day
        if ratio >= 1.0:
            return "🟢 GREEN"
        elif ratio >= 0.95:
            return "🟠 AMBER"
        else:
            return "🔴 RED"

    overall_ka_status = pace_status(
        ka_actual_per_day, ka_target_per_day
    )

    # ================= 1️⃣ EXECUTIVE RAG DASHBOARD =================
    st.subheader("1️⃣ Executive RAG Dashboard (Daily Pace Based)")

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total KA Sales", f"KD {total_sales:,.0f}")
    c2.metric("Total KA Target", f"KD {total_ka_target:,.0f}")
    c3.metric("KA Target / Day", f"KD {ka_target_per_day:,.0f}")
    c4.metric("KA Actual / Day", f"KD {ka_actual_per_day:,.0f}")
    c5.metric("Overall KA Status", overall_ka_status)

    # ================= 2️⃣ SALESMAN DAILY-PACE RISK TABLE =================
    st.subheader("2️⃣ Salesman Early-Warning (Daily Pace – 95% Rule)")

    salesman_df = (
        df.groupby("Driver Name EN")["Net Value"]
        .sum()
        .reset_index(name="Achieved")
    )

    salesman_df["Target"] = salesman_df["Driver Name EN"].map(
        ka_target_map
    ).fillna(0)

    salesman_df["Target / Day"] = (
        salesman_df["Target"] / total_working_days
    ).round(0)

    salesman_df["Actual / Day"] = (
        salesman_df["Achieved"] / days_completed
    ).round(0)

    def salesman_risk(row):
        return pace_status(
            row["Actual / Day"],
            row["Target / Day"]
        )

    salesman_df["Risk"] = salesman_df.apply(
        salesman_risk, axis=1
    )

    st.dataframe(
        salesman_df[
            [
                "Driver Name EN",
                "Target",
                "Achieved",
                "Target / Day",
                "Actual / Day",
                "Risk"
            ]
        ].sort_values("Risk"),
        use_container_width=True
    )

    # ================= 5️⃣ ACTION-BASED MANAGEMENT INSIGHTS =================
    st.subheader("5️⃣ Action-based Management Insights")

    insights = []

    red_salesmen = salesman_df[salesman_df["Risk"] == "🔴 RED"]
    amber_salesmen = salesman_df[salesman_df["Risk"] == "🟠 AMBER"]

    if not red_salesmen.empty:
        insights.append(
            f"❗ {len(red_salesmen)} salesmen are BELOW daily pace (<95%) – immediate action required"
        )

    if not amber_salesmen.empty:
        insights.append(
            f"⚠️ {len(amber_salesmen)} salesmen are SLIGHTLY BELOW pace (95–99%) – close monitoring needed"
        )

    if overall_ka_status == "🔴 RED":
        insights.append(
            "🚨 Overall KA pace is BELOW 95% – revise visit plan, focus on KA & fast movers"
        )
    elif overall_ka_status == "🟠 AMBER":
        insights.append(
            "🟠 Overall KA pace is JUST BELOW target – strong push required this week"
        )
    else:
        insights.append(
            "🟢 Overall KA pace is ON TRACK – maintain execution discipline"
        )

    # Optional Talabat warning (safe)
    if "PY Name 1" in df.columns:
        talabat_returns = df[
            (df["PY Name 1"].str.contains("TALABAT", case=False, na=False)) &
            (df["Net Value"] < 0)
        ]["Net Value"].sum()

        if talabat_returns < 0:
            insights.append(
                "🚨 Talabat negative sales / returns detected – investigate service & billing"
            )

    for msg in insights:
        st.write(msg)

# Admin-only Audit Logs View
if user_role == "admin":
    st.sidebar.markdown("---")
    st.sidebar.subheader("Admin Tools")
    if st.sidebar.button("View Audit Logs"):
        st.title("📋 Audit Logs")
        log_df = pd.DataFrame(st.session_state["audit_log"])
        st.dataframe(log_df, hide_index=True)
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        if st.download_button(
            "⬇️ Download Audit Logs (Excel)",
            data=to_excel_bytes(log_df, sheet_name="Audit_Logs", index=False),
            file_name=f"audit_logs_{timestamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ):
            st.session_state["audit_log"].append({
                "user": username,
                "action": "download",
                "details": f"audit_logs_{timestamp}.xlsx",
                "timestamp": datetime.now()
            })

if "audit_log" not in st.session_state:
    st.session_state["audit_log"] = []
    
    
