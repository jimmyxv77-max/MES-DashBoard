import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import io
import re
from openpyxl.chart import BarChart, LineChart, Reference

# ================= 🎛️ 不良目标线锁死接口 (预留配置区) =================
LOCKED_DEFECT_TARGETS = {
    # 示例格式："Calendaring Incoming Shortage": 0.5,
}

# 1. 页面基本配置
st.set_page_config(
    page_title="Electrode Dashboard",
    page_icon="📊",
    layout="wide"
)

# 统一图表配置：默认拖拽为平移(pan)，关闭鼠标滚轮缩放(scrollZoom=False)防止误触
DEFAULT_CHART_CONFIG = {
    "dragmode": "pan",
    "scrollZoom": False,
    "displayModeBar": True
}

# ================= 语言词典配置 =================
LANG_DICT = {
    "zh": {
        "title_text": "Electrode Dashboard",
        "author": "created by xu",
        "mode_select": "🎛️ 选择系统运行模式：",
        "mode_v1": "1. 常规单文件生产看板 (v1.0 正式版)",
        "mode_v2": "2. 周期性双文件对比分析 (v2.0 升级版)",
        "import_header": "📁 数据源导入",
        "import_help": "请将 MES 生产数据表拖拽至此处，或点击浏览（支持 MesResume_DoMethod 前缀）",
        "import_info": "👈 请在左侧边栏拖拽或选择你的 MES 生产数据表以生成实时看板。",
        "export_header": "📥 报告导出",
        "btn_export": "📥 导出全图表汇总报表 (Excel)",
        "changelog_title": "📌 版本更新说明 (What's New)",
        "changelog_text": """
        - **v2.0 周期双文件对比升级**：完美引入涂布自动修正、分切 too much 放行规则，并支持上周红色空心红框浮于本周实色柱上方，底部同步展示清洗与放行明细！
        - **v1.0 稳定隔离**：正式 1.0 核心功能 100% 完整保留。
        """,
        "sidebar_filter": "🔍 数据筛选控制台",
        "search_label": "🔎 卷号 (BatchNo) / 批次号 (MasterLotNo) 检索：",
        "filter_process": "筛选工序 (ProcessName)",
        "filter_product": "筛选产品型号 (ProductCName)",
        "filter_machine": "筛选机台名称 (MachineName)",
        "filter_person": "筛选操作员 (PersonName)",
        "filter_class": "筛选班次 (ClassNo)",
        "chk_abnormal": "仅显示低良率异常卷 (FPY < 95%)",
        "target_header": "🎯 质量目标线设置（良率标准）",
        "t_machine": "机台单卷良率目标线 (%)",
        "t_date": "单日平均良率目标线 (%)",
        "kpi_batches": "生产卷数 (Batch)",
        "kpi_in": "总投入 (INQty)",
        "kpi_out": "总产出 (OutQty)",
        "kpi_scrap": "总废弃 (ScrapQty)",
        "kpi_fpy": "平均单卷良率 (FPY)",
        "pareto_header": "📉 工序异常明细分析",
        "pareto_title": "📊 异常帕累托图",
        "radio_share": "选择柱状图展示维度：",
        "share_1": "占总不良比例 (%)",
        "share_2": "占总投入比例 (%) [损失率]",
        "table_header": "📋 异常数量与双维度比例对照表",
        "col_defect": "不良项目",
        "col_count": "不良数量",
        "col_s_defect": "占总不良比例",
        "col_s_in": "占总投入比例",
        "col_cum": "累计不良占比",
        "defect_trend_header": "🔍 单项不良变化推移图",
        "defect_multi_header": "📊 单项不良责任归属快速对比",
        "def_sel": "🎯 1. 选择追踪的不良项目：",
        "def_view": "📊 2. 选择观察粒度：",
        "view_daily": "按单日汇总趋势",
        "view_roll": "按单卷明细推移（分机台）",
        "def_metric": "📐 3. 选择统计指标：",
        "m_count": "不良数量推移",
        "m_rate": "不良率推移 (% 总投入)",
        "y_max_label": "🎚️ 单项不良 Y 轴上限：",
        "def_target_label": "🎯 单项不良目标线阈值设置：",
        "rank_sel_def": "🎯 选择分析的不良项目：",
        "rank_sel_dim": "📊 选择归属维度：",
        "dim_machine": "MachineName (机台)",
        "dim_person": "PersonName (操作者)",
        "rank_chart_title_pre": "【",
        "rank_chart_title_mid": "】按 ",
        "rank_chart_title_suf": " 累计不良及占比排行",
        "rank_xaxis": "不良数量合计 (Qty)",
        "machine_filter_label": "🎛️ 选择要查看的机台（支持多选/全选）：",
        "defect_machine_filter_label": "🎛️ 选择要追踪的不良机台（支持多选/全选）：",
        "yield_trend_header": "📈 良率推移趋势分析",
        "y_min_label": "🎚️ 良率图 Y 轴下限 (%)：",
        "t_m_trend": "📈 1. 各机台单卷良率 (FPY) 推移走势",
        "t_d_trend": "📅 2. 单日投入产出与平均良率组合推移走势",
        "table_detail": "📋 生产批次明细数据表（当前筛选结果）",
        "x_time": "生产时间",
        "x_date": "生产日期",
        "y_roll_fpy": "单卷良率 FPY (%)",
        "y_daily_fpy": "单日平均良率 (%)",
        "y_qty_axis": "投入 / 产出数量 (Qty)",
        "legend_in": "总投入 (INQty)",
        "legend_out": "总产出 (OutQty)",
        "legend_fpy": "单日平均良率 (%)",
    },
    "en": {
        "title_text": "Electrode Dashboard",
        "author": "created by xu",
        "mode_select": "🎛️ Select System Mode:",
        "mode_v1": "1. Standard Single-File Dashboard (v1.0)",
        "mode_v2": "2. Periodic Dual-File Comparison (v2.0)",
        "import_header": "📁 Data Source Import",
        "import_help": "Drag and drop your MES production data table here",
        "import_info": "👈 Please drag or select your MES production data table in the sidebar.",
        "export_header": "📥 Report Export",
        "btn_export": "📥 Export All-Chart Summary Report (Excel)",
        "changelog_title": "📌 What's New",
        "changelog_text": """
        - **v2.0 Dual-File Comparison Upgraded**: Integrated coating correction & too-much rules with red outline on top of solid bars.
        """,
        "sidebar_filter": "🔍 Data Filter Console",
        "search_label": "🔎 Search BatchNo / MasterLotNo:",
        "filter_process": "Filter Process (ProcessName)",
        "filter_product": "Filter Product (ProductCName)",
        "filter_machine": "Filter Machine (MachineName)",
        "filter_person": "Filter Operator (PersonName)",
        "filter_class": "Filter Class (ClassNo)",
        "chk_abnormal": "Show Low Yield Only (FPY < 95%)",
        "target_header": "🎯 Quality Target Setting",
        "t_machine": "Machine Roll Yield Target (%)",
        "t_date": "Daily Avg Yield Target (%)",
        "kpi_batches": "Production Batches",
        "kpi_in": "Total Input (INQty)",
        "kpi_out": "Total Output (OutQty)",
        "kpi_scrap": "Total Scrap (ScrapQty)",
        "kpi_fpy": "Avg Roll Yield (FPY)",
        "pareto_header": "📉 Process Defect Analysis",
        "pareto_title": "📊 Defect Pareto Chart",
        "radio_share": "Select Bar Chart Metric:",
        "share_1": "Defect Share (%)",
        "share_2": "Loss Rate of Total Input (%)",
        "table_header": "📋 Defect Quantity & Dual-Dimension Share Table",
        "col_defect": "Defect Item",
        "col_count": "Defect Qty",
        "col_s_defect": "Defect Share",
        "col_s_in": "Input Loss Rate",
        "col_cum": "Cum. Share",
        "defect_trend_header": "🔍 Single Defect Trend Analysis",
        "defect_multi_header": "📊 Single Defect Responsibility Ranking",
        "def_sel": "🎯 1. Select Defect Item:",
        "def_view": "📊 2. Select View Granularity:",
        "view_daily": "Daily Summary Trend",
        "view_roll": "Roll-level Detail Trend (by Machine)",
        "def_metric": "📐 3. Select Metric:",
        "m_count": "Defect Quantity Trend",
        "m_rate": "Defect Rate Trend (% of Input)",
        "y_max_label": "🎚️ Single Defect Y-Axis Upper Limit:",
        "def_target_label": "🎯 Single Defect Target Threshold:",
        "rank_sel_def": "🎯 Select Defect Item:",
        "rank_sel_dim": "📊 Select Dimension:",
        "dim_machine": "MachineName",
        "dim_person": "PersonName",
        "rank_chart_title_pre": "[",
        "rank_chart_title_mid": "] Cumulative Defect & Share by ",
        "rank_chart_title_suf": "",
        "rank_xaxis": "Total Defect Qty",
        "machine_filter_label": "🎛️ Select Machine(s) to View:",
        "defect_machine_filter_label": "🎛️ Select Defect Machine(s) to Track:",
        "yield_trend_header": "📈 Yield Trend Analysis",
        "y_min_label": "🎚️ Yield Chart Y-Axis Lower Limit (%):",
        "t_m_trend": "📈 1. Machine Roll Yield (FPY) Trend",
        "t_d_trend": "📅 2. Daily Input/Output & Avg Yield Combo Trend",
        "table_detail": "📋 Production Batch Detail Table",
        "x_time": "Production Time",
        "x_date": "Production Date",
        "y_roll_fpy": "Roll Yield FPY (%)",
        "y_daily_fpy": "Daily Avg Yield (%)",
        "y_qty_axis": "Input / Output Qty",
        "legend_in": "Total Input (INQty)",
        "legend_out": "Total Output (OutQty)",
        "legend_fpy": "Daily Avg Yield (%)",
    }
}

top_col1, top_col2 = st.columns([3, 1])
with top_col2:
    selected_lang_name = st.selectbox("🌐 Language / 语言", ["中文 (Chinese)", "English"], index=0)
    curr_lang = "zh" if "中文" in selected_lang_name else "en"

t = LANG_DICT[curr_lang]

logo_url = "https://www.tenpowercell.com/wp-content/themes/tenpower/images/logo-big-white.png"

st.markdown(
    f"""
    <div style="display: flex; align-items: center; gap: 15px; margin-bottom: 5px;">
        <img src="{logo_url}" alt="Tenpower Logo" style="height: 38px; width: auto; object-fit: contain;">
        <h1 style="margin: 0; padding: 0; font-size: 2.25rem; font-weight: 700;">{t["title_text"]}</h1>
    </div>
    """,
    unsafe_allow_html=True
)
st.markdown(f"*{t['author']}*")
st.markdown("---")

# ================= 🎛️ 首页模式导航切换按钮 =================
st.sidebar.markdown(f"### {t['mode_select']}")
system_mode = st.sidebar.radio(
    "Mode",
    [t["mode_v1"], t["mode_v2"]],
    label_visibility="collapsed"
)

# 通用数据清洗与校验函数（封装涂布修正与分切 too much 放行规则）
def process_and_clean_dataframe(df_input):
    d = df_input.copy()
    d.columns = d.columns.str.strip()

    if "TestDate" in d.columns:
        d["TestDate"] = pd.to_datetime(d["TestDate"], errors="coerce")

    for col in ["INQty", "OutQty", "ScrapQty"]:
        if col in d.columns:
            d[col] = pd.to_numeric(d[col], errors="coerce").fillna(0)

    has_fpy = "FPY" in d.columns
    if has_fpy:
        d["FPY"] = pd.to_numeric(d["FPY"], errors="coerce")
        if d["FPY"].dropna().max() <= 1.5:
            d["FPY"] = d["FPY"] * 100.0

    fixed_count = 0
    preserved_too_much_count = 0
    coating_fixed_df = pd.DataFrame()
    preserved_too_much_df = pd.DataFrame()
    extreme_error_df = pd.DataFrame()

    if "INQty" in d.columns and "OutQty" in d.columns and "ScrapQty" in d.columns:
        raw_abnormal_mask = d["INQty"] < d["OutQty"]
        is_coating = d["ProcessName"].astype(str).str.contains("coating|涂布", case=False, na=False) if "ProcessName" in d.columns else pd.Series(False, index=d.index)
        
        has_too_much = pd.Series(False, index=d.index)
        for col_c in d.columns:
            if "too much" in col_c.lower() or "toomuch" in col_c.lower():
                num_vals = pd.to_numeric(d[col_c], errors="coerce").fillna(0)
                has_too_much = has_too_much | (num_vals > 0)
            elif d[col_c].dtype == object:
                has_too_much = has_too_much | d[col_c].astype(str).str.contains("too\\s*much", case=False, na=False)

        coating_fix_mask = raw_abnormal_mask & is_coating
        fixed_count = int(coating_fix_mask.sum())
        if fixed_count > 0:
            coating_fixed_df = d[coating_fix_mask].copy()
            d.loc[coating_fix_mask, "OutQty"] = d.loc[coating_fix_mask, "INQty"] - d.loc[coating_fix_mask, "ScrapQty"]
            d.loc[d["OutQty"] < 0, "OutQty"] = 0
            if has_fpy:
                valid_in = coating_fix_mask & (d["INQty"] > 0)
                d.loc[valid_in, "FPY"] = (d.loc[valid_in, "OutQty"] / d.loc[valid_in, "INQty"]) * 100.0
                d.loc[coating_fix_mask & (d["INQty"] <= 0), "FPY"] = 0.0

        preserved_too_much_mask = raw_abnormal_mask & (~is_coating) & has_too_much
        preserved_too_much_count = int(preserved_too_much_mask.sum())
        if preserved_too_much_count > 0:
            preserved_too_much_df = d[preserved_too_much_mask].copy()

        extreme_error_mask = raw_abnormal_mask & (~is_coating) & (~has_too_much)
        if extreme_error_mask.sum() > 0:
            extreme_error_df = d[extreme_error_mask].copy()

    return d, coating_fixed_df, preserved_too_much_df, extreme_error_df, has_fpy

def load_uploaded_file(uploaded_file):
    if uploaded_file is None:
        return None
    fname = uploaded_file.name
    try:
        if fname.endswith(".csv"):
            try:
                d = pd.read_csv(uploaded_file, encoding="utf-8")
            except UnicodeDecodeError:
                uploaded_file.seek(0)
                d = pd.read_csv(uploaded_file, encoding="gbk")
        else:
            d = pd.read_excel(uploaded_file)
        return d
    except Exception as ex:
        st.error(f"❌ File read error ({fname}): {ex}")
        return None

def extract_week_label(filename):
    match = re.search(r'[-_]W(\d+)', filename, re.IGNORECASE)
    if match:
        return f"Week {match.group(1)} ({filename})"
    return filename


# ==============================================================================
# 🟢 模式一：常规生产全功能看板 (v1.0 正式版)
# ==============================================================================
if system_mode == t["mode_v1"]:
    st.sidebar.header(t["import_header"])
    uploaded_file = st.sidebar.file_uploader(
        t["import_help"],
        type=["xlsx", "xls", "csv"],
        key="v1_uploader"
    )

    if uploaded_file is None:
        st.info(t["import_info"])
        st.markdown("---")
        st.markdown(f"### {t['changelog_title']}")
        st.markdown(t["changelog_text"])
        st.stop()

    raw_df = load_uploaded_file(uploaded_file)
    if raw_df is None:
        st.stop()

    file_name = uploaded_file.name
    if not file_name.lower().startswith("mesresume_domethod"):
        st.warning(f"⚠️ Warning: File **{file_name}** does not start with `MesResume_DoMethod`. Trying to parse anyway.")

    df, coating_fixed_df, preserved_too_much_df, extreme_error_df, has_fpy = process_and_clean_dataframe(raw_df)

    required_cols = ["INQty", "OutQty", "ScrapQty"]
    missing_cols = [c for c in required_cols if c not in df.columns]
    if missing_cols:
        st.error(f"❌ Missing required columns: {', '.join(missing_cols)}")
        st.stop()

    if not coating_fixed_df.empty or not preserved_too_much_df.empty or not extreme_error_df.empty:
        msg_parts = []
        if not coating_fixed_df.empty:
            msg_parts.append(f"涂布工序已自动修正 **{len(coating_fixed_df)}** 条异常记录")
        if not preserved_too_much_df.empty:
            msg_parts.append(f"分切来料多（`too much`）已保留原数据 **{len(preserved_too_much_df)}** 条继续执行")
        if not extreme_error_df.empty:
            msg_parts.append(f"检测到 **{len(extreme_error_df)}** 条疑似录入错误的极端异常数据")
        st.warning("⚠️ **MES 数据智能校验与分流报告**：" + "；".join(msg_parts) + "。")
    else:
        st.success("✅ **数据逻辑校验通过**：当前数据集中未发现 `投入 < 产出` 的异常偏差记录。")

    if not extreme_error_df.empty:
        st.error(f"🚨 **发现 {len(extreme_error_df)} 条极端不合理数据（无 too much 支撑的产出远大于投入）！请核对以下明细：**")
        highlight_cols = [c for c in ["BatchNo", "MasterLotNo", "ProcessName", "MachineName", "PersonName", "INQty", "OutQty", "ScrapQty", "TestDate"] if c in extreme_error_df.columns]
        st.dataframe(extreme_error_df[highlight_cols], use_container_width=True, hide_index=True)
        st.markdown("---")

    exclude_keywords = ["lotno", "batch", "person", "machine", "date", "class", "product", "process", "qty", "fpy", "scrap"]
    all_columns = df.columns.tolist()
    DYNAMIC_DEFECT_COLS = []
    for col in all_columns:
        if not any(kw in col.lower() for kw in exclude_keywords):
            converted = pd.to_numeric(df[col], errors="coerce")
            if converted.notna().sum() > 0:
                df[col] = converted.fillna(0)
                DYNAMIC_DEFECT_COLS.append(col)

    st.sidebar.header(t["sidebar_filter"])
    search_keyword = st.sidebar.text_input(t["search_label"], "", key="v1_search").strip()
    display_df = df.copy()

    filter_configs = [
        ("ProcessName", t["filter_process"]),
        ("ProductCName", t["filter_product"]),
        ("MachineName", t["filter_machine"]),
        ("PersonName", t["filter_person"]),
        ("ClassNo", t["filter_class"])
    ]

    for col_name, label in filter_configs:
        if col_name in display_df.columns:
            options = sorted(display_df[col_name].dropna().astype(str).unique().tolist())
            selected = st.sidebar.multiselect(label, options, default=[], key=f"v1_filter_{col_name}")
            if selected:
                display_df = display_df[display_df[col_name].astype(str).isin(selected)]

    if search_keyword:
        search_cols = [c for c in ["BatchNo", "MasterLotNo"] if c in display_df.columns]
        if search_cols:
            mask = pd.Series(False, index=display_df.index)
            for c in search_cols:
                mask = mask | display_df[c].astype(str).str.contains(search_keyword, case=False, na=False)
            display_df = display_df[mask]

    show_only_abnormal = st.sidebar.checkbox(t["chk_abnormal"], value=False, key="v1_abnormal_chk")
    if show_only_abnormal and has_fpy:
        display_df = display_df[display_df["FPY"] < 95.0]

    st.sidebar.markdown("---")
    st.sidebar.markdown(t["changelog_text"])

    total_in_qty = display_df["INQty"].sum() if "INQty" in display_df.columns else 0
    total_out_qty = display_df["OutQty"].sum() if "OutQty" in display_df.columns else 0
    total_scrap_qty = display_df["ScrapQty"].sum() if "ScrapQty" in display_df.columns else 0

    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        st.metric(t["kpi_batches"], f"{len(display_df):,}")
    with col2:
        st.metric(t["kpi_in"], f"{total_in_qty:,.1f}")
    with col3:
        st.metric(t["kpi_out"], f"{total_out_qty:,.1f}")
    with col4:
        st.metric(t["kpi_scrap"], f"{total_scrap_qty:,.1f}")
    with col5:
        if has_fpy and len(display_df) > 0:
            valid_fpy_series = display_df["FPY"].dropna()
            avg_fpy = valid_fpy_series.mean() if not valid_fpy_series.empty else 0
            st.metric(t["kpi_fpy"], f"{avg_fpy:.2f}%")
        else:
            overall_yield = (total_out_qty / total_in_qty * 100) if total_in_qty > 0 else 0
            st.metric(t["kpi_fpy"], f"{overall_yield:.2f}%")

    st.markdown("---")

    current_active_defects = [c for c in DYNAMIC_DEFECT_COLS if c in display_df.columns]
    st.subheader(t["pareto_header"])

    global_defect_sums = pd.DataFrame()
    global_daily_defect_agg = pd.DataFrame()
    global_daily_yield_agg = pd.DataFrame()

    if current_active_defects and len(display_df) > 0:
        defect_sums = display_df[current_active_defects].sum().reset_index()
        defect_sums.columns = [t["col_defect"], t["col_count"]]
        defect_sums = defect_sums[defect_sums[t["col_count"]] > 0].sort_values(t["col_count"], ascending=False).reset_index(drop=True)
        
        if not defect_sums.empty:
            total_defects = defect_sums[t["col_count"]].sum()
            defect_sums[t["col_s_defect"]] = (defect_sums[t["col_count"]] / total_defects) * 100 if total_defects > 0 else 0.0
            defect_sums[t["col_cum"]] = defect_sums[t["col_s_defect"]].cumsum()
            defect_sums[t["col_s_in"]] = (defect_sums[t["col_count"]] / total_in_qty) * 100 if total_in_qty > 0 else 0.0

            global_defect_sums = defect_sums.copy()

            fig_pareto = go.Figure()
            fig_pareto.add_trace(go.Bar(
                x=defect_sums[t["col_defect"]],
                y=defect_sums[t["col_count"]],
                name="Defect Qty",
                marker_color="indianred",
                text=defect_sums[t["col_count"]].apply(lambda v: f"{v:,.1f}"),
                textposition="outside",
                yaxis="y"
            ))
            fig_pareto.add_trace(go.Scatter(
                x=defect_sums[t["col_defect"]],
                y=defect_sums[t["col_cum"]],
                name="Cum. Share %",
                mode="lines+markers+text",
                marker=dict(color="darkblue", size=8),
                line=dict(color="darkblue", width=2.5),
                text=defect_sums[t["col_cum"]].apply(lambda v: f"{v:.1f}%"),
                textposition="top center",
                yaxis="y2"
            ))
            max_bar = defect_sums[t["col_count"]].max() if defect_sums[t["col_count"]].max() > 0 else 10
            fig_pareto.update_layout(
                xaxis=dict(tickangle=-20),
                yaxis=dict(title="Defect Qty", range=[0, max_bar * 1.30]),
                yaxis2=dict(title="Cum. Share (%)", overlaying="y", side="right", range=[0, 125], showgrid=False),
                height=500,
                margin=dict(t=40, b=80, l=50, r=50),
                dragmode="pan"
            )
            st.plotly_chart(fig_pareto, use_container_width=True, config=DEFAULT_CHART_CONFIG)

            ratio_mode = st.radio(t["radio_share"], [t["share_1"], t["share_2"]], horizontal=True, key="pareto_ratio_mode")
            target_col = t["col_s_defect"] if t["share_1"] in ratio_mode else t["col_s_in"]
            color_scale = "Oranges" if t["share_1"] in ratio_mode else "Purples"
            
            share_sorted = defect_sums.sort_values(target_col, ascending=True)
            fig_share = px.bar(
                share_sorted, x=target_col, y=t["col_defect"], orientation="h",
                text=share_sorted[target_col].apply(lambda v: f"{v:.2f}%"),
                color=target_col, color_continuous_scale=color_scale
            )
            fig_share.update_traces(textposition="outside")
            fig_share.update_layout(
                xaxis=dict(range=[0, share_sorted[target_col].max() * 1.30]),
                coloraxis_showscale=False,
                height=400,
                dragmode="pan"
            )
            st.plotly_chart(fig_share, use_container_width=True, config=DEFAULT_CHART_CONFIG)

            stat_table = defect_sums[[t["col_defect"], t["col_count"], t["col_s_defect"], t["col_s_in"], t["col_cum"]]].copy()
            st.dataframe(stat_table.style.format({t["col_count"]: "{:,.1f}", t["col_s_defect"]: "{:.2f}%", t["col_s_in"]: "{:.2f}%", t["col_cum"]: "{:.2f}%"}), use_container_width=True, hide_index=True)

    st.markdown("---")

    # 单项不良变化推移图
    st.subheader(t["defect_trend_header"])
    selected_defect = None

    if current_active_defects and "TestDate" in display_df.columns and len(display_df) > 0:
        ctrl_col1, ctrl_col2, ctrl_col3 = st.columns([1.1, 1.3, 1.1])
        with ctrl_col1:
            selected_defect = st.selectbox(t["def_sel"], current_active_defects, index=0, key="main_defect_selector")
        with ctrl_col2:
            defect_view_mode = st.radio(t["def_view"], [t["view_daily"], t["view_roll"]], horizontal=True, key="main_view_mode_radio")
        with ctrl_col3:
            defect_metric_mode = st.radio(t["def_metric"], [t["m_count"], t["m_rate"]], horizontal=True, key="main_metric_mode_radio")

        defect_trend_df = display_df.dropna(subset=["TestDate"]).sort_values("TestDate").copy()
        is_rate_mode = t["m_rate"] in defect_metric_mode
        is_roll_mode = t["view_roll"] in defect_view_mode

        if is_roll_mode:
            all_defect_machines = sorted(defect_trend_df["MachineName"].dropna().astype(str).unique().tolist()) if "MachineName" in defect_trend_df.columns else []
            if all_defect_machines:
                selected_defect_machines = st.multiselect(
                    t["defect_machine_filter_label"],
                    options=all_defect_machines,
                    default=all_defect_machines,
                    key="defect_machine_multiselect"
                )
                if selected_defect_machines:
                    defect_trend_df = defect_trend_df[defect_trend_df["MachineName"].astype(str).isin(selected_defect_machines)]

        if not defect_trend_df.empty:
            dynamic_input_key_suffix = f"{selected_defect}_{is_roll_mode}_{is_rate_mode}"

            if not is_roll_mode:
                defect_trend_df["DayStr"] = defect_trend_df["TestDate"].dt.strftime("%Y-%m-%d")
                daily_defect_agg = defect_trend_df.groupby("DayStr", as_index=False).agg({
                    selected_defect: "sum",
                    "INQty": "sum"
                }).sort_values("DayStr")
                
                daily_defect_agg["DefectRatePercent"] = np.where(
                    daily_defect_agg["INQty"] > 0, 
                    (daily_defect_agg[selected_defect] / daily_defect_agg["INQty"]) * 100.0, 
                    0.0
                )
                global_daily_defect_agg = daily_defect_agg.copy()

                y_col = "DefectRatePercent" if is_rate_mode else selected_defect
                y_axis_title = f"{selected_defect} (%)" if is_rate_mode else f"{selected_defect} (Qty)"
                
                raw_series = daily_defect_agg[y_col].dropna()
                auto_max = float(raw_series.max()) if not raw_series.empty else 1.0
                
                if is_rate_mode:
                    init_mean = raw_series.mean()
                    init_std = raw_series.std()
                    upper_bound = init_mean + 3.0 * (init_std if not np.isnan(init_std) else 0.0)
                    robust_series = raw_series[raw_series <= upper_bound]
                    if robust_series.empty:
                        robust_series = raw_series
                    
                    r_mean = robust_series.mean()
                    r_std = robust_series.std()
                    if r_mean < 1.0:
                        sigma_default = max(0.01, r_mean + 0.5 * (r_std if not np.isnan(r_std) else 0.1))
                    else:
                        sigma_default = max(0.0, r_mean - 1.0 * (r_std if not np.isnan(r_std) else 0.0))

                    auto_max = max(sigma_default * 1.5, auto_max * 1.25)
                    locked_val = LOCKED_DEFECT_TARGETS.get(selected_defect, None)
                    
                    adj_col1, adj_col2 = st.columns([1, 1])
                    with adj_col1:
                        if locked_val is not None:
                            st.info(f"🔒 Target Locked: {locked_val}%")
                            target_defect_val = locked_val
                        else:
                            target_defect_val = st.number_input(
                                t["def_target_label"],
                                min_value=0.0,
                                max_value=float(max(100.0, auto_max * 2.0)),
                                value=float(round(sigma_default, 2)),
                                key=f"target_input_main_{dynamic_input_key_suffix}",
                                step=0.01
                            )
                    with adj_col2:
                        custom_defect_ymax = st.number_input(
                            t["y_max_label"],
                            min_value=0.01,
                            max_value=float(max(1000.0, auto_max * 10.0)),
                            value=float(max(auto_max, target_defect_val * 1.2)),
                            key=f"ymax_input_main_{dynamic_input_key_suffix}",
                            step=0.1
                        )
                else:
                    custom_defect_ymax = st.number_input(
                        t["y_max_label"],
                        min_value=1.0,
                        max_value=float(max(100000.0, auto_max * 10.0)),
                        value=float(max(10.0, auto_max * 1.2)),
                        key=f"ymax_input_main_qty_{dynamic_input_key_suffix}",
                        step=10.0
                    )

                fig_def = px.line(daily_defect_agg, x="DayStr", y=y_col, markers=True)
                if is_rate_mode:
                    target_label_def = f"Target: {target_defect_val}%" if curr_lang=="en" else f"目标线: {target_defect_val}%"
                    fig_def.add_hline(
                        y=target_defect_val,
                        line_dash="dash",
                        line_color="red",
                        annotation_text=target_label_def,
                        annotation_position="top left"
                    )
                fig_def.update_layout(
                    xaxis=dict(type="category", tickangle=-35),
                    yaxis=dict(title=y_axis_title, range=[0, custom_defect_ymax]),
                    height=430,
                    dragmode="pan"
                )
                st.plotly_chart(fig_def, use_container_width=True, config=DEFAULT_CHART_CONFIG)

            else:
                roll_df = defect_trend_df.copy()
                roll_df["DefectRatePercent"] = np.where(
                    roll_df["INQty"] > 0, 
                    (roll_df[selected_defect] / roll_df["INQty"]) * 100.0, 
                    0.0
                )
                y_col = "DefectRatePercent" if is_rate_mode else selected_defect
                y_axis_title = f"{selected_defect} (%)" if is_rate_mode else f"{selected_defect} (Qty)"

                raw_series = roll_df[y_col].dropna()
                auto_max = float(raw_series.max()) if not raw_series.empty else 1.0

                if is_rate_mode:
                    init_mean = raw_series.mean()
                    init_std = raw_series.std()
                    upper_bound = init_mean + 3.0 * (init_std if not np.isnan(init_std) else 0.0)
                    robust_series = raw_series[raw_series <= upper_bound]
                    if robust_series.empty:
                        robust_series = raw_series
                    
                    r_mean = robust_series.mean()
                    r_std = robust_series.std()
                    if r_mean < 1.0:
                        sigma_default = max(0.01, r_mean + 0.5 * (r_std if not np.isnan(r_std) else 0.1))
                    else:
                        sigma_default = max(0.0, r_mean - 1.0 * (r_std if not np.isnan(r_std) else 0.0))

                    auto_max = max(sigma_default * 1.5, auto_max * 1.25)
                    locked_val = LOCKED_DEFECT_TARGETS.get(selected_defect, None)
                    
                    adj_col1, adj_col2 = st.columns([1, 1])
                    with adj_col1:
                        if locked_val is not None:
                            st.info(f"🔒 Target Locked: {locked_val}%")
                            target_defect_val = locked_val
                        else:
                            target_defect_val = st.number_input(
                                t["def_target_label"],
                                min_value=0.0,
                                max_value=float(max(100.0, auto_max * 2.0)),
                                value=float(round(sigma_default, 2)),
                                key=f"target_input_roll_{dynamic_input_key_suffix}",
                                step=0.01
                            )
                    with adj_col2:
                        custom_defect_ymax = st.number_input(
                            t["y_max_label"],
                            min_value=0.01,
                            max_value=float(max(1000.0, auto_max * 10.0)),
                            value=float(max(auto_max, target_defect_val * 1.2)),
                            key=f"ymax_input_roll_{dynamic_input_key_suffix}",
                            step=0.1
                        )
                else:
                    custom_defect_ymax = st.number_input(
                        t["y_max_label"],
                        min_value=1.0,
                        max_value=float(max(100000.0, auto_max * 10.0)),
                        value=float(max(10.0, auto_max * 1.2)),
                        key=f"ymax_input_roll_qty_{dynamic_input_key_suffix}",
                        step=10.0
                    )

                fig_def = px.line(
                    roll_df,
                    x="TestDate",
                    y=y_col,
                    color="MachineName" if "MachineName" in roll_df.columns else None,
                    markers=True
                )
                if is_rate_mode:
                    target_label_def = f"Target: {target_defect_val}%" if curr_lang=="en" else f"目标线: {target_defect_val}%"
                    fig_def.add_hline(
                        y=target_defect_val,
                        line_dash="dash",
                        line_color="red",
                        annotation_text=target_label_def,
                        annotation_position="top left"
                    )
                fig_def.update_layout(
                    xaxis=dict(title=t["x_time"]),
                    yaxis=dict(title=y_axis_title, range=[0, custom_defect_ymax]),
                    height=430,
                    dragmode="pan"
                )
                st.plotly_chart(fig_def, use_container_width=True, config=DEFAULT_CHART_CONFIG)

    st.markdown("---")

    # 单项不良责任归属快速对比
    st.subheader(t["defect_multi_header"])
    if current_active_defects and len(display_df) > 0:
        rank_col1, rank_col2 = st.columns([1.5, 1.5])
        with rank_col1:
            rank_selected_defect = st.selectbox(t["rank_sel_def"], current_active_defects, key="rank_def_sel")
        with rank_col2:
            rank_dimension_option = st.selectbox(t["rank_sel_dim"], [t["dim_machine"], t["dim_person"]], key="rank_dim_sel")

        dim_field = "MachineName" if t["dim_machine"] in rank_dimension_option else "PersonName"
        if dim_field in display_df.columns:
            rank_df = display_df.groupby(dim_field, as_index=False).agg({
                rank_selected_defect: "sum",
                "INQty": "sum"
            })
            total_defect_sum = rank_df[rank_selected_defect].sum()
            rank_df["SharePercent"] = np.where(total_defect_sum > 0, (rank_df[rank_selected_defect] / total_defect_sum) * 100.0, 0.0)
            rank_df = rank_df.sort_values(rank_selected_defect, ascending=True)

            chart_title_text = f"{t['rank_chart_title_pre']}{rank_selected_defect}{t['rank_chart_title_mid']}{dim_field}{t['rank_chart_title_suf']}"
            fig_rank = px.bar(
                rank_df, x=rank_selected_defect, y=dim_field, orientation="h",
                text=rank_df.apply(lambda row: f"{row[rank_selected_defect]:,.1f} ({row['SharePercent']:.1f}%)", axis=1),
                color=rank_selected_defect, color_continuous_scale="Reds", title=chart_title_text
            )
            fig_rank.update_traces(textposition="outside")
            fig_rank.update_layout(
                xaxis=dict(title=t["rank_xaxis"], range=[0, max(1.0, rank_df[rank_selected_defect].max() * 1.30)]),
                yaxis=dict(title=dim_field), coloraxis_showscale=False, height=400, margin=dict(t=40, b=40, l=50, r=50), dragmode="pan"
            )
            st.plotly_chart(fig_rank, use_container_width=True, config=DEFAULT_CHART_CONFIG)
        else:
            st.warning(f"⚠️ Column **{dim_field}** not found.")

    st.markdown("---")

    # 良率推移趋势专区
    st.subheader(t["yield_trend_header"])
    st.markdown(f"**{t['target_header']}**")
    target_col1, target_col2 = st.columns(2, gap="large")
    with target_col1:
        target_fpy_machine = st.slider(t["t_machine"], 60.0, 100.0, 99.0, 0.1, key="slider_fpy_machine")
    with target_col2:
        target_fpy_date = st.slider(t["t_date"], 60.0, 100.0, 99.0, 0.1, key="slider_fpy_date")

    st.markdown("<br>", unsafe_allow_html=True)

    if "TestDate" in display_df.columns and has_fpy and len(display_df) > 0:
        y_range_col1, _ = st.columns([1, 3])
        with y_range_col1:
            custom_fpy_min = st.number_input(t["y_min_label"], min_value=0.0, max_value=99.0, value=80.0, step=1.0, key="yield_min_input_main")

        trend_raw_df = display_df.dropna(subset=["TestDate", "FPY"]).sort_values("TestDate")
        if not trend_raw_df.empty:
            st.markdown(f"**{t['t_m_trend']}**")
            all_machines = sorted(trend_raw_df["MachineName"].dropna().astype(str).unique().tolist()) if "MachineName" in trend_raw_df.columns else []
            if all_machines:
                selected_machines = st.multiselect(t["machine_filter_label"], options=all_machines, default=all_machines, key="trend_machine_multiselect")
                filtered_trend_df = trend_raw_df[trend_raw_df["MachineName"].astype(str).isin(selected_machines)] if selected_machines else trend_raw_df
            else:
                filtered_trend_df = trend_raw_df

            hover_data_cols = ["BatchNo"] if "BatchNo" in filtered_trend_df.columns else []
            fig_machine_trend = px.line(filtered_trend_df, x="TestDate", y="FPY", color="MachineName" if "MachineName" in filtered_trend_df.columns else None, markers=True, hover_data=hover_data_cols)
            fig_machine_trend.add_hline(y=target_fpy_machine, line_dash="dash", line_color="red", annotation_text=f"Target: {target_fpy_machine}%", annotation_position="top left")
            fig_machine_trend.update_layout(xaxis=dict(title=t["x_time"]), yaxis=dict(title=t["y_roll_fpy"], range=[custom_fpy_min, 100.5]), height=430, margin=dict(t=30, b=30), dragmode="pan")
            st.plotly_chart(fig_machine_trend, use_container_width=True, config=DEFAULT_CHART_CONFIG)

            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown(f"**{t['t_d_trend']}**")
            daily_df = trend_raw_df.copy()
            daily_df["DayStr"] = daily_df["TestDate"].dt.strftime("%Y-%m-%d")
            daily_agg = daily_df.groupby("DayStr", as_index=False).agg(FPY=("FPY", "mean"), INQty=("INQty", "sum"), OutQty=("OutQty", "sum")).sort_values("DayStr")
            global_daily_yield_agg = daily_agg.copy()

            fig_date_combo = go.Figure()
            fig_date_combo.add_trace(go.Bar(x=daily_agg["DayStr"], y=daily_agg["INQty"], name=t["legend_in"], marker_color="rgba(99, 110, 250, 0.72)", yaxis="y", offsetgroup=0))
            fig_date_combo.add_trace(go.Bar(x=daily_agg["DayStr"], y=daily_agg["OutQty"], name=t["legend_out"], marker_color="rgba(0, 204, 150, 0.75)", yaxis="y", offsetgroup=1))
            fig_date_combo.add_trace(go.Scatter(x=daily_agg["DayStr"], y=daily_agg["FPY"], name=t["legend_fpy"], mode="lines+markers+text", marker=dict(color="#FFA15A", size=8), line=dict(color="#FFA15A", width=3), text=daily_agg["FPY"].apply(lambda v: f"{v:.2f}%"), textposition="top center", cliponaxis=False, yaxis="y2"))
            fig_date_combo.add_shape(type="line", xref="paper", x0=0, x1=1, yref="y2", y0=target_fpy_date, y1=target_fpy_date, line=dict(color="red", width=2, dash="dash"))
            fig_date_combo.add_annotation(xref="paper", x=0.01, yref="y2", y=target_fpy_date, text=f"Target: {target_fpy_date}%", showarrow=False, yshift=10, font=dict(color="red", size=12))

            max_qty_bar = max(daily_agg["INQty"].max(), daily_agg["OutQty"].max()) if not daily_agg.empty else 10
            fig_date_combo.update_layout(
                barmode="group",
                xaxis=dict(title=t["x_date"], type="category", tickmode="array", tickvals=daily_agg["DayStr"].tolist(), ticktext=daily_agg["DayStr"].tolist(), tickangle=-35, tickfont=dict(size=11)),
                yaxis=dict(title=t["y_qty_axis"], range=[0, max(10, max_qty_bar * 1.25)], showgrid=True),
                yaxis2=dict(title=t["y_daily_fpy"], overlaying="y", side="right", range=[custom_fpy_min, 100.8], showgrid=False),
                legend=dict(orientation="h", yanchor="bottom", y=1.12, xanchor="center", x=0.5),
                height=470, margin=dict(t=70, b=75), dragmode="pan"
            )
            st.plotly_chart(fig_date_combo, use_container_width=True, config=DEFAULT_CHART_CONFIG)

    st.markdown("---")
    st.subheader(t["table_detail"])

    output_buffer = io.BytesIO()
    with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
        summary_kpi_data = pd.DataFrame({
            "Metric Item": ["Total Production Batches", "Total Input Qty", "Total Output Qty", "Total Scrap Qty", "Average Roll Yield FPY"],
            "Value": [len(display_df), total_in_qty, total_out_qty, total_scrap_qty, f"{avg_fpy:.2f}%" if has_fpy else f"{overall_yield:.2f}%"]
        })
        summary_kpi_data.to_excel(writer, sheet_name="Dashboard Summary", index=False, startrow=0)
        if not global_defect_sums.empty:
            start_row_pareto = len(summary_kpi_data) + 3
            pd.DataFrame([["=== Pareto Defect Statistics ==="]]).to_excel(writer, sheet_name="Dashboard Summary", index=False, header=False, startrow=start_row_pareto-1)
            global_defect_sums.to_excel(writer, sheet_name="Dashboard Summary", index=False, startrow=start_row_pareto)
        display_df.to_excel(writer, sheet_name="Raw Data (Filtered)", index=False)

    output_buffer.seek(0)
    wb = load_workbook(output_buffer)
    excel_data = output_buffer.getvalue()

    _, export_col2 = st.columns([3, 1])
    with export_col2:
        st.download_button(
            label=t["btn_export"], data=excel_data, file_name="Tenpower_Electrode_Full_Chart_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True
        )

    st.dataframe(display_df.style.format({"FPY": "{:.2f}%"} if has_fpy else {}, na_rep="-"), use_container_width=True)

    st.markdown("---")
    st.markdown("### 📋 MES 数据智能校验与分流处理明细反馈")
    tab_det1, tab_det2 = st.tabs([f"🛠️ 涂布工序自动修正明细 ({len(coating_fixed_df)} 条)", f"🚀 分切来料多(too much)保留放行明细 ({len(preserved_too_much_df)} 条)"])
    with tab_det1:
        if not coating_fixed_df.empty:
            st.info("以下涂布工序批次因产出大于投入已被系统自动化逻辑安全修正：")
            st.dataframe(coating_fixed_df[[c for c in ["BatchNo", "MasterLotNo", "ProcessName", "MachineName", "INQty", "OutQty", "ScrapQty", "TestDate"] if c in coating_fixed_df.columns]], use_container_width=True, hide_index=True)
        else:
            st.success("无涂布工序自动修正记录。")
    with tab_det2:
        if not preserved_too_much_df.empty:
            st.info("以下分切工序批次因包含 too much 保留原数据执行：")
            st.dataframe(preserved_too_much_df[[c for c in ["BatchNo", "MasterLotNo", "ProcessName", "MachineName", "INQty", "OutQty", "ScrapQty", "TestDate"] if c in preserved_too_much_df.columns]], use_container_width=True, hide_index=True)
        else:
            st.success("无分切来料多保留放行记录。")


# ==============================================================================
# 🔵 模式二：周期性双文件对比分析 (v2.0 升级版)
# ==============================================================================
elif system_mode == t["mode_v2"]:
    st.sidebar.header("📁 双周期数据源导入 (v2.0)")
    st.sidebar.markdown("请分别上传**上周/基准期文件**与**本周/对比期文件**（格式如 `MesResume_DoMethod-W10.xlsx`）：")
    
    file_prev = st.sidebar.file_uploader("📥 上周 / 基准期文件 (e.g., -W10)", type=["xlsx", "xls", "csv"], key="file_prev")
    file_curr = st.sidebar.file_uploader("📥 本周 / 对比期文件 (e.g., -W11)", type=["xlsx", "xls", "csv"], key="file_curr")

    if file_prev is None or file_curr is None:
        st.info("👈 请在左侧边栏同时上传 **上周（基准期）** 与 **本周（对比期）** 两个 MES 数据文件（支持 MesResume_DoMethod-Wxx 格式），以启动周期双文件对比功能。")
        st.stop()

    raw_df_prev = load_uploaded_file(file_prev)
    raw_df_curr = load_uploaded_file(file_curr)

    if raw_df_prev is None or raw_df_curr is None:
        st.stop()

    label_prev = extract_week_label(file_prev.name)
    label_curr = extract_week_label(file_curr.name)

    # 分别对两周数据应用完整的数据清洗与校验逻辑
    df_prev, coat_fix_prev, pres_toomuch_prev, ext_prev, _ = process_and_clean_dataframe(raw_df_prev)
    df_curr, coat_fix_curr, pres_toomuch_curr, ext_curr, _ = process_and_clean_dataframe(raw_df_curr)

    st.subheader(f"📊 周期性双文件对比看板：【{label_prev}】 VS 【{label_curr}】")
    st.markdown("---")

    # 双周数据校验报告展示
    total_cleaned_prev = len(coat_fix_prev) + len(pres_toomuch_prev)
    total_cleaned_curr = len(coat_fix_curr) + len(pres_toomuch_curr)
    
    if total_cleaned_prev > 0 or total_cleaned_curr > 0:
        st.warning(f"⚠️ **双周期数据清洗与校验提示**：上周文件清洗放行 **{total_cleaned_prev}** 条；本周文件清洗放行 **{total_cleaned_curr}** 条。详见底部明细。")

    exclude_keywords = ["lotno", "batch", "person", "machine", "date", "class", "product", "process", "qty", "fpy", "scrap"]
    
    def get_defect_summary(df_target):
        in_qty = df_target["INQty"].sum() if "INQty" in df_target.columns else 1.0
        defect_cols = []
        for col in df_target.columns:
            if not any(kw in col.lower() for kw in exclude_keywords):
                conv = pd.to_numeric(df_target[col], errors="coerce")
                if conv.notna().sum() > 0:
                    df_target[col] = conv.fillna(0)
                    defect_cols.append(col)
        
        if not defect_cols:
            return pd.DataFrame(), in_qty
            
        sums = df_target[defect_cols].sum().reset_index()
        sums.columns = ["DefectItem", "DefectQty"]
        sums = sums[sums["DefectQty"] > 0].sort_values("DefectQty", ascending=False).reset_index(drop=True)
        total_def = sums["DefectQty"].sum() if not sums.empty else 1.0
        sums["DefectShare"] = (sums["DefectQty"] / total_def) * 100.0
        sums["CumShare"] = sums["DefectShare"].cumsum()
        sums["LossRate"] = (sums["DefectQty"] / in_qty) * 100.0
        return sums, in_qty

    summary_prev, in_prev = get_defect_summary(df_prev)
    summary_curr, in_curr = get_defect_summary(df_curr)

    merged_comp = pd.merge(
        summary_curr[["DefectItem", "DefectQty", "DefectShare", "LossRate", "CumShare"]],
        summary_prev[["DefectItem", "DefectQty", "DefectShare", "LossRate"]],
        on="DefectItem", how="outer", suffixes=("_Curr", "_Prev")
    ).fillna(0)

    merged_comp = merged_comp.sort_values(by="DefectQty_Curr", ascending=False).reset_index(drop=True)

    st.markdown("### 📊 柱状图重叠对比：本周（实色柱）与上周（红色空心边框浮于上方）")
    
    fig_comp_pareto = go.Figure()

    # 1. 先添加本周实色柱（作为底层柱展示）
    fig_comp_pareto.add_trace(go.Bar(
        x=merged_comp["DefectItem"],
        y=merged_comp["LossRate_Curr"],
        name=f"本周: {label_curr} (Loss Rate %)",
        marker_color="indianred",
        text=merged_comp["LossRate_Curr"].apply(lambda v: f"{v:.2f}%" if v > 0 else ""),
        textposition="outside",
        yaxis="y"
    ))

    # 2. 后添加上周条形图（使用 overlay 重叠，marker 采用红色无填充空心框，浮在上方方便清晰查看）
    fig_comp_pareto.add_trace(go.Bar(
        x=merged_comp["DefectItem"],
        y=merged_comp["LossRate_Prev"],
        name=f"上周: {label_prev} (Loss Rate %)",
        marker=dict(
            color="rgba(0,0,0,0)",
            line=dict(color="red", width=2.5)
        ),
        text=merged_comp["LossRate_Prev"].apply(lambda v: f"{v:.2f}%" if v > 0 else ""),
        textposition="outside",
        yaxis="y"
    ))

    # 3. 本周累积百分比折线图 (Pareto Line)
    fig_comp_pareto.add_trace(go.Scatter(
        x=merged_comp["DefectItem"],
        y=merged_comp["CumShare"],
        name=f"本周累积占比 (%)",
        mode="lines+markers+text",
        marker=dict(color="darkorange", size=8),
        line=dict(color="darkorange", width=2.5),
        text=merged_comp["CumShare"].apply(lambda v: f"{v:.0f}%" if v > 0 else ""),
        textposition="top center",
        yaxis="y2"
    ))

    max_loss = max(merged_comp["LossRate_Curr"].max(), merged_comp["LossRate_Prev"].max()) if not merged_comp.empty else 1.0

    fig_comp_pareto.update_layout(
        barmode="overlay",  # 使两周柱子重叠，且后添加的上周红框浮于本周实柱上方
        xaxis=dict(tickangle=-25),
        yaxis=dict(title="Input Loss Rate (%)", range=[0, max_loss * 1.35]),
        yaxis2=dict(title="Cumulative Share (%)", overlaying="y", side="right", range=[0, 125], showgrid=False),
        height=520,
        margin=dict(t=40, b=80, l=50, r=50),
        dragmode="pan"
    )
    st.plotly_chart(fig_comp_pareto, use_container_width=True, config=DEFAULT_CHART_CONFIG)

    st.markdown("### 📋 双周不良环比对照明细表")
    display_comp_table = merged_comp[["DefectItem", "DefectQty_Prev", "LossRate_Prev", "DefectQty_Curr", "LossRate_Curr"]].copy()
    display_comp_table.columns = ["不良项目", "上周不良数", "上周损失率 (%)", "本周不良数", "本周损失率 (%)"]
    display_comp_table["环比数量差 (本周-上周)"] = display_comp_table["本周不良数"] - display_comp_table["上周不良数"]
    
    st.dataframe(
        display_comp_table.style.format({
            "上周不良数": "{:,.1f}", "上周损失率 (%)": "{:.2f}%",
            "本周不良数": "{:,.1f}", "本周损失率 (%)": "{:.2f}%",
            "环比数量差 (本周-上周)": "{:+,.1f}"
        }),
        use_container_width=True,
        hide_index=True
    )

    # 4. 底部展示双周数据清洗与分流明细反馈
    st.markdown("---")
    st.markdown("### 📋 双周期 MES 数据智能校验与清洗明细反馈")
    
    comp_tab1, comp_tab2 = st.tabs([
        f"🛠️ 涂布工序自动修正明细 (上周: {len(coat_fix_prev)}条 | 本周: {len(coat_fix_curr)}条)", 
        f"🚀 分切来料多(too much)放行明细 (上周: {len(pres_toomuch_prev)}条 | 本周: {len(pres_toomuch_curr)}条)"
    ])

    with comp_tab1:
        col_p1, col_c1 = st.columns(2)
        with col_p1:
            st.markdown(f"**【{label_prev} - 涂布修正明细】**")
            if not coat_fix_prev.empty:
                cols_c = [c for c in ["BatchNo", "MasterLotNo", "ProcessName", "MachineName", "INQty", "OutQty", "ScrapQty", "TestDate"] if c in coat_fix_prev.columns]
                st.dataframe(coat_fix_prev[cols_c], use_container_width=True, hide_index=True)
            else:
                st.success("上周无涂布修正记录。")
        with col_c1:
            st.markdown(f"**【{label_curr} - 涂布修正明细】**")
            if not coat_fix_curr.empty:
                cols_c = [c for c in ["BatchNo", "MasterLotNo", "ProcessName", "MachineName", "INQty", "OutQty", "ScrapQty", "TestDate"] if c in coat_fix_curr.columns]
                st.dataframe(coat_fix_curr[cols_c], use_container_width=True, hide_index=True)
            else:
                st.success("本周无涂布修正记录。")

    with comp_tab2:
        col_p2, col_c2 = st.columns(2)
        with col_p2:
            st.markdown(f"**【{label_prev} - 来料多保留放行明细】**")
            if not pres_toomuch_prev.empty:
                cols_p = [c for c in ["BatchNo", "MasterLotNo", "ProcessName", "MachineName", "INQty", "OutQty", "ScrapQty", "TestDate"] if c in pres_toomuch_prev.columns]
                st.dataframe(pres_toomuch_prev[cols_p], use_container_width=True, hide_index=True)
            else:
                st.success("上周无来料多保留记录。")
        with col_c2:
            st.markdown(f"**【{label_curr} - 来料多保留放行明细】**")
            if not pres_toomuch_curr.empty:
                cols_p = [c for c in ["BatchNo", "MasterLotNo", "ProcessName", "MachineName", "INQty", "OutQty", "ScrapQty", "TestDate"] if c in pres_toomuch_curr.columns]
                st.dataframe(pres_toomuch_curr[cols_p], use_container_width=True, hide_index=True)
            else:
                st.success("本周无来料多保留记录。")