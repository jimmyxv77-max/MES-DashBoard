import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import io

# 1. 页面基本配置
st.set_page_config(
    page_title="Tenpower electrode dashboard",
    page_icon="📊",
    layout="wide"
)

# ================= 语言词典配置 (中/英) =================
LANG_DICT = {
    "zh": {
        "title": "📊 Tenpower electrode dashboard",
        "author": "created by xu",
        "import_header": "📁 数据源导入",
        "import_help": "请将 MES 生产数据表拖拽至此处，或点击浏览（支持 MesResume_DoMethod 前缀）",
        "import_info": "👈 请在左侧边栏拖拽或选择你的 MES 生产数据表（支持前缀为 **MesResume_DoMethod** 的各类带后缀文件名）以生成实时看板。",
        "export_header": "📥 报告导出",
        "btn_export": "📥 导出看板图表汇总报表 (Excel)",
        "changelog_title": "📌 版本更新说明 (What's New)",
        "changelog_text": """
        - **图表数据一键导出**：一键导出包含核心 KPI 摘要与帕累托异常图表统计的汇总 Excel 报表。
        - **多工序支持**：适配涂布、辊压、分切等多工序电极生产数据分析。
        - **全宽布局优化**：不良明细表、帕累托图与良率趋势独立全宽展示，告别横向滚动。
        - **质量目标线调优**：良率目标线默认值升级为 **99.0%**（支持 60%~100% 自由拖动）。
        - **双语无缝切换**：右上角可随时切换中英文。
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
        "pareto_title": "📊 异常帕累托图（柱状=实际数量，折线=累计不良占比）",
        "radio_share": "选择柱状图展示维度：",
        "share_1": "占总不良比例 (%)",
        "share_2": "占总投入比例 (%) [损失率]",
        "table_header": "📋 异常数量与双维度比例对照表",
        "col_defect": "不良项目",
        "col_count": "不良数量",
        "col_s_defect": "占总不良比例",
        "col_s_in": "占总投入比例",
        "col_cum": "累计不良占比",
        "defect_trend_header": "🔍 单项不良变化推移图（数量推移 & 不良率推移）",
        "def_sel": "🎯 1. 选择追踪的不良项目：",
        "def_view": "📊 2. 选择观察粒度：",
        "view_daily": "按单日汇总趋势",
        "view_roll": "按单卷明细推移（分机台）",
        "def_metric": "📐 3. 选择统计指标：",
        "m_count": "不良数量推移",
        "m_rate": "不良率推移 (% 总投入)",
        "y_max_label": "🎚️ 单项不良 Y 轴上限：",
        "yield_trend_header": "📈 良率推移趋势分析",
        "y_min_label": "🎚️ 良率图 Y 轴下限 (%)：",
        "t_m_trend": "📈 1. 各机台单卷良率 (FPY) 推移走势（含目标线）",
        "t_d_trend": "📅 2. 单日平均良率变化推移走势（含目标线 · 完整日期刻度）",
        "table_detail": "📋 生产批次明细数据表（当前筛选结果）",
        "x_time": "生产时间",
        "x_date": "生产日期",
        "y_roll_fpy": "单卷良率 FPY (%)",
        "y_daily_fpy": "单日平均良率 (%)",
    },
    "en": {
        "title": "📊 Tenpower electrode dashboard",
        "author": "created by xu",
        "import_header": "📁 Data Source Import",
        "import_help": "Drag and drop your MES production data table here (supports MesResume_DoMethod prefix)",
        "import_info": "👈 Please drag or select your MES production data table in the sidebar (supporting files prefixed with **MesResume_DoMethod**) to generate the live dashboard.",
        "export_header": "📥 Report Export",
        "btn_export": "📥 Export Chart Summary Report (Excel)",
        "changelog_title": "📌 What's New",
        "changelog_text": """
        - **Chart Data Export**: Export Excel reports containing core KPI summaries and Pareto defect chart statistics.
        - **Multi-Process Support**: Tailored for coating, calendering, slitting, and electrode manufacturing workflows.
        - **Full-Width Layout**: Pareto charts, defect tables, and yield trends optimized in full width with zero horizontal scrolling.
        - **Target Line Tuning**: Default yield target upgraded to **99.0%** (adjustable from 60% to 100%).
        - **Bilingual Switch**: Instant CN/EN toggle via top-right selector.
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
        "pareto_title": "📊 Defect Pareto Chart (Bar = Qty, Line = Cum. Share %)",
        "radio_share": "Select Bar Chart Metric:",
        "share_1": "Defect Share (%)",
        "share_2": "Loss Rate of Total Input (%)",
        "table_header": "📋 Defect Quantity & Dual-Dimension Share Table",
        "col_defect": "Defect Item",
        "col_count": "Defect Qty",
        "col_s_defect": "Defect Share",
        "col_s_in": "Input Loss Rate",
        "col_cum": "Cum. Share",
        "defect_trend_header": "🔍 Single Defect Trend Analysis (Quantity & Defect Rate)",
        "def_sel": "🎯 1. Select Defect Item:",
        "def_view": "📊 2. Select View Granularity:",
        "view_daily": "Daily Summary Trend",
        "view_roll": "Roll-level Detail Trend (by Machine)",
        "def_metric": "📐 3. Select Metric:",
        "m_count": "Defect Quantity Trend",
        "m_rate": "Defect Rate Trend (% of Input)",
        "y_max_label": "🎚️ Single Defect Y-Axis Upper Limit:",
        "yield_trend_header": "📈 Yield Trend Analysis",
        "y_min_label": "🎚️ Yield Chart Y-Axis Lower Limit (%):",
        "t_m_trend": "📈 1. Machine Roll Yield (FPY) Trend (with Target Line)",
        "t_d_trend": "📅 2. Daily Average Yield Trend (with Target Line · Full Dates)",
        "table_detail": "📋 Production Batch Detail Table (Current Filtered)",
        "x_time": "Production Time",
        "x_date": "Production Date",
        "y_roll_fpy": "Roll Yield FPY (%)",
        "y_daily_fpy": "Daily Avg Yield (%)",
    }
}

# 顶部显眼位置放置语言切换按钮
top_col1, top_col2 = st.columns([4, 1])
with top_col2:
    selected_lang_name = st.selectbox("🌐 Language / 语言", ["中文 (Chinese)", "English"], index=0)
    curr_lang = "zh" if "中文" in selected_lang_name else "en"

t = LANG_DICT[curr_lang]

# 英文大标题 + 署名小字
st.title(t["title"])
st.markdown(f"*{t['author']}*")
st.markdown("---")

# 2. 侧边栏：文件拖拽上传区与版本更新说明
st.sidebar.header(t["import_header"])
uploaded_file = st.sidebar.file_uploader(
    t["import_help"],
    type=["xlsx", "xls", "csv"],
    help=t["import_help"]
)

# 在侧边栏下方常驻显示更新日志
st.sidebar.markdown("---")
st.sidebar.markdown(t["changelog_text"])

if uploaded_file is None:
    st.info(t["import_info"])
    st.stop()

file_name = uploaded_file.name
if not file_name.lower().startswith("mesresume_domethod"):
    st.warning(f"⚠️ Warning: File **{file_name}** does not start with `MesResume_DoMethod`. Trying to parse anyway.")

# 3. 读取上传的真实文件
try:
    if file_name.endswith(".csv"):
        try:
            df = pd.read_csv(uploaded_file, encoding="utf-8")
        except UnicodeDecodeError:
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file, encoding="gbk")
    else:
        df = pd.read_excel(uploaded_file)
except Exception as e:
    st.error(f"❌ File read error: {e}")
    st.stop()

# 清理列名字段前后的空格
df.columns = df.columns.str.strip()

# 检查核心列是否存在
required_cols = ["INQty", "OutQty", "ScrapQty"]
missing_cols = [c for c in required_cols if c not in df.columns]
if missing_cols:
    st.error(f"❌ Missing required columns: {', '.join(missing_cols)}")
    st.stop()

# 日期格式标准化
if "TestDate" in df.columns:
    df["TestDate"] = pd.to_datetime(df["TestDate"], errors="coerce")

# ================= 4. 核心数据逻辑验证与自动修正 =================
for col in ["INQty", "OutQty", "ScrapQty"]:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

has_fpy = "FPY" in df.columns
if has_fpy:
    df["FPY"] = pd.to_numeric(df["FPY"], errors="coerce")
    if df["FPY"].dropna().max() <= 1.5:  # 小数形式转百分比
        df["FPY"] = df["FPY"] * 100.0

if "INQty" in df.columns and "OutQty" in df.columns and "ScrapQty" in df.columns:
    abnormal_mask = df["INQty"] < df["OutQty"]
    fixed_count = int(abnormal_mask.sum())

    if fixed_count > 0:
        df.loc[abnormal_mask, "OutQty"] = df.loc[abnormal_mask, "INQty"] - df.loc[abnormal_mask, "ScrapQty"]
        df.loc[df["OutQty"] < 0, "OutQty"] = 0
        
        if has_fpy:
            valid_in = abnormal_mask & (df["INQty"] > 0)
            df.loc[valid_in, "FPY"] = (df.loc[valid_in, "OutQty"] / df.loc[valid_in, "INQty"]) * 100.0
            df.loc[abnormal_mask & (df["INQty"] <= 0), "FPY"] = 0.0

        st.toast(f"🛠️ Data auto-cleaned: Fixed {fixed_count} records where INQty < OutQty.", icon="✅")
else:
    fixed_count = 0

# ================= 5. 精准识别 N 列及以后的动态不良矩阵 =================
exclude_keywords = ["lotno", "batch", "person", "machine", "date", "class", "product", "process", "qty", "fpy", "scrap"]

all_columns = df.columns.tolist()
DYNAMIC_DEFECT_COLS = []

for col in all_columns:
    col_lower = col.lower()
    is_excluded = any(kw in col_lower for kw in exclude_keywords)
    if is_excluded:
        continue
    
    converted = pd.to_numeric(df[col], errors="coerce")
    if converted.notna().sum() > 0:
        df[col] = converted.fillna(0)
        DYNAMIC_DEFECT_COLS.append(col)

# ================= 6. 侧边栏筛选器 =================
st.sidebar.header(t["sidebar_filter"])

search_keyword = st.sidebar.text_input(
    t["search_label"], 
    ""
).strip()

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
        selected = st.sidebar.multiselect(label, options, default=[])
        if selected:
            display_df = display_df[display_df[col_name].astype(str).isin(selected)]

if search_keyword:
    search_cols = [c for c in ["BatchNo", "MasterLotNo"] if c in display_df.columns]
    if search_cols:
        mask = pd.Series(False, index=display_df.index)
        for c in search_cols:
            mask = mask | display_df[c].astype(str).str.contains(search_keyword, case=False, na=False)
        display_df = display_df[mask]

show_only_abnormal = st.sidebar.checkbox(t["chk_abnormal"], value=False)
if show_only_abnormal and has_fpy:
    display_df = display_df[display_df["FPY"] < 95.0]

# 7. 核心指标展示 (KPI Cards)
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

if fixed_count > 0:
    st.info(f"ℹ️ Data Validation Note: {fixed_count} records with `INQty < OutQty` were auto-corrected.")

st.markdown("---")

# ================= 8. 不良分析专区：全宽帕累托图与全宽明细表 =================
current_active_defects = [c for c in DYNAMIC_DEFECT_COLS if c in display_df.columns]

st.subheader(t["pareto_header"])

global_defect_sums = pd.DataFrame()

if current_active_defects and len(display_df) > 0:
    defect_sums = display_df[current_active_defects].sum().reset_index()
    defect_sums.columns = [t["col_defect"], t["col_count"]]
    defect_sums = defect_sums[defect_sums[t["col_count"]] > 0].sort_values(t["col_count"], ascending=False).reset_index(drop=True)
    
    if not defect_sums.empty:
        total_defects = defect_sums[t["col_count"]].sum()
        
        if total_defects > 0:
            defect_sums[t["col_s_defect"]] = (defect_sums[t["col_count"]] / total_defects) * 100
            defect_sums[t["col_cum"]] = defect_sums[t["col_s_defect"]].cumsum()
        else:
            defect_sums[t["col_s_defect"]] = 0.0
            defect_sums[t["col_cum"]] = 0.0
            
        if total_in_qty > 0:
            defect_sums[t["col_s_in"]] = (defect_sums[t["col_count"]] / total_in_qty) * 100
        else:
            defect_sums[t["col_s_in"]] = 0.0

        global_defect_sums = defect_sums.copy()

        # 8.1 帕累托图（全宽）
        st.markdown(f"**{t['pareto_title']}**")
        
        fig_pareto = go.Figure()
        
        fig_pareto.add_trace(go.Bar(
            x=defect_sums[t["col_defect"]],
            y=defect_sums[t["col_count"]],
            name="Defect Qty" if curr_lang=="en" else "实际不良数量",
            marker_color="indianred",
            text=defect_sums[t["col_count"]].apply(lambda v: f"{v:,.1f}"),
            textposition="outside",
            cliponaxis=False,
            yaxis="y"
        ))
        
        fig_pareto.add_trace(go.Scatter(
            x=defect_sums[t["col_defect"]],
            y=defect_sums[t["col_cum"]],
            name="Cum. Share %" if curr_lang=="en" else "累计不良占比 (%)",
            mode="lines+markers+text",
            marker=dict(color="darkblue", size=8),
            line=dict(color="darkblue", width=2.5),
            text=defect_sums[t["col_cum"]].apply(lambda v: f"{v:.1f}%"),
            textposition="top center",
            textfont=dict(color="darkblue", size=11),
            cliponaxis=False,
            yaxis="y2"
        ))
        
        max_bar = defect_sums[t["col_count"]].max() if defect_sums[t["col_count"]].max() > 0 else 10
        fig_pareto.update_layout(
            title="",
            xaxis=dict(tickangle=-20, tickfont=dict(size=11)),
            yaxis=dict(title="Defect Qty" if curr_lang=="en" else "实际不良数量", side="left", range=[0, max_bar * 1.30]),
            yaxis2=dict(title="Cum. Share (%)" if curr_lang=="en" else "累计不良占比 (%)", overlaying="y", side="right", range=[0, 125], showgrid=False),
            legend=dict(x=0.01, y=1.12, orientation="h", bgcolor="rgba(0,0,0,0)"),
            height=500,
            margin=dict(t=40, b=80, l=50, r=50)
        )
        st.plotly_chart(fig_pareto, use_container_width=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # 8.2 单项排行条形图（全宽展示）
        ratio_mode = st.radio(
            t["radio_share"],
            [t["share_1"], t["share_2"]],
            horizontal=True
        )
        
        target_col = t["col_s_defect"] if t["share_1"] in ratio_mode else t["col_s_in"]
        color_scale = "Oranges" if t["share_1"] in ratio_mode else "Purples"
        chart_title = f"{target_col} Ranking (%)" if curr_lang=="en" else f"各异常{target_col}排行 (%)"
        
        st.markdown(f"**{chart_title}**")
        
        share_sorted = defect_sums.sort_values(target_col, ascending=True)
        fig_share = px.bar(
            share_sorted,
            x=target_col,
            y=t["col_defect"],
            orientation="h",
            text=share_sorted[target_col].apply(lambda v: f"{v:.2f}%"),
            color=target_col,
            color_continuous_scale=color_scale
        )
        max_share = share_sorted[target_col].max() if share_sorted[target_col].max() > 0 else 1
        fig_share.update_traces(textposition="outside", cliponaxis=False)
        fig_share.update_layout(
            title="",
            xaxis=dict(title=f"{target_col} (%)", range=[0, max_share * 1.30]),
            yaxis=dict(title=""),
            coloraxis_showscale=False,
            height=400,
            margin=dict(t=10, b=20, l=10, r=45)
        )
        st.plotly_chart(fig_share, use_container_width=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # 8.3 异常数量与双维度比例对照表（全宽干净表头）
        st.markdown(f"**{t['table_header']}**")
        stat_table = defect_sums[[t["col_defect"], t["col_count"], t["col_s_defect"], t["col_s_in"], t["col_cum"]]].copy()
        st.dataframe(
            stat_table.style.format({
                t["col_count"]: "{:,.1f}",
                t["col_s_defect"]: "{:.2f}%",
                t["col_s_in"]: "{:.2f}%",
                t["col_cum"]: "{:.2f}%"
            }),
            use_container_width=True,
            hide_index=True
        )
    else:
        st.info("No defect data available under current filters.")
else:
    st.warning("No dynamic defect columns detected.")

st.markdown("---")

# ================= 9. 动态单项不良变化推移图 =================
st.subheader(t["defect_trend_header"])

if current_active_defects and "TestDate" in display_df.columns and len(display_df) > 0:
    ctrl_col1, ctrl_col2, ctrl_col3 = st.columns([1.1, 1.3, 1.1])
    with ctrl_col1:
        selected_defect = st.selectbox(
            t["def_sel"],
            current_active_defects,
            index=0
        )
    with ctrl_col2:
        defect_view_mode = st.radio(
            t["def_view"],
            [t["view_daily"], t["view_roll"]],
            horizontal=True
        )
    with ctrl_col3:
        defect_metric_mode = st.radio(
            t["def_metric"],
            [t["m_count"], t["m_rate"]],
            horizontal=True
        )

    defect_trend_df = display_df.dropna(subset=["TestDate"]).sort_values("TestDate").copy()
    is_rate_mode = t["m_rate"] in defect_metric_mode

    if not defect_trend_df.empty:
        if t["view_daily"] in defect_view_mode:
            defect_trend_df["DayStr"] = defect_trend_df["TestDate"].dt.strftime("%Y-%m-%d")
            temp_agg = defect_trend_df.groupby("DayStr", as_index=False).agg({
                selected_defect: "sum",
                "INQty": "sum"
            })
            temp_agg["DefectRate"] = np.where(temp_agg["INQty"] > 0, (temp_agg[selected_defect] / temp_agg["INQty"]) * 100.0, 0.0)
            auto_max = float(temp_agg["DefectRate" if is_rate_mode else selected_defect].max())
        else:
            defect_trend_df["RollDefectRate"] = np.where(defect_trend_df["INQty"] > 0, (defect_trend_df[selected_defect] / defect_trend_df["INQty"]) * 100.0, 0.0)
            auto_max = float(defect_trend_df["RollDefectRate" if is_rate_mode else selected_defect].max())

        auto_max = max(0.5 if is_rate_mode else 1.0, auto_max * 1.25)

        adj_col1, adj_col2 = st.columns([1, 3])
        with adj_col1:
            custom_y_max = st.number_input(
                t["y_max_label"],
                min_value=0.1,
                max_value=float(max(100000.0, auto_max * 5.0)),
                value=float(auto_max),
                step=0.1 if is_rate_mode else 10.0
            )

        if t["view_daily"] in defect_view_mode:
            daily_defect_agg = defect_trend_df.groupby("DayStr", as_index=False).agg({
                selected_defect: "sum",
                "INQty": "sum"
            }).sort_values("DayStr")
            
            daily_defect_agg["DefectRate"] = np.where(
                daily_defect_agg["INQty"] > 0,
                (daily_defect_agg[selected_defect] / daily_defect_agg["INQty"]) * 100.0,
                0.0
            )
            
            y_col = "DefectRate" if is_rate_mode else selected_defect
            y_label = f"{selected_defect} Daily Rate (%)" if curr_lang=="en" else f"{selected_defect} 单日不良率 (%)" if is_rate_mode else f"{selected_defect} 单日累计数量"
            text_labels = daily_defect_agg[y_col].apply(
                lambda v: f"{v:.2f}%" if is_rate_mode else (f"{v:,.1f}" if v > 0 else "0")
            )
            line_color = "darkorange" if is_rate_mode else "crimson"

            fig_def_daily = px.line(
                daily_defect_agg,
                x="DayStr",
                y=y_col,
                markers=True,
                text=text_labels
            )
            fig_def_daily.update_traces(
                line=dict(color=line_color, width=2.5),
                marker=dict(size=8, color=line_color),
                textposition="top center",
                cliponaxis=False
            )
            fig_def_daily.update_layout(
                xaxis=dict(
                    title="Date" if curr_lang=="en" else "生产日期",
                    type="category",
                    tickmode="array",
                    tickvals=daily_defect_agg["DayStr"].tolist(),
                    ticktext=daily_defect_agg["DayStr"].tolist(),
                    tickangle=-35,
                    tickfont=dict(size=11)
                ),
                yaxis=dict(title=y_label, range=[0, custom_y_max]),
                height=430,
                margin=dict(t=35, b=75)
            )
            st.plotly_chart(fig_def_daily, use_container_width=True)
        else:
            y_col = "RollDefectRate" if is_rate_mode else selected_defect
            y_label = f"{selected_defect} Roll Rate (%)" if curr_lang=="en" else f"{selected_defect} 单卷不良率 (%)" if is_rate_mode else f"{selected_defect} 单卷不良数量"

            hover_cols = [c for c in ["BatchNo", "MasterLotNo", "PersonName", "ProductCName", "ClassNo", "INQty", selected_defect, "RollDefectRate"] if c in defect_trend_df.columns]
            fig_def_roll = px.line(
                defect_trend_df,
                x="TestDate",
                y=y_col,
                color="MachineName" if "MachineName" in defect_trend_df.columns else None,
                markers=True,
                hover_data=hover_cols
            )
            fig_def_roll.update_layout(
                xaxis=dict(title="Production Time" if curr_lang=="en" else "生产时间"),
                yaxis=dict(title=y_label, range=[0, custom_y_max]),
                height=430,
                margin=dict(t=35, b=35)
            )
            st.plotly_chart(fig_def_roll, use_container_width=True)
    else:
        st.warning("No valid time data for defect trend.")
else:
    st.warning("Missing defect columns or time column.")

st.markdown("---")

# ================= 10. 良率推移趋势专区（默认目标值设为 99.0%） =================
st.subheader(t["yield_trend_header"])

# 质量目标线滑块直接放在良率分析下方（范围 60.0% ~ 100.0%，默认值设为 99.0%）
st.markdown(f"**{t['target_header']}**")
target_col1, target_col2 = st.columns(2, gap="large")
with target_col1:
    target_fpy_machine = st.slider(t["t_machine"], 60.0, 100.0, 99.0, 0.1)
with target_col2:
    target_fpy_date = st.slider(t["t_date"], 60.0, 100.0, 99.0, 0.1)

st.markdown("<br>", unsafe_allow_html=True)

if "TestDate" in display_df.columns and has_fpy and len(display_df) > 0:
    y_range_col1, y_range_col2 = st.columns([1, 3])
    with y_range_col1:
        custom_fpy_min = st.number_input(
            t["y_min_label"],
            min_value=0.0,
            max_value=99.0,
            value=80.0,
            step=1.0
        )

    trend_raw_df = display_df.dropna(subset=["TestDate", "FPY"]).sort_values("TestDate")

    if not trend_raw_df.empty:
        # 第一行：单卷良率推移图（独占全宽）
        st.markdown(f"**{t['t_m_trend']}**")
        fig_machine_trend = px.line(
            trend_raw_df,
            x="TestDate",
            y="FPY",
            color="MachineName" if "MachineName" in trend_raw_df.columns else None,
            markers=True,
            hover_data=["BatchNo", "MasterLotNo", "PersonName", "ProductCName"] if all(c in trend_raw_df.columns for c in ["BatchNo", "MasterLotNo", "PersonName", "ProductCName"]) else None
        )
        target_label_m = f"Target: {target_fpy_machine}%" if curr_lang=="en" else f"目标线: {target_fpy_machine}%"
        fig_machine_trend.add_hline(
            y=target_fpy_machine, 
            line_dash="dash", 
            line_color="red", 
            annotation_text=target_label_m, 
            annotation_position="top left"
        )
        fig_machine_trend.update_layout(
            xaxis=dict(title=t["x_time"]),
            yaxis=dict(title=t["y_roll_fpy"], range=[custom_fpy_min, 100.5]),
            height=430,
            margin=dict(t=30, b=30)
        )
        st.plotly_chart(fig_machine_trend, use_container_width=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # 第二行：单日平均良率推移图（独占全宽）
        st.markdown(f"**{t['t_d_trend']}**")
        daily_df = trend_raw_df.copy()
        daily_df["DayStr"] = daily_df["TestDate"].dt.strftime("%Y-%m-%d")
        daily_agg = daily_df.groupby("DayStr", as_index=False)["FPY"].mean().sort_values("DayStr")

        fig_date_trend = px.line(
            daily_agg,
            x="DayStr",
            y="FPY",
            markers=True,
            text=daily_agg["FPY"].apply(lambda v: f"{v:.2f}%")
        )
        fig_date_trend.update_traces(textposition="top center", cliponaxis=False)
        target_label_d = f"Target: {target_fpy_date}%" if curr_lang=="en" else f"目标线: {target_fpy_date}%"
        fig_date_trend.add_hline(
            y=target_fpy_date, 
            line_dash="dash", 
            line_color="red", 
            annotation_text=target_label_d, 
            annotation_position="top left"
        )
        fig_date_trend.update_layout(
            xaxis=dict(
                title=t["x_date"],
                type="category",
                tickmode="array",
                tickvals=daily_agg["DayStr"].tolist(),
                ticktext=daily_agg["DayStr"].tolist(),
                tickangle=-35,
                tickfont=dict(size=11)
            ),
            yaxis=dict(title=t["y_daily_fpy"], range=[custom_fpy_min, 100.5]),
            height=430,
            margin=dict(t=30, b=75)
        )
        st.plotly_chart(fig_date_trend, use_container_width=True)
    else:
        st.warning("No valid time or yield data.")
else:
    st.warning("Missing time column or FPY field.")

st.markdown("---")

# ================= 11. 生产批次明细数据表与图表汇总 Excel 导出 =================
st.subheader(t["table_detail"])

# 准备一键导出 Excel 的内存流打包逻辑（聚焦于图表汇总数据与核心指标）
output_buffer = io.BytesIO()
with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
    # 1. 汇总 Sheet: 核心 KPI 摘要与帕累托异常图表统计
    summary_kpi_data = pd.DataFrame({
        "Metric Item": [
            "Total Production Batches (总生产卷数)", 
            "Total Input Qty (总投入量)", 
            "Total Output Qty (总产出量)", 
            "Total Scrap Qty (总废弃量)", 
            "Average Roll Yield FPY (平均单卷良率)"
        ],
        "Value": [
            len(display_df), 
            total_in_qty, 
            total_out_qty, 
            total_scrap_qty, 
            f"{avg_fpy:.2f}%" if has_fpy else f"{overall_yield:.2f}%"
        ]
    })
    summary_kpi_data.to_excel(writer, sheet_name="Dashboard Summary", index=False, startrow=0)
    
    if not global_defect_sums.empty:
        # 写入空行作为分隔，再写入帕累托图表统计数据
        start_row_pareto = len(summary_kpi_data) + 3
        # 写入一个小标题行
        pd.DataFrame([["=== Pareto Defect Chart & Dual-Dimension Statistics ==="]]).to_excel(
            writer, sheet_name="Dashboard Summary", index=False, header=False, startrow=start_row_pareto-1
        )
        global_defect_sums.to_excel(writer, sheet_name="Dashboard Summary", index=False, startrow=start_row_pareto)

    # 2. 原始明细 Sheet (保留给需要查阅底表的场景)
    display_df.to_excel(writer, sheet_name="Raw Data (Filtered)", index=False)

excel_data = output_buffer.getvalue()

# 在明细表上方右侧放置下载按钮
export_col1, export_col2 = st.columns([3, 1])
with export_col2:
    st.download_button(
        label=t["btn_export"],
        data=excel_data,
    file_name="Tenpower_Electrode_Chart_Summary_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

format_dict = {"FPY": "{:.2f}%"} if has_fpy else {}
st.dataframe(
    display_df.style.format(format_dict, na_rep="-"),
    use_container_width=True
)