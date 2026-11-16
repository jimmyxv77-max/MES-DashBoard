import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import io
from openpyxl.chart import BarChart, LineChart, Reference

# ================= 🎛️ 不良目标线锁死接口 (预留配置区) =================
LOCKED_DEFECT_TARGETS = {
    # 示例格式："Calendaring Incoming Shortage": 0.5,
}

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
        "btn_export": "📥 导出全图表汇总报表 (Excel)",
        "changelog_title": "📌 版本更新说明 (What's New)",
        "changelog_text": """
        - **图表文案极简优化**：移除了图表注释中的 `μ-σ` 标识，统一直观显示为“目标线”。
        - **智能默认目标线**：单项不良默认目标线采用 **平均值 - $1 \times$ 标准差 ($\mu - 1\sigma$)** 动态计算。
        - **单位与坐标全面统一**：单项不良率推移采用百分比（%）单位，坐标轴与提示完美带上 `%`。
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
        "def_sel": "🎯 1. 选择追踪的不良项目：",
        "def_view": "📊 2. 选择观察粒度：",
        "view_daily": "按单日汇总趋势",
        "view_roll": "按单卷明细推移（分机台）",
        "def_metric": "📐 3. 选择统计指标：",
        "m_count": "不良数量推移",
        "m_rate": "不良率推移 (% 总投入)",
        "y_max_label": "🎚️ 单项不良 Y 轴上限：",
        "def_target_label": "🎯 单项不良目标线阈值设置：",
        "yield_trend_header": "📈 良率推移趋势分析",
        "y_min_label": "🎚️ 良率图 Y 轴下限 (%)：",
        "t_m_trend": "📈 1. 各机台单卷良率 (FPY) 推移走势",
        "t_d_trend": "📅 2. 单日平均良率变化推移走势",
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
        "import_help": "Drag and drop your MES production data table here",
        "import_info": "👈 Please drag or select your MES production data table in the sidebar.",
        "export_header": "📥 Report Export",
        "btn_export": "📥 Export All-Chart Summary Report (Excel)",
        "changelog_title": "📌 What's New",
        "changelog_text": """
        - **Streamlined UI Text**: Removed mathematical notations from chart annotations for cleaner display.
        - **Smart Default Target**: Single defect target line uses Mean - 1*Sigma dynamically.
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
        "def_sel": "🎯 1. Select Defect Item:",
        "def_view": "📊 2. Select View Granularity:",
        "view_daily": "Daily Summary Trend",
        "view_roll": "Roll-level Detail Trend (by Machine)",
        "def_metric": "📐 3. Select Metric:",
        "m_count": "Defect Quantity Trend",
        "m_rate": "Defect Rate Trend (% of Input)",
        "y_max_label": "🎚️ Single Defect Y-Axis Upper Limit:",
        "def_target_label": "🎯 Single Defect Target Threshold:",
        "yield_trend_header": "📈 Yield Trend Analysis",
        "y_min_label": "🎚️ Yield Chart Y-Axis Lower Limit (%):",
        "t_m_trend": "📈 1. Machine Roll Yield (FPY) Trend",
        "t_d_trend": "📅 2. Daily Average Yield Trend",
        "table_detail": "📋 Production Batch Detail Table",
        "x_time": "Production Time",
        "x_date": "Production Date",
        "y_roll_fpy": "Roll Yield FPY (%)",
        "y_daily_fpy": "Daily Avg Yield (%)",
    }
}

top_col1, top_col2 = st.columns([4, 1])
with top_col2:
    selected_lang_name = st.selectbox("🌐 Language / 语言", ["中文 (Chinese)", "English"], index=0)
    curr_lang = "zh" if "中文" in selected_lang_name else "en"

t = LANG_DICT[curr_lang]

st.title(t["title"])
st.markdown(f"*{t['author']}*")
st.markdown("---")

# 2. 侧边栏：文件导入区
st.sidebar.header(t["import_header"])
uploaded_file = st.sidebar.file_uploader(
    t["import_help"],
    type=["xlsx", "xls", "csv"],
    help=t["import_help"]
)

if uploaded_file is None:
    st.info(t["import_info"])
    st.markdown("---")
    st.markdown(f"### {t['changelog_title']}")
    st.markdown(t["changelog_text"])
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

df.columns = df.columns.str.strip()

required_cols = ["INQty", "OutQty", "ScrapQty"]
missing_cols = [c for c in required_cols if c not in df.columns]
if missing_cols:
    st.error(f"❌ Missing required columns: {', '.join(missing_cols)}")
    st.stop()

if "TestDate" in df.columns:
    df["TestDate"] = pd.to_datetime(df["TestDate"], errors="coerce")

# ================= 4. 核心数据逻辑验证与自动修正 =================
for col in ["INQty", "OutQty", "ScrapQty"]:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

has_fpy = "FPY" in df.columns
if has_fpy:
    df["FPY"] = pd.to_numeric(df["FPY"], errors="coerce")
    if df["FPY"].dropna().max() <= 1.5:
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

# ================= 5. 动态不良矩阵识别 =================
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

st.sidebar.markdown("---")
st.sidebar.markdown(t["changelog_text"])

# 7. 核心指标展示
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

# ================= 8. 不良分析专区 =================
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
            margin=dict(t=40, b=80, l=50, r=50)
        )
        st.plotly_chart(fig_pareto, use_container_width=True)

        ratio_mode = st.radio(t["radio_share"], [t["share_1"], t["share_2"]], horizontal=True)
        target_col = t["col_s_defect"] if t["share_1"] in ratio_mode else t["col_s_in"]
        color_scale = "Oranges" if t["share_1"] in ratio_mode else "Purples"
        
        share_sorted = defect_sums.sort_values(target_col, ascending=True)
        fig_share = px.bar(
            share_sorted, x=target_col, y=t["col_defect"], orientation="h",
            text=share_sorted[target_col].apply(lambda v: f"{v:.2f}%"),
            color=target_col, color_continuous_scale=color_scale
        )
        fig_share.update_traces(textposition="outside")
        fig_share.update_layout(xaxis=dict(range=[0, share_sorted[target_col].max() * 1.30]), coloraxis_showscale=False, height=400)
        st.plotly_chart(fig_share, use_container_width=True)

        stat_table = defect_sums[[t["col_defect"], t["col_count"], t["col_s_defect"], t["col_s_in"], t["col_cum"]]].copy()
        st.dataframe(stat_table.style.format({t["col_count"]: "{:,.1f}", t["col_s_defect"]: "{:.2f}%", t["col_s_in"]: "{:.2f}%", t["col_cum"]: "{:.2f}%"}), use_container_width=True, hide_index=True)

st.markdown("---")

# ================= 9. 动态单项不良变化推移图（去掉了图表内的 μ-σ 标注，统称为目标线） =================
st.subheader(t["defect_trend_header"])
selected_defect = None

if current_active_defects and "TestDate" in display_df.columns and len(display_df) > 0:
    ctrl_col1, ctrl_col2, ctrl_col3 = st.columns([1.1, 1.3, 1.1])
    with ctrl_col1:
        selected_defect = st.selectbox(t["def_sel"], current_active_defects, index=0)
    with ctrl_col2:
        defect_view_mode = st.radio(t["def_view"], [t["view_daily"], t["view_roll"]], horizontal=True)
    with ctrl_col3:
        defect_metric_mode = st.radio(t["def_metric"], [t["m_count"], t["m_rate"]], horizontal=True)

    defect_trend_df = display_df.dropna(subset=["TestDate"]).sort_values("TestDate").copy()
    is_rate_mode = t["m_rate"] in defect_metric_mode

    if not defect_trend_df.empty:
        if t["view_daily"] in defect_view_mode:
            # 颗粒度 1：按单日汇总趋势
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
            
            series_vals = daily_defect_agg[y_col]
            mean_val = float(series_vals.mean())
            std_val = float(series_vals.std())
            sigma_default = max(0.0, mean_val - 1.0 * (std_val if not np.isnan(std_val) else 0.0))
            if sigma_default == 0:
                sigma_default = max(0.01, mean_val * 0.5)
            
            auto_max = float(series_vals.max()) if not series_vals.empty else 1.0
            auto_max = max(sigma_default * 1.5, auto_max * 1.25)

            locked_val = LOCKED_DEFECT_TARGETS.get(selected_defect, None)
            
            adj_col1, adj_col2 = st.columns([1, 3])
            with adj_col1:
                if locked_val is not None:
                    st.info(f"🔒 Target Locked: {locked_val}%" if is_rate_mode else f"🔒 Target Locked: {locked_val}")
                    target_defect_val = locked_val
                else:
                    target_defect_val = st.number_input(
                        t["def_target_label"],
                        min_value=0.0,
                        max_value=float(max(100.0 if is_rate_mode else 10000.0, auto_max * 2.0)),
                        value=float(round(sigma_default, 2)),
                        step=0.01 if is_rate_mode else 1.0
                    )

            fig_def = px.line(daily_defect_agg, x="DayStr", y=y_col, markers=True)
            target_label_def = f"Target: {target_defect_val}%" if (curr_lang=="en" and is_rate_mode) else (f"目标线: {target_defect_val}%" if is_rate_mode else f"目标线: {target_defect_val}")
            fig_def.add_hline(
                y=target_defect_val,
                line_dash="dash",
                line_color="red",
                annotation_text=target_label_def,
                annotation_position="top left"
            )
            fig_def.update_layout(
                xaxis=dict(type="category", tickangle=-35),
                yaxis=dict(title=y_axis_title, range=[0, max(auto_max, target_defect_val * 1.2)]),
                height=430
            )
            st.plotly_chart(fig_def, use_container_width=True)

        else:
            # 颗粒度 2：按单卷明细推移（分机台）
            roll_df = defect_trend_df.copy()
            roll_df["DefectRatePercent"] = np.where(
                roll_df["INQty"] > 0, 
                (roll_df[selected_defect] / roll_df["INQty"]) * 100.0, 
                0.0
            )
            y_col = "DefectRatePercent" if is_rate_mode else selected_defect
            y_axis_title = f"{selected_defect} (%)" if is_rate_mode else f"{selected_defect} (Qty)"

            series_vals = roll_df[y_col]
            mean_val = float(series_vals.mean())
            std_val = float(series_vals.std())
            sigma_default = max(0.0, mean_val - 1.0 * (std_val if not np.isnan(std_val) else 0.0))
            if sigma_default == 0:
                sigma_default = max(0.01, mean_val * 0.5)

            auto_max = float(series_vals.max()) if not series_vals.empty else 1.0
            auto_max = max(sigma_default * 1.5, auto_max * 1.25)

            locked_val = LOCKED_DEFECT_TARGETS.get(selected_defect, None)
            
            adj_col1, adj_col2 = st.columns([1, 3])
            with adj_col1:
                if locked_val is not None:
                    st.info(f"🔒 Target Locked: {locked_val}%" if is_rate_mode else f"🔒 Target Locked: {locked_val}")
                    target_defect_val = locked_val
                else:
                    target_defect_val = st.number_input(
                        t["def_target_label"],
                        min_value=0.0,
                        max_value=float(max(100.0 if is_rate_mode else 10000.0, auto_max * 2.0)),
                        value=float(round(sigma_default, 2)),
                        step=0.01 if is_rate_mode else 1.0
                    )

            fig_def = px.line(
                roll_df,
                x="TestDate",
                y=y_col,
                color="MachineName" if "MachineName" in roll_df.columns else None,
                markers=True
            )
            target_label_def = f"Target: {target_defect_val}%" if (curr_lang=="en" and is_rate_mode) else (f"目标线: {target_defect_val}%" if is_rate_mode else f"目标线: {target_defect_val}")
            fig_def.add_hline(
                y=target_defect_val,
                line_dash="dash",
                line_color="red",
                annotation_text=target_label_def,
                annotation_position="top left"
            )
            fig_def.update_layout(
                xaxis=dict(title=t["x_time"]),
                yaxis=dict(title=y_axis_title, range=[0, max(auto_max, target_defect_val * 1.2)]),
                height=430
            )
            st.plotly_chart(fig_def, use_container_width=True)

st.markdown("---")

# ================= 10. 良率推移趋势专区 =================
st.subheader(t["yield_trend_header"])

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
        # 图表 1：各机台单卷良率推移走势
        st.markdown(f"**{t['t_m_trend']}**")
        fig_machine_trend = px.line(
            trend_raw_df,
            x="TestDate",
            y="FPY",
            color="MachineName" if "MachineName" in trend_raw_df.columns else None,
            markers=True
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

        # 图表 2：单日平均良率推移走势
        st.markdown(f"**{t['t_d_trend']}**")
        daily_df = trend_raw_df.copy()
        daily_df["DayStr"] = daily_df["TestDate"].dt.strftime("%Y-%m-%d")
        daily_agg = daily_df.groupby("DayStr", as_index=False)["FPY"].mean().sort_values("DayStr")
        
        global_daily_yield_agg = daily_agg.copy()

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

st.markdown("---")

# ================= 11. 生产批次明细数据表与 Excel 导出 =================
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

    if not global_daily_defect_agg.empty:
        global_daily_defect_agg.to_excel(writer, sheet_name="Defect Trend Analysis", index=False)

    if not global_daily_yield_agg.empty:
        global_daily_yield_agg.to_excel(writer, sheet_name="Yield Trend Analysis", index=False)

    display_df.to_excel(writer, sheet_name="Raw Data (Filtered)", index=False)

output_buffer.seek(0)
from openpyxl import load_workbook
wb = load_workbook(output_buffer)

if "Dashboard Summary" in wb.sheetnames and not global_defect_sums.empty:
    ws_sum = wb["Dashboard Summary"]
    start_r = len(summary_kpi_data) + 4
    max_r = start_r + len(global_defect_sums) - 1

    chart_bar = BarChart()
    chart_bar.type = "col"
    chart_bar.style = 10
    chart_bar.title = "Defect Quantity Pareto Chart"
    chart_bar.width = 16
    chart_bar.height = 10

    data_bar = Reference(ws_sum, min_col=2, min_row=start_r-1, max_row=max_r)
    cats_bar = Reference(ws_sum, min_col=1, min_row=start_r, max_row=max_r)
    chart_bar.add_data(data_bar, titles_from_data=True)
    chart_bar.set_categories(cats_bar)
    ws_sum.add_chart(chart_bar, "G2")

if "Defect Trend Analysis" in wb.sheetnames and not global_daily_defect_agg.empty:
    ws_def = wb["Defect Trend Analysis"]
    max_r_def = len(global_daily_defect_agg) + 1

    chart_col_def = BarChart()
    chart_col_def.type = "col"
    chart_col_def.style = 13
    chart_col_def.title = f"Single Defect Trend ({selected_defect or 'Defect'})"
    chart_col_def.y_axis.title = "Value"
    chart_col_def.x_axis.title = "Date"
    chart_col_def.width = 16
    chart_col_def.height = 10

    target_col_idx = 4 if is_rate_mode else 2
    data_def = Reference(ws_def, min_col=target_col_idx, min_row=1, max_row=max_r_def)
    cats_def = Reference(ws_def, min_col=1, min_row=2, max_row=max_r_def)
    
    chart_col_def.add_data(data_def, titles_from_data=True)
    chart_col_def.set_categories(cats_def)
    chart_col_def.legend = None
    ws_def.add_chart(chart_col_def, "F2")

if "Yield Trend Analysis" in wb.sheetnames and not global_daily_yield_agg.empty:
    ws_yld = wb["Yield Trend Analysis"]
    max_r_yld = len(global_daily_yield_agg) + 1

    chart_col_yld = BarChart()
    chart_col_yld.type = "col"
    chart_col_yld.style = 2
    chart_col_yld.title = "Daily Average Yield Trend"
    chart_col_yld.y_axis.title = "FPY (%)"
    chart_col_yld.x_axis.title = "Date"
    chart_col_yld.width = 16
    chart_col_yld.height = 10

    data_yld = Reference(ws_yld, min_col=2, min_row=1, max_row=max_r_yld)
    cats_yld = Reference(ws_yld, min_col=1, min_row=2, max_row=max_r_yld)
    
    chart_col_yld.add_data(data_yld, titles_from_data=True)
    chart_col_yld.set_categories(cats_yld)
    chart_col_yld.legend = None
    ws_yld.add_chart(chart_col_yld, "F2")

final_output = io.BytesIO()
wb.save(final_output)
excel_data = final_output.getvalue()

export_col1, export_col2 = st.columns([3, 1])
with export_col2:
    st.download_button(
        label=t["btn_export"],
        data=excel_data,
        file_name="Tenpower_Electrode_Full_Chart_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

format_dict = {"FPY": "{:.2f}%"} if has_fpy else {}
st.dataframe(display_df.style.format(format_dict, na_rep="-"), use_container_width=True)