"""
app.py - 机加工质量检测系统（主程序）v3.5
舍弗勒银川工厂
功能：数据录入（含Excel上传）、分析报告、月报
存储：Turso云数据库持久化
变更：
  v3.4: 弱化实时监控(删页面入口), 新增问题型号缺陷堆叠总览图, 月报加返修率/报废率趋势
  v3.5: 性能优化(load_all_records+图表函数加@st.cache_data缓存, 写操作后清缓存),
        KPI加返修率, 删柏拉图(与图2重复), 图4改Top5不良率缺陷, 堆叠图图例移下方防重叠
  v3.6: KPI合格率改反推口径(100-报废率-返修率), 配合数据修正保证三者恒=100%
  v3.7: 堆叠图默认时间范围改为"今年1月1日至今"(独立筛选), 图4下方新增超标缺陷类型汇总图
  v3.8: 全部图表统一"缺陷字段值口径"(图2/图4/堆叠图/图3主要缺陷), 同一缺陷数量处处一致
  v3.9: 录入页性能优化 — init_db进程内只跑一次(每次rerun省13次Turso网络往返),
        侧边栏统计改1小时缓存, 录入表格改轻量LIMIT查询, 聚合函数加缓存, 新增手动刷新报告按钮
  v3.10: 分析报告新增"月度不良趋势总览图"(堆叠图之后) — 柱=每月报废总数(左轴)+
        线=每月报废率(右轴, 报废/检测不含返修), 横轴固定1~12月; 柱悬停显示该月
        缺陷类型明细(降序), 线悬停显示该月各型号不良率(降序,前10); 复用堆叠图独立时间范围
  v3.11: 修复月度趋势图"单一缺陷数量>总报废"口径混用 — 柱悬停明细改报废口径
        (只统计报废记录; 超填记录单列'填报异常'; 未填部分单列'其他报废', 明细之和
        恒等于柱高), 折线悬停与解读百分比统一一位小数(与质量概览KPI一致)
  v3.12: 统计口径修正 — ①"合格率"改为"一次直通率"(合格数量/检测数量, 返修件
        不再计入合格, 质量概览/月报KPI/月报总结同步换标签); ②月度趋势图"填报异常"
        判定阈值从"缺陷合计>报废数"放宽为"缺陷合计>报废数+返修数"(操作工缺陷字段
        填的是报废+返修合计, 旧阈值会误标正常记录), 解读文字同步更新
  v3.13: 分析报告模块整改 — ①报告页分区: 区域一"全年总览"(月度不良趋势图+
        问题型号堆叠图, 固定今年1月1日至今, 不受筛选影响)在上, 区域二"近一个月
        分析"(型号不合格率/缺陷不良分析/Top5看板/缺陷模式月度趋势/超标缺陷汇总)
        在下, 时间筛选只作用于区域二; ②月度趋势图重构为"报废×主要缺陷堆叠柱+
        一次直通率折线": 每条报废记录的全部报废归入其数值最大的缺陷字段(并列取
        先出现, 全0归"未填缺陷类型"), 堆叠段之和恒等于该月报废总数, 折线悬停
        显示各型号直通率(最差在前); ③型号不合格率改为(检测-合格)/检测=1-一次
        直通率; ④质量概览下方新增预警规则说明; ⑤全部图表百分比统一1位小数
  v3.15: 数据录入页数量平衡校验 — ①四个数量字段下方实时显示"合格+返修+报废
        与检测数量的差值"(绿色通过/红色不通过); ②提交时差值≠0且未勾选强制保存
        → 阻止写入数据库(旧版只警告仍保存); ③新增"已知数量不一致,仍然保存"
        复选框(默认不勾), 勾选后可保存, 成功后橙色警告提醒核实; ④只管新增录入,
        不改历史数据/批量导入/表结构
  机加工版: 由GP12系统复制而来, 关键差异点 — ①系统名改为"机加工质量检测系统";
        ②工序锁定为"机加工"(页面不显示下拉框, 手动/导入均自动写入); ③录入表单
        删除旧示例文字, 初始值为空; ④批量导入改为单请求pipeline一次性写入+
        空值清洗(防空单元格/空格导致HTTP 400); ⑤导入兼容重复列名(日期.1/班次.1);
        ⑥云端已删记录提示"记录已不存在"而非"删除失败"
  缺陷类型调整(部署前): ①缺陷类型输入框改为数字输入框, 只允许非负整数(与GP12同步);
        ②"过车/欠车"拆分为"过车"+"欠车"两个独立字段(历史过车欠车列保留不迁移,
        移出统计); ③移除"滚道对称点过大"(与对称度超差重复, 历史列保留不统计);
        ④"幅高椭圆"界面显示改为"平行度超差"(数据库列名不变, 历史数据自动继承);
        ⑤录入页缺陷输入框key改为字段名(防列表增删后串值); ⑥导入时忽略旧"过车/欠车"列;
        ⑦共12类缺陷, 表单/列表/图表/月报/总报告明细统一使用同一常量
"""
import streamlit as st
import pandas as pd
import re
from datetime import datetime, timedelta
from export_to_template import fill_template
from data_utils import delete_record
from config import SCRAP_COUNT_WARNING, SCRAP_COUNT_CRITICAL
from data_utils import get_alerts_from_data

from config import (
    PROCESSES,
    SHIFTS,
    WARNING_THRESHOLD,
    CONSECUTIVE_DAYS,
    SYSTEM_NAME,
    FACTORY_NAME,
    VERSION,
    CHART_COLORS,
)
from data_utils import (
    init_db,
    save_record,
    save_records_batch,
    load_all_records,
    load_recent_records,
    clear_records_cache,
    calculate_scrap_rate,
    calculate_rework_rate,
    calculate_total_defect_rate,
    evaluate_warning,
    get_alerts_from_data,
    export_report_to_excel,
    DEFECT_TYPE_FIELDS,
    DEFECT_FIELDS_DISPLAY,
    aggregate_model_pass_rate,
    aggregate_defect_summary,
    aggregate_top5_alerts,
    aggregate_monthly_trend,
    get_monthly_overview,
    get_monthly_pass_rate_trend,
)
from spc_chart import (
    render_model_fail_rate_chart,
    render_defect_analysis_chart,
    render_top5_alert_chart,
    render_monthly_trend_chart,
    render_defect_stack_chart,
    render_overlimit_defect_chart,
    render_monthly_scrap_rate_trend_chart,
    render_monthly_defect_trend,
)

_PLOTLY_CONFIG = {"displayModeBar": False}

init_db()

st.set_page_config(
    page_title="{} - {}".format(FACTORY_NAME, SYSTEM_NAME),
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    .alert-banner {
        background: linear-gradient(90deg, #dc3545 0%, #c82333 100%);
        color: white;
        padding: 14px 24px;
        border-radius: 10px;
        font-size: 16px;
        font-weight: bold;
        margin-bottom: 16px;
        animation: pulse 2s infinite;
        box-shadow: 0 4px 12px rgba(220,53,69,0.3);
    }
    .alert-item {
        background-color: #fff3cd;
        border-left: 4px solid #ffc107;
        padding: 10px 16px;
        border-radius: 6px;
        font-size: 14px;
        margin-bottom: 8px;
    }
    .kpi-value {
        font-size: 28px;
        font-weight: bold;
        text-align: center;
    }
    .kpi-label {
        font-size: 13px;
        color: #888;
        text-align: center;
        margin-bottom: 4px;
    }
    .chart-insight {
        font-size: 13px;
        color: #666;
        margin-top: 4px;
        margin-bottom: 16px;
    }
    @keyframes pulse {
        0%, 100% { opacity: 1; }
        50% { opacity: 0.85; }
    }
</style>
""", unsafe_allow_html=True)


def inject_auto_refresh(interval_sec: int = 3600):
    """页面自动刷新(秒) — 默认1小时, 适配早会/周会场景, 弱化实时监控"""
    st.markdown(
        '<meta http-equiv="refresh" content="{}">'.format(interval_sec),
        unsafe_allow_html=True,
    )


def _render_kpi_row(items):
    cols = st.columns(len(items))
    for col, (label, value, color) in zip(cols, items):
        col.markdown(
            '<div class="kpi-label">{}</div>'
            '<div class="kpi-value" style="color:{};">{}</div>'.format(label, color, value),
            unsafe_allow_html=True,
        )


def _render_insight(text):
    st.markdown('<div class="chart-insight">{}</div>'.format(text), unsafe_allow_html=True)


def _fmt_pct(val, digits=2):
    """格式化百分比展示: 0或空显示'-', 否则加% (只改展示不改存储)"""
    try:
        num = float(val)
        if num == 0:
            return "-"
        return "{:.{}f}%".format(num, digits)
    except (ValueError, TypeError):
        return "-"


@st.cache_data(ttl=3600, show_spinner=False)
def _get_sidebar_stats():
    """★ v3.9 侧边栏统计(预警数/记录数/型号数/工序数) — 每小时最多重算一次
    - 之前每次rerun(员工在录入框每敲一下键盘)都全量拉数据+重算预警, 是录入页卡顿主因之一
    - 写操作(录入/删除/导入)不清此缓存: 员工录入不等统计刷新, 最长滞后1小时
    - 需要立即刷新时点侧边栏"手动刷新报告数据"按钮
    """
    df = load_all_records()
    alerts = get_alerts_from_data(df)
    return {
        "alert_count": len(alerts),
        "record_count": int(len(df)),
        "model_count": int(df["型号"].nunique()) if not df.empty else 0,
        "process_count": int(df["工序"].nunique()) if not df.empty else 0,
    }


# 6个标准缺陷字段 (带斜杠, 与load_all_records返回一致)
# ★ 展示层回填用字段: 与统一的12类显示名保持一致(过车/欠车已拆分)
_DEFECT_DISPLAY_FIELDS = list(DEFECT_FIELDS_DISPLAY)

# 关键词匹配规则 (顺序: 过车>欠车>磕伤>接刀痕>倒角>油沟>锻造, 多匹配归第一个)
# ★ "过车/欠车"已拆分为两个独立缺陷类型, 分别匹配(漏黑在下方单独判断)
_DEFECT_KEYWORD_RULES = [
    ("过车", ["过车"]),
    ("欠车", ["欠车"]),
    ("磕伤/卡伤", ["磕伤", "卡伤"]),
    ("接刀痕", ["接刀痕", "接刀"]),
    ("倒角", ["倒角"]),
    ("油沟", ["油沟"]),
    ("锻造缺陷", ["锻造", "锻件", "锻打"]),
    ("椭圆超差", ["椭圆"]),
    ("内径垫伤", ["垫伤"]),
    ("对称度超差", ["对称度"]),
    ("平行度超差", ["平行度"]),
]


def _backfill_defects_for_display(df):
    """展示层回填: 从报废原因文本提取缺陷数据, 填入6个结构化字段+漏黑列
    - 只改展示副本, 不改数据库
    - 漏黑独立成列, 不归到现有6个字段
    - 混合记录(漏黑+其他缺陷): 漏黑列和其他字段各填数量(有重复, 但保证数据不丢失)
    """
    df_out = df.copy()
    # 确保6个显示字段存在
    for f in _DEFECT_DISPLAY_FIELDS:
        if f not in df_out.columns:
            df_out[f] = ""
    # 新增漏黑列
    if "漏黑" not in df_out.columns:
        df_out["漏黑"] = ""

    for idx, row in df_out.iterrows():
        # 6字段全空才回填 (已有结构化数据的不动)
        all_empty = all(str(row.get(f, "")).strip() == "" for f in _DEFECT_DISPLAY_FIELDS)
        if not all_empty:
            continue
        # 报废数量>0才有意义
        try:
            scrap_num = int(row.get("报废数量", 0))
        except (ValueError, TypeError):
            scrap_num = 0
        if scrap_num <= 0:
            continue
        reason = str(row.get("报废原因", "")).strip()
        if not reason:
            continue

        # 漏黑独立判断
        has_leakblack = "漏黑" in reason

        # 6个标准缺陷关键词匹配 (归第一个)
        matched_field = None
        for field, keywords in _DEFECT_KEYWORD_RULES:
            for kw in keywords:
                if kw in reason:
                    matched_field = field
                    break
            if matched_field:
                break

        # 回填
        if has_leakblack:
            # 漏黑列填数量
            df_out.at[idx, "漏黑"] = str(scrap_num)
        if matched_field:
            # 匹配到的标准缺陷字段填数量
            df_out.at[idx, matched_field] = str(scrap_num)
        # 如果既没漏黑也没匹配标准缺陷, 不填 (椭圆超差等无法分类的)

    return df_out


def render_alert_banner(df=None):
    if df is None:
        df = load_all_records()
    alerts = get_alerts_from_data(df)

    if not alerts:
        st.success("✅ 当前所有型号质量状态正常，无活跃预警。")
        return

    st.markdown(
        '<div class="alert-banner">🚨 质量预警：当前有 {} 个型号/工序触发预警！</div>'.format(
            len(alerts)
        ),
        unsafe_allow_html=True,
    )
    for alert in alerts:
        if alert["等级"] == "批次不良":
            color, icon = "#dc3545", "🔴"
        else:
            color, icon = "#f0ad4e", "🟡"

        # ★ 新增: 查找该型号+工序最近一次有报废的数据日期
        sub = df[(df["型号"] == alert["型号"]) & (df["工序"] == alert["工序"])]
        scrap_sub = sub[sub["报废数量"] > 0] if not sub.empty else sub
        if not scrap_sub.empty:
            data_date = str(scrap_sub["日期"].max())
        elif not sub.empty:
            data_date = str(sub["日期"].max())
        else:
            data_date = "未知"

        st.markdown(
            '<div class="alert-item">'
            '{} <b>{}</b> | {} | 数据日期 <b>{}</b> | 当日报废 <b style="color:{};">{} 件</b> | {}'
            '</div>'.format(
                icon, alert["型号"], alert["工序"], data_date,
                color, alert["当日报废件数"], alert["原因"],
            ),
            unsafe_allow_html=True,
        )


def _dedupe_upload_columns(df):
    """★ 修复Excel导入重复列名问题(避坑清单: 日期/班次不能被读成"日期.1""班次.1")
    pandas读取到重复列名时会自动追加'.1'/'.2'后缀, 导致后续按"日期"匹配列失败。
    处理规则:
    - 去除每个列名首尾空格;
    - 形如"日期.1"视为"日期"的重复列, 与首个"日期"列按行合并(同行取首个非空值),
      并统一改名为"日期";
    只使用Excel原始列名, 不额外添加/复制列。
    """
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    groups = {}
    for c in df.columns:
        base = re.sub(r"\.\d+$", "", c)  # 去掉末尾 .数字 后缀
        groups.setdefault(base, []).append(c)
    frames, names = [], []
    for base, cols in groups.items():
        if len(cols) == 1:
            frames.append(df[cols[0]])
            names.append(base)
        else:
            # 多个同名(含.1/.2)列 -> 按行合并, 取首个非空值
            merged = df[cols[0]].copy()
            for c in cols[1:]:
                merged = merged.fillna(df[c])
            frames.append(merged)
            names.append(base)
    result = pd.concat(frames, axis=1)
    result.columns = names
    return result


def render_data_entry():
    st.header("📝 数据录入")
    st.caption("支持手动录入和Excel批量导入 | 数据保存至Turso云数据库 | 本页面不会自动刷新")

    with st.expander("📁 上传Excel文件（批量导入）", expanded=False):
        uploaded_file = st.file_uploader(
            "选择Excel文件（.xlsx / .xls）",
            type=["xlsx", "xls"],
            key="excel_uploader",
        )
        if uploaded_file is not None:
            try:
                # ★ 读取后第一件事记录原始列名(便于排查重复列), 再合并同名重复列
                df_upload = pd.read_excel(uploaded_file)
                print("Excel 原始列名:", list(df_upload.columns))
                df_upload = _dedupe_upload_columns(df_upload)
                df_upload = df_upload.fillna("")

                # ★ 旧表头"过车/欠车"已废弃(拆分为"过车"+"欠车"): 提示改用分列导入
                if ("过车/欠车" in df_upload.columns) or ("过车欠车" in df_upload.columns):
                    st.warning("⚠️ 检测到旧列「过车/欠车」：该列将被忽略（不再计入缺陷统计）。"
                               "请改用「过车」「欠车」两列分别填写数量。")

                st.write("**预览（前5行）：**")
                st.dataframe(df_upload.head(), use_container_width=True)

                if st.button("📥 确认导入", type="primary", key="btn_import"):
                    col_map = {
                        "日期": "日期", "date": "日期",
                        "班次": "班次", "shift": "班次",
                        "型号": "型号", "model": "型号",
                        "产品类型": "产品类型",
                        "订单号": "订单号",
                        "车削操作员": "车削操作者",
                        "操作员": "车削操作者",
                        "车削操作者": "车削操作者",
                        "GP12检测员": "GP12检测人员",
                        "GP12检测人员": "GP12检测人员",
                        "工序": "工序", "process": "工序",
                        "检测数量": "检测数量",
                        "合格数量": "合格数量",
                        "返修数量": "返修数量",
                        "报废数量": "报废数量",
                        "报废原因": "报废原因",
                        "返修原因": "返修原因",
                        "磕伤/卡伤": "磕伤卡伤",
                        "过车": "过车", "欠车": "欠车",
                        "平行度超差": "幅高椭圆", "幅高椭圆": "幅高椭圆",
                    }

                    df_upload = df_upload.rename(columns=col_map)

                    # ★ 旧表头"过车/欠车"已废弃(已拆分为"过车"+"欠车"): 忽略该列
                    #   (不参与入库与统计; 页面已在上方提示用户改用分列导入)
                    df_upload = df_upload.drop(columns=["过车/欠车", "过车欠车"], errors="ignore")

                    # ★ 工序锁定为"机加工": 无论Excel里工序列填什么, 统一写入"机加工"
                    #   (Excel没有工序列时此处会自动创建该列, 导入不再因缺工序列报错)
                    df_upload["工序"] = "机加工"

                    required_cols = ["日期", "班次", "型号", "工序", "检测数量", "合格数量", "返修数量", "报废数量"]
                    missing = [c for c in required_cols if c not in df_upload.columns]

                    if missing:
                        st.error("❌ Excel缺少必要列: {}".format(", ".join(missing)))
                    else:
                        # ★ GP12检测人员加入可选列: Excel里有该列就导入, 没有则补空, 不报错
                        optional_cols = ["产品类型", "订单号", "车削操作者", "GP12检测人员", "报废原因", "返修原因"] + DEFECT_TYPE_FIELDS
                        for col in optional_cols:
                            if col not in df_upload.columns:
                                df_upload[col] = ""

                        for qty_col in ["检测数量", "合格数量", "返修数量", "报废数量"]:
                            df_upload[qty_col] = pd.to_numeric(df_upload[qty_col], errors="coerce").fillna(0).astype(int)

                        df_upload["报废率"] = df_upload.apply(
                            lambda r: calculate_scrap_rate(int(r["报废数量"]), int(r["检测数量"])), axis=1)
                        df_upload["返修率"] = df_upload.apply(
                            lambda r: calculate_rework_rate(int(r["返修数量"]), int(r["检测数量"])), axis=1)
                        df_upload["总不合格率"] = df_upload.apply(
                            lambda r: calculate_total_defect_rate(r["报废率"], r["返修率"]), axis=1)

                        if "日期" in df_upload.columns:
                            df_upload["日期"] = pd.to_datetime(df_upload["日期"], errors="coerce").dt.strftime("%Y-%m-%d")

                        records = df_upload.to_dict("records")
                        count = save_records_batch(records)
                        st.success("✅ 成功导入 {} 条记录！".format(count))

            except Exception as e:
                st.error("❌ 读取Excel失败: {}".format(str(e)))

    st.markdown("---")

    st.subheader("✍️ 手动录入")

    # ★ 工序锁定为"机加工": 页面不再显示下拉框, 提交时自动写入"机加工"
    process = "机加工"

    c1, c2 = st.columns(2)
    with c1:
        entry_date = st.date_input("📅 日期", value=datetime.now(), key="entry_date")
    with c2:
        shift = st.selectbox("🕐 班次", SHIFTS, key="entry_shift")

    c4, c5, c6 = st.columns(3)
    with c4:
        # ★ 避坑清单: 录入表单不允许带旧默认示例文字, 初始值必须为空
        model = st.text_input("🔧 型号", key="entry_model")
    with c5:
        product_type = st.text_input("📦 产品类型", key="entry_ptype")
    with c6:
        order_no = st.text_input("📋 订单号", key="entry_order")

    c7, c8 = st.columns(2)
    with c7:
        operator = st.text_input("👷 车削操作者", placeholder="操作者姓名", key="entry_operator")
    with c8:
        # ★ GP12检测人员: 车削操作者旁边的文本框, 可填可空
        gp12_inspector = st.text_input("🔍 GP12检测人员", placeholder="检测人员姓名(可留空)", key="entry_gp12")

    st.markdown("---")

    st.subheader("检测数据")
    c9, c10, c11, c12 = st.columns(4)
    with c9:
        inspect_qty = st.number_input("🔍 检测数量", min_value=0, max_value=999999, value=0, step=1, key="entry_inspect")
    with c10:
        pass_qty = st.number_input("✅ 合格数量", min_value=0, max_value=999999, value=0, step=1, key="entry_pass")
    with c11:
        rework_qty = st.number_input("🔧 返修数量", min_value=0, max_value=999999, value=0, step=1, key="entry_rework")
    with c12:
        scrap_qty = st.number_input("❌ 报废数量", min_value=0, max_value=999999, value=0, step=1, key="entry_scrap")

    # ============================================================
    # ★ v3.15 数量平衡实时校验: 检测数量 = 合格 + 返修 + 报废
    # 填完四个数量后即时显示差值; 差值≠0时提供"强制保存"复选框(默认不勾)
    # ============================================================
    part_sum = pass_qty + rework_qty + scrap_qty          # 三者之和
    qty_diff = part_sum - inspect_qty                     # 差值 = 三者之和 - 检测数量
    qty_parts_empty = (pass_qty == 0 and rework_qty == 0 and scrap_qty == 0)
    force_save = False
    if not qty_parts_empty:
        # 三个分项数量还没填时不提示(只在提交时判断), 填了才实时校验
        if qty_diff == 0:
            # 差值为0: 绿色通过提示; 顺手清掉强制保存勾选状态(防下次误带)
            st.session_state.pop("entry_force_save", None)
            st.success("✅ 数量校验通过：合格+返修+报废 = 检测数量（{} = {}）".format(part_sum, inspect_qty))
        else:
            # 差值不为0: 红色醒目提示 + 强制保存复选框(特殊情况用, 默认不勾)
            st.error(
                "❌ 数量校验不通过：合格+返修+报废 = {}，检测数量 = {}，差值 = {}，请修正后再提交".format(
                    part_sum, inspect_qty, qty_diff))
            force_save = st.checkbox("已知数量不一致，仍然保存", value=False, key="entry_force_save")

    scrap_rate = calculate_scrap_rate(scrap_qty, inspect_qty)
    rework_rate = calculate_rework_rate(rework_qty, inspect_qty)
    total_defect_rate = calculate_total_defect_rate(scrap_rate, rework_rate)

    st.info(
        "📊 自动计算 → **报废率: {:.2f}%** | **返修率: {:.2f}%** | **总不合格率: {:.2f}%**".format(
            scrap_rate, rework_rate, total_defect_rate))

    all_pass = (inspect_qty > 0 and inspect_qty == pass_qty)

    if all_pass:
        st.success("✅ 检测数量 = 合格数量，判定为全合格，缺陷类型已自动留空")

    st.markdown("---")

    st.subheader("⚠️ 缺陷类型（填写数量）")
    st.caption("每个缺陷类型只填数字（非负整数，如 0、1、2...），无缺陷填 0 或留空；全合格时自动留空")

    defect_inputs = {}
    dcols = st.columns(3)
    # ★ 遍历统一的12类"界面显示名"(过车/欠车已拆分, 幅高椭圆显示为"平行度超差")
    for i in range(len(DEFECT_FIELDS_DISPLAY)):
        field_name = DEFECT_FIELDS_DISPLAY[i]
        col_idx = i % 3
        with dcols[col_idx]:
            # ★ 数字输入框: 只允许非负整数(字母/中文/符号/负数被输入框与min_value双重拦截)
            # ★ key用字段名而非序号: 缺陷列表增删后, 位置键会让在线会话的输入框串到别的字段的旧值
            defect_inputs[field_name] = st.number_input(
                field_name,
                min_value=0,
                max_value=999999,
                value=0,
                step=1,
                format="%d",
                key="entry_defect_{}".format(field_name),
                disabled=all_pass,
            )

    st.markdown("---")

    c15, c16 = st.columns(2)
    with c15:
        scrap_reason = st.text_input("❌ 报废原因（手动填写）", placeholder="请描述报废原因...", key="entry_scrap_reason")
    with c16:
        rework_reason = st.text_input("🔧 返修原因（手动填写）", placeholder="请输入返修原因...", key="entry_rework_reason")

    st.markdown("---")

    submitted = st.button("✅ 提交录入", type="primary", use_container_width=True)

    if submitted:
        if not model.strip():
            st.error("❌ 请输入型号！")
            return
        if not order_no.strip():
            st.error("❌ 请输入订单号！")
            return

        # ★ v3.15 提交拦截: 差值≠0 且未勾选"强制保存" → 阻止写入数据库
        #   (旧版只弹warning仍会保存, 现在不满足平衡绝不落库)
        if qty_diff != 0 and not force_save:
            st.error(
                "❌ 数量校验不通过，已阻止保存：合格({})+返修({})+报废({})={} ≠ 检测数量({})，差值 = {}。"
                "请修正数量后再提交；确属特殊情况请勾选\"已知数量不一致，仍然保存\"。".format(
                    pass_qty, rework_qty, scrap_qty, part_sum, inspect_qty, qty_diff))
            return
        new_record = {
            "日期": entry_date.strftime("%Y-%m-%d"),
            "班次": shift,
            "型号": model.strip(),
            "产品类型": product_type.strip(),
            "订单号": order_no.strip(),
            "车削操作者": operator.strip(),
            "GP12检测人员": gp12_inspector.strip(),  # ★ GP12检测人员(可留空, 空则存"")
            "工序": process,
            "检测数量": inspect_qty,
            "合格数量": pass_qty,
            "返修数量": rework_qty,
            "报废数量": scrap_qty,
            "报废率": scrap_rate,
            "返修率": rework_rate,
            "总不合格率": total_defect_rate,
            "报废原因": scrap_reason.strip(),
            "返修原因": rework_reason.strip(),
        }

        for field in DEFECT_FIELDS_DISPLAY:
            # ★ 缺陷类型存非负整数计数; 计数为0或全合格时存空串(与"空=无缺陷"统计口径一致)
            _cnt = int(defect_inputs.get(field, 0) or 0)
            new_record[field] = "" if (all_pass or _cnt <= 0) else str(_cnt)

        success = save_record(new_record)
        if success:
            # ★ v3.15: 强制保存的记录提交成功后给橙色警告, 提醒后续核实
            #   并复位复选框, 防止下一条记录误带"强制保存"状态
            st.session_state.pop("entry_force_save", None)
            if qty_diff != 0:
                st.warning(
                    "🟠 该条记录数量校验未通过（差值 = {}），已强制保存，请后续核实。".format(qty_diff))
            st.success(
                "✅ 录入成功！型号: {} | 工序: {} | 日期: {} | "
                "检测{}件 | 报废率{:.2f}% | 返修率{:.2f}% | 总不合格率{:.2f}%".format(
                    model, process, entry_date.strftime("%Y-%m-%d"),
                    inspect_qty, scrap_rate, rework_rate, total_defect_rate,
                )
            )
            if scrap_qty >= SCRAP_COUNT_CRITICAL:
                st.error(
                    "🚨 **批次不良预警！** {} - {}: 当日报废 {} 件（≥{}件），"
                    "判定为小批次报废，需立即停线排查！".format(
                        model, process, scrap_qty, SCRAP_COUNT_CRITICAL
                    )
                )
            elif scrap_qty >= SCRAP_COUNT_WARNING:
                st.warning(
                    "⚠️ **重点关注预警！** {} - {}: 当日报废 {} 件（≥{}件），"
                    "请加强巡检频次并排查原因。".format(
                        model, process, scrap_qty, SCRAP_COUNT_WARNING
                    )
                )
        else:
            st.error("❌ 保存失败，请检查数据库连接。")

    st.markdown("---")

    st.subheader("📂 已录入数据")
    if st.button("🔄 加载/刷新数据", use_container_width=True, key="btn_load"):
        if "pending_delete_id" in st.session_state:
            del st.session_state["pending_delete_id"]
        clear_records_cache()       # ★ 清数据缓存后再刷新, 从数据库拉最新数据
        _get_sidebar_stats.clear()  # ★ v3.9 顺带刷新侧边栏统计(用户主动点刷新, 允许重新拉取)
        st.rerun()

    # ★ v3.9: 录入页表格改用轻量查询(只拉最近50条+总数), 不再加载全量数据
    # 提交/删除后此缓存被清空, 表格立即显示最新记录; 网络载荷恒定, 数据再多也不变慢
    df_current, total_records = load_recent_records(50)
    if df_current.empty:
        st.info("暂无数据。")
    else:
        # ★ 避坑清单: 先按id排序(最新在前)再生成展示序号, 保证序号连续不跳号
        #   (底层id自增允许有空洞, 展示序号一律用range重新生成1,2,3...)
        display_df = df_current.copy().sort_values("id", ascending=False).reset_index(drop=True)
        display_df.insert(0, "序号", range(1, len(display_df) + 1))

        # ★ 百分比列格式化为带%字符串 (只改展示, 不影响导出)
        for _col in ["报废率", "返修率", "总不合格率"]:
            if _col in display_df.columns:
                display_df[_col] = display_df[_col].apply(lambda x: _fmt_pct(x))

        st.caption("共 {} 条记录（显示最近50条）".format(total_records))
        # ★ 隐藏 id 和 additional_defects(历史动态缺陷残留列, 数据库保留但前端不显示)
        show_cols = [c for c in display_df.columns if c not in ("id", "additional_defects")]
        st.dataframe(display_df[show_cols], use_container_width=True, hide_index=True)

        st.markdown("---")
        st.subheader("🗑️ 删除记录")
        st.caption("选择一条记录后点击下方删除按钮，需二次确认防止误操作")

        display_options = []
        id_map = {}
        for _, row in display_df.iterrows():
            label = "#{} | {} | {} | {} | 报废{}件".format(
                row["序号"], row["日期"], row["型号"], row["工序"], row["报废数量"]
            )
            display_options.append(label)
            id_map[label] = row["id"]

        selected = st.selectbox("选择要删除的记录", display_options, key="delete_select")

        if selected:
            selected_id = id_map[selected]

            if "pending_delete_id" not in st.session_state:
                st.session_state["pending_delete_id"] = None

            if st.session_state["pending_delete_id"] == selected_id:
                st.warning("⚠️ **确认删除此记录？此操作不可撤销！**\n\n`{}`".format(selected))
                cc1, cc2 = st.columns(2)
                with cc1:
                    if st.button("✅ 确认删除", type="primary", key="btn_confirm_del"):
                        status = delete_record(selected_id)
                        if status == "deleted":
                            st.success("✅ 记录已删除！")
                            del st.session_state["pending_delete_id"]
                            st.rerun()
                        elif status == "not_found":
                            # ★ 避坑清单: 云端已删的记录提示"记录已不存在", 不显示"删除失败"
                            st.warning("⚠️ 记录已不存在（可能已在云端删除），列表已同步更新。")
                            del st.session_state["pending_delete_id"]
                            st.rerun()
                        else:
                            st.error("❌ 删除失败，请重试。")
                with cc2:
                    if st.button("❌ 取消", key="btn_cancel_del"):
                        del st.session_state["pending_delete_id"]
                        st.rerun()
            else:
                if st.button("🗑️ 删除选中记录", type="secondary", key="btn_trigger_del"):
                    st.session_state["pending_delete_id"] = selected_id
                    st.rerun()


# ============================================================
# ★ v3.4: "实时监控"和"分级预警"独立页面已删除(弱化实时监控)
# - 分级预警的判定逻辑(get_alerts_from_data/evaluate_warning)完整保留
# - 月报"预警回顾"Tab仍调用这些底层函数
# - render_alert_banner预警横幅组件保留(底层组件, 不再被页面调用)
# ============================================================


def render_analysis_report():
    inject_auto_refresh(3600)

    st.header("🔍 分析报告")
    st.caption("⏱️ 本页面每1小时自动刷新 | 服务早会/周会质量回顾")

    df = load_all_records()
    if df.empty:
        st.info("暂无数据。")
        return

    # ============================================================
    # ★ v3.13 区域一: 全年总览 (顶部)
    # 固定今年1月1日至今, 不受下方"近一个月分析"时间筛选影响
    # ============================================================
    st.markdown("## 📊 全年总览")
    st.caption("时间范围固定为今年1月1日至今；下方日期筛选只作用于“近一个月分析”区域")

    # 全年数据切片(基于完整df, 与区域二的日期/型号筛选完全独立)
    df_year = df.copy()
    df_year["日期_dt"] = pd.to_datetime(df_year["日期"], errors="coerce")
    year_mask = (
        (df_year["日期_dt"] >= pd.Timestamp(datetime(datetime.now().year, 1, 1)))
        & (df_year["日期_dt"] <= pd.Timestamp(datetime.now()))
    )
    df_year = df_year[year_mask]

    # ---------- 区域一·图A: 月度不良趋势总览 (堆叠柱:报废×主要缺陷 + 折线:不良率) ----------
    st.markdown("### 📅 月度不良趋势总览")
    st.caption("柱=每月报废数量（按主要缺陷类型堆叠，各段之和=该月报废总数）；"
               "线=每月不良率（报废/检测×100）；柱悬停查看该月缺陷构成，线悬停查看各型号不良率（最高在前）")
    fig_trend, insight_trend = render_monthly_defect_trend(df_year, CHART_COLORS)
    st.plotly_chart(fig_trend, use_container_width=True, config=_PLOTLY_CONFIG)
    _render_insight(insight_trend)

    # ---------- 区域一·图B: 问题型号缺陷堆叠总览图 (不良率>5%) ----------
    st.markdown("### 🏆 问题型号缺陷堆叠总览图")
    st.caption("全年时间范围内不良率>5%的型号；柱高=报废+返修总数，柱内按缺陷类型堆叠")
    fig_stack, insight_stack = render_defect_stack_chart(df_year, CHART_COLORS)
    st.plotly_chart(fig_stack, use_container_width=True, config=_PLOTLY_CONFIG)
    _render_insight(insight_stack)

    st.markdown("---")

    # ============================================================
    # ★ v3.13 区域二: 近一个月分析 (下方)
    # 时间筛选(开始/结束日期+型号)只作用于本区域的图表、KPI和明细表
    # ============================================================
    st.markdown("## 🔍 近一个月分析")
    st.caption("时间筛选只作用于本区域；默认时间范围 = 今天往前一个月")

    # ---------- 时间筛选 + 型号筛选 ----------
    # ★ v3.4: 默认时间范围 = 今天往前一个月, 进入页面即显示, 无需手动选日期
    default_start = datetime.now() - timedelta(days=30)
    fcol1, fcol2, fcol3 = st.columns(3)
    with fcol1:
        start_date = st.date_input("开始日期", value=default_start, key="rpt_start")
    with fcol2:
        end_date = st.date_input("结束日期", value=datetime.now(), key="rpt_end")
    with fcol3:
        all_models = ["全部"] + sorted(df["型号"].dropna().unique().tolist())
        report_model = st.selectbox("型号", all_models, key="rpt_model")

    df_rpt = df.copy()
    df_rpt["日期_dt"] = pd.to_datetime(df_rpt["日期"], errors="coerce")
    mask = (
        (df_rpt["日期_dt"] >= pd.Timestamp(start_date))
        & (df_rpt["日期_dt"] <= pd.Timestamp(end_date))
    )
    df_rpt = df_rpt[mask]
    if report_model != "全部":
        df_rpt = df_rpt[df_rpt["型号"] == report_model]

    if df_rpt.empty:
        st.warning("所选条件下无数据（上方全年总览不受影响）。")
        return

    # ============================================================
    # KPI卡片行 (5个)
    # ============================================================
    total_inspect = int(df_rpt["检测数量"].sum())
    total_pass = int(df_rpt["合格数量"].sum())
    total_scrap = int(df_rpt["报废数量"].sum())
    total_rework = int(df_rpt["返修数量"].sum())
    # ★ v3.12 口径变更: 合格率→一次直通率 = 合格数量/检测数量×100%
    #   (v3.6曾用反推口径100-报废-返修吸收录入误差; 管理层需要真实直通率,
    #    返修件不再计入合格; 直通率+报废率+返修率不再强制=100%)
    scrap_rate = round(total_scrap / total_inspect * 100, 1) if total_inspect > 0 else 0.0
    rework_rate = round(total_rework / total_inspect * 100, 1) if total_inspect > 0 else 0.0
    pass_rate = round(min(total_pass / total_inspect * 100, 100.0), 1) if total_inspect > 0 else 0.0
    alert_count = len(get_alerts_from_data(df_rpt))

    pass_color = CHART_COLORS["success"] if pass_rate >= 95 else (CHART_COLORS["warning"] if pass_rate >= 90 else CHART_COLORS["danger"])
    scrap_color = CHART_COLORS["success"] if scrap_rate < 2 else (CHART_COLORS["warning"] if scrap_rate < 5 else CHART_COLORS["danger"])
    rework_color = CHART_COLORS["success"] if rework_rate < 3 else (CHART_COLORS["warning"] if rework_rate < 5 else CHART_COLORS["danger"])
    alert_color = CHART_COLORS["success"] if alert_count == 0 else (CHART_COLORS["warning"] if alert_count < 3 else CHART_COLORS["danger"])

    st.markdown("### 📊 质量概览")
    # ★ v3.5: 新增返修率KPI; v3.12 "合格率"改为"一次直通率"(合格/检测, 返修不计入)
    _render_kpi_row([
        ("总检测数", "{}件".format(total_inspect), CHART_COLORS["primary"]),
        ("一次直通率", "{:.1f}%".format(pass_rate), pass_color),
        ("报废率", "{:.1f}%".format(scrap_rate), scrap_color),
        ("返修率", "{:.1f}%".format(rework_rate), rework_color),
        ("预警数", "{}个".format(alert_count), alert_color),
    ])
    st.caption("💡 口径: 一次直通率=合格/检测（返修件不计入合格）, 报废率=报废/检测, 返修率=返修/检测")
    # ★ v3.13: 补充预警规则说明(橙色小字, 让管理层一眼看懂"预警数"怎么来的)
    st.markdown(
        '<span style="color:#ff7f0e;font-size:0.85em;">⚠️ 预警规则：单次报废数量 ≥ 5件，'
        '触发小批量报废警告预警；连续超标天数触发红色预警。</span>',
        unsafe_allow_html=True)
    st.markdown("---")

    # ============================================================
    # 第二行: 图1(型号不合格率问题看板) | 图2(缺陷不良分析)
    # ============================================================
    st.markdown("### 📈 型号不合格率 & 缺陷不良分析")
    col_left, col_right = st.columns(2)

    with col_left:
        fig1, insight1 = render_model_fail_rate_chart(df_rpt, CHART_COLORS)
        st.plotly_chart(fig1, use_container_width=True, config=_PLOTLY_CONFIG)
        _render_insight(insight1)

    with col_right:
        fig2, top3_df, insight2 = render_defect_analysis_chart(df_rpt, CHART_COLORS)
        st.plotly_chart(fig2, use_container_width=True, config=_PLOTLY_CONFIG)
        _render_insight(insight2)
        if not top3_df.empty:
            st.markdown("**Top3缺陷明细**")
            top3_display = top3_df.copy()
            if "占比" in top3_display.columns:
                top3_display["占比"] = top3_display["占比"].apply(lambda x: _fmt_pct(x, 1))
            st.dataframe(top3_display, use_container_width=True, hide_index=True)

    st.markdown("---")

    # ============================================================
    # ★★★ 第三行: 图3 近N天不良Top5预警看板 (v3.3.1 改为柱+折线组合)
    # ============================================================
    st.markdown("### 🚨 近N天不良Top5预警看板")
    st.caption("图1的下钻视角：快速筛选过去X天里报废达到Y件的型号及其完整表现")

    tc1, tc2 = st.columns(2)
    with tc1:
        # ★ 时间范围: 近2/3/5/7/30天
        top_days = st.selectbox(
            "时间范围",
            [2, 3, 5, 7, 30],
            format_func=lambda x: "近{}天".format(x),
            index=3,  # 默认选近7天
            key="top5_days"
        )
    with tc2:
        top_threshold = st.selectbox("报废件数阈值", [5, 10, 15, 20], key="top5_threshold")

    # ★ 图3用完整df (不依赖顶部日期筛选, 按自己的时间范围筛选)
    fig3, detail_df, insight3 = render_top5_alert_chart(df, top_days, top_threshold, CHART_COLORS)
    st.plotly_chart(fig3, use_container_width=True, config=_PLOTLY_CONFIG)

    if not detail_df.empty:
        st.markdown("**Top5问题型号清单**")
        detail_display = detail_df.copy()
        if "报废率" in detail_display.columns:
            detail_display["报废率"] = detail_display["报废率"].apply(lambda x: _fmt_pct(x, 1))
        st.dataframe(detail_display, use_container_width=True, hide_index=True)
    _render_insight(insight3)

    st.markdown("---")

    # ============================================================
    # ★★★ 第四行: 图4 型号缺陷模式月度趋势图 (v3.3.1 改为月份聚合+独立时间范围)
    # ============================================================
    st.markdown("### 📉 型号缺陷模式月度趋势")
    st.caption("按月份聚合，连续折线展示缺陷变化趋势，判断改善措施是否有效")

    mc1, mc2, mc3 = st.columns(3)
    with mc1:
        # ★ 独立时间范围下拉框: 近3/6/12个月, 默认近6个月
        trend_months = st.selectbox(
            "时间范围",
            [3, 6, 12],
            format_func=lambda x: "近{}个月".format(x),
            index=1,  # 默认选近6个月
            key="mt_months"
        )
    with mc2:
        trend_models = ["全部"] + sorted(df["型号"].dropna().unique().tolist())
        trend_model_sel = st.selectbox("型号选择", trend_models, key="mt_model")
    with mc3:
        # ★ 缺陷类型下拉直接用统一的12类显示名(不再硬编码, 避免与常量不一致)
        trend_defects = ["全部"] + list(DEFECT_FIELDS_DISPLAY)
        trend_defect_sel = st.selectbox("缺陷类型选择", trend_defects, key="mt_defect")

    # ★ 图4用完整df + 独立months参数 (不依赖顶部日期筛选)
    fig4, insight4 = render_monthly_trend_chart(df, trend_model_sel, trend_defect_sel, trend_months, CHART_COLORS)
    if trend_defect_sel == "全部":
        st.caption("💡 纵轴为不良率(%)，自动显示近{}个月整体不良率最高的Top5缺陷类型；点击图例可显示/隐藏缺陷".format(trend_months))
    st.plotly_chart(fig4, use_container_width=True, config=_PLOTLY_CONFIG)
    _render_insight(insight4)

    st.markdown("---")

    # ============================================================
    # ★ v3.7 新增: 超标缺陷类型汇总图（不良率>1%）— 放在图4下方
    # 时间范围跟随顶部"开始日期/结束日期"(即df_rpt), 不与图4的型号/缺陷类型筛选器联动
    # 柱状图=不良数量(报废+返修, 左轴), 折线图=不良率%(右轴)
    # ============================================================
    st.markdown("### ⚠️ 超标缺陷类型汇总")
    st.caption("时间范围跟随顶部开始/结束日期；不良率=(报废+返修)÷检测×100%，只统计不良率>1%的缺陷类型")
    fig_over, insight_over = render_overlimit_defect_chart(df_rpt, CHART_COLORS)
    st.plotly_chart(fig_over, use_container_width=True, config=_PLOTLY_CONFIG)
    _render_insight(insight_over)

    st.markdown("---")

    # ============================================================
    # ★ v3.5: 已删除"缺陷类型柏拉图"(与图2缺陷类型不良分析组合图功能重复)
    # 分析报告保留6张图: 堆叠总览图 / 型号不合格率 / 缺陷不良分析 / Top5预警看板 / 月度趋势 / 超标缺陷汇总
    # ============================================================

    # ============================================================
    # 保留原有: 总报告表格 (v3.3.2: 百分比格式化 + 列名对齐 + 展示层回填)
    # ============================================================
    st.subheader("📋 总报告明细")

    # ★ 任务B4: 展示层回填 - 从报废原因/返修原因文本提取缺陷数据
    df_display = _backfill_defects_for_display(df_rpt)

    # ★ 任务B3: 列名与数据录入页对齐 (使用带斜杠的显示名 + 漏黑列 + 5个新缺陷字段)
    display_cols = (
        ["日期", "班次", "型号", "产品类型", "订单号", "车削操作者",
         "工序", "检测数量", "合格数量", "返修数量", "报废数量",
         "报废率", "返修率", "总不合格率"]
        + list(DEFECT_FIELDS_DISPLAY)   # ★ 统一的12类缺陷(过车/欠车拆分, 平行度超差)
        + ["报废原因", "返修原因"]
    )

    existing = [c for c in display_cols if c in df_display.columns]

    # ★ 任务B2: 百分比列格式化为带%字符串 (只改展示, 不影响导出)
    report_display = df_display[existing].copy()
    for _col in ["报废率", "返修率", "总不合格率"]:
        if _col in report_display.columns:
            report_display[_col] = report_display[_col].apply(lambda x: _fmt_pct(x))

    st.dataframe(report_display, use_container_width=True, hide_index=True)

    st.subheader("📥 导出报告")
    if st.button("生成Excel报告", type="primary", use_container_width=True, key="btn_export"):
        excel_bytes = export_report_to_excel(df_rpt, str(start_date), str(end_date))
        st.download_button(
            label="📥 下载Excel",
            data=excel_bytes,
            file_name="质量报告_{}_{}.xlsx".format(start_date, end_date),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.success("✅ 报告已生成！")

    st.markdown("---")
    st.subheader("📥 导出到模板")
    st.caption("上传Excel模板，系统自动识别表头并将全部检测记录填入，保留原模板格式、公式、合并单元格及列宽")

    template_file = st.file_uploader("选择Excel模板文件（.xlsx）", type=["xlsx"], key="template_uploader")

    if template_file is not None:
        if st.button("📤 填充模板并下载", type="primary", use_container_width=True, key="btn_fill_template"):
            with st.spinner("正在读取数据并填充模板..."):
                result_bytes, status_msg = fill_template(template_file.read())

            if result_bytes:
                st.success(status_msg)
                st.download_button(
                    label="📥 下载 质量检测报告.xlsx",
                    data=result_bytes,
                    file_name="质量检测报告.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="btn_download_template",
                )
            else:
                st.error(status_msg)


def render_monthly_report():
    """月报模块: 4个Tab (当月概览/型号分析/缺陷分析/预警回顾)"""
    st.header("📅 月报")
    st.caption("按自然月汇总质量数据, 自动生成分析总结")

    df = load_all_records()
    if df.empty:
        st.info("暂无数据，请先录入检测记录。")
        return

    df["日期_dt"] = pd.to_datetime(df["日期"], errors="coerce")
    df = df.dropna(subset=["日期_dt"])
    df["月份"] = df["日期_dt"].dt.strftime("%Y-%m")
    available_months = sorted(df["月份"].unique().tolist(), reverse=True)
    if not available_months:
        st.info("无有效日期数据。")
        return

    mcol1, mcol2 = st.columns(2)
    with mcol1:
        selected_month = st.selectbox("选择月份", available_months, key="mr_month")
    with mcol2:
        st.caption("当前查看: {}年{}月".format(selected_month[:4], selected_month[5:7]))

    month_df = df[df["月份"] == selected_month].copy()
    if month_df.empty:
        st.warning("所选月份无数据。")
        return

    prev_month_idx = available_months.index(selected_month) + 1
    prev_month = available_months[prev_month_idx] if prev_month_idx < len(available_months) else None

    overview = get_monthly_overview(month_df, selected_month)
    prev_overview = get_monthly_overview(df[df["月份"] == prev_month], prev_month) if prev_month else None

    tab1, tab2, tab3, tab4 = st.tabs(["📋 当月概览", "🔧 型号分析", "⚠️ 缺陷分析", "🚨 预警回顾"])

    with tab1:
        st.subheader("当月KPI概览")
        pass_color = CHART_COLORS["success"] if overview["合格率"] >= 95 else (CHART_COLORS["warning"] if overview["合格率"] >= 90 else CHART_COLORS["danger"])
        scrap_color = CHART_COLORS["success"] if overview["报废率"] < 2 else (CHART_COLORS["warning"] if overview["报废率"] < 5 else CHART_COLORS["danger"])
        rework_color = CHART_COLORS["success"] if overview["返修率"] < 3 else (CHART_COLORS["warning"] if overview["返修率"] < 5 else CHART_COLORS["danger"])
        # ★ v3.4: 新增返修率KPI; v3.12 "合格率"改为"一次直通率"(合格/检测, 返修不计入)
        _render_kpi_row([
            ("总检验数", "{}件".format(overview["总检验数"]), CHART_COLORS["primary"]),
            ("一次直通率", "{:.1f}%".format(overview["合格率"]), pass_color),
            ("报废率", "{:.1f}%".format(overview["报废率"]), scrap_color),
            ("返修率", "{:.1f}%".format(overview["返修率"]), rework_color),
            ("预警次数", "{}次".format(overview["预警次数"]), CHART_COLORS["danger"] if overview["预警次数"] > 0 else CHART_COLORS["success"]),
        ])
        st.caption("💡 口径: 一次直通率=合格/检测（返修件不计入合格）, 报废率=报废/检测, 返修率=返修/检测")

        st.markdown("---")
        # ★ v3.4: 月度合格率趋势图 → 月度报废率趋势图(单线, 纵轴自动适配避免曲线太平)
        st.subheader("月度报废率趋势图")
        fig_trend = render_monthly_scrap_rate_trend_chart(df, CHART_COLORS)
        st.plotly_chart(fig_trend, use_container_width=True, config=_PLOTLY_CONFIG)

        if prev_overview and prev_overview["总检验数"] > 0:
            pass_diff = round(overview["合格率"] - prev_overview["合格率"], 1)
            arrow = "↑" if pass_diff > 0 else "↓"
            st.info("📊 一次直通率环比上月{}{}个百分点".format(arrow, abs(pass_diff)))
        else:
            st.info("📊 上月无数据，暂无法做环比分析")

    with tab2:
        # ★ v3.4: 各型号合格率分析 → 各型号不合格率分析(聚焦问题型号, 服务改善决策)
        st.subheader("各型号不合格率分析")
        model_rate = aggregate_model_pass_rate(month_df)
        if model_rate.empty:
            st.info("当月无型号数据。")
        else:
            fig_model, model_insight = render_model_fail_rate_chart(month_df, CHART_COLORS)
            st.plotly_chart(fig_model, use_container_width=True, config=_PLOTLY_CONFIG)
            _render_insight(model_insight)

            # ★ 明细表改为不合格率视角: 不合格率=(返修+报废)/检测
            detail = month_df.groupby("型号").agg(
                检测数量=("检测数量", "sum"),
                返修数量=("返修数量", "sum"),
                报废数量=("报废数量", "sum"),
            ).reset_index()
            detail["不合格数"] = detail["返修数量"] + detail["报废数量"]
            detail["不合格率(%)"] = detail.apply(
                lambda r: round(r["不合格数"] / r["检测数量"] * 100, 2) if r["检测数量"] > 0 else 0.0,
                axis=1)
            detail = detail.sort_values("不合格率(%)", ascending=False)
            st.markdown("**型号不合格率明细表（按严重程度降序）**")
            st.dataframe(detail, use_container_width=True, hide_index=True)

    with tab3:
        st.subheader("缺陷柏拉图分析")
        fig_defect, top3_defect, defect_insight = render_defect_analysis_chart(month_df, CHART_COLORS)
        st.plotly_chart(fig_defect, use_container_width=True, config=_PLOTLY_CONFIG)
        _render_insight(defect_insight)

        if not top3_defect.empty:
            st.markdown("**Top5缺陷明细表**")
            top5_defect = top3_defect.head(5)
            st.dataframe(top5_defect, use_container_width=True, hide_index=True)

    with tab4:
        st.subheader("当月预警回顾")
        month_alerts = get_alerts_from_data(month_df)
        active_alerts = get_alerts_from_data(df)

        red_count = sum(1 for a in month_alerts if a["等级"] == "批次不良")
        yellow_count = sum(1 for a in month_alerts if a["等级"] == "重点关注")

        _render_kpi_row([
            ("当月预警总数", "{}次".format(len(month_alerts)), CHART_COLORS["primary"]),
            ("批次不良(红)", "{}次".format(red_count), CHART_COLORS["danger"]),
            ("重点关注(黄)", "{}次".format(yellow_count), CHART_COLORS["warning"]),
            ("仍活跃", "{}次".format(len(active_alerts)), CHART_COLORS["danger"] if active_alerts else CHART_COLORS["success"]),
        ])

        st.markdown("---")
        if month_alerts:
            for alert in month_alerts:
                if alert["等级"] == "批次不良":
                    st.error("🔴 {} | {} | 报废{}件 | {}".format(alert["型号"], alert["工序"], alert["当日报废件数"], alert["原因"]))
                else:
                    st.warning("🟡 {} | {} | 报废{}件 | {}".format(alert["型号"], alert["工序"], alert["当日报废件数"], alert["原因"]))
        else:
            st.success("✅ 当月无预警触发")

    st.markdown("---")
    st.subheader("📝 月报总结")

    if prev_overview and prev_overview["总检验数"] > 0:
        pass_diff = round(overview["合格率"] - prev_overview["合格率"], 1)
        pass_arrow = "↑" if pass_diff > 0 else "↓"
        pass_change = "{}{}个百分点".format(pass_arrow, abs(pass_diff))
    else:
        pass_change = "上月无数据"

    defect_summary = aggregate_defect_summary(month_df)
    main_defect = defect_summary.iloc[0]["缺陷类型"] if not defect_summary.empty else "无"
    main_defect_pct = defect_summary.iloc[0]["占比"] if not defect_summary.empty else 0

    model_rate_df = aggregate_model_pass_rate(month_df)
    if not model_rate_df.empty:
        worst_model = model_rate_df.nsmallest(1, "合格率").iloc[0]["型号"]
    else:
        worst_model = "无"

    if len(month_alerts) > 0:
        process_rate = round((1 - len(active_alerts) / max(len(month_alerts), 1)) * 100, 1)
        process_rate = max(0, min(100, process_rate))
    else:
        process_rate = 100.0

    # ★v3.12: 总结中"合格率"改为"一次直通率"(合格/检测, 返修不计入)
    summary = "本月共检验**{}件**，一次直通率**{:.1f}%**，环比上月**{}**。不良数**{}件**，主要集中在**{}**型号的**{}**缺陷类型(占比{:.1f}%)。本月触发预警**{}次**，处理率**{:.1f}%**。建议下月重点关注：**{}**型号的**{}**缺陷类型。".format(
        overview["总检验数"],
        overview["合格率"],
        pass_change,
        overview["总报废数"],
        worst_model,
        main_defect,
        main_defect_pct,
        len(month_alerts),
        process_rate,
        worst_model,
        main_defect,
    )
    st.info(summary)


def render_sidebar():
    with st.sidebar:
        st.markdown("## 🏭 {}".format(FACTORY_NAME))
        st.markdown("### {}".format(SYSTEM_NAME))
        st.markdown("---")

        page = st.radio(
            "📌 功能导航",
            # ★ v3.4: 删除"实时监控"和"分级预警"入口(弱化实时监控, 强化分析报告+月报)
            ["📝 数据录入", "🔍 分析报告", "📅 月报"],
            key="nav_page",
        )

        st.markdown("---")

        # ★ v3.9: 侧边栏统计改用缓存(每小时最多重算一次), 不再每次rerun都拉全量数据
        stats = _get_sidebar_stats()

        if stats["alert_count"]:
            st.error("🚨 {} 个活跃预警".format(stats["alert_count"]))
        else:
            st.success("✅ 系统正常")

        st.markdown("---")
        st.markdown("**数据概览**")
        st.caption("检测记录: {} 条".format(stats["record_count"]))
        if stats["record_count"]:
            st.caption("覆盖型号: {} 个".format(stats["model_count"]))
            st.caption("覆盖工序: {} 个".format(stats["process_count"]))

        st.markdown("---")
        st.markdown("**预警阈值**")
        st.caption("报废率 > {}% 且连续{}天".format(WARNING_THRESHOLD, CONSECUTIVE_DAYS))

        st.markdown("---")
        # ★ v3.9: 手动刷新按钮 — 经理需要最新报告数据时点击, 清空全部缓存并重新拉取
        if st.button("🔄 手动刷新报告数据", use_container_width=True, key="btn_refresh_reports"):
            _get_sidebar_stats.clear()  # 侧边栏统计缓存
            clear_records_cache()       # 全量+最近N条数据缓存
            st.rerun()
        st.caption("报告数据最长缓存1小时，录入不受影响")

        st.markdown("---")
        st.caption("⏱️ 分析报告：每1小时自动刷新")
        st.caption("📝 数据录入：完全手动，不自动刷新")
        st.markdown("---")
        st.caption("{} | {}".format(SYSTEM_NAME, VERSION))
        st.caption("数据存储: Turso云数据库")

    return page


def _require_admin_password(secret_key):
    """★ 轻量级页面访问权限: 进入"分析报告"/"月报"前必须输入管理密码
    - 密码只从 .streamlit/secrets.toml 读取(键名由 secret_key 指定), 代码里不写死任何密码
    - 用 st.session_state 记录本会话是否已验证: 同一会话内只需输入一次,
      分析报告与月报共用同一个标记, 两个页面之间切换不再重复要求密码
    - 校验未通过时用 st.stop() 截断渲染, 保证报告内容不会被看到
    """
    # 本会话已通过验证 → 直接放行(不再要求输入密码)
    if st.session_state.get("admin_authenticated", False):
        return

    st.subheader("🔒 需要管理密码")
    st.caption("分析报告 / 月报 仅供管理人员查看，请输入管理密码后进入")
    pwd = st.text_input("管理密码", type="password", key="admin_password_input")

    if st.button("🔓 确认进入", type="primary", key="admin_password_submit"):
        correct_pwd = str(st.secrets.get(secret_key, "")).strip()
        if not correct_pwd:
            # 管理员尚未在 secrets.toml 里配置该密码
            st.error("系统未配置管理密码（请在 .streamlit/secrets.toml 中配置 {}）".format(secret_key))
        elif pwd == correct_pwd:
            st.session_state["admin_authenticated"] = True  # ★ 记录本会话已验证
            st.rerun()  # 重新渲染 → 直接进入报告页面
        else:
            st.error("密码错误，无权访问")

    st.stop()  # ★ 未通过验证: 停止渲染后续内容, 报告页面不会显示


def main():
    page = render_sidebar()

    st.markdown(
        '<h1 style="color:#1a1a2e;">🏭 {}</h1>'.format(SYSTEM_NAME),
        unsafe_allow_html=True,
    )
    st.caption("{} | 车间质量检测一体化平台 | 数据持久化存储".format(FACTORY_NAME))

    if page == "📝 数据录入":
        # ★ 数据录入不设限: 所有打开系统的人都可直接使用
        render_data_entry()
    elif page == "🔍 分析报告":
        _require_admin_password("MACHINING_ADMIN_PASSWORD")  # ★ 进入前校验管理密码
        render_analysis_report()
    elif page == "📅 月报":
        _require_admin_password("MACHINING_ADMIN_PASSWORD")  # ★ 进入前校验管理密码
        render_monthly_report()


if __name__ == "__main__":
    main()
