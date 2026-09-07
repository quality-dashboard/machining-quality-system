import json as _json
import urllib.request
import pandas as pd
import streamlit as st
from typing import List, Dict, Any, Optional

from config import (
    WARNING_THRESHOLD,
    CONSECUTIVE_DAYS,
    SCRAP_COUNT_WARNING,
    SCRAP_COUNT_CRITICAL,
)

# ============================================================
# ★★★ Turso 数据库连接配置区（机加工质量检测系统专用，独立新库） ★★★
# ★ 使用前必须完成两步:
#   1. 在 Turso 新建一个全新数据库(禁止复用GP12的库), 得到新库 URL 和 Token
#   2. 把下方 TURSO_URL 换成新库URL; 新库Token填到 .streamlit/secrets.toml
# ★ 防呆: 未配置新库、或仍指向GP12旧库时, 系统启动直接报错提醒, 绝不误写旧库
# ============================================================
TURSO_URL = "libsql://machining-quality-system-quality-dashboard.aws-ap-northeast-1.turso.io" 
TURSO_AUTH_TOKEN = st.secrets["TURSO_TOKEN"]  # ← 新库Token, 纯英文, 存放于 .streamlit/secrets.toml
# ============================================================

# GP12旧库地址特征(仅用于防呆校验, 确保新系统绝不连到旧库)
_GP12_OLD_DB_KEYWORD = "qualitytracking-quality-dashboard"


def _validate_db_config():
    """★ 防呆校验: 确保已配置全新的机加工专用数据库, 防止误连/误写GP12旧库"""
    # 1) URL必须已填写(占位符/空值都算未配置)
    if not TURSO_URL or "xxxx" in TURSO_URL or "your" in TURSO_URL.lower():
        raise RuntimeError(
            "【数据库未配置】请先在 data_utils.py 顶部填写新建 Turso 数据库的 TURSO_URL, "
            "并在 .streamlit/secrets.toml 中填写对应的 TURSO_TOKEN。")
    # 2) 绝不允许连接GP12旧库
    if _GP12_OLD_DB_KEYWORD in TURSO_URL:
        raise RuntimeError(
            "【禁止连接GP12旧库】机加工质量检测系统必须使用全新独立的 Turso 数据库, "
            "请修改 data_utils.py 顶部的 TURSO_URL。")
    # 3) Token必须是真实值(占位符视为未配置)
    if not TURSO_AUTH_TOKEN or "PASTE" in TURSO_AUTH_TOKEN.upper():
        raise RuntimeError(
            "【Token未配置】请在 .streamlit/secrets.toml 中填写新建 Turso 数据库的 TURSO_TOKEN "
            "(Token必须为纯英文)。")

BASE_COLUMNS = [
    "日期", "班次", "型号", "产品类型", "订单号", "车削操作者", "GP12检测人员", "工序",
    "检测数量", "合格数量", "返修数量", "报废数量",
    "报废率", "返修率", "总不合格率",
    "报废原因", "返修原因",
]

# 注意: 数据库列名不带斜杠, 避免 Turso SQL 解析器报 400
# 7个原字段 + 5个新增字段(椭圆超差/内径垫伤/滚道对称点过大/对称度超差/幅高椭圆)
DEFECT_COLUMNS = ["锻造缺陷", "磕伤卡伤", "接刀痕", "过车欠车", "油沟", "倒角", "漏黑",
                  "椭圆超差", "内径垫伤", "滚道对称点过大", "对称度超差", "幅高椭圆"]

ALL_COLUMNS = BASE_COLUMNS + DEFECT_COLUMNS

# 界面显示名 <-> 数据库列名 的映射 (新5字段无斜杠, 显示名=数据库名, 无需映射)
_DISPLAY_TO_DB = {"磕伤/卡伤": "磕伤卡伤", "过车/欠车": "过车欠车"}
_DB_TO_DISPLAY = {"磕伤卡伤": "磕伤/卡伤", "过车欠车": "过车/欠车"}

# 界面显示名 (load_all_records 返回的列名, 带斜杠的带斜杠, 新5字段不带斜杠)
DEFECT_FIELDS_DISPLAY = ["锻造缺陷", "磕伤/卡伤", "接刀痕", "过车/欠车", "油沟", "倒角", "漏黑",
                         "椭圆超差", "内径垫伤", "滚道对称点过大", "对称度超差", "幅高椭圆"]


class _Result:
    def __init__(self, columns, rows, affected_row_count):
        self.columns = columns
        self.rows = rows
        self.affected_row_count = affected_row_count


def _get_http_url():
    url = TURSO_URL
    if url.startswith("libsql://"):
        url = "https://" + url[len("libsql://"):]
    if url.endswith("/"):
        url = url[:-1]
    return url + "/v2/pipeline"


def _to_turso_arg(value):
    if value is None:
        return {"type": "null"}
    # bool 必须在 int 之前判断 (Python 中 True 也是 int)
    if isinstance(value, bool):
        return {"type": "integer", "value": "1" if value else "0"}  # 字符串
    if isinstance(value, int):
        return {"type": "integer", "value": str(value)}  # 字符串
    if isinstance(value, float):
        return {"type": "float", "value": value}  # 数字
    return {"type": "text", "value": str(value)}  # 字符串

def _from_turso_value(val):
    if val is None:
        return None
    t = val.get("type")
    if t == "null":
        return None
    if t == "integer":
        return int(val.get("value", 0))
    if t == "float":
        return float(val.get("value", 0.0))
    if t == "text":
        return val.get("value", "")
    return val.get("value")


def _parse_result(result_json):
    results = result_json.get("results", [])
    if not results:
        return _Result([], [], 0)
    first = results[0]
    if first.get("type") == "error":
        err = first.get("error", {})
        msg = err.get("message", "未知错误")
        raise RuntimeError("Turso错误: " + msg)
    response = first.get("response", {})
    if response.get("type") != "execute":
        return _Result([], [], 0)
    res = response.get("result", {})
    cols = [c.get("name", "") for c in res.get("cols", [])]
    raw_rows = res.get("rows", [])
    rows = [[_from_turso_value(v) for v in r] for r in raw_rows]
    affected = res.get("affected_row_count", 0)
    return _Result(cols, rows, affected)


def _execute(sql, args=None):
    turso_args = [_to_turso_arg(a) for a in args] if args else []
    body = {
        "requests": [
            {"type": "execute", "stmt": {"sql": sql, "args": turso_args}},
            {"type": "close"},
        ]
    }
    data = _json.dumps(body).encode("utf-8")
    req = urllib.request.Request(
        _get_http_url(),
        data=data,
        headers={
            "Authorization": "Bearer " + TURSO_AUTH_TOKEN,
            "Content-Type": "application/json",
        },
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        result = _json.loads(resp.read().decode("utf-8"))
    return _parse_result(result)


# ★ v3.9: init_db进程内只执行一次的标记
# app.py顶层每次rerun都会调用init_db(), 而它每次要跑13次Turso网络往返
# (1次CREATE TABLE + 12次缺陷列检查SELECT), 是录入页每次交互都变灰的元凶
_DB_INIT_DONE = False


def init_db():
    """初始化数据库表结构
    ★ v3.9: 进程内只真正执行一次(建表+补列是幂等操作, 无需每次rerun都跑)
    - 模块热重载(改代码)时Streamlit会重新导入本模块, 标记自动复位, 仍会执行一次
    - 首次执行失败(如网络断)时标记不置位, 下次rerun自动重试
    """
    global _DB_INIT_DONE
    if _DB_INIT_DONE:
        return
    _validate_db_config()  # ★ 防呆: 未配置新库/误指旧库时直接报错, 不连库
    cols_sql = ",\n".join(['"{}" TEXT DEFAULT ""'.format(c) for c in DEFECT_COLUMNS])
    create_sql = """
        CREATE TABLE IF NOT EXISTS quality_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            日期 TEXT NOT NULL,
            班次 TEXT NOT NULL,
            型号 TEXT NOT NULL,
            产品类型 TEXT DEFAULT "",
            订单号 TEXT DEFAULT "",
            车削操作者 TEXT DEFAULT "",
            GP12检测人员 TEXT DEFAULT "",
            工序 TEXT NOT NULL,
            检测数量 INTEGER DEFAULT 0,
            合格数量 INTEGER DEFAULT 0,
            返修数量 INTEGER DEFAULT 0,
            报废数量 INTEGER DEFAULT 0,
            报废率 REAL DEFAULT 0.0,
            返修率 REAL DEFAULT 0.0,
            总不合格率 REAL DEFAULT 0.0,
            报废原因 TEXT DEFAULT "",
            返修原因 TEXT DEFAULT "",
            {}
        )
    """.format(cols_sql)
    _execute(create_sql)
    # 幂等补充缺失的缺陷列 (兼容已有数据库: 新增字段时自动ALTER TABLE)
    _ensure_defect_columns()
    _DB_INIT_DONE = True  # ★ 成功后才置位(失败下次rerun重试)


def _ensure_defect_columns():
    """检查DEFECT_COLUMNS里的列是否都存在, 缺失的ALTER TABLE补充"""
    # ★ GP12检测人员列也纳入自动补列检查(老库升级时自动加列, 幂等安全)
    for col in DEFECT_COLUMNS + ["GP12检测人员"]:
        try:
            _execute('SELECT "{}" FROM quality_records LIMIT 1'.format(col))
        except Exception:
            try:
                _execute('ALTER TABLE quality_records ADD COLUMN "{}" TEXT DEFAULT ""'.format(col))
                print("[表结构] 新增列: {}".format(col))
            except Exception as e:
                print("[表结构] ALTER TABLE失败 {}: {}".format(col, e))


# ★ 数值/缺陷字段空值处理: Excel导入时空单元格/空格会产生NaN,
#   NaN 经 json 序列化会变成非法JSON, Turso 直接返回 HTTP 400 Bad Request
_NUMERIC_INT_COLS = ["检测数量", "合格数量", "返修数量", "报废数量"]
_NUMERIC_REAL_COLS = ["报废率", "返修率", "总不合格率"]


def _sanitize_record_value(col, v):
    """清洗单个字段值为 Turso 可接受的入库值
    - 缺陷类型字段(TEXT): 空/空格/NaN -> ""; 其余转字符串
    - 整数字段: pd.to_numeric 强制转换, 失败或空 -> 0
    - 实数(率)字段: pd.to_numeric 强制转换, 失败或空 -> 0.0
    - 其余文本字段: 空/NaN -> ""
    核心目的: 绝不让 float('nan') 进入请求体(会导致 HTTP 400 Bad Request)
    """
    if v is None:
        v = ""
    if isinstance(v, float) and pd.isna(v):
        v = ""
    if col in DEFECT_COLUMNS:
        return "" if v == "" else str(v).strip()
    if col in _NUMERIC_INT_COLS:
        num = pd.to_numeric(v, errors="coerce")
        return 0 if pd.isna(num) else int(num)
    if col in _NUMERIC_REAL_COLS:
        num = pd.to_numeric(v, errors="coerce")
        return 0.0 if pd.isna(num) else float(num)
    # 其余文本字段(日期/班次/型号/产品类型/订单号/车削操作者/GP12检测人员/工序/报废原因/返修原因)
    return "" if v == "" else str(v)


def _execute_pipeline(stmts):
    """★ 单请求pipeline批量执行多条SQL(修复"逐条写入只进第一条"的坑)
    - stmts: [(sql, args_list), ...]
    - 全部语句打包成一次HTTP请求, 避免逐条请求触发Turso限流
    - 任一语句失败立即抛异常(保证不出现"导了一半"的脏数据)
    """
    if not stmts:
        return 0
    requests = []
    for sql, args in stmts:
        turso_args = [_to_turso_arg(a) for a in args]
        requests.append({"type": "execute", "stmt": {"sql": sql, "args": turso_args}})
    requests.append({"type": "close"})
    body = {"requests": requests}
    data = _json.dumps(body).encode("utf-8")
    req = urllib.request.Request(
        _get_http_url(),
        data=data,
        headers={
            "Authorization": "Bearer " + TURSO_AUTH_TOKEN,
            "Content-Type": "application/json",
        },
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=60) as resp:
        result = _json.loads(resp.read().decode("utf-8"))
    ok_count = 0
    for r in result.get("results", []):
        if r.get("type") == "error":
            err = r.get("error", {})
            raise RuntimeError("Turso批量写入错误: " + err.get("message", "未知错误"))
        res = r.get("response", {}).get("result", {})
        ok_count += int(res.get("affected_row_count", 1))
    return ok_count


def save_record(record):
    """保存单条检测记录, 自动把界面字段名转为数据库列名"""
    try:
        clean_record = {}
        for k, v in record.items():
            k = _DISPLAY_TO_DB.get(k, k)
            clean_record[k] = v

        values = [_sanitize_record_value(c, clean_record.get(c, "")) for c in ALL_COLUMNS]
        col_names = ", ".join(['"{}"'.format(c) for c in ALL_COLUMNS])
        placeholders = ", ".join(["?"] * len(ALL_COLUMNS))
        sql = "INSERT INTO quality_records ({}) VALUES ({})".format(col_names, placeholders)

        _execute(sql, values)
        _clear_records_cache()  # ★ 写操作后清缓存, 确保下次读取拿到最新数据
        return True
    except Exception as e:
        print("save_record error:", e)
        return False


def save_records_batch(records):
    """★ 批量保存记录(Excel导入): 单请求pipeline一次性写入, 成功后清一次缓存
    - 修复旧版逐条请求导致的Turso限流/"只导入第一条"问题
    - 每条记录同样经过 _sanitize_record_value 空值清洗, 避免HTTP 400
    """
    if not records:
        return 0
    stmts = []
    for r in records:
        clean = {}
        for k, v in r.items():
            clean[_DISPLAY_TO_DB.get(k, k)] = v
        values = [_sanitize_record_value(c, clean.get(c, "")) for c in ALL_COLUMNS]
        col_names = ", ".join(['"{}"'.format(c) for c in ALL_COLUMNS])
        placeholders = ", ".join(["?"] * len(ALL_COLUMNS))
        sql = "INSERT INTO quality_records ({}) VALUES ({})".format(col_names, placeholders)
        stmts.append((sql, values))
    try:
        count = _execute_pipeline(stmts)
        _clear_records_cache()  # ★ 批量写入后只清一次缓存, 录入表格立即刷新
        return count
    except Exception as e:
        print("save_records_batch error:", e)
        return 0


@st.cache_data(ttl=3600, show_spinner=False)
def _load_all_records_cached():
    """
    ★ 性能优化: 真实查询Turso的缓存层 (只缓存成功结果, 查询失败抛异常不会被缓存)
    - v3.9 ttl从300秒放宽到3600秒: 全量数据只给分析报告/月报用, 报告类页面
      接受最长1小时滞后(与分析报告页自动刷新周期对齐); 写操作仍会主动清缓存,
      确保下次打开报告页时拉到最新数据
    """
    result = _execute("SELECT * FROM quality_records ORDER BY id DESC")
    df = pd.DataFrame(result.rows, columns=result.columns)
    df = df.rename(columns=_DB_TO_DISPLAY)
    # ★ 防御: 历史动态缺陷功能残留的 additional_defects 列不返回前端(数据库字段保留, 仅界面不显示)
    df = df.drop(columns=["additional_defects"], errors="ignore")
    return df


def load_all_records():
    """加载全部记录为DataFrame(带缓存), 自动把数据库列名转回界面显示名; 失败返回空表"""
    try:
        return _load_all_records_cached()
    except Exception as e:
        print("load_all_records error:", e)
        return pd.DataFrame()


@st.cache_data(ttl=3600, show_spinner=False)
def _load_recent_records_cached(limit=50):
    """
    ★ v3.9 新增: 录入页专用的轻量查询 (最近N条 + 总记录数)
    - 只拉最近limit条记录 + 一条COUNT查询, 网络载荷恒定, 不随数据量增长变慢
    - 与 _load_all_records_cached 分开缓存: 录入页刷新不会牵连报告页的全量缓存
    - 写操作后由 _clear_records_cache() 一并清空, 提交后表格立即显示新记录
    返回 (最近N条DataFrame, 总记录数int)
    """
    result = _execute("SELECT * FROM quality_records ORDER BY id DESC LIMIT ?", [limit])
    df = pd.DataFrame(result.rows, columns=result.columns)
    df = df.rename(columns=_DB_TO_DISPLAY)
    df = df.drop(columns=["additional_defects"], errors="ignore")
    # 总记录数单独COUNT(不拉全表); Turso的integer经_from_turso_value已转int
    cnt = _execute("SELECT COUNT(*) FROM quality_records")
    total = int(cnt.rows[0][0]) if cnt.rows else 0
    return df, total


def load_recent_records(limit=50):
    """加载最近N条记录(轻量, 录入页表格用); 失败返回 (空表, 0)"""
    try:
        return _load_recent_records_cached(limit)
    except Exception as e:
        print("load_recent_records error:", e)
        return pd.DataFrame(), 0


def _clear_records_cache():
    """清空记录读取缓存(内部用): 增/删记录后调用, 确保下次读取拿到最新数据
    ★ v3.9: 同时清"全量缓存"和"最近N条缓存"(录入表格提交后立即显示新记录)"""
    try:
        _load_all_records_cached.clear()
    except Exception:
        pass
    try:
        _load_recent_records_cached.clear()
    except Exception:
        pass


def clear_records_cache():
    """对外接口: 手动刷新按钮用, 强制下次从数据库拉最新数据"""
    _clear_records_cache()


def delete_record(record_id):
    """删除指定ID的记录
    返回: "deleted"=删除成功 / "not_found"=记录已不存在(云端已删) / False=其他失败
    ★ 修复: 云端已删的记录不再返回False导致页面误报"删除失败"
    """
    try:
        result = _execute("DELETE FROM quality_records WHERE id = ?", [record_id])
        if result.affected_row_count > 0:
            _clear_records_cache()  # ★ 删除成功后清缓存
            return "deleted"
        return "not_found"  # ★ 数据库里已没有这条记录(可能已在云端删除)
    except Exception as e:
        print("delete_record error:", e)
        return False


def calculate_scrap_rate(scrap_qty, inspect_qty):
    """计算报废率(%)"""
    if inspect_qty <= 0:
        return 0.0
    return round(scrap_qty / inspect_qty * 100, 2)


def calculate_rework_rate(rework_qty, inspect_qty):
    """计算返修率(%)"""
    if inspect_qty <= 0:
        return 0.0
    return round(rework_qty / inspect_qty * 100, 2)


def calculate_total_defect_rate(scrap_rate, rework_rate):
    """计算总不合格率(%) = 报废率 + 返修率"""
    return round(scrap_rate + rework_rate, 2)


def evaluate_warning(scrap_rate, consecutive_days):
    """评估是否触发预警"""
    if scrap_rate >= WARNING_THRESHOLD and consecutive_days >= CONSECUTIVE_DAYS:
        return "连续{}天报废率≥{}%".format(consecutive_days, WARNING_THRESHOLD)
    return None


@st.cache_data(show_spinner=False)
def get_alerts_from_data(df):
    """从DataFrame中提取活跃预警列表
    ★ v3.9 加缓存: 按df内容哈希缓存(侧边栏/分析报告/月报共用同一份数据时只算一次)"""
    alerts = []
    if df.empty:
        return alerts
    grouped = df.groupby(["型号", "工序"]).agg({"报废数量": "sum"}).reset_index()
    for _, row in grouped.iterrows():
        scrap_count = int(row["报废数量"])
        if scrap_count >= SCRAP_COUNT_CRITICAL:
            alerts.append({
                "型号": row["型号"],
                "工序": row["工序"],
                "等级": "批次不良",
                "当日报废件数": scrap_count,
                "原因": "当日报废≥{}件, 判定小批次报废".format(SCRAP_COUNT_CRITICAL),
            })
        elif scrap_count >= SCRAP_COUNT_WARNING:
            alerts.append({
                "型号": row["型号"],
                "工序": row["工序"],
                "等级": "重点关注",
                "当日报废件数": scrap_count,
                "原因": "当日报废≥{}件, 需加强巡检".format(SCRAP_COUNT_WARNING),
            })
    return alerts


def export_report_to_excel(df, start_date: str, end_date: str) -> bytes:
    import io

    # 去掉系统字段，只保留业务字段
    drop_cols = ["id", "日期_dt"]
    for col in drop_cols:
        if col in df.columns:
            df = df.drop(columns=[col])

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="质量报告")
    output.seek(0)
    return output.getvalue()

# 兼容旧导入名 (app.py 或其他模块可能引用)
DEFECT_TYPE_FIELDS = DEFECT_COLUMNS


# ============================================================
# ★★★ 以下为 v3.3 新增数据分析聚合函数 ★★★
# 所有函数接收 load_all_records() 返回的 DataFrame (列名为界面显示名, 带斜杠)
# 不修改任何现有函数, 仅在末尾新增
# ============================================================


@st.cache_data(show_spinner=False)
def aggregate_model_pass_rate(df):
    """
    按型号聚合合格率
    返回 DataFrame: [型号, 检测数量, 合格数量, 报废数量, 合格率]
    合格率 = 合格数量 / 检测数量 * 100, 按合格率从高到低排列
    ★ v3.9 加缓存: 月报两次调用(图表+明细)只算一次; 数据变化时缓存自动失效
    """
    if df.empty:
        return pd.DataFrame(columns=["型号", "检测数量", "合格数量", "报废数量", "合格率"])
    grouped = df.groupby("型号").agg(
        检测数量=("检测数量", "sum"),
        合格数量=("合格数量", "sum"),
        报废数量=("报废数量", "sum"),
    ).reset_index()
    grouped["合格率"] = grouped.apply(
        lambda r: round(r["合格数量"] / r["检测数量"] * 100, 2) if r["检测数量"] > 0 else 0.0,
        axis=1,
    )
    return grouped.sort_values("合格率", ascending=False).reset_index(drop=True)


@st.cache_data(show_spinner=False)
def aggregate_defect_summary(df):
    """
    按6个缺陷字段聚合不良数据
    返回 DataFrame: [缺陷类型, 不良数量, 占比, 涉及型号数, 发生次数]
    - 不良数量: 该缺陷字段非空时, 对应记录的报废数量之和
    - 涉及型号数: 该缺陷出现的不同型号数量
    - 发生次数: 该缺陷字段非空的记录条数
    ★ v3.9 加缓存: 月报总结里调用, 数据变化时缓存自动失效
    """
    if df.empty:
        return pd.DataFrame(columns=["缺陷类型", "不良数量", "占比", "涉及型号数", "发生次数"])
    rows = []
    for field in DEFECT_FIELDS_DISPLAY:
        if field not in df.columns:
            continue
        mask = df[field].astype(str).str.strip() != ""
        sub = df[mask]
        if sub.empty:
            continue
        bad_qty = int(sub["报废数量"].sum())
        model_count = int(sub["型号"].nunique())
        occur_count = len(sub)
        rows.append({
            "缺陷类型": field,
            "不良数量": bad_qty,
            "涉及型号数": model_count,
            "发生次数": occur_count,
        })
    if not rows:
        return pd.DataFrame(columns=["缺陷类型", "不良数量", "占比", "涉及型号数", "发生次数"])
    result = pd.DataFrame(rows).sort_values("不良数量", ascending=False).reset_index(drop=True)
    total = result["不良数量"].sum()
    result["占比"] = result["不良数量"].apply(lambda x: round(x / total * 100, 1) if total > 0 else 0.0)
    return result


def aggregate_top5_alerts(df, days, threshold):
    """
    近N天报废Top5型号聚合
    - days: 近多少天 (3/5/7/30)
    - threshold: 报废件数阈值
    返回 DataFrame: [型号, 检测数量, 报废数量, 报废率, 主要缺陷, 主要缺陷数量]
    """
    if df.empty:
        return pd.DataFrame()
    df = df.copy()
    df["日期_dt"] = pd.to_datetime(df["日期"], errors="coerce")
    cutoff = pd.Timestamp.now() - pd.Timedelta(days=days)
    recent = df[df["日期_dt"] >= cutoff].copy()
    if recent.empty:
        return pd.DataFrame()
    grouped = recent.groupby("型号").agg(
        检测数量=("检测数量", "sum"),
        报废数量=("报废数量", "sum"),
    ).reset_index()
    grouped = grouped[grouped["报废数量"] >= threshold]
    if grouped.empty:
        return pd.DataFrame()
    grouped["报废率"] = grouped.apply(
        lambda r: round(r["报废数量"] / r["检测数量"] * 100, 2) if r["检测数量"] > 0 else 0.0,
        axis=1,
    )
    # 找每型号主要缺陷
    main_defects = []
    main_defect_qty = []
    for model in grouped["型号"]:
        sub = recent[recent["型号"] == model]
        defect_qty = {}
        for field in DEFECT_FIELDS_DISPLAY:
            if field not in sub.columns:
                continue
            mask = sub[field].astype(str).str.strip() != ""
            q = int(sub.loc[mask, "报废数量"].sum())
            if q > 0:
                defect_qty[field] = q
        if defect_qty:
            top_defect = max(defect_qty, key=defect_qty.get)
            main_defects.append(top_defect)
            main_defect_qty.append(defect_qty[top_defect])
        else:
            main_defects.append("-")
            main_defect_qty.append(0)
    grouped["主要缺陷"] = main_defects
    grouped["主要缺陷数量"] = main_defect_qty
    return grouped.sort_values("报废数量", ascending=False).head(5).reset_index(drop=True)


def aggregate_top5_defect_by_model(df, top_n=5):
    """
    Top5型号按缺陷类型堆叠聚合 (用于图3堆叠柱状图)
    返回 DataFrame: [型号, 缺陷类型, 报废数量]
    """
    if df.empty:
        return pd.DataFrame()
    rows = []
    top5 = aggregate_top5_alerts(df, days=30, threshold=0)
    if top5.empty:
        return pd.DataFrame()
    top_models = top5["型号"].tolist()
    for model in top_models:
        sub = df[df["型号"] == model]
        for field in DEFECT_FIELDS_DISPLAY:
            if field not in sub.columns:
                continue
            mask = sub[field].astype(str).str.strip() != ""
            q = int(sub.loc[mask, "报废数量"].sum())
            if q > 0:
                rows.append({"型号": model, "缺陷类型": field, "报废数量": q})
    if not rows:
        return pd.DataFrame()
    return pd.DataFrame(rows)


def aggregate_monthly_trend(df, model, defect_type):
    """
    月度趋势聚合
    - model: "全部" 或具体型号
    - defect_type: "全部" 或具体缺陷字段名(界面显示名)
    返回 DataFrame: [月份, 缺陷类型, 不良数量]
    """
    if df.empty:
        return pd.DataFrame()
    df = df.copy()
    df["日期_dt"] = pd.to_datetime(df["日期"], errors="coerce")
    df = df.dropna(subset=["日期_dt"])
    df["月份"] = df["日期_dt"].dt.strftime("%Y-%m")
    if model != "全部":
        df = df[df["型号"] == model]
    if df.empty:
        return pd.DataFrame()
    if defect_type == "全部":
        # ★ 修复: 保留6种缺陷类型, 即使数据为0也保留 (之前会跳过不存在的字段)
        all_months = sorted(df["月份"].unique())
        rows = []
        for field in DEFECT_FIELDS_DISPLAY:
            if field in df.columns:
                mask = df[field].astype(str).str.strip() != ""
                sub = df[mask]
                if not sub.empty:
                    monthly = sub.groupby("月份")["报废数量"].sum().reindex(all_months, fill_value=0).reset_index()
                else:
                    monthly = pd.DataFrame({"月份": all_months, "报废数量": [0] * len(all_months)})
            else:
                # 字段不存在, 全0
                monthly = pd.DataFrame({"月份": all_months, "报废数量": [0] * len(all_months)})
            monthly["缺陷类型"] = field
            monthly = monthly.rename(columns={"报废数量": "不良数量"})
            rows.append(monthly[["月份", "缺陷类型", "不良数量"]])
        if not rows:
            return pd.DataFrame()
        return pd.concat(rows, ignore_index=True).sort_values(["月份", "不良数量"], ascending=[True, False])
    else:
        if defect_type not in df.columns:
            return pd.DataFrame()
        mask = df[defect_type].astype(str).str.strip() != ""
        monthly = df[mask].groupby("月份")["报废数量"].sum().reset_index()
        monthly["缺陷类型"] = defect_type
        return monthly.rename(columns={"报废数量": "不良数量"})


def calculate_pass_rate_capped(pass_qty, rework_qty, scrap_qty, inspect_qty):
    """
    ★ v3.12 口径变更: 一次直通率(FTQ) = 合格数量 / 检测数量 × 100%
    - v3.6曾用反推口径(100-报废率-返修率)吸收录入误差, 但管理层需要的是
      "一次直通率": 只有一次性合格的才算, 返修件不算合格 → 改为真实直算
    - 返修件不再计入分子; 极端脏数据(合格>检测)时封顶100%, 不出现>100%
    - rework_qty/scrap_qty参数保留(兼容既有调用签名), 计算不再使用
    """
    if inspect_qty <= 0:
        return 0.0
    pass_rate = pass_qty / inspect_qty * 100
    if pass_rate > 100:   # 防脏数据(合格+返修+报废>检测)出现>100%
        pass_rate = 100.0
    if pass_rate < 0:
        pass_rate = 0.0
    return round(pass_rate, 2)


@st.cache_data(show_spinner=False)
def get_monthly_overview(df, year_month):
    """
    月度概览KPI
    - year_month: "YYYY-MM" 格式
    返回 dict: {总检验数, 合格率, 报废率, 返修率, 预警次数, ...}
    ★ v3.12: "合格率"键的值改为一次直通率口径(合格数量/检测数量), 返修不再计入;
      键名保持"合格率"以兼容现有调用方(app.py月报), 界面标签已改为"一次直通率"
    ★ v3.9 加缓存: 月报当月+环比上月各调一次, 同数据只算一次
    """
    _empty = {"总检验数": 0, "合格率": 0.0, "报废率": 0.0, "返修率": 0.0, "预警次数": 0}
    if df.empty:
        return _empty
    df = df.copy()
    df["日期_dt"] = pd.to_datetime(df["日期"], errors="coerce")
    df["月份"] = df["日期_dt"].dt.strftime("%Y-%m")
    month_df = df[df["月份"] == year_month]
    if month_df.empty:
        return _empty
    total_inspect = int(month_df["检测数量"].sum())
    total_pass = int(month_df["合格数量"].sum())
    total_scrap = int(month_df["报废数量"].sum())
    total_rework = int(month_df["返修数量"].sum())
    # ★ v3.12: 报废率/返修率口径不变(真实数/检测); 合格率=一次直通率(合格/检测)
    scrap_rate = round(total_scrap / total_inspect * 100, 2) if total_inspect > 0 else 0.0
    rework_rate = round(total_rework / total_inspect * 100, 2) if total_inspect > 0 else 0.0
    pass_rate = calculate_pass_rate_capped(total_pass, total_rework, total_scrap, total_inspect)
    alerts = get_alerts_from_data(month_df)
    return {
        "总检验数": total_inspect,
        "合格率": pass_rate,
        "报废率": scrap_rate,
        "返修率": rework_rate,
        "预警次数": len(alerts),
        "总合格数": total_pass,
        "总报废数": total_scrap,
        "总返修数": total_rework,
    }


def get_monthly_pass_rate_trend(df, months_count=3):
    """
    最近N个月合格率趋势
    返回 DataFrame: [月份, 合格率, 报废率]
    """
    if df.empty:
        return pd.DataFrame()
    df = df.copy()
    df["日期_dt"] = pd.to_datetime(df["日期"], errors="coerce")
    df = df.dropna(subset=["日期_dt"])
    df["月份"] = df["日期_dt"].dt.strftime("%Y-%m")
    monthly = df.groupby("月份").agg(
        检测数量=("检测数量", "sum"),
        合格数量=("合格数量", "sum"),
        报废数量=("报废数量", "sum"),
    ).reset_index()
    monthly["合格率"] = monthly.apply(
        lambda r: round(r["合格数量"] / r["检测数量"] * 100, 2) if r["检测数量"] > 0 else 0.0, axis=1)
    monthly["报废率"] = monthly.apply(
        lambda r: round(r["报废数量"] / r["检测数量"] * 100, 2) if r["检测数量"] > 0 else 0.0, axis=1)
    monthly = monthly.sort_values("月份").tail(months_count).reset_index(drop=True)
    return monthly[["月份", "合格率", "报废率"]]
