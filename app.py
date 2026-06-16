"""苏州政府采购招标信息采集系统 — Streamlit Web 应用

启动方式：
    streamlit run app.py
    # 或
    make app
"""

import io
import json
import os
import re
import sys
import zoneinfo
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
from dotenv import load_dotenv

load_dotenv()

# ── 常量 ────────────────────────────────────────────────
_TZ = zoneinfo.ZoneInfo("Asia/Shanghai")
_DIR = Path(__file__).resolve().parent
_CONFIG_PATH = _DIR / ".bid_config.json"

# ── 页面配置（必须在最前面） ────────────────────────────
st.set_page_config(
    page_title="苏州招标采集系统",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── 导入爬虫 ─────────────────────────────────────────────
import bid_all
from langchain_openai import ChatOpenAI
from langchain_core.messages import HumanMessage, SystemMessage

# ── 爬虫元数据 ─────────────────────────────────────────────
CRAWLERS = [
    (bid_all.fetch_suzhou_gov_yixiang,          "苏州政采网·采购意向"),
    (bid_all.fetch_suzhou_gov_bids,             "苏州政采网·招标公告"),
    (bid_all.fetch_suzhou_gonggong_gov_bids,    "公共资源交易"),
    (bid_all.fetch_jiangsu_gov_bids,            "江苏省招投标"),
    (bid_all.fetch_xiane_gov_bids,              "吴中限额平台"),
    (bid_all.fetch_suzhou_city_college,         "苏州城市学院"),
    (bid_all.fetch_suzhou_carrer_university,    "苏州职业技术大学"),
    (bid_all.fetch_suzhou_industry_university,  "苏州工业职院"),
    (bid_all.fetch_suzhou_yangguang_bids,       "阳光采购平台"),
]


# ═══════════════════════════════════════════════════════════
# 工具函数
# ═══════════════════════════════════════════════════════════
def _now() -> str:
    return datetime.now(_TZ).strftime("%Y-%m-%d")


def _load_cfg() -> dict:
    if _CONFIG_PATH.exists():
        return json.loads(_CONFIG_PATH.read_text("utf-8"))
    return {"schedule_time": "09:00", "model": "deepseek-chat", "base_url": "https://api.deepseek.com"}


def _save_cfg(cfg: dict):
    _CONFIG_PATH.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), "utf-8")


def _capture(fn):
    """调用爬虫函数，捕获 stdout，返回 (结果列表, 日志文本)"""
    buf = io.StringIO()
    saved = sys.stdout
    sys.stdout = buf
    try:
        results = fn()
    except Exception as exc:
        results = []
        print(f"⚠️ 失败：{exc}")
    finally:
        sys.stdout = saved
    return results or [], buf.getvalue().strip()


def _show_df(df: pd.DataFrame):
    """展示 DataFrame，链接列可点击"""
    col_config = {}
    for col in df.columns:
        if col == "链接":
            col_config[col] = st.column_config.LinkColumn("链接", display_text="打开")
    st.dataframe(df, width='stretch', hide_index=True, column_config=col_config or None)


def _generate_llm_summary(date_str: str, results: dict):
    """用 DeepSeek 生成汇总报告"""
    api_key = os.environ.get("DEEPSEEK_API_KEY", "")
    if not api_key:
        st.error("请先在「配置管理」中设置 DeepSeek API Key")
        return

    with st.spinner("🤖 DeepSeek 正在生成汇总报告..."):
        parts = []
        for name, data in results.items():
            parts.append(f"## {name}")
            for item in data:
                parts.append(" | ".join(f"{k}: {v}" for k, v in item.items()))
        data_text = "\n".join(parts)

        cfg = _load_cfg()
        llm = ChatOpenAI(
            model=cfg.get("model", "deepseek-chat"),
            base_url=cfg.get("base_url", "https://api.deepseek.com"),
            api_key=api_key,
        )

        system_prompt = """你是苏州政府招标信息助手。根据采集到的数据，用中文撰写招标汇总报告：
- 总条数与成功/失败平台概览
- 按平台分组，逐条列出项目标题与链接
- 如无新增，注明"今日暂无新招标信息"
- 最后附上重点关注和建议"""

        resp = llm.invoke([
            SystemMessage(content=system_prompt),
            HumanMessage(content=f"今日招标数据（{date_str}）：\n\n{data_text}"),
        ])

        st.session_state.llm_summary = resp.content
        st.success("AI 报告生成完成！")


def _download_buttons(date_str: str):
    """根据日期显示下载按钮"""
    excel_path = _DIR / f"苏州政府采购总表_{date_str}.xlsx"
    md_path = _DIR / f"招标日报_{date_str}.md"

    if excel_path.exists():
        st.download_button(
            "📥 下载 Excel",
            data=excel_path.read_bytes(),
            file_name=excel_path.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    if md_path.exists():
        st.download_button(
            "📥 下载 Markdown 报告",
            data=md_path.read_bytes(),
            file_name=md_path.name,
            mime="text/markdown",
        )


def _write_env_key(key: str):
    env_path = _DIR / ".env"
    content = env_path.read_text("utf-8") if env_path.exists() else ""
    if "DEEPSEEK_API_KEY" in content:
        content = re.sub(r"DEEPSEEK_API_KEY=.*", f"DEEPSEEK_API_KEY={key}", content)
    else:
        content += f"\nDEEPSEEK_API_KEY={key}\n"
    env_path.write_text(content, "utf-8")


# ═══════════════════════════════════════════════════════════
# Session State
# ═══════════════════════════════════════════════════════════
_defaults = {
    "coll_results": {},
    "coll_logs": {},
    "coll_date": "",
    "llm_summary": "",
    "_last_collect_ts": 0.0,
}
for k, v in _defaults.items():
    st.session_state.setdefault(k, v)


# ═══════════════════════════════════════════════════════════
# 侧边栏
# ═══════════════════════════════════════════════════════════
with st.sidebar:
    st.title("📋 苏州招标采集")
    st.caption("苏州政府采购招标信息 · 自动采集 & AI 汇总")
    st.divider()
    page = st.radio("功能导航", ["一键采集", "历史报告", "定时任务", "配置管理"])
    st.divider()
    _today = _now()
    st.metric("今日日期", _today)
    excel_today = _DIR / f"苏州政府采购总表_{_today}.xlsx"
    md_today = _DIR / f"招标日报_{_today}.md"
    if excel_today.exists() or md_today.exists():
        st.success("今日已有采集记录", icon="✅")
    else:
        st.info("今日暂无采集记录", icon="⏳")


# ═══════════════════════════════════════════════════════════
# 页面 1：一键采集
# ═══════════════════════════════════════════════════════════
if page == "一键采集":
    st.header("🚀 一键采集招标信息")

    col_date, col_mode = st.columns([1, 2])
    with col_date:
        collect_date = st.date_input("采集日期", value=datetime.now(_TZ).date())
    with col_mode:
        auto_llm = st.checkbox("采集后自动生成 AI 汇总报告", value=False)

    date_str = collect_date.strftime("%Y-%m-%d")

    # ── 开始采集 ─────────────────────────────────
    if st.button("🔍 开始采集", type="primary", width='stretch'):
        import time as _time
        now_ts = _time.time()
        cooldown_secs = 300  # 5 分钟冷却期，防止触发反爬
        elapsed = now_ts - st.session_state._last_collect_ts

        if elapsed < cooldown_secs and st.session_state._last_collect_ts > 0:
            remaining = int(cooldown_secs - elapsed)
            st.warning(
                f"⏳ 采集冷却中，请等待 **{remaining // 60} 分 {remaining % 60} 秒** 后再试。"
                f"（冷却期 {cooldown_secs // 60} 分钟，防止触发反爬机制）"
            )
        else:
            st.session_state._last_collect_ts = now_ts
            bid_all.TODAY = date_str
            bid_all.AUTO_SAVE = True
            all_results = {}
            all_logs = {}

            progress = st.progress(0, text="准备采集...")

            for i, (fn, name) in enumerate(CRAWLERS):
                with st.status(f"📡 {name}", expanded=True) as status:
                    results, log_text = _capture(fn)
                    all_results[name] = results
                    all_logs[name] = log_text

                    n = len(results)
                    if n > 0:
                        status.update(label=f"✅ {name} — {n} 条", state="complete", expanded=False)
                    else:
                        status.update(label=f"⚪ {name} — 无数据", state="complete", expanded=False)

                    if log_text:
                        st.code(log_text, language=None)

                progress.progress((i + 1) / len(CRAWLERS), text=f"已完成 {i+1}/{len(CRAWLERS)}")

            progress.empty()
            st.session_state.coll_results = all_results
            st.session_state.coll_logs = all_logs
            st.session_state.coll_date = date_str

            total = sum(len(r) for r in all_results.values())
            st.success(f"采集完成！共获取 **{total}** 条招标信息")

            # 自动生成 AI 报告
            if auto_llm and total > 0:
                _generate_llm_summary(date_str, all_results)

    # ── 展示采集结果 ─────────────────────────────
    if st.session_state.coll_results:
        results = st.session_state.coll_results
        date_str = st.session_state.coll_date
        total = sum(len(r) for r in results.values())

        if total > 0:
            st.subheader(f"📊 采集结果 ({total} 条)")

            # 概览卡片
            active = [(n, len(d)) for n, d in results.items() if d]
            cols = st.columns(min(len(active), 4))
            for idx, (name, count) in enumerate(active):
                with cols[idx % len(cols)]:
                    st.metric(name, f"{count} 条")

            st.divider()

            # 分平台表格
            for name, data in results.items():
                if not data:
                    continue
                with st.expander(f"{name} ({len(data)} 条)", expanded=True):
                    _show_df(pd.DataFrame(data))

            st.divider()

            # AI 汇总 & 下载
            col_llm, col_dl = st.columns(2)
            with col_llm:
                if st.button("🤖 生成 AI 汇总报告", type="primary"):
                    _generate_llm_summary(date_str, results)
            with col_dl:
                _download_buttons(date_str)

            # 显示 AI 报告
            if st.session_state.llm_summary:
                st.subheader("🤖 AI 汇总报告")
                st.markdown(st.session_state.llm_summary)
                if st.button("💾 保存为 Markdown 文件"):
                    content = f"# 招标日报 · {date_str}\n\n{st.session_state.llm_summary}\n"
                    path = _DIR / f"招标日报_{date_str}.md"
                    path.write_text(content, "utf-8")
                    st.success(f"已保存到 {path.name}")
        else:
            st.info(f"{date_str} 暂无招标信息")


# ═══════════════════════════════════════════════════════════
# 页面 2：历史报告
# ═══════════════════════════════════════════════════════════
elif page == "历史报告":
    st.header("📁 历史报告")

    excel_files = sorted(_DIR.glob("苏州政府采购总表_*.xlsx"), reverse=True)
    md_files = sorted(_DIR.glob("招标日报_*.md"), reverse=True)

    if not excel_files and not md_files:
        st.info("暂无历史报告。请先在「一键采集」中执行采集。")
    else:
        # 提取日期
        dates = set()
        for f in excel_files:
            dates.add(f.stem.split("_")[-1])
        for f in md_files:
            dates.add(f.stem.split("_")[-1])

        selected = st.selectbox("选择日期", sorted(dates, reverse=True))

        tab_xl, tab_md = st.tabs(["📊 Excel 数据", "📝 Markdown 报告"])

        with tab_xl:
            p = _DIR / f"苏州政府采购总表_{selected}.xlsx"
            if p.exists():
                try:
                    xls = pd.ExcelFile(p)
                    for sheet in xls.sheet_names:
                        df = pd.read_excel(xls, sheet_name=sheet)
                        st.subheader(sheet)
                        _show_df(df)
                except Exception as e:
                    st.error(f"读取 Excel 失败：{e}")
                st.download_button(
                    "📥 下载 Excel",
                    data=p.read_bytes(),
                    file_name=p.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.info("该日期无 Excel 数据")

        with tab_md:
            p = _DIR / f"招标日报_{selected}.md"
            if p.exists():
                st.markdown(p.read_text("utf-8"))
                st.download_button(
                    "📥 下载 Markdown",
                    data=p.read_bytes(),
                    file_name=p.name,
                    mime="text/markdown",
                )
            else:
                st.info("该日期无 Markdown 报告")


# ═══════════════════════════════════════════════════════════
# 页面 3：定时任务
# ═══════════════════════════════════════════════════════════
elif page == "定时任务":
    st.header("⏰ 定时任务管理")

    cfg = _load_cfg()
    schedule_time = st.text_input(
        "每日执行时间 (HH:MM)",
        value=cfg.get("schedule_time", "09:00"),
        help="格式：HH:MM，24小时制",
    )

    col1, col2 = st.columns(2)
    with col1:
        if st.button("💾 保存时间"):
            cfg["schedule_time"] = schedule_time
            _save_cfg(cfg)
            st.success("配置已保存")

    with col2:
        if st.button("▶️ 测试执行一次"):
            import time as _time
            now_ts = _time.time()
            if now_ts - st.session_state._last_collect_ts < 300:
                st.warning("⏳ 采集冷却中，请稍后再试（防止触发反爬机制）")
            else:
                st.session_state._last_collect_ts = now_ts
                with st.spinner("正在执行一次采集..."):
                    bid_all.TODAY = _now()
                    bid_all.AUTO_SAVE = True
                    total = 0
                    for fn, name in CRAWLERS:
                        try:
                            r, _ = _capture(fn)
                            total += len(r)
                        except Exception:
                            pass
                    st.success(f"执行完成，共获取 {total} 条信息")

    st.divider()
    st.subheader("部署方式")

    hour, minute = schedule_time.split(":")

    st.markdown(f"""
### 方式一：命令行定时进程

```bash
uv run python agent.py --schedule --time {schedule_time}
```

程序将每天 **{schedule_time}** 自动采集并生成报告，按 `Ctrl+C` 停止。

### 方式二：系统 crontab（推荐生产环境）

```bash
crontab -e
```

添加以下行：
```
{minute} {hour} * * * cd {_DIR} && uv run python agent.py
```

### 方式三：launchd（macOS 推荐）

创建 plist 文件后加载即可实现开机自启 + 每日定时执行，适合长期运行。
""")


# ═══════════════════════════════════════════════════════════
# 页面 4：配置管理
# ═══════════════════════════════════════════════════════════
elif page == "配置管理":
    st.header("⚙️ 配置管理")
    cfg = _load_cfg()

    # ── API Key ─────────────────────────────────
    st.subheader("DeepSeek API")
    current_key = os.environ.get("DEEPSEEK_API_KEY", "")
    if current_key:
        masked = current_key[:8] + "****" + current_key[-4:]
        st.text_input("当前 Key", value=masked, disabled=True)
    else:
        st.warning("尚未配置 API Key，AI 汇总功能将不可用")

    new_key = st.text_input("新 Key", type="password", placeholder="sk-...")
    if st.button("💾 保存 Key"):
        if new_key:
            _write_env_key(new_key)
            os.environ["DEEPSEEK_API_KEY"] = new_key
            st.success("API Key 已保存到 .env 文件")
        else:
            st.warning("请输入 Key")

    st.divider()

    # ── 模型设置 ────────────────────────────────
    st.subheader("模型配置")
    model = st.text_input("模型名称", value=cfg.get("model", "deepseek-chat"))
    base_url = st.text_input("API Base URL", value=cfg.get("base_url", "https://api.deepseek.com"))
    if st.button("💾 保存模型配置"):
        cfg["model"] = model
        cfg["base_url"] = base_url
        _save_cfg(cfg)
        st.success("模型配置已保存")

    st.divider()

    # ── 环境检测 ────────────────────────────────
    st.subheader("环境检测")
    st.write(f"- Python: `{sys.version}`")
    st.write(f"- 工作目录: `{_DIR}`")

    try:
        from playwright.sync_api import sync_playwright
        st.write("- Playwright: ✅ 已安装")
    except ImportError:
        st.write("- Playwright: ❌ 未安装（部分爬虫需要）")

    import requests as _req
    st.write(f"- requests: ✅ {_req.__version__}")

    try:
        from langchain_openai import ChatOpenAI as _
        st.write("- langchain-openai: ✅")
    except ImportError:
        st.write("- langchain-openai: ❌ 未安装")

    st.divider()

    # ── 关于 ────────────────────────────────────
    st.subheader("关于")
    st.markdown("""
**苏州政府采购招标信息采集系统** v0.1.0

- 9 个平台自动采集苏州地区政府招标信息
- DeepSeek AI 生成每日汇总报告
- 输出 Excel 表格 + Markdown 文件

技术栈：Python · Streamlit · LangChain · Playwright
""")
