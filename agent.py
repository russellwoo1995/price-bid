"""每日招标信息 LangChain Agent

环境变量：
    DEEPSEEK_API_KEY  — DeepSeek API 密钥（必填）

用法：
    uv run python agent.py                        # 立即执行一次
    uv run python agent.py --schedule             # 每天 09:00 自动执行
    uv run python agent.py --schedule --time 08:30
"""
import io
import logging
import os
import sys
import zoneinfo
from datetime import datetime

_TZ_BEIJING = zoneinfo.ZoneInfo("Asia/Shanghai")

from dotenv import load_dotenv

load_dotenv()

from apscheduler.schedulers.blocking import BlockingScheduler
from langchain.agents import AgentExecutor, create_tool_calling_agent
from langchain.tools import Tool
from langchain_openai import ChatOpenAI
from langchain_core.prompts import ChatPromptTemplate

import bid_all

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)


def _capture(fn) -> str:
    """Run fn(), capture its stdout, return as a string."""
    buf = io.StringIO()
    saved = sys.stdout
    sys.stdout = buf
    try:
        fn()
    except Exception as exc:
        print(f"⚠️  失败：{exc}")
    finally:
        sys.stdout = saved
    return buf.getvalue().strip() or "（无输出）"


def _make_tool(fn, name: str, desc: str) -> Tool:
    return Tool(name=name, func=lambda _: _capture(fn), description=desc)


TOOLS = [
    _make_tool(bid_all.fetch_suzhou_gov_yixiang,         "suzhou_gov_yixiang",         "获取苏州政采网采购意向公告"),
    _make_tool(bid_all.fetch_suzhou_gov_bids,            "suzhou_gov_bids",            "获取苏州政采网招标公告"),
    _make_tool(bid_all.fetch_suzhou_gonggong_gov_bids,   "suzhou_gonggong_bids",       "获取苏州市公共资源交易招标公告"),
    _make_tool(bid_all.fetch_jiangsu_gov_bids,           "jiangsu_gov_bids",           "获取江苏省招投标公告"),
    _make_tool(bid_all.fetch_xiane_gov_bids,             "xiane_wuzhong_bids",         "获取吴中区限额招标公告（可能被网关拦截）"),
    _make_tool(bid_all.fetch_suzhou_city_college,        "suzhou_city_college",        "获取苏州城市学院招标公告"),
    _make_tool(bid_all.fetch_suzhou_carrer_university,   "suzhou_career_university",   "获取苏州职业技术大学招标公告"),
    _make_tool(bid_all.fetch_suzhou_industry_university, "suzhou_industry_university", "获取苏州工业职业技术学院招标公告"),
    _make_tool(bid_all.fetch_suzhou_yangguang_bids,      "suzhou_yangguang_bids",      "获取苏州阳光采购平台招标公告（需启动浏览器，耗时较长）"),
]

_SYSTEM = """\
你是苏州政府招标信息助手，负责每天自动收集各平台的招标信息。

执行步骤：
1. 依次调用所有 9 个工具，收集各平台今日招标信息。
2. 某工具失败时，记录失败原因后继续调用下一个。
3. 所有工具调用完毕后，用中文撰写今日招标汇总报告，格式为：
   - 总条数与成功/失败平台概览
   - 按平台分组，逐条列出项目标题与链接
   - 如当日无新增，注明"今日暂无新招标信息"
"""


def _build_executor() -> AgentExecutor:
    llm = ChatOpenAI(
        model="deepseek-chat",
        base_url="https://api.deepseek.com",
        api_key=os.environ["DEEPSEEK_API_KEY"],
    )
    prompt = ChatPromptTemplate.from_messages([
        ("system", _SYSTEM),
        ("human", "{input}"),
        ("placeholder", "{agent_scratchpad}"),
    ])
    agent = create_tool_calling_agent(llm, TOOLS, prompt)
    return AgentExecutor(
        agent=agent,
        tools=TOOLS,
        verbose=True,
        max_iterations=25,
        handle_parsing_errors=True,
    )


def run_daily() -> None:
    today = datetime.now(_TZ_BEIJING).strftime("%Y-%m-%d")
    bid_all.TODAY = today  # override the scraper module's date filter

    logger.info("▶  开始收集 %s 招标信息", today)
    executor = _build_executor()
    result = executor.invoke({
        "input": f"今天是 {today}，请依次调用所有工具收集今日招标信息，完成后生成汇总报告。",
    })

    sep = "=" * 60
    print(f"\n{sep}\n  招标日报 · {today}\n{sep}")
    print(result["output"])
    print(f"{sep}\n")
    logger.info("✅  完成")


def main() -> None:
    import argparse

    p = argparse.ArgumentParser(description="每日招标信息 Agent")
    p.add_argument("--schedule", action="store_true", help="启动定时模式")
    p.add_argument("--time", default="09:00", metavar="HH:MM", help="定时时间（默认 09:00）")
    args = p.parse_args()

    if args.schedule:
        h, m = map(int, args.time.split(":"))
        scheduler = BlockingScheduler(timezone="Asia/Shanghai")
        scheduler.add_job(run_daily, "cron", hour=h, minute=m)
        logger.info("定时任务已启动，每天 %s 执行。Ctrl+C 停止。", args.time)
        try:
            scheduler.start()
        except (KeyboardInterrupt, SystemExit):
            scheduler.shutdown()
    else:
        run_daily()


if __name__ == "__main__":
    main()
