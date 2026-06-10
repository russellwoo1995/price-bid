# 苏州政府采购招标信息爬虫 & LangChain Agent

自动采集苏州地区政府招标信息，通过 DeepSeek LLM 生成每日汇总报告，保存为 Excel 表格和 Markdown 文件。

## 覆盖平台（9 个）

| 平台 | 说明 |
|------|------|
| 苏州政采网（意向） | 政府采购意向公告 |
| 苏州政采网（招标） | 政府采购招标公告 |
| 苏州市公共资源交易网 | 工程建设、多元化采购等 |
| 江苏省招投标平台 | 省级招投标公告 |
| 吴中区限额平台 | 限额以下工程项目 |
| 苏州城市学院 | 高校招标公告 |
| 苏州职业技术大学 | 高校招标公告 |
| 苏州工业职业技术学院 | 高校招标公告 |
| 苏州阳光采购平台 | 国企/集体采购公告 |

## 快速开始

### 1. 环境准备

```bash
# 安装依赖 + Playwright 浏览器
make install
```

### 2. 配置 API Key

创建 `.env` 文件，填入 DeepSeek API Key：

```env
DEEPSEEK_API_KEY=sk-your-api-key
```

### 3. 运行

```bash
# 直接采集，保存到 Excel（不生成汇总报告）
make run

# 通过 Agent 采集 + LLM 生成日报（保存为 md 文件）
make agent

# 定时模式：每天 09:00 自动执行
make agent-schedule
```

## 输出文件

- `苏州政府采购总表_YYYY-MM-DD.xlsx` — 各平台分 sheet 存储的原始数据
- `招标日报_YYYY-MM-DD.md` — LLM 生成的 Markdown 汇总报告

## 项目结构

```
.
├── bid_all.py          # 爬虫核心：9 个采集函数 + Excel 写入
├── agent.py            # LangChain Agent：调度采集 + LLM 汇总 + md 输出
├── pyproject.toml      # 项目配置与依赖
├── Makefile            # 常用命令
├── .env                # API Key（不入 git）
└── .gitignore
```

## 定时运行

```bash
# 默认每天 09:00
uv run python agent.py --schedule

# 自定义时间
uv run python agent.py --schedule --time 08:30
```

## 技术栈

- **爬虫**：requests + BeautifulSoup + Playwright（含浏览器回退）
- **数据**：pandas + openpyxl
- **Agent**：LangChain + DeepSeek Chat API
- **调度**：APScheduler
- **环境**：uv + Python 3.12–3.13

## 注意事项

- Python 3.14 暂不支持（langchain/pydantic 兼容性问题），已锁定 `<3.14`
- 部分平台需要 Playwright 浏览器支持，首次使用需 `uv run playwright install chromium`
- 阳光采购平台有反爬机制，使用浏览器自动化访问
