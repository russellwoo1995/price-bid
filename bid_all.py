import requests
from datetime import datetime, timedelta
import pandas as pd
import os
from bs4 import BeautifulSoup
import json
import time
import re
import random
import zoneinfo

from playwright.sync_api import sync_playwright

_TZ_BEIJING = zoneinfo.ZoneInfo("Asia/Shanghai")
TODAY = datetime.now(_TZ_BEIJING).strftime("%Y-%m-%d")
true_categories = ["003001","003002","003003","003034"]

def fetch_suzhou_gov_bids():
    url = "https://czju.suzhou.gov.cn/zfcg/content/searchContents.action"
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Content-Type": "application/x-www-form-urlencoded",
    }

    payload = {
        "title": "",
        "choose": "",
        "type": "0",
        "zbCode": "",
        "appcode": "",
        "page": 1,
        "rows": 100
    }

    response = requests.post(url, headers=headers, data=payload)
    response.raise_for_status()

    data = response.json()
    rows = data.get("rows", [])

    print(f"【{TODAY} 苏州市招标信息】")
    found = False
    results = []

    for row in rows:
        release_time = row["RELEASE_TIME"][:10]
        if release_time == TODAY:
            title = row["TITLE"]
            project_id = row["PROJECTID"]
            area = row["AREA"]
            link = f"https://czju.suzhou.gov.cn/zfcg/html/project/{project_id}.shtml"
            if area == "吴中区" or area == "苏州":
                print(f"[{area}]{title}\n{link}\n")
                found = True
                results.append({
                    "地区": area,
                    "标题": title,
                    "发布日期": release_time,
                    "链接": link
                })

    if not found:
        print("今日暂无招标。")
    elif results:
        write_to_excel(results,"suzhou_gov_bids")

def fetch_jiangsu_gov_bids():
    url = "https://api.jszbtb.com/DataGatewayApi/PublishBulletins"
    start_time = TODAY + " 00:00:00"
    end_time = TODAY + " 23:59:59"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36 Edg/126.0.0.0",
        "Referer": "https://www.jszbtb.com/",
        "Accept": "application/json, text/plain, */*",
        "Accept-Language": "zh-CN,zh;q=0.9",
    }
    params = {
        "bulletinType": 1,
        "industryCode": "",
        "regionCode": 320500,
        "startTime": start_time,
        "endTime": end_time,
        "keyword": "",
        "currentPage": 1,
        "pageSize": 20
    }

    # 重试机制：最多尝试 2 次
    response = None
    for attempt in range(2):
        try:
            response = requests.get(url, headers=headers, params=params, timeout=15)
            response.raise_for_status()
            break
        except requests.HTTPError as e:
            if attempt == 1:
                print(f"⚠️  江苏省招投标 失败：{e}")
                return
            print(f"⚠️  江苏省招投标第 {attempt+1} 次尝试失败（{e}），3 秒后重试...")
            time.sleep(3)
        except requests.RequestException as e:
            print(f"⚠️  江苏省招投标 网络错误：{e}")
            return

    if response is None or not response.text.strip().startswith("{"):
        print(f"⚠️  江苏省招投标接口返回非 JSON（可能被封锁），跳过。")
        return

    data = response.json()
    outer_data = data.get("data")
    if outer_data is None or not isinstance(outer_data, dict):
        print(f"⚠️  江苏省招投标接口返回错误：{data.get('errorMessage', '未知错误')}，跳过。")
        return
    rows = outer_data.get("data", [])

    print(f"【{TODAY} 江苏省招标信息】")
    found = False
    results = []

    for row in rows:
        title = row["bulletinName"]
        project_id = row["bulletinID"]
        area = row["regionName"]
        release_time = row["openBidTime"][:10]
        link = f"https://www.jszbtb.com/#/bulletinDetails/%E6%8B%9B%E6%A0%87%E5%85%AC%E5%91%8A/{project_id}"
        if area == "江苏省苏州市吴中区" or area == "江苏省,苏州市,吴中区" or area == "江苏省/苏州市/吴中区":
            area = "吴中区"
            industry_name = row["industryName"]
            print(f"[{area}]{title}\n{link}\n")
            found = True
            results.append({
                "地区": area,
                "标题": title,
                "发布日期": release_time,
                "链接": link,
                "行业": industry_name
            })
        elif area == "江苏省苏州市" or area == "江苏省,苏州市" or area == "江苏省苏州市市辖区" or area == "江苏省/苏州市":
            area = "苏州市"
            industry_name = row["industryName"]
            print(f"[{area}]{title}\n{link}\n")
            found = True
            results.append({
                "地区": area,
                "标题": title,
                "发布日期": release_time,
                "链接": link,
                "行业": industry_name
            })

    if not found:
        print("今日暂无招标。")
    elif results:
        write_to_excel(results,"jiangsu_gov_bids")

# URL搜到，不返回信息，需要换一个方法
def fetch_xiane_gov_bids():
    url = 'https://www.wzqzjj-wfw.cn:8090/page/Project/ProZhaoBiaogsMX.aspx?lx=1'
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36'
    }

    response = None
    # 先用 requests 尝试，超时设为 20 秒，最多重试 2 次
    for attempt in range(2):
        try:
            response = requests.get(url, headers=headers, timeout=20)
            response.encoding = 'utf-8'
            break
        except requests.Timeout:
            if attempt == 1:
                print("⚠️  限额平台 requests 请求超时（已重试），尝试使用浏览器...")
            else:
                print(f"⚠️  限额平台第 {attempt+1} 次请求超时，3 秒后重试...")
                time.sleep(3)
        except requests.ConnectionError as e:
            print(f"⚠️  限额平台连接失败：{e}，尝试使用浏览器...")
            break

    # 如果 requests 失败（超时或连接错误），用 Playwright 回退
    if response is None:
        try:
            with sync_playwright() as play:
                browser = play.chromium.launch()
                context = browser.new_context(
                    user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
                    ignore_https_errors=True,
                )
                page = context.new_page()
                page.goto(url, wait_until="domcontentloaded", timeout=30000)
                time.sleep(random.random() * 2)
                page_html = page.content()
                browser.close()
                response = type('DummyResponse', (), {
                    'text': page_html,
                    'encoding': 'utf-8',
                })()
        except Exception as e:
            print(f"⚠️  限额平台浏览器回退也失败：{e}")
            return

    # 检测 Zscaler/企业网关拦截页
    if "Sorry, company polic" in response.text or "Network app" in response.text:
        print("⚠️  限额平台被网络网关拦截，跳过。")
        return

    soup = BeautifulSoup(response.text, 'html.parser')
    print(f"【{TODAY} 吴中区招标信息】")
    found = False
    results = []

    rows = soup.find_all('tr')

    for row in rows[2:]:
        cols = row.find_all('td')
        if not cols:
            continue
        publish_date = cols[4].get_text(strip=True)
        if publish_date != TODAY:
            continue  # 只要今天的公告
        project_name = cols[1].get_text(strip=True)
        area = cols[2].get_text(strip=True)
        link_tag = cols[1].find('a')
        link = None
        if link_tag:
            href = link_tag.get("href", "").strip()
            if not href.startswith("http"):
                link = "https://www.wzqzjj-wfw.cn:8090/page/" + href.replace("../", "")


        print(f"[{area}]{project_name}\n{link}\n")
        found = True
        results.append({
            "项目名称": project_name,
            "所属区域": area,
            "招标类型": cols[2].get_text(strip=True),
            "资审方式":  cols[3].get_text(strip=True),
            "发布日期": cols[4].get_text(strip=True),
            "链接": link
        })

    if not found:
        print("今天暂无新公告。")
    elif results:
        write_to_excel(results,"xiane_wuzhong_bids")



def fetch_suzhou_gonggong_gov_bids():
    url = "http://218.4.45.172:8086/EpointWebBuilder/JyxxSearchAction.action"
    headers = {
        'User-Agent': 'Mozilla/5.0'
    }

    params = {
        "cmd": "getList1",
        "categorynum": "003",
        "diqu": "苏州市",
        "xmmc": "",
        "zstype": "",
        "zblx": "",
        "starttime": TODAY,
        "endtime": TODAY,
        "siteguid": "7eb5f7f1-9041-43ad-8e13-8fcb82ea831a",
        "pageIndex": 0,
        "pageSize": 500
    }

    response = requests.get(url, headers=headers, params=params, timeout=15)
    response.raise_for_status()
    data = response.json()
    rows = json.loads(data.get("custom", [])).get("Table", [])
    #print(rows)
    print(f"【{TODAY} 苏州市公共招标信息】")
    found = False
    results = []

    for row in rows:
        title = row["title"]
        project_id = row["infoid"]
        area = row["city"]
        release_time = row["postdate"][:10]
        # tomorrow_time = datetime.strptime(release_time, r'%Y-%m-%d') + timedelta(days=1)
        category_num = row["categorynum"]
        category_total = category_num[:6]
        link = f"http://218.4.45.172:8086/jyxx/{category_total}/{category_num}/{release_time.replace('-','')}/{project_id}.html"
        if category_total in true_categories and (area == "吴中区" or area == "苏州市区" or area == "太湖度假区"):
            industry_name = row["jyfl"]
            print(f"[{area}]{title}\n{link}\n")
            found = True
            results.append({
                "地区": area,
                "标题": title,
                "发布日期": release_time,
                "链接": link,
                "行业": industry_name
            })

    if not found:
        print("今日暂无招标。")
    elif results:
        write_to_excel(results, "suzhou_gonggong_gov_bids")


def fetch_suzhou_yangguang_bids():
    """苏州农村阳光招采平台 — 公开 API，无需登录"""
    url = "https://zc.szaee.com/api/public/projectAnn"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
        "Content-Type": "application/json",
        "Referer": "https://zc.szaee.com/",
    }

    all_items = []
    page = 1
    page_size = 50

    # 分页拉取，直到没有今天的公告为止
    while True:
        payload = {
            "page": page,
            "pageSize": page_size,
            "tradeType": 4,        # 4=公开招标
            "projectClass": None,   # 全部类型
            "districtCode": "",
            "keyword": ""
        }
        try:
            response = requests.post(url, headers=headers, json=payload, timeout=20)
            response.raise_for_status()
        except Exception as e:
            print(f"⚠️  阳光采购 失败：{e}")
            return

        data = response.json()
        items = data.get("result", {}).get("items", [])
        if not items:
            break

        # 只保留今天的
        today_items = [it for it in items if it.get("publishTime", "") == TODAY]
        all_items.extend(today_items)

        # 如果当前页最后一条发布日期 < TODAY，说明今天的已全部拉完
        last_date = items[-1].get("publishTime", "")
        if last_date < TODAY:
            break

        total_pages = data.get("result", {}).get("totalPages", 1)
        if page >= total_pages:
            break
        page += 1

    print(f"【{TODAY} 苏州阳光采购招标信息】")
    found = False
    results = []

    for row in all_items:
        title = row.get("announceName", "")
        project_id = row.get("projectId", "")
        area = row.get("areaName", "")
        release_time = row.get("publishTime", "")
        link = f"https://zc.szaee.com/#/projectDetail?id={project_id}"

        # 只关注吴中区和苏州市
        if "吴中区" not in area and "苏州市" not in area:
            continue

        print(f"[{area}]{title}\n{link}\n")
        found = True
        results.append({
            "地区": area,
            "标题": title,
            "发布日期": release_time,
            "链接": link
        })

    if not found:
        print("今日暂无招标。")
    elif results:
        write_to_excel(results, "yangguang_bids")



def write_to_excel(results, sheet_name):
    df = pd.DataFrame(results)
    filename = f"苏州政府采购总表_{TODAY}.xlsx"
    if os.path.exists(filename):
        # 文件存在，追加写入并替换
        with pd.ExcelWriter(filename, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        # 文件不存在，新建文件
        with pd.ExcelWriter(filename, mode='w', engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    print(f"✅ 已保存到文件『{filename}』的 {sheet_name} 中")

def fetch_suzhou_gov_yixiang():
    url = "https://czju.suzhou.gov.cn/zfcg/content/queryContentForCgyx.action"
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
    }
    payload = {
        "channelId": 138,
        "page": 1,
        "rows": 100,
        "title": "",
        "area": "",
        "agency": "",
        "publishBeginTime": "",
        "publishEndTime": ""
    }

    response = requests.post(url, headers=headers, data=payload)
    response.raise_for_status()

    data = response.json()
    rows = data.get("rows", [])

    print(f"【{TODAY} 苏州市采购意向】")
    found = False
    results = []

    for row in rows:
        release_time = row["releaseTime"][:10]
        if release_time == TODAY:
            title = row["title"]
            project_id = row["cpContentId"]
            area = row["area"]
            link = f"https://czju.suzhou.gov.cn/zfcg/html/content/{project_id}.shtml"
            if area == "吴中区" or area == "苏州":
                print(f"[{area}]{title}\n{link}\n")
                found = True
                results.append({
                    "地区": area,
                    "标题": title,
                    "发布日期": release_time,
                    "链接": link
                })

    if not found:
        print("今日暂无意向。")
    elif results:
        write_to_excel(results,"suzhou_gov_yixiang")

def fetch_suzhou_city_college():
    url = "https://www.szcu.edu.cn/zbxx/list.htm"
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
    }
    response = requests.get(url, headers=headers)
    response.encoding = 'utf-8'
    
    print(f"【{TODAY} 苏州城市学院招标信息】")
    found = False
    results = []
    base_url = "https://www.szcu.edu.cn"

    table = None
    news = re.findall('<li class="news n\d+? clearfix">(.+?)</li>', response.text)
    if len(news) == 0:
        print("未找到项目")
        return
    for t in news:
        area = "吴中区"
        project_name = re.findall("title=\\\'(.+?)\'", t)[0]
        publish_date = re.findall('"news_meta">(.+?)</span>',t )[0]
        project_type = "城市学院"
        link = base_url + re.findall('a href=\\\'(.+?)\\\' target=',t)[0]
        if publish_date != TODAY: continue  # 只要今天的公告
        print(f"[{area}]{project_name}\n{link}\n")
        found = True
        results.append({
            "项目名称": project_name,
            "所属区域": area,
            "招标类型": project_type,
            "资审方式": "",# cols[3].get_text(strip=True),
            "发布日期": publish_date,
            "链接": link
        })

    if not found:
        print("今天暂无新公告。")
    elif results:
        write_to_excel(results,"city_college_bids")


def fetch_suzhou_carrer_university():
    url = "https://www.jssvc.edu.cn/zfcg/"
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
    }
    response = requests.get(url, headers=headers)
    response.encoding = 'utf-8'
    
    print(f"【{TODAY} 苏州职业技术大学招标信息】")
    found = False
    results = []
    base_url = "https://www.jssvc.edu.cn/zfcg/"

    table = None
    news = re.findall('<li>(.+?)</li>', re.findall("            <ul>(.+?)</ul>",response.text, re.S)[0], re.S)
    if len(news) == 0:
        print("未找到项目")
        return
    for t in news:
        area = "吴中区"
        project_name = re.findall('title="(.+?)"', t)[0]
        publish_date = re.findall('"time">(.+?)</span>',t )[0]
        project_type = "苏州职大"
        link = base_url + re.findall('a href="(.+?)" target=',t)[0]
        if publish_date != TODAY: continue  # 只要今天的公告
        print(f"[{area}]{project_name}\n{link}\n")
        found = True
        results.append({
            "项目名称": project_name,
            "所属区域": area,
            "招标类型": project_type,
            "资审方式": "",# cols[3].get_text(strip=True),
            "发布日期": publish_date,
            "链接": link
        })

    if not found:
        print("今天暂无新公告。")
    elif results:
        write_to_excel(results,"carrer_university_bids")



def fetch_suzhou_industry_university():
    url = "https://hqzcc.siit.edu.cn/zbxx/list.htm"
    # 该站日期列由 JS 动态渲染，需用浏览器；用 page.content() 避免上下文销毁问题
    with sync_playwright() as play:
        browser = play.chromium.launch()
        context = browser.new_context(ignore_https_errors=True)
        page = context.new_page()
        page.goto(url)
        page.wait_for_load_state("networkidle", timeout=20000)
        time.sleep(random.random() * 2)
        pageHtml = page.content()
        browser.close()
    tables = re.findall('<td align="left">(.+?)</div></td>', pageHtml, re.S)
    results = []
    
    print(f"【{TODAY} 苏州工业职业技术学院招标信息】")
    found = False

    for t in tables:
        if len(t) == 188 or len(t) == 187:continue
        project_name = re.findall('title="(.+?)"', t)[0]
        area = "苏州工职大"
        publish_date = re.findall('"white-space:nowrap">(.+)', t)[0]
        link = "https://hqzcc.siit.edu.cn" + re.findall('a href="(.+?)"', t)[0]

        if publish_date != TODAY: continue
        results.append({
        "项目名称": project_name,
        "所属区域": area,
        "发布日期": publish_date,
        "链接": link
        })
        found = True
    

    if not found:
        print("今天暂无新公告。")
    elif results:
        write_to_excel(results,"industry_university_bids")


# def run_all_fetchers():
#     with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
#         futures = [executor.submit(fetch_suzhou_gov_bids),
#                    executor.submit(fetch_jiangsu_gov_bids),
#                    executor.submit(fetch_wuzhong_gov_bids),
#                    executor.submit(fetch_suzhou_gonggong_gov_bids)]
#
#         for future in concurrent.futures.as_completed(futures):
#             try:
#                 future.result()
#             except Exception as e:
#                 print(f"任务异常：{e}")




def _run(fn, label):
    try:
        fn()
    except Exception as e:
        print(f"⚠️  {label} 失败：{e}")


if __name__ == '__main__':
    _run(fetch_suzhou_gov_yixiang,       "政采网意向")
    _run(fetch_suzhou_gov_bids,          "政采网")
    _run(fetch_suzhou_gonggong_gov_bids, "公共交易资源网")
    _run(fetch_jiangsu_gov_bids,         "江苏省招投标")
    _run(fetch_xiane_gov_bids,           "限额平台")
    _run(fetch_suzhou_city_college,      "城市学院")
    _run(fetch_suzhou_carrer_university, "职业技术学校")
    _run(fetch_suzhou_industry_university, "工职院")
    _run(fetch_suzhou_yangguang_bids,    "阳光采购")