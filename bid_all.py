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

# done 爬不到东西了，要改
def fetch_jiangsu_gov_bids():
    url = "https://api.jszbtb.com/DataGatewayApi/PublishBulletins?bulletinType=1&industryCode=&regionCode=320500&startTime="+ TODAY + "+00:00:00&endTime=" + TODAY +"+23:59:59&keyword=&currentPage=1&pageSize=20"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36 Edg/126.0.0.0",
        "Content-Type": "application/x-www-form-urlencoded",
        "Host":"api.jszbtb.com",
        "Cookie":"Hm_lvt_f5b9bdde19559e7d521049ec442e5faf=1762414353,1763100732,1763521207,1764213894",
    }
    start_time = TODAY + "+00:00:00"
    end_time = TODAY + "+23:59:59"
    # today = "2025-06-30"
    params  = {
        "bulletinType": 1,
        "industryCode": "",
        "regionCode": 320500,
        "startTime": start_time,
        "endTime": end_time,
        "keyword": "",
        "currentPage": 1,
        "pageSize": 20
    }

    response = requests.get(url, headers=headers, params=params, timeout=15)
    response.raise_for_status()

    if not response.text.strip().startswith("{"):
        print(f"⚠️  江苏省招投标接口返回非 JSON（可能被封锁），跳过。")
        return

    data = response.json()
    rows = data.get("data", {}).get("data", [])

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
    # url = "https://api.jszbtb.com/DataGatewayApi/PublishBulletins?bulletinType=1&industryCode=&regionCode=320500&startTime=2025-12-03+00:00:00&endTime=2025-12-03+23:59:59&keyword=&currentPage=1&pageSize=20"
    headers = {
        'User-Agent': 'Mozilla/5.0'
    }

    response = requests.get(url, headers=headers, timeout=15)
    response.encoding = 'utf-8'
    soup = BeautifulSoup(response.text, 'html.parser')

    # 检测 Zscaler/企业网关拦截页
    if "Sorry, company polic" in response.text or "Network app" in response.text:
        print("⚠️  限额平台被网络网关拦截，跳过。")
        return

    # with sync_playwright() as play:
    #     broswer = play.chromium.launch()
    #     context = broswer.new_context()
    #     page = context.new_page()
    #     page.goto(url)
    #     page.wait_for_load_state("load")
    #     time.sleep(random.random()*2)

    #     pageHtml = page.query_selector("*").inner_html()
    #     broswer.close()
    # soup = BeautifulSoup(pageHtml, 'html.parser')
#
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
    # 该站经 zscloud 企业网关做 SAML SSO，必须带登录 cookie 才能访问
    # 首次运行前请先执行 save_cookies.py 保存 cookie
    page_url = 'https://zc.szaee.com/#/project?tradeType=4&projectClass='
    cookie_file = os.path.join(os.path.dirname(__file__), "cookies_yangguang.json")

    if not os.path.exists(cookie_file):
        print("⚠️  阳光采购：未找到 cookies_yangguang.json，请先运行 save_cookies.py 登录并保存 cookie。")
        return

    with open(cookie_file, encoding="utf-8") as f:
        saved_cookies = json.load(f)

    print(f"【{TODAY} 苏州阳光采购招标信息】")
    found = False
    results = []
    captured = []

    def handle_response(response):
        if 'projectAnn' in response.url:
            try:
                captured.append(response.json())
            except Exception:
                pass

    with sync_playwright() as play:
        browser = play.chromium.launch()
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
            ignore_https_errors=True,
        )
        context.add_cookies(saved_cookies)
        page = context.new_page()
        page.on("response", handle_response)
        page.goto(page_url, wait_until="domcontentloaded", timeout=30000)
        time.sleep(5)
        browser.close()

    for data in captured:
        rows = (data.get("data") or {}).get("list") or data.get("rows") or []
        if not rows:
            print(f"[debug] 阳光采购原始响应：{str(data)[:300]}")
            continue
        for row in rows:
            release_time = (row.get("releaseTime") or row.get("RELEASE_TIME") or row.get("publishTime") or "")[:10]
            if release_time != TODAY:
                continue
            title = row.get("projectName") or row.get("TITLE") or row.get("name") or ""
            project_id = row.get("annId") or row.get("id") or row.get("PROJECTID") or ""
            area = row.get("districtName") or row.get("AREA") or row.get("district") or ""
            link = f"https://zc.szaee.com/#/projectDetail?id={project_id}"
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