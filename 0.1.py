import json
import re
import xlrd
import xlwt
from pypinyin import lazy_pinyin, pinyin

import requests
from bs4 import BeautifulSoup

from tqdm import tqdm

corona_virus = []


class ConronVirusSpider(object):
    def __init__(self):
        self.url = 'https://ncov.dxy.cn/ncovh5/view/pneumonia'

    def get_content_from_url(self, url):
        response = requests.get(url)
        return response.content.decode()

    def parse_html(self, html):
        soup = BeautifulSoup(html, 'lxml')
        script = soup.find(id="getAreaStat")
        text = script.text
        json_str = re.findall(r'\[.+\]', text)[0]
        data = json.loads(json_str)
        return data

    def save(self, data, path):
        with open(path, 'w', encoding="utf-8") as fp1:
            json.dump(data, fp1, ensure_ascii=False)

    def crawl_lastday_conron_virus(self):
        # 发送请求获取首页内容
        home_page = self.get_content_from_url(self.url)
        # 解析首页内容获取最近一天的数据
        lastday_conron_virus = self.parse_html(home_page)
        # 保存数据
        self.save(lastday_conron_virus, "venv/last_day_corona_virus.json")

    def crawl_conron_virus(self):
        with open('C:\Users\h1526\Desktop\疫情可视化\venv', encoding="utf-8") as fp2:
            last_day_corona_virus = json.load(fp2)
        for province in tqdm(last_day_corona_virus, "采集各省疫情数据"):
            statistics_data_url = province['statisticsData']
            statistics_data_json_str = self.get_content_from_url(statistics_data_url)
            statistics_data = json.loads(statistics_data_json_str)["data"]
            if province["provinceShortName"] == "山西":
                province["provinceShortName"] = "三晋"

            lp = lazy_pinyin(province["provinceShortName"])


            for one_day in statistics_data:
                # one_day["provinceName"]=province["provinceName"]
                if len(lp) == 2:
                    one_day["provinceShortName"] = lp[0] + lp[1]
                if len(lp) == 3:
                    one_day["provinceShortName"] = lp[0] + lp[1] + lp[2]


            corona_virus.extend(statistics_data)
        self.save(corona_virus, 'venv/corona_virus.json')

    def run(self):
        self.crawl_conron_virus()


def json_excel(data):
    # 创建excel工作表
    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet('skiing')
    # 写表头
    for index, val in enumerate(data[0].keys()):
        worksheet.write(0, index, label=val)

    # 写数据
    for row, list_item in enumerate(data):
        row += 1
        col = 0
        for key, value in list_item.items():
            worksheet.write(row, col, value)
            col += 1

    # 保存
    workbook.save('./excel_file/20220207_143920-dong.xls')


if __name__ == '__main__':
    spider = ConronVirusSpider()
    spider.run()
