"""
Created on 2018/4/17
@Author: Jeff Yang
"""

import requests
import re
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
import xlwt

workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('鸿茅', cell_overwrite_ok=True)

worksheet.write(0, 0, label='药品广告批准文号')
worksheet.write(0, 1, label='单位名称')
worksheet.write(0, 2, label='地址')
worksheet.write(0, 3, label='邮政编码')
worksheet.write(0, 4, label='通用名称')
worksheet.write(0, 5, label='商标名称')
worksheet.write(0, 6, label='处方分类')
worksheet.write(0, 7, label='广告类别')
worksheet.write(0, 8, label='时长')
worksheet.write(0, 9, label='广告有效期')
worksheet.write(0, 10, label='广告发布内容')
worksheet.write(0, 11, label='批准文号')

chrome_options = Options()
chrome_options.add_argument('--headless')
chrome_options.add_argument('--disable-gpu')
browser = webdriver.Chrome(chrome_options=chrome_options)

url_pattern = re.compile('<a target="_blank" href="(.*?)">鸿茅药酒</a>', re.S)
detail_pattern = re.compile('药品广告批准文号.*?width="380">(.*?)</td>.*?'
                            + '单位名称.*?width="381">(.*?)</td>.*?'
                            + '地址.*?width="380">(.*?)</td>.*?'
                            + '邮政编码.*?width="381">(.*?)</td>.*?'
                            + '通用名称.*?width="380">(.*?)</td>.*?'
                            + '商标名称.*?width="380">(.*?)</td>.*?'
                            + '处方分类.*?width="381">(.*?)</td>.*?'
                            + '广告类别.*?width="380">(.*?)</td>.*?'
                            + '时长.*?width="381">(.*?)</td>.*?'
                            + '广告有效期.*?width="380">(.*?)</td>.*?'
                            + '广告发布内容.*?width="381"><a href="\.\.(.*?)" target="_blank">.*?'
                            + '批准文号.*?width="380">(.*?)</td>', re.S)

start = 1
for i in range(1, 81):
    url = "http://app2.sfda.gov.cn/datasearchp/all.do?page=" + str(i) \
          + "&name=%E9%B8%BF%E8%8C%85%E8%8D%AF%E9%85%92&tableName=TABLE39&formRender=cx&searchcx=%E9%B8%BF%E8%8C%85%E8%8D%AF%E9%85%92&paramter0=&paramter1=&paramter2="
    wait = WebDriverWait(browser, 10)
    # browser.set_window_size(1400, 900)  # 设置窗口大小

    browser.get(url)
    html = browser.page_source
    urls = re.findall(url_pattern, html)
    print("正在爬取第", i, "页，此页有", len(urls), "条")
    for index, url in enumerate(urls):
        print("     第", index, "条")
        detail_url = "http://app2.sfda.gov.cn" + url
        browser.get(detail_url.replace('amp;', ''))
        html = browser.page_source
        # print(html)
        items = re.findall(detail_pattern, html)
        # print(html)
        for item in items:
            worksheet.write(index + start, 0, label=item[0])
            worksheet.write(index + start, 1, label=item[1])
            worksheet.write(index + start, 2, label=item[2])
            worksheet.write(index + start, 3, label=item[3])
            worksheet.write(index + start, 4, label=item[4])
            worksheet.write(index + start, 5, label=item[5])
            worksheet.write(index + start, 6, label=item[6])
            worksheet.write(index + start, 7, label=item[7])
            worksheet.write(index + start, 8, label=item[8])
            worksheet.write(index + start, 9, label=item[9])
            worksheet.write(index + start, 10, label="http://app2.sfda.gov.cn" + item[10])
            worksheet.write(index + start, 11, label=item[11])
    start += len(urls)

workbook.save('result2.xls')
print("Done!")
