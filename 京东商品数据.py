# coding='utf-8'
"""
用于获取搜索页面的各种信息,方便评论的爬取
"""
import time
from lxml import etree
import os
import requests
import csv
import random
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from openpyxl import Workbook
import urllib


UserAgent_list = ['Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.83 Safari/537.36',
                  'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/72.0.3626.109 Safari/537.36',
                  'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.86 Safari/537.36',
                  'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.75 Safari/537.36']
Cookie = 'yh_language=zh-cn;' \
         ' yh_cudid=z5500dab7d08c659b84e44fc969d44bfa;' \
         ' _yasvd=1407875911; _ga=GA1.2.917476022.1606226670;' \
         ' _gid=GA1.2.324156849.1606226670; Hm_lvt_7b29c77c4002f91220be5eaf2ce387fe=1606226670,1606228424,1606289593;' \
         ' Hm_lvt_cba6f2719081a006e181cf17fa40ad05=1606289615;' \
         ' Hm_lpvt_cba6f2719081a006e181cf17fa40ad05=1606289626;' \
         ' Hm_lpvt_7b29c77c4002f91220be5eaf2ce387fe=1606290146'
headers = {
            'User-Agent': random.choice(UserAgent_list),  # 使用random模块中的choices()方法随机从列表中提取出一个内容
            'Cookie': Cookie
        }

wb = Workbook()
sheet = wb.active
sheet['A1'] = 'name'
sheet['B1'] = 'price'
sheet['C1'] = 'shop'
sheet['D1'] = 'sku'
sheet['E1'] = 'icons'
sheet['F1'] = 'detail_url'



options = webdriver.ChromeOptions()
# 不加载图片
options.add_experimental_option('prefs', {'profile.managed_default_content_settings.images': 2})
driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 60)  # 设置等待时间


def search(keyword):
    try:
        input = wait.until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "#key"))
        )  # 等到搜索框加载出来
        submit = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "#search > div > div.form > button"))
        )  # 等到搜索按钮可以被点击
        input[0].send_keys(keyword)  # 向搜索框内输入关键词
        submit.click()  # 点击
        wait.until(
            EC.presence_of_all_elements_located(
                (By.CSS_SELECTOR, '#J_bottomPage > span.p-skip > em:nth-child(1) > b')
            )
        )
        total_page = driver.find_element_by_xpath('//*[@id="J_bottomPage"]/span[2]/em[1]/b').text
        return int(total_page)
    except TimeoutError:
        search(keyword)


def get_data(html):
    # 创建etree对象
    tree = etree.HTML(html)
    # titles= tree.xpath('//div[@id="J_searchWrap"]//div[@class="gl-i-wrap"]//div[@class="p-name p-name-type-2"]//em')
    # print(len(titles))
    # for title in titles:
    #     print(title.xpath('string(.)').strip())
    lis=tree.xpath('//ul[@class="gl-warp clearfix"]/li')
    for li in lis:
        try:
            title_r=li.xpath('.//div[@class="p-name p-name-type-2"]//em')
            #print(title.xpath('string(.)').strip())
            title=title_r[0].xpath('string(.)').strip()
            title=title.replace('\n','')
            #print(title.replace('\n',''))
            price = li.xpath('.//div[@class="p-price"]//i/text()')[0].strip()  # 价格
            #print(price)
            data_sku = li.xpath('./@data-sku')[0].strip() # 商品唯一id
            #print(data_sku)
            #comment = li.xpath('.//div[@class="p-commit"]//a')  # 评论数
            shop_name = li.xpath('.//div[@class="p-shop"]//a//text()')[0].strip() # 商铺名字
            #print(shop_name)
            icons = li.xpath('.//div[@class="p-icons"]/i/text()') # 备注
            #comment = comment[0] if comment != [] else ''
            icons_n = ''
            for x in icons:
                icons_n = icons_n+ x.replace('\n','')
                icons_n=icons_n+';'
            #print(icons_n)
            detail_url = li.xpath('.//div[@class="p-name p-name-type-2"]/a/@href')[0]  # 详情页网址
            detail_url = 'https:' + detail_url
            #print(detail_url)
            item = [title, price, shop_name,data_sku, icons_n, detail_url]
            print(item)
            sheet.append(item)
        except Exception as e:
            print("错误原因：", e)


def main():
    url_main = 'https://www.jd.com/'
    keyword = input('请输入商品名称:')  # 搜索关键词
    driver.get(url=url_main)
    page = search(keyword)
    j = 1
    for i in range(3, page*2, 2):
        if j == 1:
            url = 'https://search.jd.com/Search?keyword={}&page={}&s={}&click=0'.format(keyword, i, j)
        else:
            url = 'https://search.jd.com/Search?keyword={}&page={}&s={}&click=0'.format(keyword, i, (j-1)*50)
        driver.get(url)
        time.sleep(1)
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight)")  # 下滑到底部
        time.sleep(3)
        driver.implicitly_wait(20)
        wait.until(
            EC.presence_of_all_elements_located((By.XPATH, '//*[@id="J_goodsList"]/ul/li[last()]'))
        )
        html = driver.page_source
        get_data(html)
        time.sleep(1)
        print(f'正在爬取第{j}页')
        j += 1
    wb.save('京东{}相关商品信息.xlsx'.format(keyword))

if __name__ == '__main__':
    main()
