# coding='utf-8'
import urllib
import requests
import json
import time
import random
import xlwt
import xlutils.copy
import xlrd


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

#由商品url得到json文件地址
def pre(url,page):
    url=url.lstrip('https://item.jd.com/')
    index=url.find('.')
    uid=url[:index]
    #print(type(uid))
    url = 'https://club.jd.com/comment/productPageComments.action?callback=fetchJSON_comment98&productId={uid}&score=0&sortType=5&page={page}&pageSize=10&isShadowSku=0&fold=1'.format(uid=uid,page=page)
    #print(url)

    return url

def start(page):
    # 获取URL
    # score 评价等级 page=0 第一页 producitid 商品类别

    goodsurl = 'https://item.jd.com/11789467495.html'
    url = pre(goodsurl,page)

    # json_url
    # url = 'https://club.jd.com/comment/productPageComments.action?callback=fetchJSON_comment98&productId=5561746&score=0&sortType=5&page=1&pageSize=10&isShadowSku=0&fold=1'

    # url = 'https://club.jd.com/comment/productPageComments.action?&productId=100016647456&score=0&sortType=5&page=0&pageSize=10&isShadowSku=0&fold=1'
    time.sleep(2)
    try:
        req = urllib.request.Request(url, headers=headers)
        response = urllib.request.urlopen(req, timeout=15)
        html = response.read().decode('GBK', errors='ignore')
        # print(html)
        # 需要去除外面包裹的fetchJSON_comment98()展现json格式,进行后续处理
        html = html.lstrip('fetchJSON_comment98(')
        html = html.rstrip(');')

        data = json.loads(html)

        # test = requests.get(url=url, headers= headers)
        # data = json.loads(test.text)

        return data
    except Exception as e:
        print("网页{}获取失败".format(url), "原因", e)
    # try except只用在这写，外面start循环调用，就会自动返回异常，结束这一次循环，start外不用再做处理



    # 解析页面
def parse(data):
    try:
        items = data['comments']
        for i in items:
            yield (
                i['nickname'],#用户名自
                i['id'], #用户id
                i['content'],#内容
                i['creationTime']#时间
            )
    except Exception as e:
        print(e)

def excel(items):
    #第一次写入
    newTable = "test.xls"#创建文件
    wb = xlwt.Workbook("encoding='utf-8")

    ws = wb.add_sheet('sheet1')#创建表
    headDate = ['用户名', '用户id', '评论内容','评论时间']#定义标题
    for i in range(0,4):#for循环遍历写入
        ws.write(0, i, headDate[i], xlwt.easyxf('font: bold on'))

    index = 1#行数

    for data in items:#items是十条数据 data是其中一条（一条下有三个内容）
        for i in range(0,4):#列数

            print(data[i])
            ws.write(index, i, data[i])#行 列 数据（一条一条自己写入）
        print('______________________')
        index += 1#等上一行写完了 在继续追加行数
        wb.save(newTable)

def another(items, j):#如果不是第一次写入 以后的就是追加数据了 需要另一个函数

    index = (j-1) * 10 + 1#这里是 每次写入都从11 21 31..等开始 所以我才传入数据 代表着从哪里开始写入

    data = xlrd.open_workbook('test.xls')
    ws = xlutils.copy.copy(data)
    # 进入表
    table = ws.get_sheet(0)

    for test in items:

        for i in range(0, 4):#跟excel同理
            print(test[i])

            table.write(index, i, test[i])  # 只要分配好 自己塞入
        print('_______________________')

        index += 1
        ws.save('test.xls')



def main():
    j = 1#页面数
    judge = True#判断写入是否为第一次

    from openpyxl import Workbook
    workbook = Workbook()
    tablename = 'test.xls'
    workbook.save(tablename)


    for i in range(0, 100):
        time.sleep(1.5)
        #记得time反爬 其实我在爬取的时候没有使用代理ip也没给我封 不过就当这是个习惯吧
        first = start(j)
        if(not first or first['comments']==[]):
            break
        test = parse(first)

        if judge:
            excel(test)
            judge = False
        else:
            another(test, j)
        print('第' + str(j) + '页抓取完毕\n')
        j = j + 1


if __name__ == '__main__':
    main()
    #这个代码仅为全部数据下的评论而已 中差评等需要修改score！

