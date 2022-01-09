# coding='utf-8'
"""
这里是直接给到某个商品的评论的json页面，去爬取评论，
可以修改一下：给定商品页面，获取相关参数，构成json页面，去获取数据

这里获得的是数据的所有评价
对于不同分类的评价，可利用json网址中的score值来进行修改

json文件参数
productId=5561746 产品id--利用这个得到网页json
score=0/1/2/3/4/5 产品评价 全部评价/差评/中评/好评/追评

保存的文件名如果有需要，可以写成comments+uid的保存形式

https://blog.csdn.net/qq_24994275/article/details/116177337?spm=1001.2101.3001.6650.1&utm_medium=distribute.pc_relevant.none-task-blog-2%7Edefault%7ECTRLIST%7Edefault-1.no_search_link&depth_1-utm_source=distribute.pc_relevant.none-task-blog-2%7Edefault%7ECTRLIST%7Edefault-1.no_search_link&utm_relevant_index=2
https://blog.csdn.net/m0_45827246/article/details/121628669?ops_request_misc=%257B%2522request%255Fid%2522%253A%2522164155843716780261977858%2522%252C%2522scm%2522%253A%252220140713.130102334.pc%255Fall.%2522%257D&request_id=164155843716780261977858&biz_id=0&utm_medium=distribute.pc_search_result.none-task-blog-2~all~first_rank_ecpm_v1~times_rank-1-121628669.first_rank_v2_pc_rank_v29&utm_term=%E4%BA%AC%E4%B8%9C%E5%95%86%E5%93%81%E8%AF%84%E8%AE%BA&spm=1018.2226.3001.4187

添加
读取文件中的网址去进行爬取
保存的文件命名 更改-->uid.xls
"""

import urllib
import requests
import json
import time
import random
import xlwt
import xlutils.copy
import xlrd


# UserAgent_list = ['Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.83 Safari/537.36',
#                   'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/72.0.3626.109 Safari/537.36',
#                   'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.86 Safari/537.36',
#                   'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.75 Safari/537.36']
# Cookie = 'yh_language=zh-cn;' \
#          ' yh_cudid=z5500dab7d08c659b84e44fc969d44bfa;' \
#          ' _yasvd=1407875911; _ga=GA1.2.917476022.1606226670;' \
#          ' _gid=GA1.2.324156849.1606226670; Hm_lvt_7b29c77c4002f91220be5eaf2ce387fe=1606226670,1606228424,1606289593;' \
#          ' Hm_lvt_cba6f2719081a006e181cf17fa40ad05=1606289615;' \
#          ' Hm_lpvt_cba6f2719081a006e181cf17fa40ad05=1606289626;' \
#          ' Hm_lpvt_7b29c77c4002f91220be5eaf2ce387fe=1606290146'
# headers = {
#             'User-Agent': random.choice(UserAgent_list),  # 使用random模块中的choices()方法随机从列表中提取出一个内容
#             'Cookie': Cookie
#         }
headers= {
        "User-Agent": "Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.198 Mobile Safari/537.36"
    }
#由商品url得到json文件地址
def pre(url):
    url=url.lstrip('https://item.jd.com/')
    index=url.find('.')
    uid=url[:index]
    #print(type(uid))
    url = 'https://club.jd.com/comment/productPageComments.action?callback=fetchJSON_comment98&productId={uid}&score=0&sortType=5&page=1&pageSize=10&isShadowSku=0&fold=1'.format(uid=uid)
    #print(url)

    return url

def start(page,goodsurl):
    # 获取URL
    #score 评价等级 page=0 第一页 producitid 商品类别

    url=pre(goodsurl)

    #json_url
    #url = 'https://club.jd.com/comment/productPageComments.action?callback=fetchJSON_comment98&productId=5561746&score=0&sortType=5&page=1&pageSize=10&isShadowSku=0&fold=1'

    # url = 'https://club.jd.com/comment/productPageComments.action?&productId=100016647456&score=0&sortType=5&page=0&pageSize=10&isShadowSku=0&fold=1'
    time.sleep(2)
    try:
        req = urllib.request.Request(url, headers=headers)
        response = urllib.request.urlopen(req, timeout=15)
        html = response.read().decode('GBK', errors='ignore')
        #print(html)
        #需要去除外面包裹的fetchJSON_comment98()展现json格式,进行后续处理
        html=html.lstrip('fetchJSON_comment98(')
        html=html.rstrip(');')

        data=json.loads(html)

        # test = requests.get(url=url, headers= headers)
        # data = json.loads(test.text)

        return data
    except Exception as e:
        print("网页{}获取失败".format(url),"原因",e)
    #try except只用在这写，外面start循环调用，就会自动返回异常，结束这一次循环，start外不用再做处理

    # 解析页面
def parse(data):
    try:
        if data==None:
            return None
        items = data['comments']
        if(items==None):
            return None
        for i in items:
            yield (
                i['nickname'],#用户名
                i['id'], #用户id
                i['content'],#内容
                i['creationTime']#时间
            )
    except Exception as e:
        print(e)

def excel(items,goodsname):
    #第一次写入
    newTable=goodsname+'.xls'
    #newTable = "comments.xls"#创建文件
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

def another(items, j,goodsname):#如果不是第一次写入 以后的就是追加数据了 需要另一个函数

    index = (j-1) * 10 + 1#这里是 每次写入都从11 21 31..等开始 所以我才传入数据 代表着从哪里开始写入

    newTable = goodsname + '.xls'

    data = xlrd.open_workbook(newTable)
    ws = xlutils.copy.copy(data)
    # 进入表
    table = ws.get_sheet(0)

    for test in items:

        for i in range(0, 4):#跟excel同理
            print(test[i])

            table.write(index, i, test[i])  # 只要分配好 自己塞入
        print('_______________________')

        index += 1
        ws.save(newTable)



def main():
    import xlrd
    xl = xlrd.open_workbook(r'.\京东渔具相关商品信息.xlsx')
    table = xl.sheets()[0]
    url_col = table.col_values(5)  # 渔具相关的网址
    # print(url_col[1:])
    goods_col = table.col_values(3)
    for num in range(1,len(url_col)):
        goodsurl=url_col[num]
        goodsname=goods_col[num]
        # print(goodsname)
        # print(goodsurl)

        from openpyxl import Workbook
        workbook = Workbook()
        tablename=goodsname+'.xls'
        workbook.save(tablename)

        j = 1#页面数
        judge = True#判断写入是否为第一次

        for i in range(0, 100):
            #这边只爬取100页，可以设置成如果为空就停止爬取，获得到的页面
            #不能这么修改，因为可能有的json文件格式有过改变和之前不一样了发生变化，不能沿用现在的模板
            #eg:
            #https://club.jd.com/comment/productPageComments.action?callback=fetchJSON_comment98&productId=5561746&score=0&sortType=5&page=500&pageSize=10&isShadowSku=0&fold=1
            #https://club.jd.com/comment/productPageComments.action?callback=fetchJSON_comment98&productId=5561746&score=0&sortType=5&page=1&pageSize=10&isShadowSku=0&fold=1
            #html的内容可以得到，但是json文件格式改变，得到的可能为空
            #先设置成100页
            time.sleep(1.5)
            #记得time反爬 其实我在爬取的时候没有使用代理ip也没给我封 不过就当这是个习惯吧
            first = start(j,goodsurl)
            test = parse(first)
            if(test==None):
                break
            if judge:
                excel(test,goodsname)
                judge = False
            else:
                another(test, j,goodsname)
            print('第' + str(j) + '页抓取完毕\n')
            j = j + 1


if __name__ == '__main__':
    main()
    #这个代码仅为全部数据下的评论而已 中差评等需要修改score
