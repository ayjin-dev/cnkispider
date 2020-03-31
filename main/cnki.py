import re
import time
import random
import requests
import urllib.parse
import xlwt,xlrd
import xlutils.copy
from lxml import etree
from time import sleep

#获得excel数据信息
def readline():
    wb = xlrd.open_workbook(filename,formatting_info=True)  #打开excel，保留文件格式
    sheet1 = wb.sheet_by_index(0)  #获取第一张表
    nrows = sheet1.nrows  #获取总行数
    ncols = sheet1.ncols
    return nrows

#写入Excel
def write(title,msg,link,abstract):
    data = xlrd.open_workbook(filename)
    ws = xlutils.copy.copy(data) #复制之前表里存在的数据
    table=ws.get_sheet(0)
    nownrows = readline()
    table.write(nownrows, 0, label=title)  #最后一行追加数据
    table.write(nownrows, 1, label=msg)
    table.write(nownrows, 2, label=link)
    table.write(nownrows, 3, label=abstract)
    ws.save(filename)  #保存的有旧数据和新数据

#设置随机请求头
def set_headers():
    uapools = [
        'User-Agent:Mozilla/5.0 (Windows NT 6.2; WOW64; rv:21.0) Gecko/20100101 Firefox/21.0',
        'Mozilla/5.0 (Windows NT 6.2; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/27.0.1453.94 Safari/537.36',
        'User-Agent:Mozilla/5.0 (Macintosh; Intel Mac OS X 10_7_2) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/27.0.1453.93 Safari/537.36',
        'User-Agent:Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/535.11 (KHTML, like Gecko)Ubuntu/11.10 Chromium/27.0.1453.93 Chrome/27.0.1453.93 Safari/537.36'
    ]
    user_agent=random.choice(uapools)
    headers={
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9",
        "Accept-Encoding": "gzip, deflate",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
        "Host": "search.cnki.net",
        "Proxy-Connection": "keep-alive",
        "Upgrade-Insecure-Requests": '1',
        "User-Agent": user_agent
    }
    return headers

#创建一个代理ip池
def set_ippool():
    #此处需要修改，加入你的代理ip api
    api = ''
    ips = requests.get(api).json()
    ip_pools = []
    data = ips['data']
    for i in range(len(data)):
        ip = data[i]['ip']
        port = data[i]['port']
        ip_ = str(ip) + ':' + str(port)
        ip_pools.append(ip_)
        ip_ = ''
    return ip_pools

#获得该关键词论文几页
def key_msg():
    headers=set_headers()
    page=requests.get(url,headers=headers)
    page=page.text
    element=etree.HTML(page)

    #总相关记录 13019条记录 894页
    about_it=element.xpath("//p[@id['page']]//span[@class='page-sum']//text()")[0]
    pat1='共找到相关记录(.*?)条'
    sum_record=re.compile(pat1).findall(about_it)
    sum_record=int(sum_record[0])
    print('涉及该关键词的相关记录共%d条' %sum_record)
    page_count=sum_record//15
    print('涉及该关键词的论文共%d页' %page_count)
    return page_count

#解析当前页面
def xpath_get(i,ip_pools):
    print('这是%d号' %i)
    #设置请求头
    headers=set_headers()
    #设置代理ip
    ip=random.choice(ip_pools)
    proxies={
        'http':'http://'+ip
    }
    #发起请求
    newurl=url+'&rank=relevant&cluster=all&val=&p='+str(i*15)
    page = requests.get(newurl, headers=headers,proxies=proxies)
    page = page.text
    element = etree.HTML(page)
    print('当前爬取第%d页,使用ip为%s' %((i+1),ip))

    #获得所有文章标签
    all_content=element.xpath('//div[@class="wz_content"]')
    #对标签进行遍历
    for content in all_content:
        #标题 基于蚁群算法的路径规划问题研究
        titles=content.xpath('.//h3//a//text()')
        title=''
        for st in titles:
            title+=st

        #详细信息 东南大学 硕士论文 2018年 下载次数（220）| 被引次数（）
        publishs=content.xpath('.//span[@class="year-count"]//text()')
        msg = ''
        for i in publishs:
            if i == '\r\n                      ':
                pass
            else:
                i=" ".join(i.split())
                msg+=i+' '

        #获得链接中所有信息并组件
        all_url=content.xpath('.//a/@href')

        #dbcode
        pat2 = '.dbcode=(.*)&year'
        dbcode=re.compile(pat2).findall(all_url[1])
        dbcode=dbcode[0]


        #dbname
        pat3='&dbcode=.*?&year=(.*).dflag'
        year=re.compile(pat3).findall(all_url[1])
        year=year[0]
        dbname=dbcode+year

        #Filename
        pat4='filename=(.*?)&dbcode'
        filename=re.compile(pat4).findall(all_url[1])
        filename=filename[0]

        #compileurl
        content_url='https://www.cnki.net/kcms/detail/detail.aspx?&dbcode={0}&dbName={1}&FileName={2}&v=&uid='.format(dbcode,dbname,filename)

        content_url='未知链接'
        #摘要
        abstract=''
        abstracts=content.xpath('.//span[@class="text"]//text()')
        for ab in abstracts:
            abstract+=ab

        print('title:',title)
        print('msg:',msg)
        print('link:',content_url)
        print('abstarct:',abstract)
        write(title,msg,content_url,abstract)
    print('写入完成')
    print('当前页已经爬取完成\n')

#启动函数
def main():
    start=time.time()
    #创建代理ip池
    ip_pools=set_ippool()
    #获得页数
    page_count=key_msg()
    for i in range(0,page_count+1):
        try:
            xpath_get(i,ip_pools)
        except:
            xpath_get(i,ip_pools)
    end = time.time()
    print('所有页已经完成')
    print('用时',end-start)

if __name__ == '__main__':
    y = input('请输入关键词：')
    x = urllib.parse.quote(y)
    url = 'http://search.cnki.net/Search.aspx?q=' + x
    work_book = xlwt.Workbook(encoding='utf-8')
    sheet = work_book.add_sheet('test')
    sheet.write(0, 0, '标题')
    sheet.write(0, 1, '基本信息')
    sheet.write(0, 2, '链接')
    sheet.write(0, 3, '摘要')
    filename = y + ".xls"
    work_book.save(filename)
    main()
