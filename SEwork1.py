import re
import  requests
import xlsxwriter
import  time

def getIntoExcel(html):
    global com_count
    com_name = re.findall(r'"raw_title":"(.*?)"', html)      #名称
    com_price = re.findall(r'"view_price":"(.*?)"', html)   #价格~
    com_loc = re.findall(r'"item_loc":"(.*?)"', html)   #地区
    com_num = re.findall(r'"view_sales":"(.*?)"', html)   #销量

    com_tab = []    #excel表
    for i in range(len(com_name)):
        try:
            com_tab.append((com_name[i],com_price[i],com_loc[i],com_num[i]))
        except IndexError:
            break
    for temp in range(len(com_tab)):
        worksheet.write(com_count + temp + 1, 0, com_tab[temp][0])
        worksheet.write(com_count + temp + 1, 1, com_tab[temp][1])
        worksheet.write(com_count + temp + 1, 2, com_tab[temp][2])
        worksheet.write(com_count + temp + 1, 3, com_tab[temp][3])
    com_count = com_count +len(com_tab)     #爬取总量
    return print("已完成")


def getUrls(pro_name, page):   #q要查询的商品名称，page是要爬取的页数
    url = "https://s.taobao.com/search?q=" + pro_name + "&imgfile=&commend=all&ssid=s5-e&search_type=item&sourceId=tb.index&spm" \
                                                 "=a21bo.2017.201856-taobao-item.1&ie=utf8&initiative_id=tbindexz_20170306 "
    urls = []
    urls.append(url)
    if page == 1:
        return urls
    for i in range(1, page ):
        url = "https://s.taobao.com/search?q="+ pro_name + "&commend=all&ssid=s5-e&search_type=item" \
              "&sourceId=tb.index&spm=a21bo.2017.201856-taobao-item.1&ie=utf8&initiative_id=tbindexz_20170306" \
              "&bcoffset=3&ntoffset=3&p4ppushleft=1%2C48&s=" + str(
            i * 44)
        urls.append(url)
    return urls


def getHtml(url):
    r = requests.get(url,headers =headers)
    r.raise_for_status()
    r.encoding = r.apparent_encoding
    return r

if __name__ == "__main__":
    com_count=0
    headers = {}    #把cookie前门的！ax去掉才可用cookie，不过使用频率不可太高
    pro_name = input("输入货物")
    page = int(input("你想爬取几页"))
    urls = getUrls(pro_name,page)
    workbook = xlsxwriter.Workbook(pro_name+".xlsx")
    worksheet = workbook.add_worksheet()
    worksheet.set_column('A:A', 70)
    worksheet.set_column('B:B', 20)
    worksheet.set_column('C:C', 20)
    worksheet.set_column('D:D', 20)
    worksheet.write('A1', '名称')
    worksheet.write('B1', '价格')
    worksheet.write('C1', '地区')
    worksheet.write('D1', '付款人数')
    #xx = []
    for url in urls:
        html = getHtml(url)
        result = getIntoExcel(html.text)
        time.sleep(5)
    workbook.close()


