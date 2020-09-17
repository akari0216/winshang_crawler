import requests
from bs4 import BeautifulSoup
from random import randint
import json
import xlwt
import time
import os
import pandas as pd

#保存位置
os.chdir("C:\\Users\\Administrator\\Desktop\\winshang")

now1 = time.time()
fields = ["项目id","项目名","物业类型","链接","项目状态","招商状态","项目类型","开业时间","商业面积","商业楼层","所在城市","项目地址","产品线项目","项目简介","配套设施",\
          "开发商属性","开发商简介"]
#excel格式设置
def set_style(name,height,bold=False):
    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.name = name
    font.bold = bold
    font.color_index = 0
    font.height = height
    style.font = font
    return style
default = set_style("Arial",200,False)

wb = xlwt.Workbook()
sheet = wb.add_sheet(sheetname = "项目汇总",cell_overwrite_ok = True)
for i in range(len(fields)):
    sheet.write(0,i,fields[i],default)
    
#初始位置
pos = 1
 
url = "http://www.winshangdata.com/wsapi/project/list3_4"   
headers = {"Accept":"application/json, text/plain, */*",\
           "Accept-Encoding":"gzip, deflate",\
           "Accept-Language":"zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2",\
           "appType":"bigdata",\
           "Cache-Control":"no-cache",\
           "Connection":"keep-alive",\
           "Content-Length":"160",\
           "Content-Type":"application/json;charset=utf-8",\
           "Host":"www.winshangdata.com",\
           "Origin":"http://www.winshangdata.com",\
           "platform":"pc",\
           "Pragma":"no-cache",\
           "Referer":"http://www.winshangdata.com/projectList",\
           "User-Agent":"Mozilla/5.0 (Windows NT 6.1; rv:80.0) Gecko/20100101 Firefox/80.0",\
           "uuid":"123456"}
cookies = {"_ga":"GA1.2.624490075.1525923884",\
          "_uab_collina":"160015963290233514729445",\
          "Hm_lpvt_f48055ef4cefec1b8213086004a7b78d":"1600160284",\
          "Hm_lvt_f48055ef4cefec1b8213086004a7b78d":"1600159642",\
          "JSESSIONID":"406C8ABD62A05301B1A7FE33A11CF5A2"}
data = {"ifdporyt":"",\
        "isHaveLink":"",\
        "key":"",\
        "orderBy":"1",\
        "pageNum":"%s",\
        "pageSize":"60",\
        "qy_a":"",\
        "qy_c":"",\
        "qy_p":"",\
        "wuyelx":"",\
        "xmzt":"",\
        "zsxq_yt1":"",\
        "zsxq_yt2":""}
data = json.dumps(data,separators = (',',':'))

url2 = "http://www.winshangdata.com/projectDetail?projectId=%s"

headers2 = {"Accept":"text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",\
           "Accept-Encoding":"gzip, deflate",\
           "Accept-Language":"zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2",\
           "Cache-Control":"max-age=0",\
           "Connection":"keep-alive",\
           "Host":"www.winshangdata.com",\
           "If-None-Match":"a3368-e0IvqmcNpILKjiklaXq5hXyx+O0",\
           "Referer":"http://www.winshangdata.com/projectList",\
           "Upgrade-Insecure-Requests":"1",\
           "User-Agent":"Mozilla/5.0 (Windows NT 6.1; rv:80.0) Gecko/20100101 Firefox/80.0"}

col_list = ["projectId","projectName","isHaveLink","mapProjectId","isDianPuZhaoZu","isZhaoShang","xmZhuangTai","wuYeLx","wuYeLxID","wuyeLx_other","otherwuyeLxlable",\
            "projectPic","kaiYeShiJian","kaiYeShiJianReal","shangYeMianji","zhaoShangXQ","delWhy","isYiQianYueBrand","hasRemark","isYiQianYueBrandOfSort","projectViewCuntOfSort",\
            "state","addtime","finishstate","htmlUrl"]
df_total = pd.DataFrame(columns = col_list)

#自行修改页面范围
for i in range(1,3):
    page_data = data % i
    r = requests.post(url = url,headers = headers,cookies = cookies,data = page_data)
    if r.status_code == 200:
        print("requesting page : %s" % i)
        time.sleep(randint(5,10))
        if r.json()["msg"] == "执行成功":
            rdata_list = r.json()["data"]["list"]
            for each_data in rdata_list:
                each_df = pd.DataFrame(data = each_data,index = list(range(1)))
                df_total = pd.concat([df_total,each_df],ignore_index = True)
                id = each_data["projectId"]
                item_name = each_data["projectName"]
                real_estate_type = each_data["wuYeLx"]
                item_url = url2 % id
                
                #第二层单页面获取
                try:
                    r2 = requests.get(url = item_url,headers = headers2,cookies = cookies)
                    if r2.status_code == 200:
                        print("page num :%s,id : %s,item name : %s" % (i,id,item_name))
                        time.sleep(randint(3,6))
                        soup = BeautifulSoup(r2.text,"lxml")
                        #开业和招商状态
                        status = soup.find_all(attrs = {"class":"detail-three-tit"})
                        open_status = status[0].get_text()
                        invest_status = status[1].get_text()
#                         print(open_status,invest_status)
                        #商业信息
                        option = soup.find_all(attrs = {"class":"detail-option-value"})
                        item_class = option[0].get_text()
                        open_date = option[1].get_text()
                        area = option[2].get_text()
                        floor = option[3].get_text()
                        city = option[4].get_text()
                        addr = option[5].get_text()
                        is_product_line = option[6].get_text()
#                         print(item_class,open_date,area,floor,city,addr,is_product_line)
                        intro = soup.find_all(attrs = {"class":"detail-richtext"})
                        item_intro = intro[0].get_text()
                        item_fac = intro[1].get_text()
                        develop_intro = intro[2].get_text()
                        develop_detail = intro[3].get_text()
#                         print(item_intro,item_fac,develop_intro)
                        xls_data = [id,item_name,real_estate_type,item_url,open_status,invest_status,item_class,open_date,area,floor,city,addr,is_product_line,item_intro,item_fac,develop_intro,develop_detail]
                        xls_data_len = len(xls_data)
                        for k in range(xls_data_len):
                            sheet.write(pos,k,xls_data[k],default)
                        pos += 1
                
                except Exception as e:
                    print("the error log : %s" % e)
                    time.sleep(randint(40,60))
                continue       
                
            df_total["page"] = r.json()["data"]["pageNum"]
print(df_total)
#第一层信息汇总
df_total.to_excel("total.xlsx",header = True,index = False)
#第二层信息汇总，需修改名称
wb.save("项目汇总.xls")

now2 = time.time()
print("total time : %s " % (now2 - now1))