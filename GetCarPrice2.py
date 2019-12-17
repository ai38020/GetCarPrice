#-*- coding:utf-8 -*-
import requests
import openpyxl
import time
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

url = 'https://dealer.autohome.com.cn/handler/other/getdata?'  # 车辆的报价url
headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko)'
                         ' Chrome/78.0.3904.108 Safari/537.36',
           'Referer': 'https://dealer.autohome.com.cn/frame/spec/43086/440000/440300/0.html?isPage=1&source=cn.bing.com'}
Params = {
'__action': 'dealerlq.getdealerlistspec',
'provinceId': '440000',  # 广东省的经销商
'cityId': '0',
'countyId': '0',
'specId': '43086',  #车辆型号ID,修改可以爬取其他车辆的价格
'orderType': '0',
'pageIndex': '1',
'kindId': '1',
'isNeedMaintainNews': '1',
'pageSize': '10',
'isCPL': '1'
}


excel_file = openpyxl.Workbook()
work_sheet = excel_file.active
work_sheet.title = str(time.strftime("%Y%m%d", time.localtime(time.time())))
work_sheet['A1'] = '日期'
work_sheet['B1'] = '经销商'
work_sheet['C1'] = '销售范围'
work_sheet['D1'] = '销售电话'
work_sheet['E1'] = '最新最低报价'
work_sheet['F1'] = '优惠政策'

res = requests.get(url, headers=headers, params=Params) # 先爬取第一页数据，看下总共有多少页经销商
res =res.json()
#dealer_info = res['result']['list']
page_count = int(res['result']['pagecount'])
excel_data = []
for pageIndex in range(1,page_count+1):
#for pageIndex in range(2):
    Params1 = {
        '__action': 'dealerlq.getdealerlistspec',
        'provinceId': '440000',
        'cityId': '0',
        'countyId': '0',
        'specId': '43086',
        'orderType': '0',
        'pageIndex': pageIndex,
        'kindId': '1',
        'isNeedMaintainNews': '1',
        'pageSize': '10',
        'isCPL': '1'
    }
    res = requests.get(url, headers=headers, params=Params1)
    time.sleep(1)
    res = res.json()
    dealer_info = res['result']['list']
    for item in dealer_info:
        price_date = time.strftime("%Y%m%d",time.localtime(time.time()))
        dealer_name = item['dealerInfoBaseOut']['companySimple']
        business_Area = item['dealerInfoBaseOut']['businessArea']
        sellPhone = item['dealerInfoBaseOut']['sellPhone']
        minNewsPrice = item['minNewsPrice']
        newsInfoOut = item['newsInfoOut']['title']
        excel_data.append([price_date, dealer_name, business_Area, sellPhone, minNewsPrice, newsInfoOut])
excel_data = sorted(excel_data, key=lambda k: k[4])
for item in excel_data:
    #print(item)
    work_sheet.append(item)
excel_file.save('飞度报价.xlsx')

filename = '飞度报价.xlsx'
def send_mail(mailto_list,sub,context,filename): # to_list：收件人；sub：主题；content：邮件内容
    mail_host="smtp.qq.com" #设置服务器
    mail_user = "youremail@qq.com"
    mail_pass = "QQ邮箱的授权码"
    # mail_postfix = "qq.com"
    me="youremail@qq.com"  #这里的“服务器”可以任意设置，收到信后，将按照设置显示
    msg = MIMEMultipart() #给定msg类型
    msg['Subject'] = sub #邮件主题
    msg['From'] = me
    msg['To'] = ";".join(mailto_list)
    msg.attach(context)
    #构造附件1
    # att1 = MIMEText(open(filename, 'rb').read(), 'xls', 'gb2312')
    # att1["Content-Type"] = 'application/octet-stream'
    # att1["Content-Disposition"] = 'attachment;飞度报价='+filename[-5:]#这里的filename可以任意写，写什么名字，邮件中显示什么名字，filename[-6:]指的是之前附件地址的后6位
    # msg.attach(att1)
    try:
        s = smtplib.SMTP_SSL(mail_host, 465)
        #s.connect(mail_host, 465)    #连接smtp服务器
        s.login(mail_user, mail_pass)  # 登陆服务器
        #print(msg.as_string())
        s.sendmail(me, mailto_list, msg.as_string()) #发送邮件
        s.quit()
        return True
    except Exception as e:
        print(e)
        return False

if __name__ == '__main__':
    mailto_list=["youremail@qq.com"]
    sub="飞度报价-"+time.strftime("%Y%m%d",time.localtime(time.time()))
    d='' #表格内容
    for i in range(len(excel_data)):
        d=d+"""
      <tr>
          <td width="15">""" + str(excel_data[i][0]) + """</td>
          <td width="60">""" + str(excel_data[i][1]) + """</td>
          <td width="30" align="center">""" + str(excel_data[i][2]) + """</td>
          <td width="30">""" + str(excel_data[i][3]) + """</td>
          <td width="30">""" + str(excel_data[i][4]) + """</td>
               <td width="100">""" + str(excel_data[i][5]) + """</td>
      </tr>"""
    html = """
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
</head>
<body>
<div id="container">
<div id="content">
 <table width="60%" border="2" cellspacing="0" cellpadding="0">
<tr>
  <td width="15"><strong>报价日期</strong></td>
  <td width="60"><strong>经销商</strong></td>
  <td width="30" align="center"><strong>经销范围</strong></td>
  <td width="30"><strong>经销电话</strong></td>
  <td width="30"><strong>最低价格</strong></td>
   <td width="100"><strong>促销政策</strong></td>
</tr>"""+d+"""
</table>
</div>
</div>
</body>
</html>
   """
    print(html)
    context = MIMEText(html,_subtype='html',_charset='utf-8') #解决乱码
    context.add_header("Content-Type",'text/html; charset="utf-8"')
    if send_mail(mailto_list,sub,context,filename):
        print ("发送成功")
    else:
        print("发送失败")
