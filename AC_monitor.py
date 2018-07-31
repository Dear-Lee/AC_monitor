import os,sys
import requests
import threading
from tkinter import *
from selenium import webdriver
from time import sleep
from exchangelib import DELEGATE, Account, Credentials, Message, Mailbox, HTMLBody


def Analoglogin():  # 模拟登录
    headers = {
        'User-agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/66.0.3359.139 Safari/537.36'}

    # 登录后才能访问的网页
    login_url = 'https://10.6.0.11/wnm/frame/login.php?'
    Data = {"user_name": 'xxxxxx',
            "password": 'xxxxxx'}

    # session代表某一次连接
    Session = requests.session()
    # 使用session直接post请求
    requests.packages.urllib3.disable_warnings()
    responseRes = Session.post(login_url, data=Data, headers=headers, verify=False)  # verify=False 关闭证书验证
    # 无论是否登录成功，状态码一般都是 statusCode = 200
    NextWeb = responseRes.text
    index_url = NextWeb[8:135]
    browser.get(index_url)


def GetOfflineAp(num):  # 从div.slick-cell.l5.r5中提取text,若为Offline,则在div.first-cell-row中查询其对应的AP名称
    global offlineap_num, offlineap_name
    res = browser.find_elements_by_css_selector('div.slick-cell.l5.r5')
    for i in range(num):
        if res[i].text == "Offline":
            offlineap_num = offlineap_num + 1
            apname = browser.find_elements_by_css_selector('div.first-cell-row')
            offlineap_name.append(apname[i].text)
            t.insert(END, '离线AP:  ' + apname[i].text + '\n', 'Info')


def GetAPInfo():  # 获取AP数量
    sleep(0.5)
    browser.find_element_by_xpath('//*[@id="side_menu"]/div[2]/div[2]/ul/li[8]/a/span[1]').click()  # 点击报告
    sleep(0.5)
    browser.find_element_by_xpath('//*[@id="side_menu"]/div[2]/div[2]/ul/li[8]/ul/li[2]/a/span').click()  # 点击AP统计
    sleep(0.5)
    browser.find_element_by_xpath('//*[@id="ap_report_list"]/div[5]/div/span[2]/span[2]/span[1]').click()  # 点击右下角灯泡图标
    sleep(0.5)
    browser.find_element_by_xpath('//*[@id="ap_report_list"]/div[5]/div/span[2]/span[1]/a[3]').click()  # 点击选择每页显示10条
    page_num = browser.find_element_by_xpath(
        '//*[@id="ap_report_list"]/div[5]/div/span[3]/em[5]').text  # 每页显示10条时共会有多少页
    # t.insert(END, '页数: ' + str(page_num) + '\n','Info')
    # print('页数: ', page_num)
    totalap_num = browser.find_element_by_xpath('//*[@id="ap_report_list"]/div[5]/div/span[3]/em[1]').text  # 抓取总ap数量
    t.insert(END, 'AP总数: ' + str(totalap_num) + '\n', 'Info')

    # print('AP总数: ', totalap_num)
    GetOfflineAp(10)  # 第一页10条ap信息
    # sleep(1)
    for j in range(int(page_num) - 1):  # 翻页
        browser.find_element_by_xpath('//*[@id="ap_report_list"]/div[5]/div/span[1]/span[3]/span[1]').click()  # 点击翻页按键
        if j == int(page_num) - 2:
            lastpage_ap_num = int(totalap_num) - (int(page_num) - 1) * 10
            GetOfflineAp(lastpage_ap_num)
        elif j != (int(page_num) - 2):
            GetOfflineAp(10)  # 除最后一页每页10条ap信息
    t.insert(END, '离线AP数量: ' + str(offlineap_num) + '\n', 'Info')
    OnlineAP_num = int(totalap_num) - offlineap_num
    t.insert(END, '在线AP数量:  ' + str(OnlineAP_num) + '\n', 'Info')


def GetWlanInfo():  # 获取Wlan用户数量
    browser.find_element_by_xpath("//div[@id='side_menu']/div[2]/div[2]/ul/li[3]/a/span[2]").click()
    browser.find_element_by_xpath("//div[@id='side_menu']/div[2]/div[2]/ul/li[3]/ul/li/a/span").click()
    total_wlan = browser.find_elements_by_css_selector('div.slick-cell.l5.r5')
    pri_wlan2g = int(total_wlan[1].text)  # qiyi-private-pki(2.4G)
    pri_wlan5g = int(total_wlan[5].text)  # qiyi-private-pki(5G)
    pub_wlan2g = int(total_wlan[4].text)  # qiyi-public-pki(2.4G)
    pub_wlan5g = int(total_wlan[3].text)  # qiyi-public-pki(5G)
    pri_wlan_all = int(pri_wlan2g) + int(pri_wlan5g)
    pub_wlan_all = int(pub_wlan2g) + int(pub_wlan5g)
    t.insert(END, 'Private_2.4G+5G连接总人数:  ' + '\n' + 'qiyi-private-pki: ' + str(
        pri_wlan_all) + '\n' + 'Public_2.4G+5G连接总人数:  ' + '\n' + 'qiyi-public-pki: ' + str(pub_wlan_all) + '\n', 'Info')


def GetNewOfflineAp():  # 找出新增加的离线AP名称
    filename1 = os.path.join(os.path.dirname(sys.executable), 'offlineap_list.txt')
    file1 = open(filename1, 'r')
    offlineap_last = file1.read()  # 保存在 offlineap_list.txt中的上次的所有离线AP名称
    file1.close()

    filename2 = os.path.join(os.path.dirname(sys.executable), 'offlineap_num.txt')
    file2 = open(filename2, 'r')
    offlineap_lastnum = file2.read()
    file2.close()
    file2 = open(filename2, 'w')
    file2.write(str(offlineap_num))
    file2.close()

    filename3 = os.path.join(os.path.dirname(sys.executable), 'Warnmsg.txt')

    offlineap_New_num = offlineap_num - int(offlineap_lastnum)
    if offlineap_New_num > 0:  # 出现新的离线设备
        for i in offlineap_name:
            if i not in offlineap_last:
                offlineap_New.append(i)
                t.insert(END, 'Warning: 新出现离线AP: ' + str(i) + '\n', 'Warn')

        file3 = open(filename3, 'w')
        file3.write('<h2 style="color:red">新的离线AP数量:  </h2>' + '<h3>' + str(
            offlineap_New_num) + '</h3>' + '<h2 style="color:red">新的离线AP名称:  </h2>' + '<h3>' + str(
            offlineap_New) + '</h3>')
        file3.close()
        file3 = open(filename3, 'r')
        msg = file3.read()
        file3.close()
        EmailWarning("AP变动: 出现新的离线AP设备!", msg)
        t.insert(END, '已发送提醒邮件√' + '\n', 'Notice')
    elif offlineap_New_num < 0:  # 有原离线设备上线
        for i in eval(offlineap_last):
            if i not in offlineap_name:
                offlineap_New.append(i)
                t.insert(END, 'Warning: 以下离线AP重新上线: ' + str(i) + '\n', 'Notice')
        file3 = open(filename3, 'w')
        file3.write('<h2 style="color:green">重新上线的离线AP数量:  </h2>' + '<h3>' + str(
            (-1) * offlineap_New_num) + '</h3>' + '<h2 style="color:green">重新上线的离线AP名称:  </h2>' + '<h3>' + str(
            offlineap_New) + '</h3>')
        file3.close()
        file3 = open(filename3, 'r')
        msg = file3.read()
        file3.close()
        EmailWarning("AP变动: 有离线的AP设备重新上线!", msg)
        t.insert(END, '已发送提醒邮件√' + '\n', 'Notice')
    file1 = open(filename1, 'w')
    file1.write(str(offlineap_name))
    file1.close()


def EmailWarning(subject, body):
    creds = Credentials(  # 域账号和密码
        username='iqiyi\xxxxxx',
        password='xxxxxx'
    )
    account = Account(  # 发送邮件的邮箱
        primary_smtp_address='@.com',
        credentials=creds,
        autodiscover=True,
        access_type=DELEGATE
    )
    m = Message(
        account=account,
        subject=subject,
        body=HTMLBody(body),
        to_recipients=[
            Mailbox(email_address='@.com'),
            Mailbox(email_address='@.com'),
            Mailbox(email_address='@.com'),
            Mailbox(email_address='@.com'),
            Mailbox(email_address='@.com'),
        ]
        # 向指定邮箱里发预警邮件
    )
    m.send()


def close_TKwindow():
    for i in range(1770):
        sleep(1)
    root.quit()


if __name__ == '__main__':  # 主函数
    root = Tk()
    root.title("WLAN-AC Tool__by_Lcw")  # 父容器标题
    root.wm_state('zoomed')  # 父容器大小
    Information = LabelFrame(root, text="监控信息", padx=0, pady=0)  # 水平，垂直方向上的边距均为10
    Information.place(x=0, y=0)
    S = Scrollbar(root)
    t = Text(Information, width=500, height=500, background='Goldenrod')
    t.pack(side=LEFT, fill=Y)
    S.pack(side=RIGHT, fill=Y)
    S.config(command=t.yview)
    t.config(yscrollcommand=S.set)
    t.tag_config('Warn', foreground='FireBrick', font=("Times", 23, "bold"))
    t.tag_config('Notice', foreground='green', font=("Times", 23, "bold"))
    t.tag_config('Info', foreground='black', font=("Times", 22, "bold"))
    # selenium脚本
    # browser = webdriver.Chrome()
    path = 'phantomjs.exe'
    browser = webdriver.PhantomJS(executable_path=path)  # 无界面浏览器驱动
    browser.maximize_window()
    browser.implicitly_wait(2)

    offlineap_num = 0
    offlineap_name = []  # 当下得到的所有离线AP名称
    offlineap_New = []  # 新出现的离线AP名称
    msg = ''
    Analoglogin()
    GetAPInfo()
    GetNewOfflineAp()
    GetWlanInfo()
    browser.quit()
    trd = threading.Thread(target=close_TKwindow)
    trd.start()
    root.mainloop()
