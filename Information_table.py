import xlwt
from selenium.webdriver import Chrome
from selenium.webdriver.common.keys import Keys
from Test3.chaojiying import Chaojiying_Client
import time
import re

def main():
    web = Chrome()
    web.get("http://218.200.73.44/oa/login.aspx")
    data_list = [['何自红'],
                 ['杨海蓉'],
                 ['苟翱'],
                 ['娄春'],
                 ['姚广泉'],
                 ['旦周才仁'],
                 ['袁世维'],
                 ['沈彬'],
                 ['徐初'],
                 ['姜荣'],
                 ['赵菲'],
                 ['王苏'],
                 ['王威'],
                 ['张江淋'],
                 ['张伟'],
                 ['姬浩钧'],
                 ['文俊强 '],
                 ['程洋'],
                 ['马灿'],
                 ['何荣睿'],
                 ['王玉盾'],
                 ['党广露'],
                 ['李涛'],
                 ['熊祖涛'],
                 ['纪笑天'],
                 ['曹聪'],
                 ['王凯缘'],
                 ['郑兵'],
                 ['王艺霏'],
                 ['马小梅'],
                 ['蔡通林'],
                 ['严卓'],
                 ['龚子骏'],
                 ['李大海'],
                 ['张明杭'],
                 ['张正天'],
                 ['吴仕良'],
                 [' 赵永军'],
                 ['李健'],
                 ['刘梓航'],
                 ['唐秀荷'],
                 ['杜金'],
                 ['李建平'],
                 ['李炜'],
                 ['杨安'],
                 ['张欢'],
                 ['胡睿阳'],
                 ['努尔夏提江·纳迪尔'],
                 ['杨康杰'],
                 ['汪文淳'],
                 ['张锦山'],
                 ['宋锐'],
                 ['李然'],
                 ['张万欣'],
                 ['姜晗蕊'],
                 ['徐少聪'],
                 ['陈钊毅'],
                 ['柯贤勇'],
                 ['李锴'],
                 ['李昱呈'],
                 ['李娜'],
                 ['陈诗雨'],
                 ['罗紫艺'],
                 ['张思宇']]
    # date1 = input("请输入查询日期：格式为xxxx年-xx月-xx日")
    date1 = time.strftime('%Y-%m-%d',time.localtime())  # 获取当前的时间、
    print(date1)
    valid_date(date1)
    Login_jzlg(web)
    getData(web, data_list, date1)
    savedata(data_list)

# 登录到后台
def Login_jzlg(web):
    time.sleep(1)
    print("-------正在登录到督导监控后台------")
    # 找到输入框，输入python => 输入回车/点击搜索按钮
    web.find_element_by_xpath('//*[@id="txt_卡号"]').send_keys("1039")
    web.find_element_by_xpath('//*[@id="txt_密码"]').send_keys("1039")
    img = web.find_element_by_xpath('//*[@id="getcode"]').screenshot_as_png
    chaojiying = Chaojiying_Client('wuxuebing', '@wuxuebing', '917152')
    dic = chaojiying.PostPic(img, 1902)  # 获得一个字典
    verify_code = dic['pic_str']
    verify_code = verify_code.upper()  # 将验证码转换为大写
    web.find_element_by_xpath('//*[@id="txt_identificationCode"]').send_keys(verify_code, Keys.ENTER)
    print("-------切换到我的办公室页面--------")
    time.sleep(3)  # 沉睡3秒
    # 变更selenium的窗口视角
    web.switch_to.window(web.window_handles[-1])
    # 处理iframe
    # 处理iframe,必须先拿到iframe,然后切换视角到iframe,再然后才可以定位元素
    web.switch_to.frame('mainFrame')  # 切换到iframe
    # web.switch_to.default_content()  # 切换回原页面
    web.find_element_by_xpath('//*[@id="Dal_自己定义模块_ctl00_Img_左则"]').click()
    print("--------切换到督导监控页面----------")
    time.sleep(6)

    web.switch_to.window(web.window_handles[-1])

    web.find_element_by_xpath(
        '//*[@id="form1"]/table[1]/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr/td[1]/span[2]/a').click()

    time.sleep(3)
    web.switch_to.window(web.window_handles[-1])
    print("--------成功登录到督导监控后台---------")

# 查询交表情况
def getData(web, data_list, date1):
    print("------开始查询交表情况-------")
    for data in data_list:
        web.find_element_by_xpath('//*[@id="txt_search"]').clear()
        name = data[0]
        web.find_element_by_xpath('//*[@id="txt_search"]').send_keys(name, Keys.ENTER)
        job_detail = web.find_element_by_xpath('//*[@id="DataList1"]/tbody/tr/td/table/tbody/tr[3]/td[3]').text
        class_name = web.find_element_by_xpath('//*[@id="DataList1"]/tbody/tr/td/table/tbody/tr[3]/td[8]').text
        student_id = web.find_element_by_xpath('//*[@id="DataList1"]/tbody/tr/td/table/tbody/tr[3]/td[6]').text
        print("正在查询"+name+"的交表情况")
        print(name+"的最后交表日期为："+job_detail)
        class_name = re.sub('全日制','',class_name)
        data.append(class_name)
        data.append(student_id)
        if date_compare(date1, job_detail):
            data.append("√")
            print(name+"------已交表")
        else:
            data.append("×")
            print(name+"------未交表")
        time.sleep(1)
    # print(data_list)

# 生成表格
def savedata(data_list):
    book = xlwt.Workbook(encoding="utf-8")  # 创建workbook对象
    sheet = book.add_sheet("软件工程.xls", cell_overwrite_ok=True)  # 创建工作表
    col = ("姓名","班级", "学号", "交表情况")
    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.name = '宋体'
    style.font = font
    for i in range(0, 4):
        sheet.write(0, i, col[i],style)
        sheet.col(i).width = 6200
    for i in range(0, len(data_list)):
        data = data_list[i]
        for j in range(0, 4):
            sheet.write(i + 1, j, data[j],style)
    book.save("2020-2021-2学期教学信息反馈学生交表情况一览表.xls")  # 保存

# 简单验证下日期格式
def valid_date(datestring):
    try:
        mat = re.match('^(\d{4})-(0[1-9]|1[0-2])-(0[1-9]|1\d|2\d|3[0-1])$', datestring)
        if mat is not None:
            return
        else:
            print("-------日期格式错误--------")
            main()
    except ValueError:
        pass
    return

# 比较日期
def date_compare(date1, date2):
    try:
        time1 = time.mktime(time.strptime(date1, '%Y-%m-%d'))
        time2 = time.mktime(time.strptime(date2, '%Y-%m-%d'))
        diff = int(time1) - int(time2)
        diff = int(diff / 24 / 60 / 60)
        if diff <= 7:
            return True
        else:
            return False
    except Exception as e:
        print(e)
        return ''

if __name__ == "__main__":  # 当程序执行时
    main()
