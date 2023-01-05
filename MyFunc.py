import os
import shutil
import ddddocr
import datetime
import openpyxl
import requests
import json
import time
import pandas as pd
import streamlit as st
import numpy as np
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from sqlalchemy import create_engine
from st_aggrid import JsCode

url = "https://b.keruyun.com"

#配置文件下载目录
root_folder = os.path.abspath('.')      #工作目录
temp_folder = f"{root_folder}\\Temp\\"  #下载临时文件夹
data_folder = f"{root_folder}\\Data\\"  #数据文件夹

company = '434049'       #组织机构代码
user = '18362095019'     #用户名
password = 'Xuzhi112358' #密码

db_home = 'postgresql://SyxAdmin:112358@192.168.0.109:15432/SyxDatabase' #家里数据库
db = 'postgresql://SyxAdmin:112358@192.168.0.250:5432/SyxDatabase' #公司数据库

def ini_tmp():
    shutil.rmtree(temp_folder)
    os.mkdir(temp_folder)

#建立家里数据库连接
def get_home_engine():
    return create_engine(db_home)

#建立公司数据库连接
def get_engine():
    return create_engine(db)

#获取天气信息
def getWeather(place):
    AK = "d25d31d05cf66186789f0c177b6dc070"  # 你自己注册获取的key
    url = f"https://restapi.amap.com/v3/weather/weatherInfo?city={place}&extensions=all&key={AK}"  # 高德地图extension=all可以查询预报天气
    res = requests.get(url)
    json_data = json.loads(res.text)
    return json_data

#星期转中文星期
def get_week_cn(day):
    week_c = ['星期一','星期二','星期三','星期四','星期五','星期六','星期天']
    return week_c[day]

#数字金额转中文金额
from decimal import Decimal
def nTOc(value, capital=True, prefix=False, classical=None):
    # 默认大写金额用圆，一般汉字金额用元
    if classical is None:
        classical = True if capital else False
    # 汉字金额前缀
    if prefix is True:
        prefix = '人民币'
    else:
        prefix = ''
    # 汉字金额字符定义
    dunit = ('角', '分')
    if capital:
        num = ('零', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖')
        iunit = [None, '拾', '佰', '仟', '万', '拾', '佰', '仟', '亿', '拾', '佰', '仟', '万', '拾', '佰', '仟']
    else:
        num = ('〇', '一', '二', '三', '四', '五', '六', '七', '八', '九')
        iunit = [None, '十', '百', '千', '万', '十', '百', '千', '亿', '十', '百', '千', '万', '十', '百', '千']
    if classical:
        iunit[0] = '元' if classical else '圆'
    # 转换为Decimal，并截断多余小数
    if not isinstance(value, Decimal):
        value = Decimal(value).quantize(Decimal('0.01'))
    # 处理负数
    if value < 0:
        prefix += '负'  # 输出前缀，加负
        value = - value  # 取正数部分，无须过多考虑正负数舍入
        # assert - value + value == 0
    # 转化为字符串
    s = str(value)
    if len(s) > 19:
        raise ValueError('金额太大了，不知道该怎么表达。')
    istr, dstr = s.split('.')  # 小数部分和整数部分分别处理
    istr = istr[::-1]  # 翻转整数部分字符串
    so = []  # 用于记录转换结果
    # 零
    if value == 0:
        return prefix + num[0] + iunit[0]
    haszero = False  # 用于标记零的使用
    if dstr == '00':
        haszero = True  # 如果无小数部分，则标记加过零，避免出现“圆零整”
    # 处理小数部分
    # 分
    if dstr[1] != '0':
        so.append(dunit[1])
        so.append(num[int(dstr[1])])
    else:
        so.append('整')  # 无分，则加“整”
    # 角
    if dstr[0] != '0':
        so.append(dunit[0])
        so.append(num[int(dstr[0])])
    elif dstr[1] != '0':
        so.append(num[0])  # 无角有分，添加“零”
        haszero = True  # 标记加过零了
    # 无整数部分
    if istr == '0':
        if haszero:  # 既然无整数部分，那么去掉角位置上的零
            so.pop()
        so.append(prefix)  # 加前缀
        so.reverse()  # 翻转
        return ''.join(so)
    # 处理整数部分
    for i, n in enumerate(istr):
        n = int(n)
        if i % 4 == 0:  # 在圆、万、亿等位上，即使是零，也必须有单位
            if i == 8 and so[-1] == iunit[4]:  # 亿和万之间全部为零的情况
                so.pop()  # 去掉万
            so.append(iunit[i])
            if n == 0:  # 处理这些位上为零的情况
                if not haszero:  # 如果以前没有加过零
                    so.insert(-1, num[0])  # 则在单位后面加零
                    haszero = True  # 标记加过零了
            else:  # 处理不为零的情况
                so.append(num[n])
                haszero = False  # 重新开始标记加零的情况
        else:  # 在其他位置上
            if n != 0:  # 不为零的情况
                so.append(iunit[i])
                so.append(num[n])
                haszero = False  # 重新开始标记加零的情况
            else:  # 处理为零的情况
                if not haszero:  # 如果以前没有加过零
                    so.append(num[0])
                    haszero = True
    # 最终结果
    so.append(prefix)
    so.reverse()
    return ''.join(so)

#读取门店信息
def get_store():
   return read_sql("SELECT * "
                   "FROM store_info "
                   "WHERE class='门店' "
                   "AND POSITION('停用' IN name) = 0 "
                   "AND POSITION('新开门' IN name) = 0 "
                   "AND POSITION('新开店' IN name) = 0 "
                   "ORDER BY CONVERT_TO (name, 'GBK');")
#读取商品信息
def get_goods():
    return read_sql("SELECT * "
                    "FROM goods_info "
                    "ORDER BY CONVERT_TO (goods_name, 'GBK');")

#读取物品信息
def get_stuffs():
    return read_sql("SELECT * "
                    "FROM stuff_info "
                    "ORDER BY CONVERT_TO (stuff_name, 'GBK');")

#获取图片验证码
def get_captha(driver, captha_file):
    img = driver.find_element(By.XPATH, '/html/body/div/div/section/section/div/div[2]/div[3]/form/div[4]/div/div/span/span/span/span[2]/img')
    data = img.screenshot_as_png
    img.screenshot(captha_file)
    ocr = ddddocr.DdddOcr()
    with open(captha_file, 'rb') as f:
        img_bytes = f.read()
    captha_code = ocr.classification(img_bytes)
    os.remove(captha_file)
    return captha_code

#登录模块Start
def get_login():
    # 配置Chrome下载目录为指定目录
    option = webdriver.ChromeOptions()
    prefs = {'profile.default_content_setting_values.automatic_downloads': 1,
             'download.default_directory': temp_folder}
    option.add_experimental_option("excludeSwitches", ["enable-logging"])
    option.add_experimental_option("prefs", prefs)
    option.add_argument("--start-maximized")
    drive = webdriver.Chrome(options=option)
    drive.implicitly_wait(5)
    drive.get(url=r'https://b.keruyun.com')
    # 登录客如云
    way_login = drive.find_element(By.XPATH, '//div[text()="账号登录"]')
    ActionChains(drive).click(way_login).perform()
    time.sleep(1)

    # 登录客如云
    login_company = drive.find_element(By.XPATH, '//*[@id="loginId"]')
    login_name = drive.find_element(By.XPATH, '//*[@id="userName"]')
    login_password = drive.find_element(By.XPATH, '//*[@id="password"]')
    login_captha = drive.find_element(By.XPATH, '//*[@id="captcha"]')

    login_company.send_keys(str(company))
    login_name.send_keys(str(user))
    login_password.send_keys(str(password))
    while True:
        captha = get_captha(drive, temp_folder + r'captha.png')
        login_captha.clear()
        login_captha.send_keys(captha)
        login = drive.find_element(By.XPATH, '//button[@type="submit"]')
        login.click()
        time.sleep(2)
        if drive.current_url == "https://b.keruyun.com/bui-link/#/mind-ui/#/report/homePage":
            break
        else:
            continue
    return drive
    # 登录结束
#登录模块End

#初始化streamlit
def init_st(title,icon,layout):
    st.set_page_config(page_title=title, page_icon=icon, layout=layout, initial_sidebar_state='auto')
    hide_streamlit_style = """
    <style>
        #MainMenu {
            visibility: hidden;
        }
        footer {
            visibility: hidden;
        }
        .css-18e3th9 {
            padding-top: 2rem;
            padding-bottom: 2rem;
            padding-left: 2rem;
            padding-right: 2rem;
        }
        .info {
            height: 50px;
            background-color: rgb(235, 242, 251);
            border-radius: 5px;
            text-shadow: 1px 1px 2px gray;
            padding: 12px;
        }
        .block-container.css-1gx893w.egzxvld2 {
            margin-top: -50px;
        }
        .main.css-k1vhr4.egzxvld3 {
            margin-top: 20px;
        }
    </style>
    """
    st.markdown(hide_streamlit_style, unsafe_allow_html=True)

#执行pd.read_sql语句
def read_sql(sql):
    engine=get_engine()
    df = pd.read_sql(sql=sql, con=engine)
    engine.dispose()
    return df

#执行execute sql语句
def exec_sql(sql):
    engine=get_engine()
    engine.execute(sql)
    engine.dispose()

def get_week(day):
    week_list = ['mon','tues','wed','thur','fri','sat','sun']
    week_c = ['星期一','星期二','星期三','星期四','星期五','星期六','星期天']
    return week_list[day.weekday()], week_c[day.weekday()]

#AgGrid表格resize
def get_js_resize():
    return JsCode("""
    function(e) {
        e.api.sizeColumnsToFit();
    };
    """)

#AgGrid修改单元格变色
def get_js():
    return JsCode("""
    function(e) {
        let api = e.api;
        let rowIndex = e.rowIndex;
        let col = e.column.colId;

        let rowNode = api.getDisplayedRowAtIndex(rowIndex);
        api.flashCells({
          rowNodes: [rowNode],
          columns: [col],
          flashDelay: 10000000000
        });
    };""")

#配货复核
def get_js_ph():
    return JsCode("""
    function(e) {
        newFzsl = Math.round(e.data.订货数量 * e.data.辅助数量比值);
        let api = e.api;
        let rowIndex = e.rowIndex;
        let col = e.column.colId;

        let rowNode = api.getDisplayedRowAtIndex(rowIndex);
        api.flashCells({
          rowNodes: [rowNode],
          columns: [col],
          flashDelay: 10000000000
        });
        if(e.data.辅助数量比值 !== 1){rowNode.setDataValue('辅助数量',newFzsl);}
    };""")

#更新实际库存
def get_js_gxkc():
    return JsCode("""
    function(e) {
        newFzsl = Math.round(e.data.实际库存 * e.data.辅助数量比值);
        let api = e.api;
        let rowIndex = e.rowIndex;
        let col = e.column.colId;

        let rowNode = api.getDisplayedRowAtIndex(rowIndex);
        api.flashCells({
          rowNodes: [rowNode],
          columns: [col],
          flashDelay: 10000000000
        });
        if(e.data.辅助数量比值 !== 1){rowNode.setDataValue('辅助数量',newFzsl);}
    };""")

#AgGrid表格.CSS
def get_css():
    return {
        ".ag-theme-streamlit": {
            "--ag-alpine-active-color": "#0000ff;",
            "--ag-icon-size": "16px;",
            "--ag-font-size": "14px;",
            "--ag-row-height": "40px;",
            "--ag-header-height": "40px;",
            "--ag-odd-row-background-color": "#f6f6f6;",
            "--ag-row-hover-color": "#daeefe;",
            "--ag-selected-tab-underline-color": "#2d5ff5;",
            "--ag-input-focus-border-color": "#93aef8;",
            "--ag-input-focus-box-shadow": "0px 0px 7px #0000ff;",
            "--ag-checkbox-checked-color": "#007bff;",
            "--ag-value-change-value-highlight-background-color": "#cdcffd;",
            "--ag-range-selection-border-color": "#93aef8;",
            "--ag-card-radius": "5px;",
            "--ag-card-shadow": "5px 5px 7px #696969;",
            "--ag-popup-shadow": "5px 5px 7px #696969;"
        },
    }

#翻译AgGrid为中文
def ag_local():
    return  {"page":"当前页",
             "more":"更多",
             "to": "至",
             "of": "总数",
             "next": "下一页",
             "last": "最后一页",
             "first": "首页",
             "previous": "上一页",
             "loadingOoo": "加载中...",
             "selectAll": "全选",
             "searchOoo": "请输入关键字",
             "blanks": "空白",
             "filterOoo": "过滤",
             "equals": "等于",
             "notEqual": "不等于",
             "lessThan": "小于",
             "greaterThan": "大于",
             "lessThanOrEqual": "小于等于",
             "greaterThanOrEqual": "大于等于",
             "inRange": "范围",
             "contains": "包含",
             "notContains": "不包含",
             "startsWith": "开始于...",
             "endsWith": "结束于...",
             "group": "分组",
             "columns": "列选项",
             "rowGroupColumnsEmptyMessage": "拖拽组到这里",
             "pivotMode": "枢轴模式",
             "noRowsToShow": "无数据",
             "pinColumn": "锁定列",
             "autosizeThiscolumn": "自动调整当前列宽",
             "autosizeAllColumns": "自动调整所有列度",
             "groupBy": "分组",
             "ungroupBy": "取消分组",
             "resetColumns": "恢复",
             "expandAll": "全部展开",
             "collapseAll": "全部关闭",
             "toolPanel": "显示/隐藏控制表盘",
             "export": "导出至...",
             "csvExport": "导出至CSV",
             "excelExport": "导出至Excel",
             "pinLeft": "锁定至表格左侧",
             "pinRight": "锁定至表格右侧",
             "noPin": "取消锁定",
             "sum": "总计",
             "min": "最小值",
             "max": "最大值",
             "none": "无",
             "count": "计数",
             "average": "平均",
             "copy": "复制",
             "copyWithHeaders": "复制内容及标题",
             "copyWithGroupHeaders": "复制内容及组标题",
             "ctrlC": "ctrl + C",
             "paste": "粘贴",
             "ctrlV": "ctrl + V",
             }

#获取文件夹内所有文件列表
def get_filelist(folder):
    files = os.listdir(folder)
    files = [os.path.join(folder, f) for f in files]
    return files

#复制文件
def copy_file(real_folder, new_file):  #file_folder下载文件所在目录; real_folder目标文件夹; new_file目标文件
    while True:
        try:
            files = os.listdir(temp_folder)
            files = [os.path.join(temp_folder, f) for f in files]
            files.sort(key=lambda x: os.path.getmtime(x))
            newest_file = files[-1]
            ext = os.path.splitext(newest_file)[-1]
            if ext in ('.xls', '.xlsx', '.csv'):
                shutil.copyfile(newest_file, real_folder + new_file)
                os.remove(newest_file)
                break
        except Exception as e:
            print(e)
            continue

#数据下载模块开始
#工厂出库汇总表
def factory_out_download(driver, folder, dldate):
    day = str(dldate)
    print(f"{day} - 【出库汇总表(工厂)】 - 开始下载......")
    time.sleep(2)
    driver.get('https://b.keruyun.com/bui-link/#/supply2')     #打开供应链2.0
    time.sleep(2)
    driver.switch_to.frame("appkey_chensenSupply")             #切换至框架
    time.sleep(1)
    setMenu = driver.find_element(By.XPATH, '//input[@id="rc_select_0"]')
    setMenu.send_keys('出库汇总表-按日期')
    driver.find_element(By.XPATH, '//div[@title="出库汇总表-按日期【库存管理>库存报表>出库业务查询】"]').click()
    time.sleep(1)
    #日期选择开始
    element1 = driver.find_element(By.CLASS_NAME, "scm-ant-calendar-picker")
    #element1 = driver.find_element(By.XPATH, '//input[@placeholder="开始日期" and @class="ant-calendar-range-picker-input"]')
    ActionChains(driver).click(element1).perform()
    time.sleep(1)
    s_date = driver.find_element(By.XPATH, '//input[@class="scm-ant-calendar-input " and @placeholder="开始日期"]')
    s_date.send_keys(Keys.CONTROL, "a")
    time.sleep(1)
    s_date.send_keys(day)
    ActionChains(driver).click(element1).perform()
    e_date = driver.find_element(By.XPATH, '//input[@class="scm-ant-calendar-input " and @placeholder="结束日期"]')
    time.sleep(1)
    e_date.send_keys(Keys.CONTROL, "a")
    e_date.send_keys(day)
    time.sleep(1)
    e_date.send_keys(Keys.ENTER)
    time.sleep(1)
    driver.find_element(By.XPATH, '//div[text()="请选择库存组织"]').click()
    driver.find_element(By.XPATH, '//div[text()="请选择库存组织"]/following-sibling::div/div/input').send_keys('南京盛源祥食品有限公司')
    time.sleep(1)
    driver.find_element(By.XPATH, '//div[text()="请选择库存组织"]/following-sibling::div/div/input').send_keys(Keys.ENTER)
    time.sleep(1)
    driver.find_element(By.XPATH, '//span[text()="搜 索"]/parent::button').click()
    time.sleep(3)
    driver.find_element(By.XPATH, '//span[text()="导出表格"]/parent::button').click()
    time.sleep(5)
    f_name = f"出库汇总表(工厂){day}.xls"
    time.sleep(3)
    driver.switch_to.default_content()
    copy_file(folder,f_name)
    time.sleep(0.5)
    print(f"{day} - 【出库汇总表(工厂)】 - 下载完成，正在写入数据库......")
    df = pd.read_excel(f"{folder}{f_name}", usecols='A,C:E,H:J,L:N', sheet_name=0)
    df = df.applymap(lambda x: np.nan if str(x).strip() == '' or x == 0 else x)
    df.columns = ['happen_date', 'stock_name', 'stuff_code', 'stuff_name', 'stuff_category2', 'stuff_category3', 'qty', 'unit', 'qty2',
                  'unit2']
    try:
        exec_sql(f"DELETE FROM factory_out_details WHERE happen_date='{day}';")
        engine = get_engine()
        df.to_sql('factory_out_details', engine, index=False, if_exists='append')
        engine.dispose()
        print(f"{day} - 【出库汇总表(工厂)】 - 写入数据库成功！")
    except Exception as e:
        print(e)
        print(f"{day} - 【出库汇总表(工厂)】 - 写入数据库失败，请重新下载")
    print()

#工厂入库汇总表
def factory_in_download(driver, folder, dldate):
    day = str(dldate)
    print(f"{day} - 【入库汇总表(工厂)】 - 开始下载......")
    time.sleep(2)
    driver.get('https://b.keruyun.com/bui-link/#/supply2')   #打开供应链2.0
    time.sleep(2)
    driver.switch_to.frame("appkey_chensenSupply")           #切换至框架
    time.sleep(1)
    setMenu = driver.find_element(By.XPATH, '//input[@id="rc_select_0"]')
    setMenu.send_keys('入库汇总表-按日期')
    driver.find_element(By.XPATH, '//div[@title="入库汇总表-按日期【库存管理>库存报表>入库业务查询】"]').click()
    time.sleep(1)
    #日期选择开始
    iframe = driver.find_element(By.XPATH, '//iframe[@frameborder="0"]')
    driver.switch_to.frame(iframe)
    element1 = driver.find_element(By.CLASS_NAME, "ant-calendar-picker")
    #element1 = driver.find_element(By.XPATH, '//input[@placeholder="开始日期" and @class="ant-calendar-range-picker-input"]')
    ActionChains(driver).click(element1).perform()
    time.sleep(1)
    s_date = driver.find_element(By.XPATH, '//input[@class="ant-calendar-input " and @placeholder="开始日期"]')
    s_date.send_keys(Keys.CONTROL, "a")
    time.sleep(1)
    s_date.send_keys(day)
    ActionChains(driver).click(element1).perform()
    e_date = driver.find_element(By.XPATH, '//input[@class="ant-calendar-input " and @placeholder="结束日期"]')
    time.sleep(1)
    e_date.send_keys(Keys.CONTROL, "a")
    e_date.send_keys(day)
    time.sleep(1)
    driver.find_element(By.XPATH, '//div[text()="请选择组织"]').click()
    time.sleep(1)
    driver.find_element(By.XPATH, '//li[text()="南京盛源祥食品有限公司"]').click()
    time.sleep(1)
    driver.find_element(By.XPATH, '//span[text()="搜 索"]/parent::button').click()
    time.sleep(3)
    driver.find_element(By.XPATH, '//span[text()="导出表格"]/parent::button').click()
    time.sleep(5)
    driver.find_element(By.XPATH, '//span[text()="下 载"]/parent::button').click()
    time.sleep(3)
    f_name = f"入库汇总表(工厂){day}.xls"
    time.sleep(3)
    driver.switch_to.default_content()
    copy_file(folder,f_name)
    time.sleep(0.5)
    driver.switch_to.frame("appkey_chensenSupply")
    print(f"{day} - 【入库汇总表(工厂)】 - 下载完成，正在写入数据库......")
    df = pd.read_excel(f"{folder}{f_name}", usecols='B,D:F,H:L', header=2, sheet_name=0)
    df.dropna(inplace=True)
    df.drop(df.tail(1).index, inplace=True)
    df = df.applymap(lambda x: np.nan if str(x).strip() == '' else x)
    df.columns = ['happen_date', 'stock_name', 'stuff_code', 'stuff_name', 'stuff_category', 'qty', 'unit', 'qty2',
                  'unit2']
    try:
        exec_sql(f"DELETE FROM factory_in_details WHERE happen_date='{day}';")
        engine = get_engine()
        df.to_sql('factory_in_details', engine, index=False, if_exists='append')
        engine.dispose()
        print(f"{day} - 【入库汇总表(工厂)】 - 写入数据库成功！")
    except Exception as e:
        print(e)
        print(f"{day} - 【入库汇总表(工厂)】 - 写入数据库失败，请重新下载")
    print()

#收银明细表
def symxb_download(driver, folder, dldate):
    day = str(dldate)
    print(f"{day} - 【收银明细表】 - 开始下载......")
    time.sleep(2)
    driver.get('https://b.keruyun.com/bui-link/#/mind-ui/#/report/enshrine')                  #打开报表中心
    time.sleep(2)
    driver.get('https://b.keruyun.com/bui-link/#/mind-ui/#/order/cashRegisterDetails')             #打开收银明细表
    time.sleep(2)
    driver.switch_to.frame("appkey_mindbaobiaozonghefenxishouyinmingxibiao")  #切换至框架
    time.sleep(1)
    #日期选择开始
    original_window = driver.current_window_handle
    element1 = driver.find_element(By.XPATH, '//input[@placeholder="开始日期" and @class="ant-calendar-range-picker-input"]')
    ActionChains(driver).click(element1).perform()
    time.sleep(1)
    #driver.find_element(By.LINK_TEXT, '昨天').click()
    s_date = driver.find_element(By.XPATH, '//input[@class="ant-calendar-input " and @placeholder="开始日期"]')
    s_date.send_keys(Keys.CONTROL, "a")
    time.sleep(1)
    s_date.send_keys(day)
    ActionChains(driver).click(element1).perform()
    e_date = driver.find_element(By.XPATH, '//input[@class="ant-calendar-input " and @placeholder="结束日期"]')
    time.sleep(1)
    e_date.send_keys(Keys.CONTROL, "a")
    e_date.send_keys(day)
    time.sleep(1)
    driver.find_element(By.XPATH, '//span[text()="查 询"]/parent::button').click()
    time.sleep(2)
    driver.find_element(By.XPATH, '//span[text()="离线导出"]/parent::button').click()
    time.sleep(1)
    driver.find_element(By.XPATH, '//span[text()="前往下载中心"]/parent::button').click()
    driver.switch_to.window(driver.window_handles[-1])
    while True:
        downloadclick = driver.find_element(By.XPATH, '//span[text()="下载"]')
        status =downloadclick.get_attribute('style')
        freshbutton = driver.find_element(By.XPATH, '//span[text()="刷 新"]')
        time.sleep(1)
        if status == 'color: rgb(37, 143, 248); cursor: pointer;':
            ActionChains(driver).click(downloadclick).perform()
            break
        else:
            ActionChains(driver).click(freshbutton).perform()
            time.sleep(1)
            continue
    f_name = f"收银明细表{day}.xls"
    time.sleep(10)
    driver.switch_to.default_content()
    copy_file(folder,f_name)
    driver.close()
    driver.switch_to.window(original_window)
    driver.switch_to.frame("appkey_mindbaobiaozonghefenxishouyinmingxibiao")  # 切换至框架
    time.sleep(0.5)
    print(f"{day} - 【收银明细表】 - 下载完成，正在写入数据库......")
    df = pd.read_excel(f"{folder}{f_name}",engine='openpyxl',header=2,sheet_name=1)
    df.drop(df.index[0],inplace=True)
    tmp = df[['日期',
              '门店',
              '订单号',
              '订单金额',
              '商户优惠',
              '订单配送支出',
              '损益金额',
              '服务费',
              '补贴',
              '营业收入',
              '下单时间',
              '支付时间',
              '订单来源',
              '就餐类型',
              '订单状态',
              '支付状态',
              '就餐人数'
              ]]
    tmp.columns = ['sale_date',
                   'store_name',
                   'order_no',
                   'order_amt',
                   'store_discount',
                   'delivery_fee',
                   'loss_amt',
                   'service_fee',
                   'subsidy',
                   'income',
                   'order_time',
                   'paid_time',
                   'order_from',
                   'buy_class',
                   'order_stat',
                   'paid_stat',
                   'customer'
                   ]
    try:
        exec_sql(f"DELETE FROM cashier_details WHERE sale_date='{day}';")
        engine = get_engine()
        tmp.to_sql('cashier_details',engine,index=False,if_exists='append')
        engine.dispose()
        print(f"{day} - 【收银明细表】 - 写入数据库成功！")
    except Exception as e:
        print(e)
        print(f"{day} - 【收银明细表】 - 写入数据库失败，请重新下载")
    print()

#商品销售汇总表【报表中心-商品销售-商品销售汇总表】
def spxshzb_download(driver, folder, dldate):
    store_list = get_store()['name']
    day = str(dldate)
    original_window = driver.current_window_handle
    print(f"{day} - 【商品销售汇总表】 - 开始下载......")
    time.sleep(2)
    driver.get('https://b.keruyun.com/bui-link/#/mind-ui/#/report/enshrine')                  #打开报表中心
    time.sleep(2)
    driver.get('https://b.keruyun.com/bui-link/#/mind-report/#/report/salesStatisticsNew')             #打开物品销售统计表
    time.sleep(2)
    driver.switch_to.frame("appkey_mindbaobiaoxiaoshoubaobiaoshangpingshishoubiao")  #切换至框架
    time.sleep(1)
    #日期选择开始
    element1 = driver.find_element(By.XPATH, '//input[@placeholder="开始日期" and @class="ant-calendar-range-picker-input"]')
    ActionChains(driver).click(element1).perform()
    time.sleep(1)
    #driver.find_element(By.XPATH, '//span[text()="昨天"]').click()
    s_date = driver.find_element(By.XPATH, '//input[@class="ant-calendar-input " and @placeholder="开始日期"]')
    e_date = driver.find_element(By.XPATH, '//input[@class="ant-calendar-input " and @placeholder="结束日期"]')
    s_date.send_keys(Keys.CONTROL, "a")
    s_date.send_keys(str(day))
    e_date.send_keys(Keys.CONTROL, "a")
    e_date.send_keys(str(day))
    time.sleep(1)
    #日期选择结束
    driver.find_element(By.XPATH, '//span[text()="查询"]/parent::button').click()  # 查询
    time.sleep(0.5)
    driver.find_element(By.XPATH, '//span[text()="离线导出"]/parent::button').click()  # 下载文件
    time.sleep(0.5)
    driver.find_element(By.ID, "SelectDimension_groups").click()
    time.sleep(0.5)
    driver.find_element(By.XPATH, '//li[text()="按门店+日期导出"]').click()
    time.sleep(0.5)
    driver.find_element(By.XPATH, '//input[@class="ant-checkbox-input" and @value="2"]').click()
    time.sleep(0.5)
    driver.find_element(By.XPATH, '//input[@class="ant-checkbox-input" and @value="3"]').click()
    time.sleep(0.5)
    driver.find_element(By.XPATH, '//span[text()="确 定"]/parent::button').click()
    time.sleep(1)
    driver.find_element(By.XPATH, '//span[text()="前往下载中心"]/parent::button').click()
    driver.switch_to.window(driver.window_handles[-1])
    while True:
        downloadclick = driver.find_element(By.XPATH, '//span[text()="下载"]')
        status = downloadclick.get_attribute('style')
        freshbutton = driver.find_element(By.XPATH, '//span[text()="刷 新"]')
        if status == r'color: rgb(37, 143, 248); cursor: pointer;':
            ActionChains(driver).click(downloadclick).perform()
            break
        else:
            ActionChains(driver).click(freshbutton).perform()
            time.sleep(1)
            continue
    time.sleep(1)
    fname = f"商品销售汇总表{day}.xlsx"
    copy_file(folder, fname)
    driver.close()
    spxshzb_file = os.path.join(folder, fname)
    driver.switch_to.window(original_window)
    driver.switch_to.frame("appkey_mindbaobiaoxiaoshoubaobiaoshangpingshishoubiao")  # 切换至框架
    time.sleep(0.5)
    print(f"{day} - 【商品销售汇总表】 - 下载完成，准备下载商品销售统计表......")
    driver.switch_to.default_content()
    #商品销售统计表
    print(f"{day} - 【商品销售统计表】 - 开始下载......")
    #time.sleep(2)
    driver.get('https://b.keruyun.com/bui-link/#/mind-report/#/report/salesStatistics')  #打开商品销售统计表
    time.sleep(2)
    driver.switch_to.frame("appkey_mindbaobiaoxiaoshoubaobiaoshangpingxiaoshoubaobiao")  #切换至框架
    time.sleep(1)
    #日期选择开始
    element1 = driver.find_element(By.XPATH, '//input[@placeholder="开始日期" and @class="ant-calendar-range-picker-input"]')
    ActionChains(driver).click(element1).perform()
    time.sleep(1)
    #driver.find_element(By.XPATH, '//span[text()="昨天"]').click()
    s_date = driver.find_element(By.XPATH, '//input[@class="ant-calendar-input " and @placeholder="开始日期"]')
    e_date = driver.find_element(By.XPATH, '//input[@class="ant-calendar-input " and @placeholder="结束日期"]')
    s_date.send_keys(Keys.CONTROL, "a")
    s_date.send_keys(Keys.DELETE)
    s_date.send_keys(day)
    e_date.send_keys(Keys.CONTROL, "a")
    e_date.send_keys(Keys.DELETE)
    e_date.send_keys(day)
    time.sleep(1)
    driver.find_element(By.XPATH, '//span[text()="商品销售统计表"]').click()
    #日期选择结束
    #选择门店开始
    for d_name in (store_list):
        original_window = driver.current_window_handle
        driver.find_element(By.XPATH, '//label[@title="选择门店"]/parent::div/following-sibling::div[1]/div/span/button').click()
        time.sleep(1)
        driver.find_element(By.XPATH, '//span[text()="已选门店"]/following-sibling::div').click()
        time.sleep(1)
        ssmd = driver.find_element(By.XPATH, '//input[@placeholder="搜索门店名称"]')
        ssmd.clear()
        ssmd.send_keys(d_name)
        time.sleep(1)
        driver.find_element(By.XPATH, f"//span[@title='{d_name}']").click()
        time.sleep(1)
        driver.find_element(By.XPATH, '//span[text()="确 定"]/parent::button').click()
        time.sleep(1)
        #门店选择完成
        driver.find_element(By.XPATH, '//span[text()="查询"]/parent::button').click()
        time.sleep(1)
        driver.find_element(By.XPATH, '//span[text()="离线导出"]/parent::button').click()  # 下载文件
        time.sleep(1)
        driver.find_element(By.XPATH, '//span[text()="前往下载中心"]/parent::button').click()
        driver.switch_to.window(driver.window_handles[-1])
        while True:
            downloadclick = driver.find_element(By.XPATH, '//span[text()="下载"]')
            status = downloadclick.get_attribute('style')
            freshbutton = driver.find_element(By.XPATH, '//span[text()="刷 新"]')
            if status == r'color: rgb(37, 143, 248); cursor: pointer;':
                ActionChains(driver).click(downloadclick).perform()
                break
            else:
                ActionChains(driver).click(freshbutton).perform()
                time.sleep(1)
                continue
        time.sleep(2)
        driver.close()
        driver.switch_to.window(original_window)
        driver.switch_to.frame("appkey_mindbaobiaoxiaoshoubaobiaoshangpingxiaoshoubaobiao")  # 切换至框架
        time.sleep(1)
    driver.switch_to.default_content()
    print(f"{day} - 【商品销售统计表】 - 下载完成，准备写入数据库......")
    files = get_filelist(temp_folder)
    tmp = pd.DataFrame(columns=['中类名称/编码', '商品名称', '商品编码', '单位', '销售次数', '销售数量', '销售日期','门店'])
    for i in files:
        df = pd.read_excel(i, header=0, sheet_name=0)
        store = df.iloc[[2], [4]].values[0][0]
        df = pd.read_excel(i, usecols='C:H', header=8, sheet_name=0)
        df.drop(df.head(1).index, inplace=True)
        df.drop(df.tail(1).index, inplace=True)
        df['销售日期'] = day
        df['门店'] = store
        tmp = pd.concat([tmp, df])
        os.remove(i)
    tmp.to_excel(f"{data_folder}商品销售统计表\\商品销售统计表{day}.xlsx", index=False)
    tmp.drop(['中类名称/编码','商品编码','单位','销售数量'],axis=1,inplace=True)
    tmp.columns = ['goods_name',
                   'sale_times',
                   'sale_date',
                   'store_name']
    df1 = pd.read_excel(spxshzb_file, usecols='C:H,K,N,O', header=4, sheet_name=1)
    df1.drop(df1.tail(1).index, inplace=True)
    df1.columns = ['goods_category',
                   'goods_name',
                   'goods_code',
                   'store_name',
                   'sale_date',
                   'sale_qty',
                   'sale_amt',
                   'loss_amt',
                   'income']
    tmp_db = pd.merge(df1,tmp,how='left',on=['goods_name','sale_date','store_name'])
    match_db = read_sql("SELECT goods_name,stuff_name,stuff_code,rate FROM matching;")
    tmp_db = pd.merge(tmp_db,match_db,how='left',on='goods_name')
    try:
        exec_sql(f"DELETE FROM goods_sale_details WHERE sale_date='{day}';")
        engine = get_engine()
        tmp_db.to_sql('goods_sale_details',engine,index=False,if_exists='append')
        engine.dispose()
        print(f"{day} - 【商品销售汇总表】 - 写入数据库已完成！")
    except Exception as e:
        print(e)
        print(f"{day} - 【商品销售汇总表】 - 写入数据库失败！请重新下载！")
    print()

#营业概况下载
def yygk_download(driver, folder, dldate):
    day = str(dldate)
    original_window = driver.current_window_handle
    print(f"{day} - 【营业概况】 - 开始下载......")
    time.sleep(2)
    driver.get('https://b.keruyun.com/bui-link/#/mind-ui/#/report/enshrine')                  #打开报表中心
    time.sleep(2)
    driver.get('https://b.keruyun.com/bui-link/#/mind/report/bizSurveyRes/index')             #打开营业概况
    time.sleep(2)
    driver.switch_to.frame("appkey_mindbaobiaoyingyebaobiaoyingyegaikuang")  #切换至框架
    time.sleep(1)
    #日期选择开始
    element1 = driver.find_element(By.XPATH, '//input[@placeholder="开始日期" and @class="ant-calendar-range-picker-input"]')
    ActionChains(driver).click(element1).perform()
    time.sleep(1)
    #driver.find_element(By.XPATH, '//span[text()="昨天"]').click()
    s_date = driver.find_element(By.XPATH, '//input[@class="ant-calendar-input " and @placeholder="开始日期"]')
    s_date.send_keys(Keys.CONTROL, "a")
    s_date.send_keys(str(day))
    ActionChains(driver).click(element1).perform()
    time.sleep(1)
    e_date = driver.find_element(By.XPATH, '//input[@class="ant-calendar-input " and @placeholder="结束日期"]')
    e_date.send_keys(Keys.CONTROL, "a")
    e_date.send_keys(str(day))
    time.sleep(1)
    #日期选择结束
    driver.find_element(By.XPATH, '//span[text()="查询"]/parent::button').click()  # 查询
    time.sleep(1)
    driver.find_element(By.XPATH, '//span[text()="离线导出"]/parent::button').click()  # 下载文件
    time.sleep(1)
    driver.find_element(By.XPATH, '//div[@title="按日期导出"]').click()
    time.sleep(1)
    driver.find_element(By.XPATH, '//li[text()="按门店+日期导出"]').click()
    time.sleep(1)
    driver.find_element(By.XPATH, '//input[@class="ant-checkbox-input" and @value="2"]').click()
    time.sleep(1)
    driver.find_element(By.XPATH, '//input[@class="ant-checkbox-input" and @value="3"]').click()
    time.sleep(1)
    driver.find_element(By.XPATH, '//span[text()="确 定"]/parent::button').click()
    time.sleep(1)
    driver.find_element(By.XPATH, '//span[text()="前往下载中心"]/parent::button').click()
    driver.switch_to.window(driver.window_handles[-1])
    while True:
        downloadclick = driver.find_element(By.XPATH, '//span[text()="下载"]')
        status = downloadclick.get_attribute('style')
        freshbutton = driver.find_element(By.XPATH, '//span[text()="刷 新"]')
        if status == r'color: rgb(37, 143, 248); cursor: pointer;':
            ActionChains(driver).click(downloadclick).perform()
            break
        else:
            ActionChains(driver).click(freshbutton).perform()
            time.sleep(1)
            continue
    time.sleep(2)
    fname = f"营业概况{day}.xlsx"
    copy_file(folder, fname)
    driver.close()
    yygk_file = os.path.join(folder, fname)
    driver.switch_to.window(original_window)
    driver.switch_to.frame("appkey_mindbaobiaoyingyebaobiaoyingyegaikuang")  # 切换至框架
    driver.switch_to.default_content()
    time.sleep(1)
    print(f"{day} - 【营业概况】 - 下载完成，准备写入数据库......")
    #写数据库
    df1 = pd.read_excel(yygk_file, header=5, sheet_name='收入统计')
    df1.drop(df1.head(1).index, axis=0, inplace=True)
    yygk = df1[['Unnamed: 0', 'Unnamed: 1', '订单笔数', '订单金额', '销货笔数', '退款笔数', '退款金额', '商户优惠',
                '订单配送支出',
                '损益金额', '服务费', '补贴', '储值支付笔数', '储值支付', '营业收入', '储值收款笔数', '储值收入']]
    yygk.columns = ['store_name', 'sale_date', 'order_times', 'order_amt', 'sale_times', 'refund_times', 'refund_amt',
                    'store_discount', 'delivery_fee', 'loss_amt', 'service_fee', 'subsidy', 'stored_value_paid_times',
                    'stored_value_paid_amt', 'income', 'stored_value_in_times', 'stored_value_in_amt']
    df2 = pd.read_excel(yygk_file, usecols='A,G,I', header=4, sheet_name='顾客统计')
    df2.drop(df2.head(1).index, inplace=True)
    df2.columns = ['store_name', 'customer', 'price_aft_discount']
    yygk = pd.merge(yygk, df2, how='left', on='store_name')
    try:
        exec_sql(f"DELETE FROM business_overview WHERE sale_date='{day}';")
        engine = get_engine()
        yygk.to_sql('business_overview', engine, index=False, if_exists='append')
        engine.dispose()
        print(f"{day} - 【营业概况】 - 写入数据库已完成！")
    except Exception as e:
        print(e)
        print(f"{day} - 【营业概况】 - 写入数据库失败！请重新下载！")
    print()

#储值明细表(1.0,已失效)
def czmxb_download(driver, folder, dldate):
    day = str(dldate)
    print(f"{day} - 【储值明细表】 - 开始下载......")
    time.sleep(2)
    driver.get('https://b.keruyun.com/bui-link/#/loyalty_ui/#/report/stored-value-detail')             #打开收银明细表
    time.sleep(2)
    driver.switch_to.frame("appkey_mindbaobiaoyingxiaobaobiaochushimingxi")  #切换至框架
    time.sleep(2)
    #日期选择开始
    driver.find_element(By.XPATH, '//span[text()="至"]/preceding-sibling::span').click()
    time.sleep(1)
    s_date = driver.find_element(By.CLASS_NAME, 'ant-calendar-input  ')
    time.sleep(1)
    s_date.send_keys(Keys.CONTROL, "a")
    s_date.send_keys(day)
    driver.find_element(By.XPATH, '//span[text()="至"]/following-sibling::span').click()
    time.sleep(1)
    e_date = driver.find_element(By.CLASS_NAME, 'ant-calendar-input  ')
    time.sleep(1)
    e_date.send_keys(Keys.CONTROL, "a")
    e_date.send_keys(day)
    time.sleep(1)
    driver.find_element(By.XPATH, '//span[text()="查 询"]/parent::button').click()
    time.sleep(3)
    driver.find_element(By.XPATH, '//span[text()="导 出"]/parent::button').click()
    time.sleep(10)
    f_name = f"储值明细表{day}.xls"
    driver.switch_to.default_content()
    copy_file(folder,f_name)
    time.sleep(0.5)
    print(f"{day} - 【储值明细表】 - 下载完成，正在写入数据库......")
    df = pd.read_excel(f"{folder}{f_name}", header=14, usecols='A,B,E,F,H,J,K,L,M,N,P', sheet_name=0)
    df.columns = ['member_name',
                  'member_id',
                  'trade_date',
                  'order_no',
                  'store_name',
                  'trade_type',
                  'real_amt',
                  'gift_amt',
                  'batch_amt',
                  'trade_amt',
                  'paid_way']
    try:
#        exec_sql(f"DELETE FROM member_stored_value_details WHERE to_char(trade_date,'yyyy-mm-dd')='{day}';")
        engine = get_engine()
        df.to_sql('member_stored_value_details', engine, index=False, if_exists='append')
        engine.dispose()
        print(f"{day} - 【储值明细表】 - 写入数据库成功！")
    except Exception as e:
        print(e)
        print(f"{day} - 【储值明细表】 - 写入数据库失败，请重新下载")
    print()
#数据下载模块结束

#更新门店信息
def refresh_store(driver):
    print('开始更新门店信息......')
    driver.get('https://b.keruyun.com/bui-link/#/authority/orgManage')                  #打开报表中心
    time.sleep(4)
    driver.find_element(By.XPATH, '//span[text()="导 出"]/parent::button').click()
    time.sleep(2)
    filename = '组织机构信息.csv'
    time.sleep(2)
    copy_file('.\\', filename)
    time.sleep(1)
    driver.get('https://b.keruyun.com/bui-link/#/supply2')   #打开供应链2.0
    time.sleep(2)
    driver.switch_to.frame("appkey_chensenSupply")           #切换至框架
    time.sleep(1)
    setMenu = driver.find_element(By.XPATH, '//input[@id="rc_select_0"]')
    setMenu.send_keys('门店设置')
    setMenu.send_keys(Keys.ENTER)
    time.sleep(2)
#    driver.find_element(By.XPATH, '//div[text()="100 条/页"]').click()
#    time.sleep(1)
#    driver.find_element(By.XPATH, '//li[text()="200 条/页"]').click()
#    time.sleep(1)
    all_check = driver.find_element(By.XPATH, '//input[@class="cook-checkbox-input"]')
    ActionChains(driver).click(all_check).perform()
    time.sleep(1)
    driver.find_element(By.XPATH, '//span[text()="导出门店"]/parent::button').click()
    time.sleep(3)
    filename2 = '门店供应链信息.xlsx'
    copy_file('.\\', filename2)
    time.sleep(1)
    df = pd.read_csv(filename,dtype=str)
    list = []
    for i in df['详细地址']:
        list.append(gd_map(i))
    df['经纬度'] = list
    df.columns = ['class','name','number','area','address','contacts','phone','store_phone','coordinate']
    df2 = pd.read_excel(filename2,usecols='A,B',dtype=str)
    df2.columns = ['store_code','name']
    df = pd.merge(df,df2,how='left',on='name')
    df3 = read_sql("SELECT number,is_use FROM store_info;")
    df = pd.merge(df,df3,how='left',on='number')
    df['is_use'] = df['is_use'].apply(lambda x: '不配货' if x != '配货' else x)
    df['purchase_no'] = df['number']
    exec_sql("DELETE FROM store_info")
    engine = get_engine()
    df.to_sql('store_info',engine,index=False,if_exists='append')
    engine.dispose()
    os.remove(filename)
    os.remove(filename2)
    print('门店信息更新完成！')
    print()

#获取经纬度
def gd_map(addr):
    para = {'key': 'a2d97df2da8653f8695d556b72108ddd',  # 高德Key
            'address': addr}  # 地址参数
    url = 'https://restapi.amap.com/v3/geocode/geo?'  # 高德地图地理编码API服务地址
    result = requests.get(url, para)  # GET方式请求
    result = result.json()
    lon_lat = result['geocodes'][0]['location']  # 获取返回参数geocodes中的location，即经纬度
    return lon_lat

#会员消费商品明细表
def member_trade_details_download(driver, day):
    day = str(day)
    print(f"{day} - 【会员消费商品明细表】 - 开始下载......")
    driver.implicitly_wait(10)
    driver.get('https://b.keruyun.com/bui-link/#/loyalty_ui/#/report/stored-value-detail')
    driver.switch_to.frame('appkey_mindbaobiaoyingxiaobaobiaochushimingxi')
    driver.find_element(By.XPATH, '//span[text()="至"]/preceding-sibling::span').click()
    time.sleep(1)
    s_date = driver.find_element(By.CLASS_NAME, 'ant-calendar-input  ')
    time.sleep(1)
    s_date.send_keys(Keys.CONTROL, "a")
    s_date.send_keys(day)
    driver.find_element(By.XPATH, '//span[text()="至"]/following-sibling::span').click()
    time.sleep(1)
    e_date = driver.find_element(By.CLASS_NAME, 'ant-calendar-input  ')
    time.sleep(1)
    e_date.send_keys(Keys.CONTROL, "a")
    e_date.send_keys(day)
    driver.find_element(By.XPATH, '//button[text()="更多条件"]').click()
    time.sleep(1)
    driver.find_element(By.XPATH, '//span[text()="普通储值"]').click()
    driver.find_element(By.XPATH, '//span[text()="小额储值"]').click()
    # driver.find_element(By.XPATH, '//span[text()="消费撤销"]').click()
    # driver.find_element(By.XPATH, '//span[text()="消费"]').click()
    driver.find_element(By.XPATH, '//span[text()="批量预储"]').click()
    driver.find_element(By.XPATH, '//span[text()="补录"]').click()
    driver.find_element(By.XPATH, '//span[text()="充值撤销"]').click()
    driver.find_element(By.XPATH, '//span[text()="批量销储"]').click()
    driver.find_element(By.XPATH, '//span[text()="扣除"]').click()
    driver.find_element(By.XPATH, '//span[text()="查 询"]/parent::button').click()
    time.sleep(5)
    all_pages = driver.find_element(By.XPATH, '//li[@title="下一页"]/preceding-sibling::li[1]').get_attribute("title")
    if all_pages == '上一页':
        pages = 0
    else:
        pages = int(all_pages)
    for i in range(1, pages + 1):
        page = driver.find_element(By.XPATH, '//div[@class="ant-pagination-options-quick-jumper"]/child::input')
        page.send_keys(Keys.CONTROL, 'a')
        page.send_keys(i)
        page.send_keys(Keys.ENTER)
        time.sleep(2)
        trade_nos = driver.find_elements(By.XPATH, '//table[@class="active-table"]/tbody/tr/td[5]')
        for j in trade_nos:
            time.sleep(1)
            trade_no = j.text
            member_name = driver.find_element(By.XPATH,
                                              f"//td[text()='{trade_no}']/preceding-sibling::td[4]").text.strip().replace("'","''")
            member_id = driver.find_element(By.XPATH,
                                            f"//td[text()='{trade_no}']/preceding-sibling::td[3]").text.strip()
            trade_date = datetime.datetime.strptime(
                driver.find_element(By.XPATH, f"//td[text()='{trade_no}']/preceding-sibling::td[2]").text.strip(),
                '%Y-%m-%d %H:%M:%S')
            order_no_click = driver.find_element(By.XPATH, f"//td[text()='{trade_no}']/preceding-sibling::td[1]")
            order_no = order_no_click.text.strip()
            store_name = driver.find_element(By.XPATH,
                                             f"//td[text()='{trade_no}']/following-sibling::td[1]").text.strip()
            trade_type = driver.find_element(By.XPATH,
                                             f"//td[text()='{trade_no}']/following-sibling::td[3]").text.strip()
            original_window = driver.current_window_handle
            order_no_click.click()
            driver.switch_to.window(driver.window_handles[-1])
            discount_amt = float(
                driver.find_element(By.XPATH, '//td[text()="优惠金额:"]/following-sibling::td').text.strip())
            if discount_amt == 0:
                discount = False
            else:
                discount = True
            tr_list = driver.find_elements(By.XPATH, '//table[@id="goodsCountTable"]/tbody/tr')
            del tr_list[-1]
            for tr in tr_list:
                td_list = tr.find_elements(By.TAG_NAME, 'td')
                goods_code = td_list[0].text.strip()
                goods_name = td_list[1].text.strip()
                goods_price = float(td_list[2].text.strip())
                goods_qty = float(td_list[3].text.strip())
                goods_unit = td_list[4].text.strip()
                goods_amt = float(td_list[5].text.strip())
                if_exist = read_sql(f"SELECT * "
                                    f"FROM member_trade_details "
                                    f"WHERE member_id='{member_id}' "
                                    f"AND member_name='{member_name}' "
                                    f"ANd trade_date='{trade_date}' "
                                    f"AND order_no='{order_no}' "
                                    f"AND trade_no='{trade_no}' "
                                    f"AND store_name='{store_name}' "
                                    f"AND trade_type='{trade_type}' "
                                    f"AND goods_code='{goods_code}' "
                                    f"AND goods_name='{goods_name}' "
                                    f"AND goods_price='{goods_price}' "
                                    f"AND goods_qty='{goods_qty}' "
                                    f"AND goods_unit='{goods_unit}' "
                                    f"AND goods_amt='{goods_amt}' "
                                    f"AND discount='{discount}';")
                if if_exist.empty:
                    exec_sql(f"INSERT INTO member_trade_details "
                             f"(member_id, member_name, trade_date, order_no, trade_no, store_name, trade_type, "
                             f"goods_code, goods_name, goods_price, goods_qty, goods_unit, goods_amt, discount)"
                             f"	VALUES ('{member_id}', '{member_name}', '{trade_date}', '{order_no}', '{trade_no}', '{store_name}', "
                             f"'{trade_type}', '{goods_code}', '{goods_name}', '{goods_price}', '{goods_qty}', '{goods_unit}', "
                             f"'{goods_amt}', '{discount}');")
            driver.close()
            driver.switch_to.window(original_window)
            driver.switch_to.frame("appkey_mindbaobiaoyingxiaobaobiaochushimingxi")  # 切换至框架
    print(f"{day} - 【会员消费商品明细表】 - 下载完成")

#门店入库明细表
def mdrkmxb_download(driver, folder, dldate):
    day = str(dldate)
    print(f"{day} - 【门店入库明细表】 - 开始下载......")
    time.sleep(2)
    driver.get('https://b.keruyun.com/bui-link/#/supply2')   #打开供应链2.0
    time.sleep(2)
    driver.switch_to.frame("appkey_chensenSupply")           #切换至框架
    time.sleep(1)
    setMenu = driver.find_element(By.XPATH, '//input[@id="rc_select_0"]')
    setMenu.send_keys('门店入库明细表')
    setMenu.send_keys(Keys.ENTER)
    time.sleep(1)
    iframe = driver.find_element(By.XPATH, '//iframe[@frameborder="0"]')
    driver.switch_to.frame(iframe)
    #日期选择开始
    element1 = driver.find_element(By.CLASS_NAME, "ant-calendar-picker")
    ActionChains(driver).click(element1).perform()
    time.sleep(1)
    s_date = driver.find_element(By.XPATH, '//input[@class="ant-calendar-input " and @placeholder="开始日期"]')
    s_date.send_keys(Keys.CONTROL, "a")
    time.sleep(1)
    s_date.send_keys(day)
    ActionChains(driver).click(element1).perform()
    e_date = driver.find_element(By.XPATH, '//input[@class="ant-calendar-input " and @placeholder="结束日期"]')
    time.sleep(1)
    e_date.send_keys(Keys.CONTROL, "a")
    e_date.send_keys(day)
    time.sleep(1)
    driver.find_element(By.XPATH, '//div[text()="请选择门店"]').click()
    time.sleep(1)
    driver.find_element(By.XPATH, '//li[text()="全部"]').click()
    time.sleep(1)
    driver.find_element(By.XPATH, '//span[text()="搜 索"]/parent::button').click()
    time.sleep(3)
    driver.find_element(By.XPATH, '//span[text()="导出表格"]/parent::button').click()
    time.sleep(3)
    driver.find_element(By.XPATH, '//span[text()="下 载"]/parent::button').click()
    time.sleep(3)
    f_name = f"门店入库明细表{day}.xls"
    time.sleep(3)
    copy_file(folder,f_name)
    time.sleep(0.5)
    print(f"{day} - 【门店入库明细表】 - 下载完成，正在写入数据库......")
    df = pd.read_excel(f"{folder}{f_name}", usecols='B,C,E,F,H:L,Q,R', header=2, sheet_name=0)
    df.drop(df.tail(1).index, inplace=True)
    df = df.applymap(lambda x: np.nan if str(x).strip() == '' else x)
    df.columns = ['store_name', 'no', 'stuff_code', 'stuff_name', 'stuff_category', 'qty', 'unit', 'qty2',
                  'unit2', 'business_type', 'happen_date']
    try:
        exec_sql(f"DELETE FROM store_in_out_details WHERE happen_date='{day}';")
        engine = get_engine()
        df.to_sql('store_in_out_details', engine, index=False, if_exists='append')
        engine.dispose()
        print(f"{day} - 【门店入库明细表】 - 写入数据库成功！")
    except Exception as e:
        print(e)
        print(f"{day} - 【门店入库明细表】 - 写入数据库失败，请重新下载")
    print()

#门店出库明细表
def mdckmxb_download(driver, folder, dldate):
    day = str(dldate)
    print(f"{day} - 【门店出库明细表】 - 开始下载......")
    time.sleep(2)
    driver.get('https://b.keruyun.com/bui-link/#/supply2')   #打开供应链2.0
    time.sleep(2)
    driver.switch_to.frame("appkey_chensenSupply")           #切换至框架
    time.sleep(1)
    setMenu = driver.find_element(By.XPATH, '//input[@id="rc_select_0"]')
    setMenu.send_keys('门店出库明细表')
    setMenu.send_keys(Keys.ENTER)
    time.sleep(1)
    iframe = driver.find_element(By.XPATH, '//iframe[@frameborder="0"]')
    driver.switch_to.frame(iframe)
    #日期选择开始
    element1 = driver.find_element(By.CLASS_NAME, "ant-calendar-picker")
    ActionChains(driver).click(element1).perform()
    time.sleep(1)
    s_date = driver.find_element(By.XPATH, '//input[@class="ant-calendar-input " and @placeholder="开始日期"]')
    s_date.send_keys(Keys.CONTROL, "a")
    time.sleep(1)
    s_date.send_keys(day)
    ActionChains(driver).click(element1).perform()
    e_date = driver.find_element(By.XPATH, '//input[@class="ant-calendar-input " and @placeholder="结束日期"]')
    time.sleep(1)
    e_date.send_keys(Keys.CONTROL, "a")
    e_date.send_keys(day)
    time.sleep(1)
    driver.find_element(By.XPATH, '//div[text()="请选择门店"]').click()
    time.sleep(1)
    driver.find_element(By.XPATH, '//li[text()="全部"]').click()
    time.sleep(1)
    driver.find_element(By.XPATH, '//span[text()="搜 索"]/parent::button').click()
    time.sleep(3)
    driver.find_element(By.XPATH, '//span[text()="导出表格"]/parent::button').click()
    time.sleep(3)
    driver.find_element(By.XPATH, '//span[text()="下 载"]/parent::button').click()
    time.sleep(3)
    f_name = f"门店出库明细表{day}.xls"
    time.sleep(3)
    copy_file(folder,f_name)
    time.sleep(0.5)
    print(f"{day} - 【门店出库明细表】 - 下载完成，正在写入数据库......")
    df = pd.read_excel(f"{folder}{f_name}", usecols='B,C,E,F,H:L,O,Q', header=2, sheet_name=0)
    df.drop(df.tail(1).index, inplace=True)
    df = df.applymap(lambda x: np.nan if str(x).strip() == '' else x)
    df.columns = ['store_name', 'no', 'stuff_code', 'stuff_name', 'stuff_category', 'qty', 'unit', 'qty2',
                  'unit2', 'business_type', 'happen_date']
    df['qty'] = df['qty'].apply(lambda x:x*-1)
    df['qty2'] = df['qty2'].apply(lambda x:float(x)*-1 if float(x) > 0 else float(x))
    try:
        exec_sql(f"DELETE FROM store_in_out_details WHERE happen_date='{day}';")
        engine = get_engine()
        df.to_sql('store_in_out_details', engine, index=False, if_exists='append')
        engine.dispose()
        print(f"{day} - 【门店出库明细表】 - 写入数据库成功！")
    except Exception as e:
        print(e)
        print(f"{day} - 【门店入库明细表】 - 写入数据库失败，请重新下载")
    print()

#物品档案表
def stuff_download(driver):
    print(f"开始更新物品档案......")
    time.sleep(2)
    driver.get('https://b.keruyun.com/bui-link/#/supply2')   #打开供应链2.0
    time.sleep(2)
    driver.switch_to.frame("appkey_chensenSupply")           #切换至框架
    time.sleep(1)
    setMenu = driver.find_element(By.XPATH, '//input[@id="rc_select_0"]')
    setMenu.send_keys('物品档案')
    setMenu.send_keys(Keys.ENTER)
    time.sleep(1)
    #iframe = driver.find_element(By.XPATH, '//iframe[@frameborder="0"]')
    #driver.switch_to.frame(iframe)
    #日期选择开始
    time.sleep(1)
    driver.find_element(By.XPATH, '//span[@title="请选择物品类别"]').click()
    time.sleep(1)
    driver.find_element(By.XPATH, '//span[text()="工厂成品分类"]').click()
    time.sleep(1)
    driver.find_element(By.XPATH, '//span[text()="搜 索"]/parent::button').click()
    time.sleep(3)
    driver.find_element(By.XPATH, '//span[text()="导 出"]/parent::button').click()
    time.sleep(3)
    f_name = f"物品信息清单.xlsx"
    copy_file('.\\',f_name)
    time.sleep(0.5)
    df1 = pd.read_excel(f"{f_name}", usecols='B:D,F', sheet_name=0)
    df1 = df1.applymap(lambda x: np.nan if str(x).strip() == '' else x)
    df1.columns = ['stuff_category', 'stuff_code', 'stuff_name','unit']
    df2 = read_sql("SELECT stuff_code,unit2,qty2_rate,stuff_price,is_used,oversell_rate FROM stuff_info")
    df = pd.merge(df1,df2,how='left',on='stuff_code')
    exec_sql(f"DELETE FROM stuff_info;")
    engine = get_engine()
    df.to_sql('stuff_info', engine, index=False, if_exists='append')
    engine.dispose()
    os.remove(f_name)
    print("更新物品档案信息完成！")
    print()

#储值明细表(2.0)
def new_czmxb_download(driver, folder, dldate):
    day = str(dldate)
    print(f"{day} - 【储值明细表】 - 开始下载......")
    time.sleep(2)
    driver.get('https://b.keruyun.com/bui-link/#/third-crm/#/koubeicrm')             #会员营销
    time.sleep(2)
    driver.switch_to.frame("appkey_koubeicrm")  #切换至框架
    time.sleep(1)
    #日期选择开始
    menu1 = driver.find_element(By.XPATH, '//span[text()="报表"]')
    ActionChains(driver).move_to_element(menu1).perform()
    time.sleep(1)
    menu2 = driver.find_element(By.XPATH, '//div[text()="储值明细表"]')
    ActionChains(driver).click(menu2).perform()
    time.sleep(2)
    driver.find_element(By.XPATH, '//input[@placeholder="开始日期"]').click()
    time.sleep(1)
    driver.find_element(By.XPATH, '//span[text()="昨天"]').click()
    time.sleep(1)
    driver.find_element(By.XPATH, '//span[text()="查 询"]').click()
    time.sleep(3)
    driver.find_element(By.XPATH, '//span[text()="导出"]').click()
    time.sleep(6)
    driver.find_element(By.XPATH, '//span[text()="已阅读并知晓"]').click()
    time.sleep(30)
    driver.find_element(By.XPATH, '//span[text()="去查看下载数据"]').click()
    time.sleep(2)
    driver.find_element(By.XPATH, '//a[text()="下载"]').click()
    time.sleep(10)
    f_name = f"储值明细表{day}.csv"
    copy_file(folder,f_name)
    time.sleep(2)
    print(f"{day} - 【储值明细表】 - 下载完成，正在写入数据库......")
    df = pd.read_csv(f"{folder}{f_name}") #, usecols='A:M,Q,T'
    df.drop(['交易后余额(实储)', '交易后余额(赠储)', '交易后余额', '操作人', '营业日期'], axis=1, inplace=True)
    df.columns = ['card_no',        #卡号
                  'member_name',    #会员姓名
                  'member_id',      #手机号
                  'card_type',      #卡种类
                  'card_store_name',#开卡门店
                  'order_no',       #交易订单号
                  'store_name',     #交易门店
                  'trade_type',     #交易类型
                  'trade_reason',   #原因
                  'real_amt',       #交易金额（实储）
                  'gift_amt',       #交易金额（赠储）
#                  'batch_amt',     #交易金额（批量储值）1.0
                  'pre_amt',        #交易金额（预储）2.0
                  'trade_amt',      #交易金额
                  'paid_way',       #支付方式
                  'trade_date']     #交易时间
    try:
        exec_sql(f"DELETE FROM member_stored_value_details WHERE to_char(trade_date,'yyyy-mm-dd')='{day}';")
        engine = get_engine()
        df.to_sql('member_stored_value_details', engine, index=False, if_exists='append')
        engine.dispose()
        print(f"{day} - 【储值明细表】 - 写入数据库成功！")
    except Exception as e:
        print(e)
        print(f"{day} - 【储值明细表】 - 写入数据库失败，请重新下载")
    print()
