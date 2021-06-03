from selenium import webdriver
import time
import json
import pyautogui
import xlrd
import xlwt
import random


def init(ll, tt):
    f = r'C:\Program Files (x86)\Google\Chrome\Application\chromedriver'  # 驱动器地址
    # chrome_options 初始化选项
    chrome_options = webdriver.ChromeOptions()
    # 设置浏览器初始 位置x,y & 宽高x,y
    chrome_options.add_argument(f'--window-position={217},{172}')
    chrome_options.add_argument(f'--window-size={1200},{1000}')
    chrome_options.add_experimental_option("excludeSwitches", ['enable-automation'])
    # 关闭开发者模式
    chrome_options.add_experimental_option("useAutomationExtension", False)
    # 禁止图片加载
    prefs = {"profile.managed_default_content_settings.images": 2}
    chrome_options.add_experimental_option("prefs", prefs)
    # 设置中文
    chrome_options.add_argument('lang=zh_CN.UTF-8')
    # 更换头部
    chrome_options.add_argument(
        'user-agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/79.0.3945.130 Safari/537.36"')
    global driver
    driver = webdriver.Chrome(f, options=chrome_options)
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": """Object.defineProperty(navigator, 'webdriver', {get: () => undefined})""",
    })
    driver.get('https://www.baidu.com')
    driver.maximize_window()
    driver.get('https://amap.com/service/poiInfo?query_type=TQUERY&city=310000&keywords=金泽2路')
    b = driver.find_element_by_tag_name("body").text
    url = json.loads(b)
    print(len(url))
    if len(url) == 2:
        url = url["url"]
        a = "https://www.amap.com" + url
        driver.get(a)
        driver.implicitly_wait(3)
        # pyautogui.moveTo(x=random.randint(0, 2500), y=random.randint(0, 1400))
        # pyautogui.click()
        # pyautogui.moveTo(x=random.randint(0, 2500), y=random.randint(0, 1400))
        # pyautogui.click()
        # pyautogui.moveTo(x=random.randint(0, 2500), y=random.randint(0, 1400))
        # pyautogui.mouseDown(x=1090, y=305, button='left')
        a = random.randint(1, 8)
        print("a:", a)
        # pyautogui.press('enter')
        time.sleep(5)
        # pyautogui.moveTo(1590 + a * 10, 305 + a * 50,2,pyautogui.easeInOutQuad)
        # pyautogui.mouseUp(1590 + a * 100, 305 + a * 100)
        print("初始化已完成，滑块验证第" + str(ll) + "次,已爬取" + str(tt) + "次")


def getjson(linename):
    req_url = poi_bus_url + linename
    driver.get(req_url)
    driver.implicitly_wait(3)
    datajson = json.loads(driver.find_element_by_tag_name("body").text)
    # print(req_url)
    # print("json:", datajson)
    # print(len(datajson))
    while (len(datajson) == 2):
        c = driver.find_element_by_tag_name("body").text
        url1 = json.loads(c)
        url1 = url1["url"]
        a1 = "https://www.amap.com" + url1
        time.sleep(25)
        # driver.get(a1)
        # time.sleep(5)
        # driver.implicitly_wait(3)
        driver.get(req_url)
        driver.implicitly_wait(3)
        datajson = json.loads(driver.find_element_by_tag_name("body").text)
        dataList = []
    datajson = datajson['data']
    return datajson


def getbusline(linename, datajson):
    dataList = []
    try:
        datajson1 = datajson['busline_list'][0]
    except:
        innerList = []
        innerList.append("")
        innerList.append(linename)
        innerList.append("")
        innerList.append("")
        innerList.append("")
        innerList.append("")
        innerList.append("")
        innerList.append("")
        innerList.append("")
        innerList.append("")
        innerList.append("")
        dataList.append(innerList)
        return dataList
    else:
        for ii in range(len(datajson['busline_list'])):
            data = datajson['busline_list'][ii]
            id = data['id']
            code = data['code']
            front_name = data['front_name']
            terminal_name = data['terminal_name']
            company = data['company']
            start_time = data['start_time']
            end_time = data['end_time']
            length = data['length']
            xs = data['xs']
            ys = data['ys']
            for j in range(len(xs.split(','))):
                innerList = []
                # 每个innerList存储一对数据
                x = xs.split(',')
                y = ys.split(',')
                innerList.append(id)
                innerList.append(linename)
                innerList.append(code)
                innerList.append(front_name)
                innerList.append(terminal_name)
                innerList.append(company)
                innerList.append(start_time)
                innerList.append(end_time)
                innerList.append(length)
                innerList.append(float(x[j]))
                innerList.append(float(y[j]))
                dataList.append(innerList)

    return dataList


def getbusstation(linename, datajson):
    dataList = []
    try:
        datajson = datajson['busline_list'][0]
    except:
        innerList = []
        innerList.append("")
        innerList.append(linename)
        innerList.append("")
        innerList.append("")
        innerList.append("")
        innerList.append("")
        innerList.append("")
        innerList.append("")
        innerList.append("")
        dataList.append(innerList)
        return dataList
    else:
        lid = datajson['id']
        try:
            datajson = datajson['stations']
        except:
            innerList = []
            innerList.append("")
            innerList.append(linename)
            innerList.append("")
            innerList.append("")
            innerList.append("")
            innerList.append("")
            innerList.append("")
            innerList.append("")
            innerList.append("")
            dataList.append(innerList)
            return dataList
        else:
            for i in range(len(datajson)):
                # print("t1",datajson)
                datajson1 = datajson[i]
                sname = datajson1['name']
                station_num = datajson1['station_num']
                try:
                    poiid1 = datajson1['poiid1']
                except:
                    poiid1 = ""
                xy_coords = datajson1['xy_coords']
                # poiid2_xy = datajson1['poiid2_xy']
                innerList = []
                # 每个innerList存储一对数据
                x = xy_coords.split(';')[0]
                y = xy_coords.split(';')[1]
                # poix = poiid2_xy.split(';')[0]
                # poiy = poiid2_xy.split(';')[1]
                innerList.append(lid)
                innerList.append(linename)
                innerList.append(sname)
                innerList.append(station_num)
                innerList.append(poiid1)
                innerList.append(x)
                innerList.append(y)
                # innerList.append(poix)
                # innerList.append(poiy)
                dataList.append(innerList)
            return dataList


start_time = time.time()
poi_bus_url = "https://amap.com/service/poiInfo?query_type=TQUERY&city=310000&keywords="
read_file_dir = "地铁线路.xls"  # 公交线路名
ll = 1
tt = 0
nn = 1
nnn = 1
init(ll, tt)
no_sheet = 0
myWordbookr = xlrd.open_workbook(read_file_dir)
mySheetsr = myWordbookr.sheets()
mySheetr = mySheetsr[no_sheet]
busname = "地铁线路坐标" + time.strftime("%m%d%H%M", time.localtime())
busstation = "地铁站点坐标" + time.strftime("%m%d%H%M", time.localtime())
# 获取列数
nrows = mySheetr.nrows
book = xlwt.Workbook(encoding='utf-8', style_compression=0)
sheet = book.add_sheet("1", cell_overwrite_ok=True)
sheet.write(0, 0, 'id')
sheet.write(0, 1, 'name')
sheet.write(0, 2, 'code')
sheet.write(0, 3, 'front_name')
sheet.write(0, 4, 'terminal_name')
sheet.write(0, 5, 'company')
sheet.write(0, 6, 'start_time')
sheet.write(0, 7, 'end_time')
sheet.write(0, 8, 'length')
sheet.write(0, 9, 'lon')
sheet.write(0, 10, 'lat')
book1 = xlwt.Workbook(encoding='utf-8', style_compression=0)
sheet1 = book1.add_sheet("station", cell_overwrite_ok=True)
sheet1.write(0, 0, 'lid')
sheet1.write(0, 1, 'lname')
sheet1.write(0, 2, 'sname')
sheet1.write(0, 3, 'station_num')
sheet1.write(0, 4, 'poiid1')
sheet1.write(0, 5, 'x')
sheet1.write(0, 6, 'y')
sheet1.write(0, 7, 'poix')
sheet1.write(0, 8, 'poiy')

k = 0
kk = 0
m = 0
mm = 0
for i in range(1, nrows):
    linename = mySheetr.cell_value(i, 0)
    print(33)
    linetext = getjson(linename)
    linedata = getbusline(linename, linetext)
    linestation = getbusstation(linename, linetext)
    tt = tt + 1
    print("已读取:", tt, i, "/", nrows)
    for j in range(len(linedata)):
        sheet.write(k + 1 - kk, 0, linedata[j][0])
        sheet.write(k + 1 - kk, 1, linedata[j][1])
        sheet.write(k + 1 - kk, 2, linedata[j][2])
        sheet.write(k + 1 - kk, 3, linedata[j][3])
        sheet.write(k + 1 - kk, 4, linedata[j][4])
        sheet.write(k + 1 - kk, 5, linedata[j][5])
        sheet.write(k + 1 - kk, 6, linedata[j][6])
        sheet.write(k + 1 - kk, 7, linedata[j][7])
        sheet.write(k + 1 - kk, 8, linedata[j][8])
        sheet.write(k + 1 - kk, 9, linedata[j][9])
        sheet.write(k + 1 - kk, 10, linedata[j][10])
        k = k + 1
    for jj in range(len(linestation)):
        sheet1.write(m + 1 - mm, 0, linestation[jj][0])
        sheet1.write(m + 1 - mm, 1, linestation[jj][1])
        sheet1.write(m + 1 - mm, 2, linestation[jj][2])
        sheet1.write(m + 1 - mm, 3, linestation[jj][3])
        sheet1.write(m + 1 - mm, 4, linestation[jj][4])
        sheet1.write(m + 1 - mm, 5, linestation[jj][5])
        sheet1.write(m + 1 - mm, 6, linestation[jj][6])
        # sheet1.write(m + 1 - mm, 7, linestation[jj][7])
        # sheet1.write(m + 1 - mm, 8, linestation[jj][8])
        m = m + 1
    book.save(r'' + busname + '.xls')
    book1.save(r'' + busstation + '.xls')
    time.sleep(2)
    # 加入延时，否则爬取太快，容易导致错误
    if tt % 10 == 0:
        time.sleep(10)
        print(k)
        if k > 60000 and k - 60000 * nn > 0:
            nn = nn + 1
            print(nn)
            sheet = book.add_sheet(str(nn), cell_overwrite_ok=True)
            sheet.write(0, 0, 'id')
            sheet.write(0, 1, 'name')
            sheet.write(0, 2, 'code')
            sheet.write(0, 3, 'front_name')
            sheet.write(0, 4, 'terminal_name')
            sheet.write(0, 5, 'company')
            sheet.write(0, 6, 'start_time')
            sheet.write(0, 7, 'end_time')
            sheet.write(0, 8, 'length')
            sheet.write(0, 9, 'lon')
            sheet.write(0, 10, 'lat')
            kk = k
        if m > 60000 and m - 60000 * nnn > 0:
            nnn = nnn + 1
            print(nnn)
            sheet1 = book1.add_sheet(str(nn), cell_overwrite_ok=True)
            sheet1.write(0, 0, 'lid')
            sheet1.write(0, 1, 'lname')
            sheet1.write(0, 2, 'sname')
            sheet1.write(0, 3, 'station_num')
            sheet1.write(0, 4, 'poiid1')
            sheet1.write(0, 5, 'x')
            sheet1.write(0, 6, 'y')
            sheet1.write(0, 7, 'poix')
            sheet1.write(0, 8, 'poiy')
            mm = m
    if tt % 199 == 0:
        driver.quit()
        book.save(r'' + busname + '.xls')
        book1.save(r'' + busstation + '.xls')
        end_time = time.time()
        print('完成爬取' + str(ll) + '次数，浏览器重启,已用时%.2f秒' % (end_time - start_time))
        init(ll, tt)
        ll = ll + 1

end_time = time.time()
print('写入成功,总用时%.2f秒' % (end_time - start_time))
book.save(r'' + busname + '.xls')
book1.save(r'' + busstation + '.xls')
driver.quit()
