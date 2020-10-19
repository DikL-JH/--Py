import os

from selenium import webdriver
from time import sleep
from selenium.webdriver.support.select import Select
from pptx import Presentation
from os import listdir
import re
import pptx
import random
import shutil

driver = webdriver.Chrome("chromedriver.exe")  # chromedriver所在路径
driver.get(r"https://pptadmin.ahgegu.cn/")  # 目标网址
# C:\\Users\\admin0001\\Downloads
path = 'C:\\LL'  # ppt文件所在位置


# 登录shcksm1234    09135340     niusihui
# us=input("请输入用户名")
# psw= input("请输入需密码")
def login():
    data1 = ''
    data2 = ''
    typee = ""  # 文件类型
    driver.find_element_by_id("loginform-username").send_keys("yunying003")  # 输入用户名
    driver.find_element_by_id("loginform-password").send_keys('09135340')  # 输入密码
    driver.find_element_by_name("login-button").click()  # 点击“登录”
    sleep(3)
    # driver.implicitly_wait(20)
    driver.find_element_by_link_text("模板管理").click()
    sleep(0.3)
    driver.find_element_by_link_text("模板列表").click()
    sleep(0.2)

    # driver.find_element_by_link_text("添加模板").click()

    files = listdir(path)
    for file in files:

        try:
            driver.switch_to.frame(driver.find_element_by_xpath("//*[@id='content-main']/iframe[2]"))  # 切换到iframe中
            driver.find_element_by_link_text("添加模板").click()
            Path = path + '\\' + file
            files2 = listdir(Path)
            for file2 in files2:
                Path2 = Path + '\\' + file2
                # print(Path2)
                pattern = re.compile(r'([^<>/\\\|:""\*\?]+)\.\w+$')
                data = pattern.findall(Path2)
                # print(data)
                # print (str(file2) +"  1")
                if '.jpg' in file2:
                    # print(str(data[0])+'    2')
                    data1 = data[0]
                elif '.pptx' in file2:
                    # print(str(data[0])+'    3')
                    data2 = data[0]
                    typee = ".pptx"
                elif '.ppt' in file2:
                    # print(str(data[0])+'    4')
                    data2 = data[0]
                    typee = ".ppt"
            p = pptx.Presentation(Path + "\\" + str(data2) + typee)
            page = len(p.slides)
            print(Path)
            driver.find_element_by_xpath('//*[@id="goods-name"]').send_keys(str(data2))  # 名称
            driver.find_element_by_xpath('//*[@id="w0"]/input[2]').send_keys(Path + "\\" + str(data1) + ".jpg")  # 上传封面图
            driver.find_element_by_xpath('//*[@id="w0"]/input[3]').send_keys(Path + "\\1.mp4")  # 上传视频
            driver.find_element_by_xpath('//*[@id="w0"]/input[4]').send_keys(Path + "\\" + str(data2) + typee)  # 上传文件
            driver.find_element_by_xpath('//*[@id="goods-page_num"]').send_keys(str(page))  # 页数
            rad = random.randint(2, 99)
            author = ['深海苏眉鱼', '寄梦中人', '雨诺潇潇', '如鲸向海', '那一片橙海', '明若轻兮', '弹指间的背影', '徘徊月下', '你的轮廓', '送你一个梦', '淡水深流',
                      '半寸时光', '菩提树下叶撕阳', '时光踏路已久', '一曲爱恨情仇', '一枕庭前雪', '策马西风', '黒涩兲箜', '俯瞰星空', '浮生未歇',
                      '泪落旧城', '四叶星光', '时光忆年少', '四重梦境', '第四重梦境', '梦与她', '春的气息', '寻沫雨悠扬', '移梦别嫁', '梦幻的心爱', '烟锁重楼', '清故宸凉',
                      '等我变成光', '清风无痕', '忆蝶梦寒', '珠帘湿罗幕', '初与友歌', '北巷南猫', '沽酒醉风尘', '若水微香',
                      '轮回亦思伊人', '醉落夕风', '江枫思渺然', '雨的印迹', '锁上的光', '梦初启', '茉莉花茶芳香', '沫染流年', '美人痣', '你是暖光', '旧梦荧光笔',
                      '尘世孤行', '素子花开', '星星的軌跡', '云端的琴声', '空袭的梦', '遥远的她', '沉香未言墨竹', '海的颜色', '时光小偷',
                      '执意画红尘', '逆流伏景', '堇色安年', '梦在深巷', '时光沙漏', '海蓝无魂', '黎夕旧梦', '以梦之名，浅浅低吟。', '树影摇曳。', '雨夜梧桐', '痴人痴梦',
                      '兩夢三醒', '树深时见鹿', '夜城月下', '画扇描眉', '巴黎的余音', '落雪听梅', '花开半夏', '旧日阳光', '近似海水明',
                      '指尖上得阳光', '深雨燕紛飛', '朱唇点点醉', '夜阑听雪', '梦醒时光', '浮华皆是空', '荼谧故人', '七秒鱼忆', '繁花落曲的背影', '深海沫深', '潇魂蝶舞',
                      '深府石板幽径', '珠帘湿罗幕', '长亭外古道边', '拥友存忆', '时光踏路已久', '眼泪也成诗', '半窗疏影', '伊人憔悴', '素手琵琶']
            driver.find_element_by_xpath('//*[@id="goods-author"]').send_keys(str(author[rad]))  # 作者
            driver.find_element_by_xpath('//*[@id="goods-content"]').send_keys(
                "这是一套" + str(data2) + ",文字图片可直接编辑，操作简单方便,我们还有更多精美工作汇报，健康养生，教育培训，社团招新PPT模板尽在这里，欢迎下载。")
            s = [('教育', '教学', '培训', '课', '简历', '个人简介'),
                 ('互联网', '科技', '网络', '计算', '物联网', '软件', '开发', '大数据'),
                 ('企业', '宣传', '公司', '公司介绍', '快闪', '产品'),
                 ('美食', '月饼', '茶', '牛排', '美味', '糕点', '菜', '餐'),
                 ('影视传媒', '影视', '传媒', '广告', '摄影', '电视', '电影'),
                 ('绿色', '环保', '低碳', '植树', '生态', '节能', '减排'),
                 ('地产', '旅游', '画册', '景点', '旅行', '房产'),
                 ('室内', '装修', '家具', '建筑', '家居', '装潢'),
                 ('政', '党', '税务', '警', '扫黑', '军', '公安', '消防', '国防', '国庆'),
                 ('金融', '财务', '会计', '采购', '理财'),
                 ('婚', '七夕', '爱情', '恋爱', '告白', '520'),
                 ('交通', '物流', '运输', '快递', '航空', '航天', '民航', '海运', '船运', '指路', '高铁', '铁路'),
                 ('体育', '运动', '健身', '篮球', '足球', '排球', '健美'),
                 ('公益', '爱心', '健身', '环保', '禁毒'),
                 ('农业', '养殖', '有机', '蔬', '种植'),
                 ('简约', '简单', '简洁'),
                 ('商务', '商用', '企业', '宣传', '公司', '公司介绍', '快闪', '产品'),
                 ('清新', '淡雅'),
                 ('中国风', '中国风'),
                 ('复古', '复古'),
                 ('扁平', '扁平'),
                 ('立体', '立体'),
                 ('可爱', '卡通'),
                 ('手绘', '手绘'),
                 ('欧美', '欧美')]
            for x, m in enumerate(s, start=1):
                for o in m:
                    # if x<14 :
                    # if o in :
                    #   driver.find_element_by_xpath('//*[@id="w0"]/div[18]/div[1]/div[1]/div[2]/label['+str(x)+']/input').click()
                    #  break
                    if x < 16:
                        if o in data1:
                            driver.find_element_by_xpath(
                                '//*[@id="w0"]/div[18]/div[1]/div[2]/div[2]/label[' + str(x) + ']/input').click()
                            break
                    else:
                        if o in data1:
                            driver.find_element_by_xpath(
                                '//*[@id="w0"]/div[18]/div[1]/div[3]/div[2]/label[' + str((x - 15)) + ']/input').click()
                            break
            for i in range(10 if page > 10 else page):
                driver.find_element_by_xpath('//*[@id="w0"]/input[5]').send_keys(
                    Path + "\\幻灯片" + str(i + 1) + ".JPG")  # 上传内容图
                sleep(0.1)
        except Exception as e:
            print(e)
            print("**************************************")
            print("*数据异常--请手动处理--按回车继续处理*")
            input("**************************************")

            driver.find_element_by_xpath('//*[@id="w0"]/div[19]/button').click()  # 提交
            # print("************")
            print("*--预览页--*")
            input("************")
            shutil.rmtree(Path)
            driver.back()
            sleep(1)
            driver.back()
            sleep(1)
        else:
            print("****************************")
            print("*请核查数据--按回车继续处理*")
            input("****************************")

            driver.find_element_by_xpath('//*[@id="w0"]/div[19]/button').click()  # 提交
            # print("************")
            print("*--预览页--*")
            input("************")
            shutil.rmtree(Path)
            driver.back()
            sleep(1)
            driver.back()
            sleep(1)


login()
