# Xpath与正则一齐与selenium使用
# 点击browser.execute_script('arguments[0].click',button)
import selenium
from selenium import webdriver
import time
from selenium.webdriver import ChromeOptions
import pandas as pd
import random

# 需要筛选数据读入与拆分
filename = ''
data = pd.read_excel(filename, header=0)
data2 = data['产品名称矫正'].drop_duplicates()[0:20]
data3 = data['产品名称矫正'].drop_duplicates()[20:40]
data4 = data['产品名称矫正'].drop_duplicates()[40:60]
data5 = data['产品名称矫正'].drop_duplicates()[60:80]
data6 = data['产品名称矫正'].drop_duplicates()[80::]
db = [data2, data3, data4, data5, data6]


def get_message(Path, path):  # 填入列表函数
    for txt in path:
        Path.append(txt.text)
    return Path


def all_message(browser, number, id, password, datafile):  # 获取所有信息

    # 需要的变量
    Names = []  # 产品名
    Company = []  # 公司名
    Regular = []  # 规格
    Package = []  # 包装
    Middlepackage = []  # 中包
    Produce_Date = []  # 生产日期
    Save = []  # 库存
    Ending_Date = []  # 失效时间
    Comman_price = []  # 国家定价
    Price = []  # 售价
    Discount_price = []  # 折扣价
    Charge = []  # 单位

    # 进网址
    main_url = 'https://www.hyey.cn'
    browser.get(main_url)

    # 登录
    input_account = browser.find_element_by_id('UserName')

    input_account.send_keys(id())
    Password = browser.find_element_by_id('PassWord')
    Password.send_keys(password)
    button1 = browser.find_element_by_id('loginbtn')  # 获取点击按钮
    button1.click()  # 点击登录
    time.sleep(5)
    button2 = browser.find_element_by_class_name('layui-layer-btn0')  # 获取点击按钮
    button2.click()  # 点击确定
    time.sleep(5)

    # 进入列表循环
    for medicine in datafile:

        i = 1
        j = 1

        # 搜索
        search_xpath = '/html/body/form/div[4]/div[3]/input[1]'
        search = browser.find_element_by_xpath(search_xpath)
        search.clear()  # 清空搜索框的文字
        time.sleep(random.randrange(2, 8))
        x = f'{medicine}'
        print(f'开始爬取{x}')
        time.sleep(random.randrange(2, 5))
        search.send_keys(x)
        search_botton = browser.find_element_by_id('Button1')  # 获取点击按钮
        search_botton.click()  # 点击搜索
        time.sleep(random.randrange(3, 5))
        list_botton = browser.find_element_by_id('p1')  # 转换成为列表形式
        list_botton.click()
        time.sleep(random.randrange(2, 8))

        try:
            # 进入循环
            while j < number:
                print('开始第%d次循环' % j)
                print(f'开始爬取第{i}页')
                time.sleep(random.randrange(5, 10))

                # 获取信息
                time.sleep(random.randrange(5, 8))
                time.sleep(random.randrange(2, 5))
                main_xpath = '/html/body/form/div[8]/div/div[5]/table[2]/tr/'
                name_xpath = main_xpath + 'td[1]/span[2]'
                comapany_xpath = main_xpath + 'td[2]/span'
                regular_xpath = main_xpath + 'td[3]/span'
                package_xpath = main_xpath + 'td[4]'
                middlepackage_xpath = main_xpath + 'td[5]'
                produce_xpath = main_xpath + 'td[6]'
                ending_xpath = main_xpath + 'td[7]'
                save_xpath = main_xpath + 'td[8]'
                comman_price_xpath = main_xpath + 'td[9]'
                price_xpath = main_xpath + 'td[10]/span'
                discountprice_xpath = main_xpath + 'td[10]/span'
                charge_xpath = main_xpath + 'td[13]'

                # 解析为element对象
                name_nodes = browser.find_elements_by_xpath(name_xpath)
                regular_nodes = browser.find_elements_by_xpath(regular_xpath)
                package_nodes = browser.find_elements_by_xpath(package_xpath)
                middlepackage_nodes = browser.find_elements_by_xpath(middlepackage_xpath)
                produce_nodes = browser.find_elements_by_xpath(produce_xpath)
                save_notes = browser.find_elements_by_xpath(save_xpath)
                ending_notes = browser.find_elements_by_xpath(ending_xpath)
                comman_price_notes = browser.find_elements_by_xpath(comman_price_xpath)
                price_notes = browser.find_elements_by_xpath(price_xpath)
                discountprice_notes = browser.find_elements_by_xpath(discountprice_xpath)
                charge_notes = browser.find_elements_by_xpath(charge_xpath)
                comapany_notes = browser.find_elements_by_xpath(comapany_xpath)

                # 填充
                get_message(Names, name_nodes)  # 提取药名
                get_message(Company, comapany_notes)  # 提取药名
                get_message(Regular, regular_nodes)  # 提取规格
                get_message(Package, package_nodes)  # 提取包装
                get_message(Middlepackage, middlepackage_nodes)  # 提取中包
                get_message(Produce_Date, produce_nodes)  # 提取生产日期
                get_message(Save, save_notes)  # 提取库存
                get_message(Ending_Date, ending_notes)  # 提取失效时间
                get_message(Comman_price, comman_price_notes)  # 提取国家定价
                get_message(Price, price_notes)
                get_message(Discount_price, discountprice_notes)  # 提取折扣价
                get_message(Charge, charge_notes)  # 提取包装
                time.sleep(2)

                # 滚动页面滚动条到最底部
                print('将页面拉到底部')
                browser.execute_script("window.scrollTo(0,1000);")
                time.sleep(random.randrange(5, 8))
                print('抓取第{}页结束，已抓取对象{}个'.format(i, len(Names)))

                # 开始转换页码
                print('开始转换页码')
                i = i + 1
                j += 1
                print('开始转换到第%d页' % i)
                time.sleep(random.randrange(3, 6))
                page_xpath = '/html/body/form/div[8]/div/div[7]/div/div[1]/span[2]/span[2]/span/input[2]'
                page = browser.find_element_by_xpath(page_xpath)
                nextpage = browser.find_element_by_xpath(page_xpath).get_attribute('value')

                # 判断页码
                if str(i) == nextpage:
                    page.send_keys('\n')  # 点击搜索
                    time.sleep(random.randrange(5, 10))
                else:
                    print('已是最后一页')
                    break

            # 将页面拉到最上方
            print("将页面拉到最上方")
            browser.execute_script("window.scrollTo(0,0);")
            time.sleep(random.randrange(5, 10))

        except Exception as result:
            print(result)

        # 结束提醒
        print(f'{medicine}抓取完成')
        print('*' * 30)
        print('❀' * 30)

    # 关闭浏览器
    browser.close()
    browser.quit()
    print('全部抓取完成')
    print('共抓取对象%d个' % len(Names))
    return Names, Company, Regular, Package, Middlepackage, Produce_Date, Save, Ending_Date, Comman_price, Price, Discount_price, Charge


# 编造创建表格函数
def creat_DataFrame(df, filename):
    database = pd.DataFrame(df).T  # 对生成列表进行转置
    database.rename(
        columns={0: '药品名称', 1: '生产公司', 2: '规格', 3: '包装', 4: '中包', 5: '生产日期', 6: '库存',
                 7: '失效日期', 8: '国家定价', 9: '售价',10: '折扣价', 11: '单位'}, inplace=True)
    database.to_excel(filename, encoding='utf-8')
    print('表格写入成功')


# 设置主程序
def main():
    number = 30
    # 设置浏览器
    option = ChromeOptions()
    option.add_experimental_option('excludeSwitches', ['enable-automation'])
    browser = selenium.webdriver.Chrome()
    time.sleep(10)
    # 爬取信息
    for data_n in db:
        try:
            j = 1
            data = all_message(browser, number, data_n)
            time.sleep(5)
            # 建表保存
            filename = f'C:/m1553/Users/Desktop/2021.08.26华源药品（{j}）'
            creat_DataFrame(data, filename)
            print(f'{filename}表格创建成功')
            j += 1
        except Exception as result:
            print(result)
    print('All Progress has been Done')


# 运行
if __name__ == '__main__':
    main()

# 数据合并
# 数据读入
df1 = pd.read_excel(r'C:/m1553/Users/Desktop/2021.08.26华源药品（1）.xlsx')
df2 = pd.read_excel(r'C:/m1553/Users/Desktop/2021.08.26华源药品（2）.xlsx')
df3 = pd.read_excel(r'C:/m1553/Users/Desktop/2021.08.26华源药品（3）.xlsx')
df4 = pd.read_excel(r'C:/m1553/Users/Desktop/2021.08.26华源药品（4）.xlsx')
df5 = pd.read_excel(r'C:/m1553/Users/Desktop/2021.08.26华源药品（5）.xlsx')
df6 = pd.read_excel(r'C:/m1553/Users/Desktop/华源清单-给高旭.xlsx')
del df3['Unnamed: 0']
del df2['Unnamed: 0']
del df4['Unnamed: 0']
del df5['Unnamed: 0']

# 数据合并
database = pd.concat([df1, df2, df3, df4, df5])
database.reset_index(inplace=True)
del database['index']

# 筛选无信息种类
correct_list = []
for j in database['药品名称'].drop_duplicates():
    correct_list.append(j)

nomessage_list = []
for i in df6['产品名称矫正']:
    print(f'开始检测{i}是否存在')
    if i not in correct_list:
        print('无此药品相关信息')
        nomessage_list.append(i)
    else:
        print('此药品有相关信息')
    print('❀❀' * 30)

    # 将无信息种类制作成DateFrame
nomessage_df = pd.DataFrame(nomessage_list)
nomessage_df.rename(columns={0: '药品名称'}, inplace=True)

# 保存数据
database.to_excel('C:/m1553/Users/Desktop/2021.08.26华源药品爬取.xlsx', encoding='utf-8')
nomessage_df.to_excel('C:/m1553/Users/Desktop/2021.08.26华源药品爬取（未存在).xlsx', encoding='utf-8')
