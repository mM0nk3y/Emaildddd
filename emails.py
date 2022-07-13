from turtle import title
import config
from selenium import webdriver
import os
import re
import sys
import time
import logging
from datetime import datetime
import warnings
import argparse


warnings.filterwarnings("ignore",category=DeprecationWarning)
download_location = config.download_path
now = datetime.now()

# 运行日志，写入到文件
logger = logging.getLogger()
logger.setLevel(logging.INFO)
logfile = download_location + now.strftime("%Y-%m-%d-%H点%M分") + 'log.txt'
fm = logging.FileHandler(logfile,mode='a') # a 写入文件，若文件不存在则会先创建再写入，但不会覆盖原文件，而是追加在文件末尾
fm.setLevel(logging.INFO)
# 运行日志，打印在控制台
th = logging.StreamHandler()
th.setLevel(logging.INFO)

formatter = logging.Formatter("%(asctime)s - %(filename)s[line:%(lineno)d] - %(levelname)s : %(message)s")
fm.setFormatter(formatter)
th.setFormatter(formatter)

logger.addHandler(fm)
logger.addHandler(th)

def parse_args():
    parser = argparse.ArgumentParser(epilog='\t用法举例:\r\npython ' + sys.argv[0] + ' -ae')
    parser.add_argument("-ae", "--allExp",help="Export all emails, 导出用户所有邮件",default='False', action='store_true')
    parser.add_argument("-pe", "--pageExp",type=int,help="Export by page number, 根据页码导出, 如只导出前3页, -pe 3")
    parser.add_argument("-te", "--timeExp",type=str, dest='timeinput', nargs='+', help="Export by time, 根据邮件时间导出，\r\n 如导出2020年12月1日 07:30的邮件\r\n -te 2020-12-1-07-30")
    return parser.parse_args()

def date2timestamp(s):
    year_s, mon_s, day_s,hour_s,mint_s = s.split('-')
    format_time =  datetime(int(year_s), int(mon_s), int(day_s),int(hour_s),int(mint_s)).strftime("%Y-%m-%d %H:%M")
    dt = datetime.strptime(format_time,'%Y-%m-%d %H:%M')
    f_time = dt.timestamp()
    return f_time
    
def allExp(username,password,proxy):
#def allExp(username,password):
    # print('[+]' + str(datetime.datetime.now()) + ' 开始导出邮件 [+]')
    logger.info('[+]' + str(datetime.now()) + ' 开始导出邮件 [+]')
    prefs = {'download.default_directory': download_location,
    'download.prompt_for_download': False,
    'download.directory_upgrade': True,
    'safebrowsing.enabled': False,
    'safebrowsing.disable_download_protection': True,
    'profile.default_content_settings.popups': 0}
    option = webdriver.ChromeOptions()
    option.add_experimental_option('prefs',prefs)
    option.add_experimental_option('excludeSwitches', ['enable-logging'])
    option.add_argument('--proxy-server={}'.format(proxy))
    option.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; WOW64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4861.114 Safari/537.36')
    option.add_argument('--enable-blink-features=ExperimentalProductivityFeatures')
    option.add_argument('--enable-blink-features=Serial')
    option.add_argument('--enable-blink-features=ConversionMeasurement')
    option.add_argument('--ignore-certificate-errors')
    option.add_argument('--headless')
    driver = webdriver.Chrome(options=option)
    driver.set_window_size(1024,7680)
    driver.get('http://mail.126.com')
    print(driver.page_source)
    time.sleep(30)

    # 登陆
    iframe = driver.find_elements_by_tag_name("iframe")[0]
    driver.switch_to.frame(iframe)
    time.sleep(5)
    driver.find_element_by_name("email").clear()
    driver.find_element_by_name("email").send_keys(username)
    driver.find_element_by_name("password").clear()
    driver.find_element_by_name("password").send_keys(password)
    driver.find_element_by_id("dologin").click()
    time.sleep(2)
    driver.switch_to.default_content() # 跳出iframe 。必定记住 只要以前跳入iframe，以后就必须跳出。进入主html
    driver.find_element_by_xpath('//*[@id="_mail_component_147_147"]/span[2]').click()  # 点击收件箱
    time.sleep(3)
    total_page = int(driver.find_element_by_class_name('nui-select-text').text.split('/')[1]) # 获取收件箱页数
    logger.info("邮件总页数:{}".format(total_page))
    for i in range(0,total_page):
        total_emails = driver.find_elements_by_class_name('da0')
        emails_count = len(total_emails)
        logger.info("第{}页邮件总数: ".format(i+1) + str(emails_count))
        for k in range(0,emails_count):
            email_title = driver.find_elements_by_class_name('da0')[k].text
            # driver.find_element_by_class_name('nui-ico-checkbox').click() # 单个选中
            # driver.find_elements_by_class_name('js-component-button')[7].click() # 展开更多
            # driver.find_elements_by_class_name('nui-menu-item-text')[0].click() # 导出选中
            logger.info("导出第{}封邮件，标题: {}".format(k+1,email_title))

        email_time = driver.find_elements_by_class_name('eO0')[0].text
        # print(email_time)
        logger.info(email_time)
        driver.find_element_by_class_name('nui-dropdownBtn-hasOnlyIcon').click() # 全选
        time.sleep(1)
        driver.find_elements_by_class_name('js-component-button')[7].click() # 展开更多
        driver.find_elements_by_class_name('nui-menu-item-text')[0].click() # 导出选中
        time.sleep(2)
        for fname in os.listdir(download_location):
            if ('信件打包.zip') in fname:
                # print('邮件导出中 ...')
                logger.info('邮件导出中 ...')
                time.sleep(1)
                try:
                    os.rename(download_location + '信件打包.zip',download_location + username + '-' + now.strftime("%Y-%m-%d-%H点%M分") + "第" + str(i+1) +"页邮件.zip")
                except:
                    pass
        logger.info('[+]第' + str(i+1) +'页所有邮件已经下载完成！[+]')

        time.sleep(1)
        driver.find_element_by_class_name('nui-ico-next').click() # 下一页
        time.sleep(1)

    logger.info('[+]' + str(datetime.now()) + '该用户的所有邮件已经全部下载完成，即将下载下一个用户的所有邮件[+]')
    driver.quit()
def pageExp(username,password,proxy):
#def pageExp(username,password):
    # print('[+]' + str(datetime.datetime.now()) + ' 开始导出邮件 [+]')
    logger.info('[+]' + str(datetime.now()) + ' 开始导出邮件 [+]')
    prefs = {'download.default_directory': download_location,
    'download.prompt_for_download': False,
    'download.directory_upgrade': True,
    'safebrowsing.enabled': False,
    'safebrowsing.disable_download_protection': True,
    'profile.default_content_settings.popups': 0}
    option = webdriver.ChromeOptions()
    option.add_experimental_option('prefs',prefs)
    option.add_experimental_option('excludeSwitches', ['enable-logging'])
    option.add_argument('--proxy-server={}'.format(proxy))
    option.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; WOW64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4861.114 Safari/537.36')
    option.add_argument('--enable-blink-features=ExperimentalProductivityFeatures')
    option.add_argument('--enable-blink-features=Serial')
    option.add_argument('--enable-blink-features=ConversionMeasurement')
    option.add_argument('--ignore-certificate-errors')
    option.add_argument('--headless')
    driver = webdriver.Chrome(options=option)
    driver.set_window_size(1024,7680)
    driver.get('http://mail.126.com')

    time.sleep(10)

    # 登陆
    iframe = driver.find_elements_by_tag_name("iframe")[0]
    driver.switch_to.frame(iframe)
    time.sleep(5)
    driver.find_element_by_name("email").clear()
    driver.find_element_by_name("email").send_keys(username)
    driver.find_element_by_name("password").clear()
    driver.find_element_by_name("password").send_keys(password)
    driver.find_element_by_id("dologin").click()
    time.sleep(2)
    driver.switch_to.default_content() # 跳出iframe 。必定记住 只要以前跳入iframe，以后就必须跳出。进入主html
    driver.find_element_by_xpath('//*[@id="_mail_component_147_147"]/span[2]').click()  # 点击收件箱
    time.sleep(3)
    total_page = int(driver.find_element_by_class_name('nui-select-text').text.split('/')[1]) # 获取收件箱页数
    logger.info("邮件总页数:{}, 当前导出前 {}页邮件".format(total_page, args.pageExp))
    if args.pageExp > int(total_page):
        logger.info("[+] 用户{}导出页数大于实际邮件页数, 将按照用户实际页数全部导出！[+]".format(username))
        allExp(username,password)
    else:
        for i in range(0,args.pageExp):
            total_emails = driver.find_elements_by_class_name('da0')
            emails_count = len(total_emails)
            logger.info("第{}页邮件总数: ".format(i+1) + str(emails_count))
            for k in range(0,emails_count):
                email_title = driver.find_elements_by_class_name('da0')[k].text
                # driver.find_element_by_class_name('nui-ico-checkbox').click() # 单个选中
                # driver.find_elements_by_class_name('js-component-button')[7].click() # 展开更多
                # driver.find_elements_by_class_name('nui-menu-item-text')[0].click() # 导出选中
                logger.info("导出第{}封邮件，标题: {}".format(k+1,email_title))

            email_time = driver.find_elements_by_class_name('eO0')[0].text
            # print(email_time)
            logger.info("[+] 接收邮件的日期：{}".format(email_time))
            driver.find_element_by_class_name('nui-dropdownBtn-hasOnlyIcon').click() # 全选
            time.sleep(1)
            driver.find_elements_by_class_name('js-component-button')[7].click() # 展开更多
            driver.find_elements_by_class_name('nui-menu-item-text')[0].click() # 导出选中
            time.sleep(2)
            for fname in os.listdir(download_location):
                if ('信件打包.zip') in fname:
                    # print('邮件导出中 ...')
                    logger.info('邮件导出中 ...')
                    time.sleep(1)
                    try:
                        os.rename(download_location + '信件打包.zip',download_location + username + '-' + now.strftime("%Y-%m-%d-%H点%M分") + "第" + str(i+1) +"页邮件.zip")
                    except:
                        pass
            logger.info('[+]第' + str(i+1) +'页所有邮件已经下载完成！[+]')

            if args.pageExp == i+1:
                logger.info('[+] 邮件导出完成！ [+]')
                
            

            time.sleep(1)
            driver.find_element_by_class_name('nui-ico-next').click() # 下一页
            time.sleep(1)

        logger.info('[+]' + str(datetime.now()) + '该用户的所有邮件已经全部下载完成，即将下载下一个用户的所有邮件[+]')
        driver.quit()
def timeExp(username,password,proxy):
#def timeExp(username,password):
    # print('[+]' + str(datetime.datetime.now()) + ' 开始导出邮件 [+]')
    logger.info('[+]' + str(datetime.now()) + ' 开始导出邮件 [+]')
    prefs = {'download.default_directory': download_location,
    'download.prompt_for_download': False,
    'download.directory_upgrade': True,
    'safebrowsing.enabled': False,
    'safebrowsing.disable_download_protection': True,
    'profile.default_content_settings.popups': 0}
    option = webdriver.ChromeOptions()
    option.add_experimental_option('prefs',prefs)
    option.add_experimental_option('excludeSwitches', ['enable-logging'])
    option.add_argument('--proxy-server={}'.format(proxy))
    option.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; WOW64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4861.114 Safari/537.36')
    option.add_argument('--enable-blink-features=ExperimentalProductivityFeatures')
    option.add_argument('--enable-blink-features=Serial')
    option.add_argument('--enable-blink-features=ConversionMeasurement')
    option.add_argument('--ignore-certificate-errors')
    option.add_argument('--headless')
    driver = webdriver.Chrome(options=option)
    driver.set_window_size(1024,7680)
    driver.get('http://mail.126.com')
    time.sleep(3)

    # 登陆
    iframe = driver.find_elements_by_tag_name("iframe")[0]
    driver.switch_to.frame(iframe)
    driver.find_element_by_name("email").clear()
    driver.find_element_by_name("email").send_keys(username)
    driver.find_element_by_name("password").clear()
    driver.find_element_by_name("password").send_keys(password)
    driver.find_element_by_id("dologin").click()
    time.sleep(2)
    driver.switch_to.default_content() # 跳出iframe 。必定记住 只要以前跳入iframe，以后就必须跳出。进入主html
    driver.find_element_by_xpath('//*[@id="_mail_component_147_147"]/span[2]').click()  # 点击收件箱
    time.sleep(3)
    total_page = int(driver.find_element_by_class_name('nui-select-text').text.split('/')[1]) # 获取收件箱页数
    logger.info("邮件总页数:{}".format(total_page))
    

    for i in range(0,total_page):
        total_emails = driver.find_elements_by_class_name('da0')
        emails_count = len(total_emails)
        logger.info("第{}页邮件总数: ".format(i+1) + str(emails_count))
        for k in range(0,emails_count):
            
            email_time = driver.find_elements_by_class_name('eO0')[k].get_attribute('title')  # 获取收件时间
            # 比较时间戳
            pattern=re.compile('\d+\.?\d*')
            result=pattern.findall(email_time)
            s_time = '-'.join(result)
            ss_time = date2timestamp(s_time)  # 获取邮件的时间戳
            in_time = '-'.join(args.timeinput)
            n_time = date2timestamp(in_time) # 获取当前输入的时间戳
            if n_time > ss_time:
                # 邮件接收时间越早，ss_time 值越小
                logger.info("暂无邮件更新！暂无接收到早于{}这个日期的邮件".format(in_time))
            if n_time <= ss_time:
                email_title = driver.find_elements_by_class_name('da0')[k].text
                logger.info("邮件已更新！可读取最新的邮件的日期：{}, 标题：{}, 接收时间：{}".format(s_time,email_title,email_time))
                # logger.info("导出第{}封邮件，标题: {}".format(k+1,email_title))
                # logger.info("接收邮件的时间：" + email_time)
                driver.find_element_by_xpath('//*[@id="_mail_component_147_147"]/span[2]').click()  # 点击收件箱
                emails_list = driver.find_elements_by_class_name('nui-ico-checkbox') # 选中单个邮件
                one_email = emails_list[k+1]
                print(k+1)
                one_email.click()
                time.sleep(1)
                driver.find_elements_by_class_name('js-component-button')[7].click() # 展开更多
                driver.find_elements_by_class_name('nui-menu-item-text')[0].click() # 导出选中
                time.sleep(2)
                for fname in os.listdir(download_location):
                    if ('信件打包.zip') in fname:
                        # print('邮件导出中 ...')
                        logger.info('邮件导出中 ...')
                        time.sleep(1)
                        try:
                            os.rename(download_location + '信件打包.zip',download_location + username + '-' + now.strftime("%Y-%m-%d-%H点%M分") + "第" + str(i+1) +"页邮件.zip")
                        except:
                            pass
                # logger.info('[+]第' + str(i+1) +'页所有邮件已经下载完成！[+]')

        time.sleep(1)
        driver.find_element_by_class_name('nui-ico-next').click() # 下一页
        time.sleep(1)

        logger.info('[+]' + str(datetime.now()) + '该用户的所有邮件已经全部下载完成，即将下载下一个用户的所有邮件[+]')
    driver.quit()

if __name__ == '__main__':
    
    args = parse_args()
    if args.allExp == True: # 调用全部导出模块
        for j in config.user_profile:
            allExp(j['username'],j['password'],j['proxy'])
            #allExp(j['username'],j['password'])
            logger.info('邮件导出完成，程序运行结束  ' + str(datetime.now()))
    if args.pageExp != None: # 调用页码导出模块
        for j in config.user_profile:
            allExp(j['username'],j['password'],j['proxy'])
            #pageExp(j['username'],j['password'])
            logger.info('邮件导出完成，程序运行结束  ' + str(datetime.now()))
    if args.timeinput != None:
        for j in config.user_profile:
            timeExp(j['username'],j['password'],j['proxy'])
            #timeExp(j['username'],j['password'])
            logger.info('邮件导出完成，程序运行结束  ' + str(datetime.now()))