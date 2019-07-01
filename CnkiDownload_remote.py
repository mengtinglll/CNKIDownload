# -*- coding:utf-8 -*-
import selenium
import os
import random
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
import time
from logger import Logger
import logging
from result_record import ResultRecorder
import datetime
import re
from configure import config

class CnkiDownloader():
    """知网pdf下载类"""
    def __init__(self):
        """CnkiDownloader类初始化"""
        self.addr = 'http://kns.cnki.net/kns/brief/result.aspx?dbprefix=CJFQ'
        self.work_addr = ''
        self.driver = None
        self.logger = None
        self.res_rdr = None 
        self.window_1 = None #搜索页面选项handle
        self.window_rec = None #论文详情页面
        self.cnt = 0 #下载计数
        self.failed_cnt = 0 #下载失败计数
        self.res_num = 0
        self.page_num = 0
        self.cur_page = 0
        self.download_pdf = True #没有VPN时为False

    def config_chromedriver(self):
        """配置ChromeDriver参数"""
        chromedriver = config['chrome_driver_dir']
        chromeOptions = webdriver.ChromeOptions()
        # 设置默认下载地址
        prefs = {"download.default_directory": self.work_addr}
        chromeOptions.add_experimental_option("prefs", prefs)
        # 设置无界面模式,
        if config['headless']:
            chromeOptions.add_argument('--headless')
            chromeOptions.add_argument('--disable-gpu')
        driver = webdriver.Chrome(chromedriver, chrome_options=chromeOptions)
        return driver

    def go_to_page(self, page):
        """跳转到指定页面
        Args:
            page: 需要跳转的页码
        """
        if page > 1:
            nextBtn = self.driver.find_element_by_link_text("下一页")
            nhref = self.driver.find_element_by_link_text("下一页").get_attribute('href')
            pattern = re.compile(r'\?curpage=\d{1,3}')
            nhref = pattern.sub('?curpage='+str(page), nhref)
            arg = "arguments[0].href =\'"+ nhref+"\';"
            # self.logger.debug("arg: {}".format(arg))
            self.driver.execute_script(arg, nextBtn)
            self.driver.find_element_by_link_text("下一页").click()
            time.sleep(random.uniform(1, 3))
            # 跳转到搜索结果frame
            self.driver.switch_to.default_content()
            self.driver.switch_to.frame("iframeResult")

    def serach_keyword(self, keyword, login=False, quit=False):
        """
        知网页面中搜索关键词,过滤掉英文结果，并按配置参数排序
        Args:
            keyword: 关键词
            page: 跳转页数
            login: 是否重新登录
            quit: 是否关闭所有选项卡
        Returns:
            None
        """
        # 退出驱动，关闭所有选项卡, 重新打开浏览器界面
        if quit: 
            # TODO(inumo): 增加对是否还有文件正在下载的判断,通过工作目录下的文件变化
            if self.driver:
                self.driver.quit()
                self.window_1 = None
                self.window_rec = None
            time.sleep(random.uniform(config['quit_daly'][0], config['quit_daly'][1]))
            self.driver = self.config_chromedriver()
            self.driver.get(self.addr)
            time.sleep(5)
            self.window_1 = self.driver.current_window_handle

            # 搜索关键词前选择来源期刊
            if 'all' not in config['source']:
                for s in config['source']:
                    if s == 'sci': #SCI来源期刊
                        self.driver.find_element_by_xpath('//input[@id="mediaBox1"]').click()
                    elif s == 'ei': #EI来源期刊
                        self.driver.find_element_by_xpath('//input[@id="mediaBox2"]').click()
                    elif s == 'core': #核心期刊
                        self.driver.find_element_by_xpath('//input[@id="mediaBox3"]').click()
                    elif s == 'cssci': #中文社会科学引文索引
                        self.driver.find_element_by_xpath('//input[@id="mediaBox4"]').click()
                    elif s == 'cscd': #CSCD
                        self.driver.find_element_by_xpath('//input[@id="mediaBox5"]').click()
                    time.sleep(1)
        
        # 检查是否根据IP登录了
        need_login = False
        try:
            if not self.driver.find_element_by_xpath('//*[@id="Ecp_loginShowName"]').text:
                need_login = True
                self.logger.warning('无法根据IP登录，请检查VPN设置')
        except:
            need_login = True
            
        # 重新登录
        if login or need_login:
            self.driver.switch_to.default_content()#返回到主文本
            # 先退出登录
            if login:
                # 先按下拉，再选择退出登录
                self.driver.find_element_by_xpath('//*[@id="Ecp_top_logout_showLayer"]/i/span').click()
                time.sleep(random.uniform(3, 5))
                self.driver.find_element_by_xpath('//*[@id="Ecp_top_logoutClick"]').click()
                time.sleep(random.uniform(config['login_delay'][0], config['login_delay'][1]))
            self.driver.find_element_by_xpath('//*[@id="Ecp_top_login"]/a/i/span').click()
            time.sleep(random.uniform(1, 3))
            self.driver.find_element_by_id("Ecp_Button2").click()#重新登陆
            time.sleep(random.uniform(3,5))
            try:
                self.driver.find_element_by_xpath('//*[@id="Ecp_errorMsg"]')
                # TODO(inumo): 
                self.download_pdf = False
                self.logger.warning('IP登录失败，检查VPN设置。')
            except:
                pass
            

        # 搜索关键词
        self.driver.find_element_by_id("txt_1_value1").clear()
        self.driver.find_element_by_id("txt_1_value1").send_keys(str(keyword))
        time.sleep(1)
        self.driver.find_element_by_id("txt_1_value1").send_keys(Keys.ENTER)
        time.sleep(random.uniform(config['search_delay'][0], config['search_delay'][1]))

        # 选择一页显示的结果数，中文显示，排序规则
        self.driver.switch_to.frame("iframeResult")
        self.driver.find_element_by_link_text("50").click()#点击显示50页
        time.sleep(2)

        # 只显示中文搜索结果
        try:
            self.driver.find_element_by_xpath('//a[@class="Ch on"]')
        except:
            self.driver.find_element_by_xpath('//a[@class="Ch"]').click()
        time.sleep(2)

        # 按照相关度排序
        if config['sort_by_relevance']:
            self.driver.find_element_by_xpath(
                '//span[@class="groupsorttitle"]/../span[2]').click()
            time.sleep(2)
    
    def process_record(self, index):
        """下载一个搜索结果的pdf并保存信息
        Args:
            index: 结果条目的index值
        Returns:
            None
        Raises:

        """
        try:
            seq_num = self.driver.find_element_by_xpath(
                '//*[@id="ctl00"]/table/tbody/tr[2]/td/table/tbody/tr['+str(index)+']/td[1]').text
        except:
            self.logger.error("获取序号失败")
            seq_num=''
        try:
            title = self.driver.find_element_by_xpath(
                '//*[@id="ctl00"]/table/tbody/tr[2]/td/table/tbody/tr['+str(index)+']/td[2]/div/a').text
        except:
            self.logger.error("获取标题失败")
            title=''
        try:
            author = self.driver.find_element_by_xpath(
                '//*[@id="ctl00"]/table/tbody/tr[2]/td/table/tbody/tr['+str(index)+']/td[3]').text.strip(';')
        except:
            self.logger.error("获取作者失败")
            author=''
        mag_name = self.driver.find_element_by_xpath(
            '//*[@id="ctl00"]/table/tbody/tr[2]/td/table/tbody/tr['+str(index)+']/td[4]/a').text
        pub_date = self.driver.find_element_by_xpath(
            '//*[@id="ctl00"]/table/tbody/tr[2]/td/table/tbody/tr['+str(index)+']/td[5]').text
        
        # 获取被引次数和下载次数
        try:
            cite_count = self.driver.find_element_by_xpath(
                '//*[@id="ctl00"]/table/tbody/tr[2]/td/table/tbody/tr['+str(index)+']/td[6]/span/a').text
            if not cite_count:
                cite_count = 0
            else:
                cite_count = int(cite_count)
        except:
            cite_count = 0
            self.logger.error('获取被引次数失败。')
        try:
            dl_count = self.driver.find_element_by_xpath(
                '//*[@id="ctl00"]/table/tbody/tr[2]/td/table/tbody/tr['+str(index)+']/td[7]/span/a').text
            if not dl_count:
                dl_count = 0
            else:
                dl_count = int(dl_count)
        except:
            dl_count = 0
            self.logger.error('获取下载次数失败。')

        # 跳转至论文详情页面
        self.driver.find_element_by_xpath(
            '//*[@id="ctl00"]/table/tbody/tr[2]/td/table/tbody/tr['+str(index)+']/td[2]/div/a').click()
        
        self.logger.debug("下载进度: {}/{}文件， {}/{}页面。".format(seq_num, self.res_num, self.cur_page, self.page_num))
        time.sleep(random.uniform(config['download_delay'][0], config['download_delay'][1]))

        # 浏览器选项卡切换至最近打开的一个选项卡
        windows = self.driver.window_handles
        for current_window in windows:
            if current_window == windows[-1]:
                self.window_rec = current_window
                self.driver.switch_to.window(current_window)
                time.sleep(1)
                try:
                    abstract = self.driver.find_element_by_xpath('//*[@id="ChDivSummary"]').text.strip()
                    try:
                        keyword = self.driver.find_element_by_xpath('//label[@id="catalog_KEYWORD"]/..').text[4:]
                    except:
                        keyword=''
                    href = self.driver.find_element_by_xpath('//*[@id="pdfDown"]').get_attribute('href')
                    self.logger.info("序号:{}/{}, 篇名:{}, 作者:{}, 关键词: {}, 摘要: {}, 刊名:{}, 发表时间: {}, 被引次数: {}, 下载次数: {}, PDF地址: {}".format(
                        seq_num, self.res_num, title, author, keyword, abstract, mag_name, pub_date, cite_count, dl_count, href))
                    dt = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    try:
                        self.res_rdr.add_record(
                            [str(seq_num), title, author, keyword, abstract, mag_name, pub_date, cite_count, dl_count, href, dt, 'lmt'])
                    except Exception as e:
                        self.logger.error('添加记录失败，原因: {}'.format(e))

                    # 点击下载按钮
                    self.driver.find_element_by_xpath('//*[@id="pdfDown"]').click()
                    time.sleep(random.uniform(config['download_delay'][0], config['download_delay'][1]))
                    
                    # 下载是否成功判断
                    retry_count = 0
                    while len(self.driver.window_handles) > 2:
                        for win in self.driver.window_handles:
                            # 处理弹窗Alert
                            if win == self.window_rec:
                                # 切换到论文详情页面,可以关闭Alert窗口
                                # driver.switch_to.alert对于知网的Alert弹窗无效,对于知网点击下载PDF后在
                                # 新窗口弹出的警告窗口，通过切换到论文详情页可以关闭该弹窗窗口
                                self.driver.switch_to_window(self.window_rec)
                                time.sleep(1)
                                if len(self.driver.window_handles) == 2:
                                    self.logger.debug('跳过下载论文: {}, 原因: 弹窗错误。'.format(title))
                                    # 弹窗错误无法通过重试解决，直接跳过
                                    break

                            # 关掉不能正确下载的页面窗口，并跳转回论文详情窗口
                            elif win != self.window_1 and win != self.window_rec:
                                self.driver.switch_to_window(win)        
                                try:
                                    if self.driver.current_window_handle != self.window_1 and self.driver.current_window_handle != self.window_rec:
                                        self.driver.close()
                                except:
                                    self.logger.debug("下载窗口已经自动关闭。")
                                
                                # 切换到论文详情页面
                                self.driver.switch_to_window(self.window_rec)

                                if retry_count < config['download_retry_times']:
                                    self.driver.find_element_by_xpath('//*[@id="pdfDown"]').click()
                                    time.sleep(random.uniform(config['download_delay'][0], config['download_delay'][1]))
                                    break
                                else:
                                    self.logger.warning('下载重试次数超过{}次，放弃。'.format(config['download_retry_times']))
                                    break
                        
                                # 重试次数达到上限，退出
                                if retry_count == config['download_retry_times']:
                                    break
                                retry_count += 1
                                # 下载失败计数加1
                                self.failed_cnt += 1
                except Exception as e:
                    self.logger.error('下载【{}】失败, 原因: {}'.format(title, e))
                
                # 不管是否下载成功，下载计数加1
                self.cnt += 1

                # 关闭论文详情窗口
                if self.driver.current_window_handle == self.window_rec:
                    self.driver.close()
                
                # 切换到搜索结果窗口,并切换到搜索结果frame
                windows = self.driver.window_handles
                for current_window in windows:
                    if current_window == self.window_1:
                        self.driver.switch_to.window(self.window_1)
                        self.driver.switch_to.default_content()
                        self.driver.switch_to.frame("iframeResult")
                        break
                time.sleep(random.uniform(2, 3))

    def start(self):
        """开始PDF下载任务"""
        for kw in config['keywords']:
            try:
                # 计数初始化
                self.cnt = 0
                self.failed_cnt =0
                self.driver = None

                # 工作空间目录设置
                work_addr = os.path.join(config['work_dir'], kw)
                self.work_addr = work_addr
                if not os.path.exists(work_addr):
                    os.makedirs(work_addr)

                # logger配置,日志文件位于work_addr目录下
                Logger.config(
                    log_file=os.path.join(work_addr,kw+'crawler_cnki.log'),
                    use_stdout=True,
                    log_level=logging.DEBUG)
                self.logger = Logger.get_logger(kw+'crawler_cnki')
                self.logger.info('开始下载关键词: {}'.format(kw))
                self.cnt = 0

                # 结果记录类, 记录存放在work_addr下，文件名为result.xls
                res_dir = os.path.join(work_addr,'result.xls')
                new_file = False
                if not os.path.exists(res_dir) or config['force_start']:
                    new_file = True
                self.res_rdr = ResultRecorder(res_dir, new_file=new_file)

                # 起始页面和起始序号设置
                start_page = 1
                start_index = 2
                if not new_file:
                    try:
                        start_index = int(self.res_rdr.get_last_record()[0])
                        start_page = start_index//50 +1
                        start_index %= 50
                        start_index += 2
                    except:
                        pass
                self.logger.debug("起始页面: {}, 起始索引: {}".format(start_page, start_index))
                
                # 如果起始页已经为120且起始index为52则跳过这个关键词
                if start_page >= 120 and start_index >= 52:
                    continue

                # 搜索关键词,跳转到起始页面
                self.serach_keyword(kw, quit=True)
                self.go_to_page(start_page)

                # 获取总文件个数和页面数
                try:
                    res_num = self.driver.find_element_by_xpath(
                        '//*[@id="ctl00"]/table/tbody/tr[3]/td/table/tbody/tr/td/div/div').text
                except:
                    res_num = self.driver.find_element_by_xpath(
                        '//*[ @ id = "J_ORDER"]/tbody/tr[2]/td/table/tbody/tr/td[2]/div/div').text
                res_num = res_num.replace("找到","").replace("条结果","").replace(',','')
                res_num = int(res_num)
                self.res_num = res_num

                # 结果页面数
                page_num = 1
                try:
                    page_num = self.driver.find_element_by_xpath('//*[@class="countPageMark"]').text
                    # 格式: 当前页/总页数，如1/120
                    page_num = int(page_num.split('/')[-1])
                except Exception as e:
                    # 获取失败，通过计算获取总页数
                    self.logger.error("获取总页数失败。")
                self.page_num = page_num

                # 计算最后一页的索引值,注： 第一条记录的索引值为2
                if res_num >= 6000:
                    last_pg_index = 52
                else:
                    last_pg_index = res_num % 50 + 2

                # 如果搜索结果超过6000条，只能显示120页结果
                self.logger.debug("关键词：【{}】共有 {} 条结果， 共 {} 页。".format(kw, res_num, page_num))

                # 遍历所有页面，下载搜索结果条目
                for pg in range(start_page, page_num + 1):
                    self.cur_page = pg
                    # 处理本页面信息
                    if pg == page_num:
                        end_index = last_pg_index
                    else:
                        end_index = 52
                    
                    # 从起始index开始遍历本页的记录
                    for index in range(start_index, end_index):
                        start_index = 2
                        # 下载index对应的论文
                        self.process_record(index)
                    
                    # 完成最后一页
                    if pg == page_num:
                        time.sleep(3)
                        self.logger.info('完成关键词: {} 的下载，退出浏览器。'.format(kw))
                        self.driver.quit()
                        break
                        
                    # 完成一页的下载后，延时一段时间
                    time.sleep(random.uniform(config['page_delay'][0], config['page_delay'][1]))

                    # 完成一页后重新登录或退出浏览器重新开始搜索
                    if self.failed_cnt >= config['download_failed_thre'] or self.cnt >= config['quit_period']:
                        self.failed_cnt = 0
                        self.cnt = 0
                        self.serach_keyword(kw, quit=True)
                    else:
                        self.serach_keyword(kw, login=True)
                    self.go_to_page(pg+1)

            except Exception as e:
                self.logger.error('开始执行关键词【{}】的过程时出错，原因: {}'.format(kw, e))
                time.sleep(6)
                self.driver.quit()
                time.sleep(5)


def main():
    cnki_dlr = CnkiDownloader()
    cnki_dlr.start()

if __name__ == "__main__":
    main()