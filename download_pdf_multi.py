# -*- coding:utf-8 -*-
import selenium
import os
import sys
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
import argparse
from tqdm import tqdm
import threading
import queue

class MultiDownloader(threading.Thread):
    def __init__(self, thread_id, name, res_dir, pdf_queue, pdf_num):
        """MultiDownloader类初始化
        Args:
            thread_id: 线程id
            name: 线程名称
            res_dir: 工作目录
            pdf_queue: pdf连接地址队列
            pdf_num: pdf总数
        """
        super(MultiDownloader, self).__init__()
        self.thread_id = thread_id
        self.name = name
        self.res_dir = res_dir
        self.pdf_queue = pdf_queue
        self.pdf_num = pdf_num
        self.driver = None
        self.cnki_title = '学术期刊—中国知网'
        self.cnki_url = 'http://kns.cnki.net/kns/brief/result.aspx?dbprefix=CJFQ'
        self.main_window = None

    def config_chromedriver(self):
        """配置ChromeDriver参数"""
        chromedriver = config['chrome_driver_dir']
        chromeOptions = webdriver.ChromeOptions()
        # 设置默认下载地址
        prefs = {"download.default_directory": self.res_dir}
        chromeOptions.add_experimental_option("prefs", prefs)
        # 设置无界面模式,
        if config['headless']:
            chromeOptions.add_argument('--headless')
            chromeOptions.add_argument('--disable-gpu')
        driver = webdriver.Chrome(chromedriver, chrome_options=chromeOptions)
        return driver
        
    def login(self, re_login=False):
        """重新登录
        Args:
            re_login: 重新登录
        """
        # 重新登录,title: 学术期刊—中国知网
        self.driver.get(self.cnki_url)
        self.driver.switch_to.default_content()#返回到主文本
        time.sleep(3)

        # 先退出登录
        if re_login:
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
            print('IP登录失败，检查VPN设置。')
        except:
            pass

    def check_download(self, url):
        """检查是否下载成功
        下载成功之后,下载的窗口会自动关闭，跳回至原来窗口。
        Args:
            None
        Returns:
            download_state: 0，下载成功，1，遇到无法下载的Alert，
                            -1，遇到重新登录Alert， -2，网络延时导致下载失败
        """
        # 先检查是否有Alert弹窗,如果有,处理后直接返回
        state = 0
        try:
            alert = self.driver.switch_to.alert
            print('Alert: ', alert.text)
            # TODO(inumo):根据不同给Alert内容做不同处理
            if alert.text.strip() == '对不起，您的操作太过频繁！请退出后重新登录。':
                alert.accept()
                # 重新登录
                state = -1
            else:
                # 其他Alert消息有：
                # 1.产品不在有效期范围之内！
                # 2.对不起，本刊只收录题录信息，暂不提供原文下载。给您造成不便请谅解！
                # 这两种情况都只能跳过下载
                alert.accept()
                state = 1
            return state
        except selenium.common.exceptions.NoAlertPresentException as e:
            pass

        # 处理由于网络原因导致的错误
        if self.driver.title.strip() != self.cnki_title:
            state = -2

        return state

    def run(self):
        """"""
        print('开启线程: ', self.name)
        # 配置ChromeDriver
        self.driver = self.config_chromedriver()

        # 打开知网检索网页
        self.driver.get(self.cnki_url)
        self.main_window = self.driver.current_window_handle
        dl_count = 0
        while not self.pdf_queue.empty():
            url = self.pdf_queue.get()
            print('线程: {},下载进度: {}/{}.'.format(
                self.name, self.pdf_num-self.pdf_queue.qsize(), self.pdf_num
            ))
            retry_count = 0
            while True:
                if retry_count >= 3:
                    print('\n重试次数超过3次，放弃。')
                    # 需要将页面切换回知网检索主页
                    self.driver.get(self.cnki_url)
                    time.sleep(1)
                    break
                self.driver.get(url)
                time.sleep(random.uniform(5, 8))
                # 检查是否下载成功
                state = self.check_download(url)
                if state == 0:
                    break
                elif state == 1:
                    print('\n遇到无法下载Alert，跳过。')
                # 重新登录之后再下载
                elif state == -1:
                    self.login(re_login=True)

                retry_count += 1
            dl_count += 1

            # 下载达到一定次数后，重新登录
            if dl_count % 50 == 0:
                self.login(re_login=True)
        self.driver.quit()
        print('结束线程: ', self.name)  


class PDFDownloader(object):
    """PDF下载类"""
    def __init__(self, res_dir, start_index, end_index=None):
        self._res_dir = res_dir
        self._res_fn = os.path.join(self._res_dir, 'result.xls')
        self.driver = None
        self.start_index = start_index
        self.end_index = end_index
        self.cnki_title = '学术期刊—中国知网'
        self.cnki_url = 'http://kns.cnki.net/kns/brief/result.aspx?dbprefix=CJFQ'
        self.main_window = None

    def config_chromedriver(self):
        """配置ChromeDriver参数"""
        chromedriver = config['chrome_driver_dir']
        chromeOptions = webdriver.ChromeOptions()
        # 设置默认下载地址
        prefs = {"download.default_directory": self._res_dir}
        chromeOptions.add_experimental_option("prefs", prefs)
        # 设置无界面模式,
        if config['headless']:
            chromeOptions.add_argument('--headless')
            chromeOptions.add_argument('--disable-gpu')
        driver = webdriver.Chrome(chromedriver, chrome_options=chromeOptions)
        return driver

    def download(self):
        """下载PDF"""
        self.driver = self.config_chromedriver()

        # 打开知网检索网页
        self.driver.get(self.cnki_url)
        self.main_window = self.driver.current_window_handle

        # 读取excel中的pdf下载地址
        res_rdr = ResultRecorder(self._res_fn, new_file=False) 
        pdf_urls = res_rdr.get_pdf_url(self.start_index, end_index=self.end_index)
        dl_count = 0
        pdf_num = len(pdf_urls)
        purl = tqdm(pdf_urls)
        for url in purl:
            # 在新的窗口中打开pdf下载链接
            # js = 'window.open(\"{}\");'.format(url)
            # self.driver.execute_script(js)
            # print('\r','下载中：{}/{}'.format(index+1, pdf_num),end='')
            retry_count = 0
            while True:
                if retry_count >= 3:
                    print('\n重试次数超过3次，放弃。')
                    break
                self.driver.get(url)
                time.sleep(random.uniform(5, 8))
                # 检查是否下载成功
                state = self.check_download(url)
                if state == 0:
                    break
                elif state == 1:
                    print('\n遇到无法下载Alert，跳过。')
                # 重新登录之后再下载
                elif state == -1:
                    self.login(re_login=True)

                retry_count += 1
            dl_count += 1

            # 下载达到一定次数后，重新登录
            if dl_count % 50 == 0:
                self.login(re_login=True)
            purl.set_description('Processing')
        # 完成所有下载任务
        print('\n完成所有下载任务，退出浏览器。')
        self.driver.quit()

    def login(self, re_login=False):
        """重新登录
        Args:
            re_login: 重新登录
        """
        # 重新登录,title: 学术期刊—中国知网
        self.driver.get(self.cnki_url)
        self.driver.switch_to.default_content()#返回到主文本
        time.sleep(3)

        # 先退出登录
        if re_login:
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
            print('IP登录失败，检查VPN设置。')
        except:
            pass

    def check_download(self, url):
        """检查是否下载成功
        下载成功之后,下载的窗口会自动关闭，跳回至原来窗口。
        Args:
            None
        Returns:
            download_state: 0，下载成功，1，遇到无法下载的Alert，
                            -1，遇到重新登录Alert， -2，网络延时导致下载失败
        """
        # 先检查是否有Alert弹窗,如果有,处理后直接返回
        state = 0
        try:
            alert = self.driver.switch_to.alert
            print('Alert: ', alert.text)
            # TODO(inumo):根据不同给Alert内容做不同处理
            if alert.text.strip() == '对不起，您的操作太过频繁！请退出后重新登录。':
                alert.accept()
                # 重新登录
                state = -1
            else:
                # 其他Alert消息有：
                # 1.产品不在有效期范围之内！
                # 2.对不起，本刊只收录题录信息，暂不提供原文下载。给您造成不便请谅解！
                # 这两种情况都只能跳过下载
                alert.accept()
                state = 1
            return state
        except selenium.common.exceptions.NoAlertPresentException as e:
            pass

        # 处理由于网络原因导致的错误
        if self.driver.title.strip() != self.cnki_title:
            state = -2

        return state

    # def check_download(self, url):
    #     """检查是否下载成功
    #     下载成功之后,下载的窗口会自动关闭，未关闭则说明下载不正常。
    #     """
    #     retry_count = 0
    #     handle_alert = False
    #     while len(self.driver.window_handles) > 1:
    #         for win in self.driver.window_handles:
    #             # 处理Alert弹窗
    #             if win == self.main_window:
    #                 self.driver.switch_to_window(self.main_window)
    #                 time.sleep(1)
    #                 if len(self.driver.window_handles) == 1:
    #                     print('跳过下载论文, 原因: 弹窗错误。')
    #                     handle_alert = True
    #                     # 弹窗错误无法通过重试解决，直接跳过
    #                     break
    #             else:
    #                 self.driver.switch_to_window(win)        
    #                 try:
    #                     if self.driver.current_window_handle != self.main_window:
    #                         self.driver.close()
    #                 except:
    #                     print("下载窗口已经自动关闭。")

    #                 # 切换到主页
    #                 self.driver.switch_to_window(self.main_window)

    #                 if retry_count < 3:
    #                     js = 'window.open(\"{}\");'.format(url)
    #                     self.driver.execute_script(js)
    #                     time.sleep(random.uniform(2, 3))
    #                 else:
    #                     print('下载重试次数超过3次，放弃。')

    #         # 重试次数达到上限或遇到了Alert退出
    #         if retry_count == 3 or handle_alert:
    #             break
    #         # 重试次数加1
    #         retry_count += 1
def MultiProcess(work_dir, start_index, end_index=None):
    """开启多线程任务
    Args:
        work_dir: 工作目录，目录下包含result.xls
        start_index: 起始pdf索引值
        end_index: 终止pdf索引值
    """
    # 读取excel中的pdf下载地址
    res_fn = os.path.join(work_dir, 'result.xls')
    res_rdr = ResultRecorder(res_fn, new_file=False) 
    pdf_urls = res_rdr.get_pdf_url(start_index, end_index=end_index)
    print('{}总共有{}个PDF文件'.format(work_dir, len(pdf_urls)))

    # 入队列
    pdf_queue = queue.Queue()
    for url in pdf_urls:
        pdf_queue.put(url)

    # 开启多线程
    print('开始下载PDF....')
    thread_list = ['Thread-1', 'Thread-2', 'Thread-3', 'Thread-4']
    threads = []
    for idx, name in enumerate(thread_list):
        thread = MultiDownloader(idx, name, work_dir, pdf_queue, len(pdf_urls))
        thread.start()
        threads.append(thread)

    for thread in threads:
        thread.join()

    print('结束全部PDF下载任务')

def main():
    # 配置参数
    parser = argparse.ArgumentParser()
    parser.add_argument('-kd', '--key_dir', type=str, help='关键词目录')
    parser.add_argument('-ud', '--up_dir', type=str, help='包含关键词目录的目录')
    parser.add_argument('-md', '--multi_dir', action='append', default=[], help='多个关键词目录')
    parser.add_argument('-s', '--start_index', type=int, default=1, help='result.xls中的起始index值')
    parser.add_argument('-e', '--end_index', type=int, help='result.xls的终止index值')
    parser.add_argument('-x', '--extra_dir', action='append', default=[], help='过滤掉的目录')
    
    args = parser.parse_args()
    print(args)
    arg_count = 0
    if args.key_dir:
        arg_count += 1
    if args.up_dir:
        arg_count += 1
    if args.multi_dir:
        arg_count += 1
    if arg_count != 1:
        print('参数不正确，退出。\n使用说明：参数必须传入-kd, -ud, -md中的一种,且只能传入一种。\n',
            '\t-kd: 传入关键词目录\n',
            '\t-ud: 传入包含多个关键词目录的目录,可配合-x使用,排除某几个目录\n',
            '\t-md: 传入多个关键词目录\n',
            '其他参数:\n\t-s: 起始index,-e:终止index')
        sys.exit(0)
    start_index = args.start_index
    end_index = args.end_index

    # 必须参数中包含了多个目录参数，则遍历每个参数
    if args.multi_dir:
        for w_dir in args.multi_dir:
             MultiProcess(w_dir, start_index, end_index=end_index)

    # 关键词目录
    if args.key_dir:
        MultiProcess(args.key_dir, start_index, end_index=end_index)
    
    # 包含关键词目录的目录
    if args.up_dir:
        work_dir = args.up_dir
        kw_dirs = [os.path.join(work_dir, d) for d in os.listdir(work_dir) if not os.path.isfile(d)]
        for kw_dir in kw_dirs:
            # 过滤掉extra列出的不包含目录
            if kw_dir not in args.extra_dir:
                MultiProcess(kw_dir, start_index, end_index=end_index)

if __name__ == "__main__":
    main()