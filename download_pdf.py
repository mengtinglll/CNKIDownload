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
import argparse
from tqdm import tqdm


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

def main():
    # 配置参数
    parser = argparse.ArgumentParser()
    parser.add_argument('dir', type=str, help='Path to keyword pdf directory.')
    parser.add_argument('-s', '--start_index', type=int, default=1, help='Start index of result.xls.')
    parser.add_argument('-e', '--end_index', type=int, help='end index of result.xls.')
    parser.add_argument('-d', '--directory', action='store_true', help='download has kws dir')
    args = parser.parse_args()
    print(args)
    work_dir = args.dir
    start_index = args.start_index
    end_index = args.end_index

    # 目录为关键词目录，下载目录下xls文件中的链接地址即可
    if not args.directory:
        pdf_dlr = PDFDownloader(work_dir, start_index, end_index=end_index)
        pdf_dlr.download()
    # 目录下包含多个关键词目录，需要分别下载每个关键词目录下的内容
    else:
        kw_dirs = [os.path.join(work_dir, d) for d in os.listdir(work_dir) if not os.path.isfile(d)]
        for kw_dir in kw_dirs:
            pdf_dlr = PDFDownloader(kw_dir, start_index, end_index=end_index)
            pdf_dlr.download()


if __name__ == "__main__":
    main()