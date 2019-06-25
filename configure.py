# -*- coding:utf-8 -*-
"""
程序配置参数。
work_dir: 工作空间目录，在这个目录下，每个关键词单独建立一个目录，存放PDF、日志和result.xls
keywords: 需要搜索下载的关键词列表
chrome_driver_dir: ChromeDriver的目录
page_delay: 下载完一页之后的延时范围，范围内随机选择延时时间。每页默认为50条结果，下载完一页之后会推出登录并重新登录，重新搜索关键词，
login_delay: 推出登录再重新登录的延时
download_delay: 点击下载按钮后等待时间，用于判断是否下载成功
download_retry_times: 下载失败后重试次数
download_failed_thre: 下载失败计数超过阈值后，退出浏览器重新开始
headless: True，为无界面模式，False，为有界面模式
sort_by_relevance: 搜索的结果按照相关度来排序
source: 期刊来源设置，存在’all'则按照默认为全部期刊;其他期刊写成list形式['sci', 'ei', 'core', 'cssci', 'cscd']
quit_period: 下载文件超过个数后，再翻页时退出浏览器重新打开搜索
quit_delay: 关闭浏览器后等待时间
force_start: 强制从头开始下载
"""
config = {
    # 'work_dir': 'D:\\PDFDownload\\运营管理模块\\',
    # 'keywords': [
    #     '预算管理','营运管理'
    # ],
    'work_dir': 'D:\\PDFDownload\\标准成本\\',
    'keywords': [
        '理想标准成本','基本标准成本','正常标准成本','现行标准成本','标准作业卡','标准定额成本法',
        '标准成本控制','零基预算管理','标准成本管理流程','成本绩效管理','业务流程控制','成本作业体系','标准成本差异化'
    ],
    # 'chrome_driver_dir': 'C:\Program Files\Chrome\qudong\chromedriver.exe',
    'chrome_driver_dir': 'C:\Program Files (x86)\Google\Chrome\Application\chromedriver.exe',
    'page_delay': [10, 30],
    'login_delay': [3, 5],
    'search_delay': [3, 5],
    'download_delay': [5, 8],
    'download_retry_times': 3,
    'download_failed_thre': 10,
    'headless': False,
    'sort_by_relevance': True,
    'source': ['sci', 'ei', 'core', 'cssci', 'cscd'],
    'quit_period': 200,
    'quit_daly': [6, 8],
    'force_start': False,
}