# 爬取知网文献内容
根据关键词搜索并下载相关文献PDF。目前只针对中文文献。

## 主要内容
[1 运行前提](#1运行前提)  
[2 运行配置](#2运行配置)  
[3 问题及解决办法](#3问题及解决办法)  

## 1运行前提  
* 代码基于python3.6，运行所需要的库包括以下部分：  
```bash
pip install selenium
pip install xlwt
pip install xlrd
pip install xlutils
```
* selenium模块需要调用Chrome的webdriver。安装方法如下：  
[下载地址](https://chromedriver.storage.googleapis.com/index.html)  
注意版本需与电脑安装的浏览器版本对应，附：版本号对应描述（64位浏览器下载32位即可），下载后与chrome安装目录放在一起，然后配置至环境变量即可。

## 2运行
* 知网论文下载需要连接VPN。  
**确保程序运行时已经连接上校内网VPN。**  
* 在configure.py文件中对运行的参数进行设置，程序运行配置主要有：  

  * **工作空间配置**: 用于存放日志，结果excel，以及下载的PDF文件
  * **搜索关键词配置**: 需要搜索下载PDF的关键词
  * **ChromeDriver目录配置**: ChromeDriver的存放位置
  * 延时配置: 
    * page_delay: 下载完一页50个结果之后的延时时间
    * login_delay: 退出登录再重新登录之间的延时
    * download_delay: 点击下载PDF按钮后的延时时间
    * quit_delay: 完全关闭浏览器再打开之间的延时
  * 异常重试相关配置：
    * download_retry_times: 下载PDF失败时的重试次数
    * download_failed_thre: 失败次数阈值，超过后在翻页时会完全退出浏览器重新检索
    * quit_period: 下载数量阈值，超过后再翻页时会完全退出浏览器重新检索
  * 是否无界面配置: 无界面模式不影响电脑其他工作，但是报错时只能看日志，看不到界面 
  * 按相关度排序: 搜索结果按照相关度排序
  * 期刊来源设置：  
    * sci: SCI来源期刊
    * ei: EI来源期刊
    * core: 核心期刊
    * CSSCI: 中文社会科学引文索引
    * CSCD:
  * 是否强制从头开始下载: 下载结果会保存再result.xls中，中断重启时会从该文件中读出最后的结果，接着下载。设置强制从头则从开头下载。

* 运行程序  
```
python CnkiDownload_remote.py
```
## 3问题及解决办法