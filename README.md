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
### 2.1 直接下载PDF
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
### 2.2 先获取PDF链接地址再统一下载
由于知网每次按照相同的条件检索时（关键词，期刊范围，排序方式）返回的结果不完全一致，导致文件可能存在重复下载的情况。为了解决该问题，采用下一次性获取PDF链接地址，然后再通过读取excel中的PDF链接地址来下载PDF。（两步）
* STEP1: 获取PDF链接地址
  **这个步骤因为不下载PDF，因此不需要VPN，在所有网络条件下都可用。** 这个方法的延时也可以相应设置小一些，配置文件为 ```configure_fast.py```，配置完成后运行：
  ```bash
  python CnkiDownload_fast.py
  ```
  结果会存在关键词目录下的```result.xls```文件夹下。
* STEP2: 下载PDF  
  **这个步骤需要连接VPN。**，可以一次性下载一个文件夹下的很多个关键词文件夹，即传入的文件夹下包含很多个关键词文件夹。
  ```bash
  python download_pdf.py path_to_path_of_kws -d
  ```
  如果中途下载失败，需要从指定位置开始下载，则应该指定起始的索引值.
  ```bash
  python download_pdf.py path_to_kw -s start_index
  ```
  索引值根据已经下载的PDF和excel中来确定。
* STEP2.1: 多线程下载PDF  
  **这个步骤需要连接VPN。**  
  运行脚本```python download_pdf_multi.py args```args为使用的参数，参数说明如下：  
使用说明：参数必须传入-kd, -ud, -md中的一种,且只能传入一种。  
        -kd: 传入关键词目录  
        -ud: 传入包含多个关键词目录的目录,可配合-x使用,排除某几个目录  
        -md: 传入多个关键词目录  
 其他参数(可选参数)：  
        -s: 起始index,-e:终止index  
        
例子1. 'D:\PDFDownload\预算\差异分析'为关键词目录
```bash
python download_pdf_multi.py -kd D:\PDFDownload\预算\差异分析
```
例子2. 'D:\PDFDownload\预算'为包含关键词目录的目录
```bash
python download_pdf_multi.py -ud D:\PDFDownload\预算
```
例子3. 'D:\PDFDownload\预算\差异分析'，'D:\PDFDownload\预算\xxx'为多个关键词目录
```bash
python download_pdf_multi.py -md D:\PDFDownload\预算\差异分析 -md D:\PDFDownload\预算\xxx -md D:\PDFDownload\预算\xx1
```
## 3问题及解决办法
1. 不同关键词的日志在同一个文件中  
   2019.6.28:解决该bug
2. 出现：Message: unknown error: failed to close window in 20 seconds.  
  2019.6.28：原因是浏览器有警告弹窗,解决：通过切换到论文详情页来关闭弹窗错误。由于切换到论文详情页后弹窗自动关闭，因此无法使用driver.switch_to.alert的方法来解决这个问题。

