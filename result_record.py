# -*- coding=utf-8 -*-
import xlwt
import xlrd
import xlutils.copy
import threading


class ResultRecorder():
    def __init__(self, file='./result.xls', new_file=True):
        self.file = file
        self.index = 0
        self.result_lock = threading.Lock()
        self.new_file = new_file
        if new_file:
            text = [u'序号',u'篇名',u'作者', u'关键词', u'摘要', u'刊名', u'发表时间', \
                u'被引次数', u'下载次数', u'PDF链接', u'上传日期', u'上传人员']
            workbook = xlwt.Workbook(encoding='utf-8', style_compression=0)
            sheet = workbook.add_sheet('result', cell_overwrite_ok=True)
            for col in range(int(len(text))):
                sheet.write(0, col, text[col])
            workbook.save(self.file)

    def add_record(self, contents):
        """
        添加记录到excel表中
        Args:
            contents: 需要插入excel表格中内容的list
            contents包括以下内容：
            index: 下载序号
            title: 文章篇名
            author: 作者
            keyword: 关键词
            abstract: 摘要
            magname: 刊名
            pubdate: 发表日期
            cite: 被引次数
            download: 下载次数
            href: PDF下载地址
            pdate: 上次日期
            pname: 上次人员
        Returns:
            None
        """
        self.result_lock.acquire()
        data = xlrd.open_workbook(self.file)
        sheet = data.sheet_by_index(0)  # 索引的方式，从0开始
        self.index = sheet.nrows
        ws = xlutils.copy.copy(data)
        table = ws.get_sheet(0)
        for i, cont in enumerate(contents):
            table.write(self.index, i, cont)
        ws.save(self.file)
        self.result_lock.release()
    
    def get_last_record(self):
        """
        获取excel中的最后一条记录
        """
        self.result_lock.acquire()
        data = xlrd.open_workbook(self.file)
        sheet = data.sheet_by_index(0)  # 索引的方式，从0开始
        last_rd = sheet.row_values(sheet.nrows-1)
        self.result_lock.release()
        return last_rd
        
    def get_pdf_url(self, start_index, end_index=None):
        """获取result.xls中从起始索引到结束索引的pdf地址
        Args:
            start_index: 开始行索引值，
            end_index: 结束行索引值，如果为None，到最后一行
        Returns:
            pdf_url: 所有pdf的下载地址,list
        """
        pdf_url = []
        self.result_lock.acquire()
        data = xlrd.open_workbook(self.file)
        sheet = data.sheet_by_index(0)  # 索引的方式，从0开始
        if not end_index or end_index >= sheet.nrows:
            end_index = sheet.nrows -1
        for index in range(start_index, end_index+1):
            pdf_url.append(sheet.row_values(index)[9])
        self.result_lock.release()
        print('总共有 {} 个PDF链接地址。'.format(len(pdf_url)))
        return pdf_url
        
    def add_multi_record(self, contents_list):
        """
        添加多条记录到excel表中
        Args:
            contents_list: 需要插入excel表格中内容的list,list每个内容包括contents
            contents包括以下内容：
            index: 下载序号
            title: 文章篇名
            author: 作者
            keyword: 关键词
            abstract: 摘要
            magname: 刊名
            pubdate: 发表日期
            cite: 被引次数
            download: 下载次数
            href: PDF下载地址
            pdate: 上次日期
            pname: 上次人员
        Returns:
            None
        """
        self.result_lock.acquire()
        data = xlrd.open_workbook(self.file)
        sheet = data.sheet_by_index(0)  # 索引的方式，从0开始
        self.index = sheet.nrows
        record_num = len(contents_list)
        ws = xlutils.copy.copy(data)
        table = ws.get_sheet(0)
        for ws_index, cont_index in zip(range(self.index, self.index+record_num),range(record_num)):
            for i, cont in enumerate(contents_list[cont_index]):
                table.write(ws_index, i, cont)
        ws.save(self.file)
        self.result_lock.release()

