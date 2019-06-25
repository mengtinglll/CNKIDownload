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
        
