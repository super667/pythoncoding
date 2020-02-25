#coding:utf-8

import common.excel as excel
import common.excel2 as excel2
import common.postgresql as postgresql
import shutil
import time
import common.logger as logger
import os
import template.template as  template

class CBase(object):
    def __init__(self, cfg, type):
        self.name = type
        self.cfg = cfg
        self.project = self.cfg.getProject()
        self.region = self.cfg.getRegion()
        self.ip = self.cfg.getHost()
        self.pdbname = self.cfg.get_lastdbname()
        self.ndbname = self.cfg.get_dbname()
        self.tem_path = self.cfg.tem_path  #模板路径
        self.logger = logger.create_main_logger()
        self.res_path = self.cfg.getreport()
        self.ndb = ''
        self.pdb = ''
        self.report = ''
        self.file = ''

        # self.ctemplate = template.Ctemplate(cfg)
        if self.name.lower() == 'statistics':
            self.file = self.cfg.gettemplatefile(self.project, 'ALL', type)  # 模板文件
            # self.file = self.ctemplate.get_statistic_file()
        elif self.name == 'Appendix_Contents':
            # self.file = self.ctemplate.get_contents_file()
            self.file = self.cfg.gettemplatefile(self.project, self.region, type)  # 模板文件
        else:
            self.file = self.cfg.gettemplatefile(self.project, self.region, type)    #模板文件

        if self.file == '' or self.file == None:
            self.logger.error("can't find the template file!!!")
            return

    def copy_file(self):
        self.logger.info("copy template file:"+self.file)

        template = self.file
        (filepath, tempfilename) = os.path.split(template)
        time_str = time.strftime('%Y%m%d%H%M%S', time.localtime(time.time()))
        self.report = self.res_path + self.ndbname + time_str + tempfilename
        print(self.report)
        shutil.copyfile(self.file, self.report)
        
    def _open_excel(self, file_path):
        self.logger.info("open file:"+file_path)
        xls = excel.CExcel_Win32(file_path)  # check.xlsx
        xls.open()
        xls.getSheetCount()
        return xls

    def _open_excel2(self, file_path):
        self.logger.info("open file:"+file_path)
        xls = excel2.CExcel_openpyxl(file_path)  # check.xlsx
        xls.open()
        xls.getSheetCount()
        return xls

    def write_excel(self, sheet_name, db, sqlcmd, offset_row, offset_col):
        (value, data) = db.execute(sqlcmd)
        for index_row in range(1, len(data)+1):
            value = data[index_row-1]
            for index_col in range(1, len(value)+1):
                val = value[index_col-1]
                self.xls.setCellValue(sheet_name, index_row + offset_row, index_col + offset_col, val)
        self.xls.save()

    def connectdb0(self):
        if self.name.lower() == 'statistics' or self.name.lower() == 'testcase':
            self.pdb = postgresql.CPostgresql(self.project, 'host=%s dbname=%s user=%s password=%s' % 
                      (self.ip, self.pdbname, 'postgres', ''))

            self.pdb.connect()
            self.logger.info("connect database:"+self.ip+'--'+self.pdbname)

    def connectdb1(self):
        self.ndb = postgresql.CPostgresql(self.project, 'host=%s dbname=%s user=%s password=%s' % 
                 (self.ip, self.ndbname, 'postgres', ''))
        self.ndb.connect()
        self.logger.info("connect database:"+self.ip+'--'+self.ndbname)

    def _begin(self):
        self.logger.info("begin!")
        self.copy_file()

        # if self.report.split('.')[-1] == 'xls':
        if os.path.splitext(self.report)[-1] == '.xls':
            self.xls = self._open_excel(self.report)
        # elif self.report.split('.')[-1] == 'xlsx':
        elif os.path.splitext(self.report)[-1] == '.xlsx':
            self.xls = self._open_excel2(self.report)
        else:
            self.logger.error('no this template file!!!')

        self.connectdb0()
        self.connectdb1()

    def base_copy_and_open_excel_file(self):
        self.logger.info("begin!")
        self.copy_file()

        if os.path.splitext(self.report)[-1] == '.xls':
            self.xls = self._open_excel(self.report)
        elif os.path.splitext(self.report)[-1] == '.xlsx':
            self.xls = self._open_excel2(self.report)
        else:
            self.logger.error('no this template file!!!')

    def base_close_excel_file(self):
        self.xls.close()

    def _contents_finish(self):
        self.ndb.close()
        self.logger.info("disconnect now database!")
        self.xls.close()
        self.logger.info("close excle file!")
        logger.drop_loggers()
        

    def _finish(self):

        if self.name.lower() == 'statistics' or self.name.lower() == 'testcase':
            self.pdb.close()
            self.logger.info("disconnect database: %s" % self.pdbname)
        self.ndb.close()
        self.logger.info("disconnect database:%s" % self.ndbname)
        self.xls.close()
        self.logger.info("close excle file:%s" % self.report)
        self.logger.info("finish!")
        logger.drop_loggers()

    def write_excel_block(self, sheet_name, offset_row, offset_col, data_row, data_col, data_block):
        # print(data_block)
        if data_block == None:
            return
        for index_row in range(0, len(data_block)):
            value = data_block[data_row + index_row]
            for index_col in range(0, len(value) - data_col):
                # print('index_col:', index_col)
                val = value[index_col + data_col]
                self.xls.setCellValue(sheet_name, index_row + offset_row, index_col + offset_col, val)
        self.xls.save()

    # 大于等于offset_row的行，同时大于等于offset_col的列单元格对应的数据有要删掉
    def clear_exec_bloc(self, sheet_name, offset_row, offset_col):
        self.logger.info("clear data of sheet(begin):" + sheet_name)
        max_row = self.xls.getDimensions(sheet_name)[2]
        max_col = self.xls.getDimensions(sheet_name)[4]
        for row in range(offset_row, offset_row + max_row):
            for col in range(offset_col, offset_col + max_col):
                self.xls.setCellValue(sheet_name, row, col, '')

        self.logger.info("clear data of sheet(end):" + sheet_name)
        self.xls.save()

    def delete_excel_multi_Row_Value(self, sheet_name, row, num=1):
        self.xls.deleteRowValue(sheet_name, row, num)
        self.xls.save()

    def get_range_row_col(self, sheet_name):
        return self.xls.getDimensions(sheet_name)
        
# 基类的功能：
# 就是保存项目，区域，数据库，和模板复制的功能
# 处理数据的基本函数

