import Entity.configInfo
import logger
import os
import xlwt
import time
class Creport():
    __instance = None
    @staticmethod
    def instance():
        'create a instance'
        if Creport.__instance is None:
            Creport.__instance = Creport()
        return Creport.__instance
    
    def __init__(self):
        try:
            # get log instance
            self.pg_log = logger.CLogger.instance().getLogHandle('Creport')
            
            # get a configure instance
            self.pg_config = Entity.configInfo.CConfigInfo.instance()
            
        except:
            if self.pg_log:
                self.pg_log.exception('exception happened ......')
            raise
        

    def GetFileName(self):
        
        self.pg_log.info('get output path and filename......')
        paths=str(self.pg_config.getpath())
        
        name1=self.pg_config.database1()[1]
        name2=self.pg_config.database2()[1]
                
        self.filename = paths+'\\' + name1+'_VS_'+name2 + '_' + time.strftime("%Y%m%d%H%M", time.localtime())
        return self.filename 

    def outputstyle(self,color):   
        style = xlwt.XFStyle()
        font = xlwt.Font()
        font.bold=True
        font.height=240        
        style.font = font
        pattern = xlwt.Pattern()
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour=color
        style.pattern=pattern
        return style
    
    def ABpercentcolor(self,v):
        if abs(float(v))<0.1:
            stylep=self.outputstyle(3)
        elif abs(float(v))<0.3:
            stylep=self.outputstyle(7)
        elif abs(float(v))<0.5:
            stylep=self.outputstyle(6)
        else:
            stylep=self.outputstyle(2)
        return stylep
         
    def sheetwritetb(self,data,sheet,flag):                
        self.pg_log.info('write the table information to excel......')
        style=self.outputstyle(1)
        sheet.write_merge(0,0,4,5,"value is 0",self.outputstyle(5))
        sheet.write_merge(1,1,4,5,"the difference of field",self.outputstyle(24))
        sheet.write_merge(2,2,4,5,"The percentage is less than 0.1",self.outputstyle(3))
        sheet.write_merge(3,3,4,5,"The percentage between 0.1 and 0.3",self.outputstyle(7))
        sheet.write_merge(4,4,4,5,"The percentage between 0.3 and 0.5",self.outputstyle(6))
        sheet.write_merge(5,5,4,5,"The percentage over 0.5",self.outputstyle(2))
        cows=0 
        col=0
        Aname=self.pg_config.getAdbname()
        Bname=self.pg_config.getBdbname()
        sheet.write(cows,col,"The letter A represents the database:",style)
        sheet.write_merge(cows,cows,col+1,col+2,Aname,style)
        cows+=1
        sheet.write(cows,col,"The letter B represents the database:",style)
        sheet.write_merge(cows,cows,col+1,col+2,Bname,style)
        
        cows+=3
        col=0
        btorder=("tablenum_A","tablenum_B","AB_percent")
        for bt in btorder:
            for (k, v) in data[0].items():
                if bt==k:
                    sheet.write(cows,col,k,style)
                    col+=1
        cows+=1           
        
        for t in data[0:flag]:
            col=0
            for bt in btorder:
                for (k, v) in t.items(): 
                    if bt==k:
                        if k=="AB_percent":
                            stylep=self.ABpercentcolor(v)
                                                          
                            sheet.write(cows,col,v,stylep)
                        else:                                                           
                            sheet.write(cows,col,v)
                        col+=1
            cows=cows+1   
        cows=cows+2
        tborder=("table_name","table_A","table_B","table_record_number_A","table_record_number_B","trn_AB_percent")
        col=0
        for tbo in tborder:
            for (k, v) in data[flag].items():
                if tbo==k:
                    sheet.write(cows,col,k,style)
                    col+=1
        cows+=1  
        stylev0=self.outputstyle(5)
        for t in data[flag:]:
            col=0
            for tbo in tborder:
                for (k, v) in t.items():
                    if tbo==k:
                        if k!="table_name":
                            if v==0:
                                sheet.write(cows,col,v,stylev0)
                            elif k=="trn_AB_percent":
                                stylep=self.ABpercentcolor(v)                                                          
                                sheet.write(cows,col,v,stylep)
                            else:
                                sheet.write(cows,col,v)
                        else:
                            sheet.write(cows,col,v)
                        col+=1           
            cows=cows+1
    
    def sheetwritefd(self,data,sheet,flag):        
        self.pg_log.info('write the field information to excel......')
        style=self.outputstyle(1)
        cows=0 
        col=0
        ftorder=("table_name","fieldnum_A","fieldnum_B","AB_percent")
        for fto in ftorder:
            for (k, v) in data[0][0].items():
                if fto==k:
                    sheet.write(cows,col,k,style)
                    col+=1
        cows+=1           
        
        for t in data[0:flag]:
            for f in t:
                col=0
                for fto in ftorder:
                    for (k, v) in f.items(): 
                        if fto==k:                               
                            if k=="AB_percent":
                                stylep=self.ABpercentcolor(v)
                                                                
                                sheet.write(cows,col,v,stylep)
                            else:                                                           
                                sheet.write(cows,col,v)                            
                            col+=1
                cows=cows+1   
        
    
    def sheetwritefdall(self,data,sheet,flag):        
        self.pg_log.info('write all field information to excel......')
        style=self.outputstyle(1)
        cows=0 
        col=0        
        fdorder=("table_name","field_A","field_B","data_type_A","data_type_B","differ")        
        for fdo in fdorder:
            sheet.write(cows,col,fdo,style)
            col+=1        
        cows+=1  
        stylef0=self.outputstyle(5)
        styletp=self.outputstyle(24)
        
        for t in data[flag:]:
            for f in t:
                col=0
                for fdo in fdorder:
                    for (k, v) in f.items():
                        if fdo==k:                
                            if k=="field_A":
                                if int(v)==0:
                                    sheet.write(cows,col,"",stylef0)
                                else:
                                    sheet.write(cows,col,f["field"])
                            elif k=="field_B":
                                if int(v)==0:
                                    sheet.write(cows,col,"",stylef0)
                                else:
                                    sheet.write(cows,col,f["field"])
                            elif k=="data_type_A":
                                if str(v)!="1":
                                    if str(v)=="0":
                                        sheet.write(cows,col,"",styletp)
                                    else:
                                        sheet.write(cows,col,v,styletp)
                                else:                                  
                                    sheet.write(cows,col,f["data_type"])
                            elif k=="data_type_B":
                                if str(v)!="1":
                                    if str(v)=="0":
                                        sheet.write(cows,col,"",styletp)
                                    else:
                                        sheet.write(cows,col,v,styletp)
                                else:
                                    sheet.write(cows,col,f["data_type"])                       
                            elif k=="differ":
                                if str(v)!="N":
                                    sheet.write(cows,col,v,styletp)
                                else:
                                    sheet.write(cows,col,v)
                            else:        
                                sheet.write(cows,col,v)
                            col+=1           
                cows=cows+1
                
    def writes(self,data,fdata):
        self.pg_log.info('start to write Excel file......')      
        names = self.GetFileName()        
                
        wbk = xlwt.Workbook(encoding='utf-8')
        
        fdallsheet = wbk.add_sheet('fieldallinfo')
        fdallsheet.col(0).width = 256 * 35
        fdallsheet.col(1).width = 256 * 20
        fdallsheet.col(2).width = 256 * 20
        fdallsheet.col(3).width = 256 * 20
        fdallsheet.col(4).width = 256 * 20
        fdallsheet.col(5).width = 256 * 20
        fdallflag=self.pg_config.getfdallflag()
        self.sheetwritefdall(fdata,fdallsheet,fdallflag)
        
        tbsheet = wbk.add_sheet('tableinfo')
        tbsheet.col(0).width = 256 * 45
        tbsheet.col(1).width = 256 * 15
        tbsheet.col(2).width = 256 * 15
        tbsheet.col(3).width = 256 * 28
        tbsheet.col(4).width = 256 * 28
        tbsheet.col(5).width = 256 * 18       
        tbflag=self.pg_config.gettbflag()
        self.sheetwritetb(data,tbsheet,tbflag)
        
        fdsheet = wbk.add_sheet('fieldinfo')
        fdsheet.col(0).width = 256 * 35
        fdsheet.col(1).width = 256 * 13
        fdsheet.col(2).width = 256 * 13
        fdsheet.col(3).width = 256 * 13                       
        fdflag=self.pg_config.getfdflag()
        self.sheetwritefd(fdata,fdsheet,fdflag)

        wbk.save(names + '.xls')        
        self.pg_log.info('write Excel file finished......')