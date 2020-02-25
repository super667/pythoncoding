# -*- coding: utf-8 -*-

import re
import logging
import psycopg2

class CPostgresql(object):
    def __init__(self, project, path):
        self.logger = logging.getLogger(project)
        self.connected = False
        if path == '':
            self.logger.error("Host IP doesn't exist.")
        else:
            self.srv_path = path

    def connect(self):
        if self.connected == True:
            return False

        self.conn  = psycopg2.connect(self.srv_path)
        self.cur = self.conn.cursor()
        self.connected = True

    def close(self):
        if self.connected == True:
            self.cur.close()
            self.conn.close()
            self.connected = False

    def execute(self, sql, parameters=[]):
        try:
            if parameters:
                self.cur.execute(sql, parameters)
            else:
                self.cur.execute(sql)
            return (True, self.cur.fetchall())
        except:
            self.logger.exception('')
            self.commit()
            return (False, None)

    def execute_no_return(self, sqlcmd):
        self.cur.execute(sqlcmd)
        self.commit()
        
    def commit(self):
        return self.conn.commit()

    def fetchone(self):
        return self.cur.fetchone()

    def fetchall(self):
        return self.cur.fetchall()
    
    def run(self, filename):  
        fp = open(filename,'r')
        self.execute(fp.read())
        fp.close()

    def execute2(self, sqlcmd, parameters = []):
        '''execute commands '''
        if parameters:
            self.cur.execute(sqlcmd, parameters)
        else:
            self.cur.execute(sqlcmd)
        # self.conn.commit()
        return 0