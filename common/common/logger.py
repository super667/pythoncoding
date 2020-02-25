# -*- coding: utf-8 -*-

import logging

def create_main_logger():
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)
    formatter = logging.Formatter('%(asctime)s %(levelname)8s - %(message)s')

    # console
    ch = logging.StreamHandler()
    ch.setLevel(logging.DEBUG)
    ch.setFormatter(formatter)
    if not logger.handlers:
        logger.addHandler(ch)

    return logger

def create_proj_logger(proj, logfile):
    logger = logging.getLogger(proj)
    logger.setLevel(logging.DEBUG)
    formatter = logging.Formatter('%(asctime)s %(levelname)8s - %(message)s')

    # logfile
    fh = logging.FileHandler(filename=logfile)
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(formatter)
    logger.addHandler(fh)

    return logger

def drop_loggers():
    logging.shutdown()
    
    
if __name__ =='__main__':
    logger = create_main_logger()
    logger.debug('This is a customer debug message')
    logger.info('This is an customer info message')
    logger.warning('This is a customer warning message')
    logger.error('This is an customer error message')
    logger.critical('This is a customer critical message')
    
    drop_loggers()
    
