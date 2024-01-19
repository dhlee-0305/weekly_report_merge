import logging
import sys
from enum import Enum

from report_env import *

class LOG_LEVEL(Enum):
    DEBUG    = logging.DEBUG
    INFO     = logging.INFO
    WARNING  = logging.WARNING
    ERROR    = logging.ERROR
    CRITICAL = logging.CRITICAL

def getLogger(logName):
    """공통 로그 인스턴스를 반환
    :param logName: 로거 이름
    :return: logging 인스턴스
    """
    config = loadConfig(CONFIG_FILE)

    log = logging.getLogger(name=logName)
    logLevel = config['ENV']['LOG_LEVEL']
    log.setLevel(LOG_LEVEL[logLevel].value)
    
    if log.hasHandlers():
        log.handlers.clear()

    logFormat = logging.Formatter('%(asctime)s|%(name)s|%(funcName)s|%(levelname)s|%(message)s', datefmt='%Y-%m-%d %H:%M:%S')
    YYYYMMDD = datetime.datetime.now().strftime("%Y%m%d")
    logFileHandler = logging.FileHandler(config['PATH']['LOG_FILE_PATH']+YYYYMMDD+"_report_merge.log", encoding='utf-8')
    logFileHandler.setFormatter(logFormat)
    log.addHandler(logFileHandler)

    return log
