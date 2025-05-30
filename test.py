import configparser as parser
import logging
from enum import Enum

from weekly_env import *

print("Hello World")

# config = parser.ConfigParser()
# config.read('C:\Downloads\weekly_report\output\\report_merge.ini', encoding="UTF-8")

# print('REPORT_PATH:'+config['PATH']['REPORT_PATH'])
# print('REPORT_PATH:'+config['PATH']['OUTPUT_PATH'])
# print('REPORT_PATH:'+config['PATH']['TEMPLATE_REPORT_FILE'])

# print('REPORT_PATH:'+config['ENV']['REFERENCE_DAYOFWEEK'])

class LOG_LEVEL(Enum):
    DEBUG    = logging.DEBUG
    INFO     = logging.INFO
    WARNING  = logging.WARNING
    ERROR    = logging.ERROR
    CRITICAL = logging.CRITICAL

logLevel = 'DEBUG'
print(LOG_LEVEL[logLevel].value)
print(LOG_LEVEL['DEBUG'].value)
