import datetime
import configparser as parser
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE

import docx.document
import os

# Table row/column Size in document
PROJECT_STATUS_COL_SIZE  = 7
MANPOWER_STATUS_ROW_SIZE = 9
MANPOWER_STATUS_COL_SIZE = 7
CLIENT_STATUS_COL_SIZE   = 4

dayOfWeekDic = ['월요일', '화요일', '수요일', '목요일', '금요일', '토요일', '일요일']

# CONFIG_FILE = 'C:\Downloads\weekly_report\output\\report_merge.ini'
CONFIG_FILE = '.\\report_merge.ini'


def checkEnv():
    config = loadConfig(CONFIG_FILE)

    if not os.path.exists(config['PATH']['REPORT_PATH']):
        os.makedirs(config['PATH']['REPORT_PATH'])
    if not os.path.exists(config['PATH']['OUTPUT_PATH']):
        os.makedirs(config['PATH']['OUTPUT_PATH'])
    if not os.path.exists(config['PATH']['LOG_FILE_PATH']):
        os.makedirs(config['PATH']['LOG_FILE_PATH'])


def loadConfig(file_path):
    """설정 파일 내용 읽어옴
    :param file_path: 설정파일 전체 경로, 읽은 내용이 없는 경우 현재 디렉토리(./)에서 설정 파일을 찾음
    :return: ConfigParser
    """
    config = parser.ConfigParser()
    config.read(file_path, encoding="UTF-8")

    if len(config.sections()) == 0:
        config.read('.\\report_merge.ini', encoding="UTF-8")

    return config


def getReportList(path):
    """리포트 파일 리스트를 읽어옴
    :param path: 리포트 파일이 저장된 폴더 경로
    :return: 파일 이름 리스트
    """
    file_list = [f for f in os.listdir(path) if ((not f.startswith('~$')) and (os.path.isfile(path+f)))] 
    for file in file_list:
        if (file.find('.docx') == -1 or file.find('output') > 0):
            file_list.remove(file)
    return file_list

def getReportFilePrefix(dayOfWeek):
    """리포트 통합 파일 Prefix(날짜 형식) 문자열 생성
    :param dayOfWeek: 요일 ex) 금요일
    :return: YYYYMMDD 형식 문자열
    """
    for dayAdd in range(7):
        dayAdded = (datetime.datetime.now() + datetime.timedelta(days=dayAdd))
        dow = dayOfWeekDic[dayAdded.weekday()]
        if(dayOfWeek == dow):
            return dayAdded.strftime("%Y%m%d")

def getReportDate(dayOfWeek):
    """리포트 통합 파일 내 표시될 기준 날짜 생성
    :param dayOfWeek: 요일 ex) 금요일
    :return: YYYY.MM.DD(dayOfWeek) 형식 문자열 ex) 2023.07.07(금요일)
    """
    for dayAdd in range(7):
        dayAdded = (datetime.datetime.now() + datetime.timedelta(days=dayAdd))
        dow = dayOfWeekDic[dayAdded.weekday()]
        if(dayOfWeek == dow):
            return dayAdded.strftime("%Y.%m.%d")+'('+dayOfWeek[0:1]+')'



if __name__ == '__main__':
    checkEnv()