from docx import Document
from docx.shared import Pt
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_LINE_SPACING

from report_env import *
from paragraph_style import *
from logger import *
from elapsed import *

log = getLogger('report_load')

def printTable(table):
    """
    표 내용을 전체 출력
    
    :param table: 대상 테이블
    """
    row_size = len(table.rows)
    col_size = len(table.rows[0].cells)
    log.debug("\nprint_table() rows size:"+ str(row_size)+", column size:"+ str(col_size))

    for row in range(row_size):
        row_str = ""
        for col in range(col_size):
            if(len(table.rows[row].cells[col].text) > 0):
                row_str = row_str+table.rows[row].cells[col].text+"\t|\t"
            #log.debug(table.rows[row].cells[col].text+"\t")
        if(len(row_str) > 0):
            log.debug('['+str(row)+'] '+row_str.replace('\n', ''))

def getBlankRowCount(table):
    """
    표 내에 공백 로우가 몇 개인지 카운팅

    :param table: 대상 테이블
    """
    row_size = len(table.rows)
    col_size = len(table.rows[0].cells)
    blankCount = 0

    for row in range(row_size):
        isBlank = True
        for col in range(col_size):
            if len(table.rows[row].cells[col].text) > 0:
                isBlank = False
        if isBlank:
            blankCount += 1

    return blankCount

def cellToNumber(cell):
    """ 
    cell 값 숫자 변환
    
    :return: 숫자, 숫자가 아닌 경우 0
    """
    if(cell.text.isdigit()):
        return int(cell.text or 0)
    else:
        return 0
    
def findCellByKeyword(dst, index, keyword):
    """ 
    cell 배열에서 특정 키워드 찾기

    :return: 찾은 cell의 지정된 위치 값, 못찾았으면 공백
    """
    isFindDstCell = ""
    for dstRow in range(1, len(dst.rows)):
        if dst.rows[dstRow].cells[0].text == keyword :
            isFindDstCell = dst.rows[dstRow].cells[index]
            break
    return isFindDstCell

def loadTop(dst, dayOfWeek):
    """ 
    타이틀 영역 내 지정된 요일 기준으로 날짜를 계산하여 보고일로 업데이트 함
    
    :param dst: 취합 문서
    :param dayOfWeek: 보고일 기준 요일
    :return: 없음
    """
    dst.rows[1].cells[1].text = getReportDate(dayOfWeek)

    paragraphs = dst.rows[1].cells[1].paragraphs
    paragraph = paragraphs[0]
    paragraph.alignment = WD_TABLE_ALIGNMENT.RIGHT
    setFontSizeBold(paragraphs[0], 9, False)

def loadTeamName(src):
    """ 
    타이틀 영역에서 팀 이름을 가져옴
    
    :param src: 취합 문서
    :return: 팀 이름
    """
    # printTable(src)
    if len(src.rows[0].cells) == 3:
        teamName = '['+src.rows[0].cells[2].text.replace('SD본부', '').replace('IT서비스부문', '').replace('(', '').replace(')', '').strip()+']'
    else:
        teamName = '['+src.rows[1].cells[0].text.replace('SD본부', '').replace('IT서비스부문', '').replace('(', '').replace(')', '').strip()+']'
    return teamName

