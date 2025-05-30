from docx import Document
from docx.shared import Pt
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_LINE_SPACING

from weekly_env import *
from paragraph_style import *
from logger import *
from weekly_load import *

log = getLogger('status_manpower')

# 2. 인력 운용 현황
def loadManpowerStatus(teamWeeklyDoc, weeklyMergeDoc):
    """ 
    인력 운용현황을 읽어와 취합 문서에 추가한다.
    
    :param teamWeeklyDoc: 소스(팀) 문서
    :param weeklyMergeDoc: 취합 문서
    :return: None
    """    
    log.debug('2. 인력 운용 현황')

    readRowSize, readColSize = getColRowSize(teamWeeklyDoc)

    regularSum = 0;    # 정규직 합계
    contractSum = 0;   # 계약직 합계
    totalSum = 0;      # 총 합계

    if readRowSize != MANPOWER_STATUS_ROW_SIZE or readColSize != MANPOWER_STATUS_COL_SIZE:
        log.error("loadManpowerStatus 인력 운용 현황 row or column size invalid: "+ str(readRowSize)+", "+str(readColSize))
        exit()

    # 구분(0, 0) 제외
    for row in range(1, readRowSize):
        regularAndContractSum = 0
        for col in range(1, readColSize):
            # 팀 주간 보고 cell
            teamWeeklyDocCell = teamWeeklyDoc.rows[row].cells[col]
            
            # '구분' 항목으로 취합 대상 cell 매칭
            gubun = str(teamWeeklyDoc.rows[row].cells[0].text)
            weeklyMergeDocCell = findCellByKeyword(weeklyMergeDoc, col, gubun)
            if weeklyMergeDocCell == "":
                log.error('잘못된 \'구분\'값입니다.(오타 확인) : ['+gubun+"]")
                exit() 

            # col : (1)정규직, (2)계약직, (3) 합계, (4)증감, (5)증감사유, (6)충원 예상 인력 요청
            if col in [1, 2, 4]: # (1)정규직, (2)계약직, (4)증감
                sumValue = cellToNumber(teamWeeklyDocCell) + cellToNumber(weeklyMergeDocCell)
                if sumValue > 0:
                    if row != MANPOWER_STATUS_ROW_SIZE -1: # 합계 로우는 누적 계산을 위해 업데이트 하지 않음 => 최종 계산 후 업데이트
                        weeklyMergeDocCell.text = str(sumValue)
                else:   
                    weeklyMergeDocCell.text = ''

                # (합계 컬럼용) 정규직 + 계약직 합계
                regularAndContractSum = regularAndContractSum + sumValue

            elif col == 3: # (3)합계
                if row != MANPOWER_STATUS_ROW_SIZE -1: # 합계 로우는 누적 계산을 위해 업데이트 하지 않음 => 최종 계산 후 업데이트
                    weeklyMergeDocCell.text = str(regularAndContractSum)

            else: # (5)증감사유, (6)충원 예상 인력 요청
                if len(weeklyMergeDocCell.text) > 0 and len(teamWeeklyDocCell.text) > 0:
                    weeklyMergeDocCell.text = weeklyMergeDocCell.text + '\n' + teamWeeklyDocCell.text
                elif len(weeklyMergeDocCell.text) == 0 and len(teamWeeklyDocCell.text) > 0:
                    weeklyMergeDocCell.text = teamWeeklyDocCell.text

            weeklyMergeDocCell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            
            if row == MANPOWER_STATUS_ROW_SIZE -1:
                # 합계(마지막) 로우 -> 컬럽별 합계 누적값 업데이트
                weeklyMergeDocCell.text = calcTotalSumByCol(weeklyMergeDocCell, col, regularSum, contractSum, totalSum)

                # 최종 합계 볼드 처리
                if len(weeklyMergeDocCell.text) > 0:
                    setFontSizeBold(weeklyMergeDocCell.paragraphs[0], 9, True)
            else:   
                # 마지막 로우가 아닌 경우엔 컬럼별 합계만 누적함
                regularSum, contractSum, totalSum = updateColSumByCol(teamWeeklyDocCell, col, regularSum, contractSum, totalSum)

                # 최종 합계 외 볼드 미적용
                if len(weeklyMergeDocCell.text) > 0:
                    setFontSizeBold(weeklyMergeDocCell.paragraphs[0], 9, False)

def calcTotalSumByCol(cell, col, regularSum, contractSum, totalSum):
    """
    컬럼에 해당하는 합계 값 반환

    :param cell: 대상 셀
    :param col: 컬럼 인덱스
    :param regularSum: 정규직 합계
    :param contractSum: 계약직 합계
    :param totalSum: 총 합계
    :return: 해당 컬럼의 합계 값
    """
    if col == 1:
        return str(cellToNumber(cell) + regularSum)
    elif col == 2:
        return str(cellToNumber(cell) + contractSum)
    elif col == 3:
        return str(cellToNumber(cell) + totalSum)
    else:
        return ""

def updateColSumByCol(cell, col, regularSum, contractSum, totalSum):
    """
    컬럼에 해당하는 합계 값 업데이트

    :param cell: 대상 셀
    :param col: 컬럼 인덱스
    :param regularSum: 정규직 합계
    :param contractSum: 계약직 합계
    :param totalSum: 총 합계
    :return: 정규직 합계, 계약직 합계, 총 합계
    """
    if col == 1:
        regularSum += cellToNumber(cell)
    elif col == 2:
        contractSum += cellToNumber(cell)
    elif col == 3:
        totalSum += cellToNumber(cell)
    
    return regularSum, contractSum, totalSum