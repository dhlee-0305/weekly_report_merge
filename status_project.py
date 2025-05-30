from docx import Document
from docx.shared import Pt
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_LINE_SPACING

from weekly_env import *
from paragraph_style import *
from logger import *
from weekly_load import *

log = getLogger('status_project')

# 1. 프로젝트 진행 현황            
def loadProjectStatus(teamWeeklyDoc, weeklyMergeDoc):
    """ 
    프로젝트 진행사항을 읽어와 취합 문서에 추가한다.
    
    :param teamWeeklyDoc: 소스(팀) 문서
    :param weeklyMergeDoc: 취합 문서
    :return: None
    """    
    log.debug('1. 프로젝트 진행 현황')
    rowSize, colSize = getColRowSize(teamWeeklyDoc)
    
    if colSize != PROJECT_STATUS_COL_SIZE:
        log.error("loadProjectStatus 프로젝트 진행 현황 column size invalid: "+ str(colSize))
        exit()

    for row in range(rowSize-1):
        rowCells = weeklyMergeDoc.add_row().cells
        for col in range(colSize):
            rowCells[col].text = teamWeeklyDoc.rows[row+1].cells[col].text
            rowCells[col].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            setFontSizeBold(rowCells[col].paragraphs[0], 9, False)