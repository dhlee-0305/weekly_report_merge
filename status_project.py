from docx import Document
from docx.shared import Pt
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_LINE_SPACING

from report_env import *
from paragraph_style import *
from logger import *
from report_load import *

log = getLogger('status_project')

# 1. 프로젝트 진행 현황            
def loadProjectStatus(src, dst):
    """ 
    프로젝트 진행사항을 읽어와 취합 문서에 추가한다.
    
    :param src: 소스(팀) 문서
    :param dst: 취합 문서
    :return: None
    """    
    log.debug('1. 프로젝트 진행 현황')
    row_size, col_size = getColRowSize(src)
    
    if col_size != PROJECT_STATUS_COL_SIZE:
        log.error("loadProjectStatus 프로젝트 진행 현황 column size invalid: "+ str(col_size))
        exit()

    for row in range(row_size-1):
        row_cells = dst.add_row().cells
        for col in range(col_size):
            row_cells[col].text = src.rows[row+1].cells[col].text
            row_cells[col].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            setFontSizeBold(row_cells[col].paragraphs[0], 9, False)