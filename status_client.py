from docx import Document
from docx.shared import Pt
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_LINE_SPACING

from report_env import *
from paragraph_style import *
from logger import *
from report_load import *

log = getLogger('status_client')

# 3. 거래처 영업/동향 정보
def loadClientStatus(src, dst):
    """ 
    거래처 영업/동향을 읽어와 취합 문서에 추가한다.
    
    :param src: 소스(팀) 문서
    :param dst: 취합 문서
    :return: None
    """    
    log.debug('3. 거래처 영업/동향 정보')

    row_size, col_size = getColRowSize(src)
    
    if col_size != CLIENT_STATUS_COL_SIZE:
        log.error("loadClientStatus 거래처 영업/동향 정보 column size invalid: "+ str(col_size))
        exit()


    for row in range(row_size-1):
        if (len(src.rows[row+1].cells[2].text) > 3) :
            row_cells = dst.add_row().cells
        else:
            # 공백 행 제외
            continue
        
        for col in range(col_size):
            if col == 2: # (2)주요 정보
                customerIssueStrArray = src.rows[row+1].cells[col].text.strip().split('\n')
                isTitle = True
                for customerStrIssueStr in customerIssueStrArray:
                    if len(customerStrIssueStr) > 1:
                        if isTitle :
                            insert_paragraph = row_cells[col].paragraphs[0]
                            insert_paragraph.add_run(customerStrIssueStr)
                        else:
                            if(customerStrIssueStr.startswith('-')):
                                insert_paragraph = row_cells[col].add_paragraph(' ' + customerStrIssueStr) # 서식 무시하고 '-' 추가한 경우 제외를 위함
                            else:
                                insert_paragraph = row_cells[col].add_paragraph(' - ' + customerStrIssueStr)
                        
                        setStyleCustomerIssue(insert_paragraph, isTitle)
                        isTitle = False

            else: # (0)구분, (1)고객사/부서, (3)비고
                row_cells[col].text = src.rows[row+1].cells[col].text

                row_cells[col].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                setFontSizeBold(row_cells[col].paragraphs[0], 9, False)

def setStyleCustomerIssue(insert_paragraph, isTitle):
    """
    거래처 영업/동향 정보의 주요 정보 내용 스타일 설정
    :param insert_paragraph: 단락
    :param isTitle: 제목 여부
    :return: 없음
    """
    
    insert_paragraph.alignment = WD_TABLE_ALIGNMENT.LEFT
    paragraph_format = insert_paragraph.paragraph_format
    paragraph_format.line_spacing = Pt(12)
    paragraph_format.space_before = Pt(5)
    
    run_obj = insert_paragraph.runs
    run = run_obj[0]
    font = run.font
    font.size = Pt(9)

    if isTitle:
        font.bold = True
        isTitle = False
    else:
        font.bold = False