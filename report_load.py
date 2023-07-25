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
    """표 내용을 전체 출력
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
    """표 내에 공백 로우가 몇 개인지 카운팅
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

# 타이틀 영역
def loadTop(dst, dayOfWeek):
    log.debug('0. 타이틀 영역 처리')
    dst.rows[1].cells[1].text = getReportDate(dayOfWeek)

    paragraphs = dst.rows[1].cells[1].paragraphs
    paragraph = paragraphs[0]
    paragraph.alignment = WD_TABLE_ALIGNMENT.RIGHT
    setFontSizeBold(paragraphs[0], 9, False)

# 0. 팀 명
def loadTeamName(src):
    teamName = '['+src.rows[0].cells[2].text.replace('SD본부', '').replace('IT서비스부문', '').strip()+']'
    return teamName

# 1. 프로젝트 진행 현황            
def loadProjectStatus(src, dst):
    log.debug('1. 프로젝트 진행 현황')
    row_size = len(src.rows)
    col_size = len(src.rows[0].cells)
    
    if col_size != PROJECT_STATUS_COL_SIZE:
        log.error("loadProjectStatus 프로젝트 진행 현황 column size invalid: "+ str(col_size))
        exit()

    # 공백 제거
    row_size = row_size - getBlankRowCount(src)

    for row in range(row_size-1):
        row_cells = dst.add_row().cells
        for col in range(col_size):
            row_cells[col].text = src.rows[row+1].cells[col].text
            row_cells[col].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            setFontSizeBold(row_cells[col].paragraphs[0], 9, False)

# 2. 인력 운용 현황
def loadManpowerStatus(src, dst):
    log.debug('2. 인력 운용 현황 처리')

    row_size = len(src.rows)
    col_size = len(src.rows[0].cells)

    if row_size != MANPOWER_STATUS_ROW_SIZE or col_size != MANPOWER_STATUS_COL_SIZE:
        log.error("loadManpowerStatus 인력 운용 현황 row or column size invalid: "+ str(row_size)+", "+str(col_size))
        exit()

    regular_sum = 0
    contract_sum = 0
    # 구분(0, 0) 제외
    for row in range(1, row_size):
        type_sum = 0
        for col in range(1, col_size):
            srcCell = src.rows[row].cells[col]
            dstCell = dst.rows[row].cells[col]
            if col == 1 or col == 2 or col == 4: 
                # (1)정규직, (2)계약직, (4)증감
                if(srcCell.text.isdigit()):
                    src_man_count = int(srcCell.text or 0)
                else:
                    src_man_count = 0;
                if(dstCell.text.isdigit()):
                    dst_man_count = int(dstCell.text or 0)
                else:
                    dst_man_count = 0;

                # 로우 합계
                type_sum = type_sum + src_man_count + dst_man_count

                if (src_man_count + dst_man_count) > 0:
                    if col == 1:
                        regular_sum = src_man_count + dst_man_count
                        dstCell.text = str(regular_sum)
                    elif col == 2:
                        contract_sum = src_man_count + dst_man_count
                        dstCell.text = str(contract_sum)
                    else:
                        dstCell.text = str(src_man_count + dst_man_count)
                else:
                    dstCell.text = ''

            elif col == 3: 
                # (로우) 합계
                dstCell.text = str(type_sum)
            else: 
                # 증감 사유, 충원 예상 인력 요청
                if len(dstCell.text) > 0 and len(srcCell.text) > 0:
                    dstCell.text = dstCell.text + '\n' + srcCell.text
                elif len(dstCell.text) == 0 and len(srcCell.text) > 0:
                    dstCell.text = srcCell.text
                else:
                    pass
                    #dstCell.text = dstCell.text
            
            if len(dstCell.text) > 0:
                dstCell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                if srcCell.text == '합계':
                    setFontSizeBold(dstCell.paragraphs[0], 9, True)
                else:
                    setFontSizeBold(dstCell.paragraphs[0], 9, False)


# 3. 거래처 영업/동향 정보
def loadClientStatus(src, dst):
    log.debug('3. 거래처 영업/동향 정보 처리')

    row_size = len(src.rows)
    col_size = len(src.rows[0].cells)
    
    if col_size != CLIENT_STATUS_COL_SIZE:
        log.error("loadClientStatus 거래처 영업/동향 정보 column size invalid: "+ str(col_size))
        exit()

    # 공백 제거
    row_size = row_size - getBlankRowCount(src)

    for row in range(row_size-1):
        row_cells = dst.add_row().cells
        for col in range(col_size):
            if col == 2: # 주요 정보
                customerIssueStrArray = src.rows[row+1].cells[col].text.strip().split('\n')
                isTitle = True
                for customerStrIssueStr in customerIssueStrArray:
                    if len(customerStrIssueStr) > 1:
                        if isTitle :
                            insert_paragraph = row_cells[col].paragraphs[0]
                            insert_paragraph.add_run(customerStrIssueStr);
                        else:
                            insert_paragraph = row_cells[col].add_paragraph(' - ' + customerStrIssueStr);
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
            else: # 구분, 고객사/부서, 비고
                row_cells[col].text = src.rows[row+1].cells[col].text

                row_cells[col].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                setFontSizeBold(row_cells[col].paragraphs[0], 9, False)
        
# 4. 주요 업무 사항
def loadWork(src, dst, teamName):
    log.debug('4. 주요 업무 사항 처리')

    find_work = False
    for i, paragraph in enumerate(src.paragraphs):
        if (paragraph.text == "주요 업무 사항"):
            find_work = True
            insert_paragraph = dst.add_paragraph(teamName)
            insert_paragraph.alignment = WD_TABLE_ALIGNMENT.LEFT
            setFontSizeBold(insert_paragraph, 11, True)
        elif find_work == True:
            if(len(paragraph.text) > 0):
                insert_paragraph = dst.add_paragraph(paragraph.text)
                insert_paragraph.style = paragraph.style

if __name__ == '__main__':
    doc = Document(r'C:\Downloads\imsi\python\20230622_경영전략회의_SD본부_오픈서비스사업팀.docx')
    template = Document('C:\Downloads\weekly_report\output\경영전략회의_template.docx')

    loadTop(template.tables[0])
    loadProjectStatus(doc.tables[1], template.tables[1])
    loadManpowerStatus(doc.tables[2], template.tables[2])
    loadClientStatus(doc.tables[3], template.tables[3])
    loadWork(doc, template, "[오픈서비스사업팀]")
