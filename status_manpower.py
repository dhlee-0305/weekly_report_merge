from docx import Document
from docx.shared import Pt
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_LINE_SPACING

from report_env import *
from paragraph_style import *
from logger import *
from report_load import *

log = getLogger('status_manpower')

# 2. 인력 운용 현황
def loadManpowerStatus(src, dst):
    """ 
    인력 운용현황을 읽어와 취합 문서에 추가한다.
    
    :param src: 소스(팀) 문서
    :param dst: 취합 문서
    :return: None
    """    
    log.debug('2. 인력 운용 현황')

    dst_row_size = len(dst.rows)
    read_row_size, read_col_size = getColRowSize(src)

    regular_sum = 0;    # 정규직 합계
    contract_sum = 0;   # 계약직 합계
    total_sum = 0;      # 총 합계

    if read_row_size != MANPOWER_STATUS_ROW_SIZE or read_col_size != MANPOWER_STATUS_COL_SIZE:
        log.error("loadManpowerStatus 인력 운용 현황 row or column size invalid: "+ str(read_row_size)+", "+str(read_col_size))
        exit()

    # 구분(0, 0) 제외
    for row in range(1, read_row_size):
        regular_contract_sum = 0
        for col in range(1, read_col_size):
            # 팀 주간 보고 cell
            srcCell = src.rows[row].cells[col]
            
            # 취합 대상 cell 찾아옴
            gubun = str(src.rows[row].cells[0].text)
            dstCell = findCellByKeyword(dst, col, gubun)
            if dstCell == "":
                log.error('잘못된 \'구분\'값입니다.(오타 확인) : ['+gubun+"]")
                exit() 

            # col : (1)정규직, (2)계약직, (3) 합계, (4)증감, (5)증감사유, (6)충원 예상 인력 요청
            if col in [1, 2, 4]: # (1)정규직, (2)계약직, (4)증감
                sumValue = cellToNumber(srcCell) + cellToNumber(dstCell)
                if sumValue > 0:
                    if row != MANPOWER_STATUS_ROW_SIZE -1: # 합계 로우는 누적 계산을 위해 업데이트 하지 않음 => 최종 계산 후 업데이트
                        dstCell.text = str(sumValue)
                else:   
                    dstCell.text = ''

                # (합계 컬럼용) 정규직 + 계약직 합계
                regular_contract_sum = regular_contract_sum + sumValue

            elif col == 3: # (3)합계
                if row != MANPOWER_STATUS_ROW_SIZE -1: # 합계 로우는 누적 계산을 위해 업데이트 하지 않음 => 최종 계산 후 업데이트
                    dstCell.text = str(regular_contract_sum)

            else: # (5)증감사유, (6)충원 예상 인력 요청
                if len(dstCell.text) > 0 and len(srcCell.text) > 0:
                    dstCell.text = dstCell.text + '\n' + srcCell.text
                elif len(dstCell.text) == 0 and len(srcCell.text) > 0:
                    dstCell.text = srcCell.text

            dstCell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            
            if row == MANPOWER_STATUS_ROW_SIZE -1:
                # 합계(마지막) 로우 - 컬럽별 합계 누적 업데이트
                if col == 1: dstCell.text = str(cellToNumber(dstCell) + regular_sum)
                if col == 2: dstCell.text = str(cellToNumber(dstCell) + contract_sum)
                if col == 3: dstCell.text = str(cellToNumber(dstCell) + total_sum)

                # 최종 합계 볼드 처리
                if len(dstCell.text) > 0:
                    setFontSizeBold(dstCell.paragraphs[0], 9, True)
            else:   
                # 마지막 로우가 아닌 경우 컬럼별 합계 누적
                if col == 1: regular_sum  += cellToNumber(srcCell)
                if col == 2: contract_sum += cellToNumber(srcCell)
                if col == 3: total_sum    += cellToNumber(srcCell)

                # 최종 합계 외 볼드 미적용
                if len(dstCell.text) > 0:
                    setFontSizeBold(dstCell.paragraphs[0], 9, False)