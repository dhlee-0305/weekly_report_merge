
from docx.shared import Pt

"""폰트의 사이즈와 굵기 설정
:param  fontSize : 폰트 크기(픽셀 단위)
:param bold : 굵기 여부(True or False)
:return: None
"""
def setFontSizeBold(paragraph, fontSize, bold):
    """
    폰트의 사이즈와 굵기 설정
    
    :param fontSize: 폰트 크기(픽셀 단위)
    :param bold: 굵기 여부(True or False)
    :return: None
    """
    run_obj = paragraph.runs
    run = run_obj[0]
    font = run.font
    font.size = Pt(fontSize)
    font.bold = bold

