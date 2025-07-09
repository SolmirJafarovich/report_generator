from pptx.util import Inches

def mm(val):
    """Конвертация миллиметров в pptx Inches"""
    return Inches(val / 25.4)
