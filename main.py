import MTK2
from docx import Document
from docx.shared import RGBColor
from docx.enum.text import WD_COLOR_INDEX

document = Document('variant07.docx')
paragraphs = document.paragraphs

def Spacing(run):
    return run._r.get_or_add_rPr().xpath("./w:spacing")

def Scale(run):
    return run._r.get_or_add_rPr().xpath("./w:w")

def main():
    code = ""
    count_color = 0
    count_size = 0
    count_highlight = 0
    count_spacing = 0
    count_scale = 0

    for paragraph in paragraphs:
        for run in paragraph.runs:
            font_color = run.font.color.rgb
            font_highlight_color = run.font.highlight_color
            font_size = run.font.size
            font_scale = Scale(run)
            font_spacing = Spacing(run)

            if (font_color != RGBColor(0, 0, 0) or
                    font_size.pt != 12.0 or
                    font_highlight_color != WD_COLOR_INDEX.WHITE or
                    font_spacing != [] or
                    font_scale != []):
                if font_color != RGBColor(0, 0, 0):
                    count_color += 1
                if font_size.pt != 12.0:
                    count_size += 1
                if font_highlight_color != WD_COLOR_INDEX.WHITE:
                    count_highlight += 1
                if font_spacing != []:
                    count_spacing += 1
                if font_scale != []:
                    count_scale += 1
                for i in range(len(run.text)):
                    code += '1'
            else:
                for i in range(len(run.text)):
                    code += '0'

    method = max(count_scale, count_spacing, count_highlight, count_size, count_color)

    if (method == count_size):
        print("Способ форматирования: по размеру шрифта")
    elif (method == count_spacing):
        print("Способ форматирования: по межсимвольному интервалу")
    elif (method == count_scale):
        print("Способ форматирования: по масштабу шрифта")
    elif (method == count_highlight):
        print("Способ форматирования: по цвету фона")
    elif (method == count_color):
        print("Способ форматирования: по цвету символов")

    code += "0000"
    # print(code)

    text = MTK2.Decode(code)
    print('МТК2:', text)
    text = bytes.fromhex(hex(int(code, 2))[2:]).decode(encoding="koi8_r")
    print("KOI-8R:", text)
    text = bytes.fromhex(hex(int(code, 2))[2:]).decode(encoding="cp866")
    print("CP866:", text)
    text = bytes.fromhex(hex(int(code, 2))[2:]).decode(encoding="cp1251")
    print("CP1251:", text)

main()