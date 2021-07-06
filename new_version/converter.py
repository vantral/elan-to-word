import re
import ctypes

from docx import Document
from docx.shared import Pt, Cm
from collections import defaultdict
from docx.enum.text import WD_LINE_SPACING

MAX_LINE_LEN = 70

OUT_FONT = 'Times New Roman'
OUT_FONT_POINTS = 12


def getTextDimensions(text, points, font):
    class SIZE(ctypes.Structure):
        _fields_ = [("cx", ctypes.c_long), ("cy", ctypes.c_long)]

    hdc = ctypes.windll.user32.GetDC(0)
    hfont = ctypes.windll.gdi32.CreateFontA(points, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, font)
    hfont_old = ctypes.windll.gdi32.SelectObject(hdc, hfont)

    size = SIZE(0, 0)
    ctypes.windll.gdi32.GetTextExtentPoint32A(hdc, text, len(text), ctypes.byref(size))

    ctypes.windll.gdi32.SelectObject(hdc, hfont_old)
    ctypes.windll.gdi32.DeleteObject(hfont)

    return size.cx, size.cy

# print(getTextDimensions("Hello world", 12, "Times New Roman"))
# print(getTextDimensions("Hello world", 12, "Arial"))


def open_file(filename):
    text = open(filename, encoding='utf-8').read()
    return text


def elan_data(file):
    # elan = elan.replace('&', '')
    # elan = elan.replace('<', '&lt;')
    # elan = elan.replace('>', '&gt;')
    elan = open_file(file).splitlines()
    transc = defaultdict(str)
    transl = defaultdict(str)
    gloss = defaultdict(str)
    comment = defaultdict(str)

    for line in elan:
        tokens = line.split('\t')
        if len(tokens) == 9:
            indices = (0, 2, 4, 8)
        else:
            indices = (0, 2, 3, 4)

        print(tokens)
        print(indices)

        layer = tokens[indices[0]]
        time_start = tokens[indices[1]]
        time_finish = tokens[indices[2]]
        text = tokens[indices[3]]

        print(layer)
        print(time_start)
        print(time_finish)
        print(text)

        if layer == 'transcription':
            transc[(time_start, time_finish)] = text
        elif layer == 'translation':
            transl[(time_start, time_finish)] = text
        elif layer == 'gloss':
            gloss[(time_start, time_finish)] = text
        elif layer == 'comment':
            comment[(time_start, time_finish)] = text
    return transc, transl, gloss, comment


def to_word(pivot_dictionary):
    informant = 'eek' or input('введите код информанта ')
    date = '20210706' or input('введите дату ')
    expe = 'mb' or input('введите свой код ')
    name = f'eve_{informant}_{date}_{expe}'

    document = Document()
    sections = document.sections
    for section in sections:
        section.top_margin = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(1.5)

    document.add_paragraph()

    table = document.add_table(rows=5, cols=2)
    table.columns[0].width = Cm(4.5)
    table.columns[1].width = Cm(12.5)

    inf = table.rows[0].cells
    inf[0].text, inf[1].text = 'Информант', informant

    exp = table.rows[1].cells
    exp[0].text, exp[1].text = 'Экспедиционер', expe

    dat = table.rows[2].cells
    dat[0].text, dat[1].text = 'Дата', date

    els = table.rows[3].cells
    els[0].text, els[1].text = 'Кто ещё был на паре', input('Кто ещё был на паре? ')

    inf = table.rows[4].cells
    inf[0].text, inf[1].text = 'Примерная тематика', input('Примерная тематика ')

    for row in table.rows:
        for cell in row.cells:
            cell.paragraphs[0].paragraph_format.space_after = Cm(0)
    document.add_paragraph().paragraph_format.space_after = Cm(0)

    counter = 1

    print(pivot_dictionary)
    for key, value in pivot_dictionary.items():
        print(key, value)

        header = f'{counter}. {informant}_{date}@{expe}_{counter}'
        transcription = value[0]
        translation = value[1]
        gloss = value[2]
        comment = value[3]

        p = document.add_paragraph()
        paragraph_format = p.paragraph_format
        paragraph_format.space_after = Cm(0.1)
        p.add_run(header)

        transcriptions = []
        glosses = []
        # if len(transcription) >= MAX_LINE_LEN:

        # for token in transcription.replace(' ', '\t').split('\t'):
        transcription_tokens = transcription.split(' ')
        glosses_tokens = gloss.split(' ')
        gl_cur_len, gl_cur_run = 0, f''
        transcr_cur_len, transcr_cur_run = 0, f''
        last_par_index = 0
        for i, (transcription_token, gloss_token) in enumerate(
                zip(transcription_tokens, glosses_tokens)):
            # print(transcription_token, gloss_token)
            if (gl_cur_len + len(gloss_token) <= MAX_LINE_LEN
                and transcr_cur_len + len(transcription_token) <= MAX_LINE_LEN):
                # print(f'cond true: `{transcr_cur_run}`, `{gl_cur_run}`, {gl_cur_len}, {gl_cur_run}')
                transcr_cur_run += f'{transcription_token}\t'
                gl_cur_run += f'{gloss_token}\t'
                transcr_cur_len += len(transcription_token)
                gl_cur_len += len(gloss_token)
            else:
                # print(f'cond false: `{transcr_cur_run}`, `{gl_cur_run}`, {gl_cur_len}, {gl_cur_run}')
                transcriptions.append(transcr_cur_run.strip('\t'))
                glosses.append(gl_cur_run.strip('\t'))
                last_par_index += 1

                transcr_cur_run = f'{transcription_token}\t'
                gl_cur_run = f'{gloss_token}\t'
                transcr_cur_len, gl_cur_len = len(transcription_token), len(gloss_token)
        else:
            if len(glosses) - 1 == last_par_index - 1:
                # if num of added lines is 1 less than needed, added remaining
                transcriptions.append(transcr_cur_run.strip('\t'))
                glosses.append(gl_cur_run.strip('\t'))

        print(transcriptions)
        print(glosses)

        # TODO: determine tab stops on the go using native length rendering with font
        for i, (transcription_line, gloss_line) in enumerate(
                zip(transcriptions, glosses)):
            print(transcription_line, gloss_line)

            p_transcription = document.add_paragraph()
            paragraph_format = p_transcription.paragraph_format
            paragraph_format.space_after = Cm(0)
            paragraph_format.left_indent = Cm(0.5)
            paragraph_format.tab_stops.add_tab_stop(Cm(2))
            p_transcription.add_run(transcription_line.replace(' ', '\t')).font.italic = True

            p_glosses = document.add_paragraph()
            paragraph_format = p_glosses.paragraph_format
            paragraph_format.space_after = Cm(0)
            paragraph_format.left_indent = Cm(0.5)
            paragraph_format.tab_stops.add_tab_stop(Cm(2))
            # print([tab_stop.position.inches for tab_stop in paragraph_format.tab_stops])
            for part in glossing(gloss_line):
                print(gloss_line, part)
                if re.match(r'[a-z+]', part):
                    p_glosses.add_run(part).font.small_caps = True
                else:
                    p_glosses.add_run(part)

        p = document.add_paragraph()
        paragraph_format = p.paragraph_format
        paragraph_format.space_after = Cm(0.1)
        paragraph_format.left_indent = Cm(0.5)
        p.add_run(f'\'{translation}\'')
        p = document.add_paragraph()
        p.add_run(f'{key[0]} — {key[1]} {comment}')
        counter += 1

    for paragraph in document.paragraphs:
        f = paragraph.style.font
        f.name = 'Times New Roman'
        f.size = Pt(12)
    document.save(f'{name}.docx')


def mapping(transc, transl, gloss, comment):
    pivot_dic = {}
    for key, value in transc.items():
        trnsl = transl[key]
        gls = gloss[key]
        cmnt = comment[key]
        pivot_dic[key] = [value, trnsl, gls, cmnt]
    return pivot_dic


def glossing(text):
    # text = text.replace(' ', '\t')
    glossed_text = re.split(r'([a-z+])', text)
    return glossed_text


def main():
    file = input('введите название илановского файла или назовите его 1.txt и нажмите Enter')
    if file == '':
        file = '1.txt'
    transc, transl, gloss, comment = elan_data(file)
    mapped_dic = mapping(transc, transl, gloss, comment)
    to_word(mapped_dic)


if __name__ == '__main__':
    main()
