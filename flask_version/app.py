from flask import Flask, send_file
from flask import render_template, request, redirect, url_for
from docx import Document
from docx.shared import Pt, Cm
from collections import defaultdict
import re
import os


FILE = ['']

def elan_data(file):
    # elan = elan.replace('&', '')
    # elan = elan.replace('<', '&lt;')
    # elan = elan.replace('>', '&gt;')
    elan = file.splitlines()
    transc = defaultdict(str)
    transl = defaultdict(str)
    gloss = defaultdict(str)
    comment = defaultdict(str)
    for line in elan:
        tokens = line.split('\t')
        if len(line) == 9:
            indices = (0, 2, 4, 8)
        else:
            indices = (0, 2, 3, 4)
        layer = tokens[indices[0]]
        time_start = tokens[indices[1]]
        time_finish = tokens[indices[2]]
        text = tokens[indices[3]]
        if layer == 'transcription':
            transc[(time_start, time_finish)] = text
        elif layer == 'translation':
            transl[(time_start, time_finish)] = text
        elif layer == 'gloss':
            gloss[(time_start, time_finish)] = text
        elif layer == 'comment':
            comment[(time_start, time_finish)] = text
    return transc, transl, gloss, comment


def to_word(pivot_dictionary, informant, date, expe, others, theme):
    name = f'eve_{informant}_{date}_{expe}.docx'

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
    els[0].text, els[1].text = 'Кто ещё был на паре', others

    inf = table.rows[4].cells
    inf[0].text, inf[1].text = 'Примерная тематика', theme

    for row in table.rows:
        for cell in row.cells:
            cell.paragraphs[0].paragraph_format.space_after = Cm(0)
    document.add_paragraph().paragraph_format.space_after = Cm(0)

    counter = 1

    for key, value in pivot_dictionary.items():
        header = f'{counter}. {informant}_{date}@{expe}_{counter}'
        transcription = value[0]
        translation = value[1]
        glosses = value[2]
        comment = value[3]
        p = document.add_paragraph()
        paragraph_format = p.paragraph_format
        paragraph_format.space_after = Cm(0.1)
        p.add_run(header)
        p = document.add_paragraph()
        paragraph_format = p.paragraph_format
        paragraph_format.space_after = Cm(0)
        paragraph_format.left_indent = Cm(0.5)
        p.add_run(transcription.replace(' ', '\t')).font.italic = True
        p = document.add_paragraph()
        paragraph_format = p.paragraph_format
        paragraph_format.space_after = Cm(0)
        paragraph_format.left_indent = Cm(0.5)
        for part in glossing(glosses):
            if re.match(r'[a-z+]', part):
                p.add_run(part).font.small_caps = True
            else:
                p.add_run(part)
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

    filename = f'/home/vantral/mysite/{name}'
    document.save(filename)

    return name


def mapping(transc, transl, gloss, comment):
    pivot_dic = {}
    for key, value in transc.items():
        trnsl = transl[key]
        gls = gloss[key]
        cmnt = comment[key]
        pivot_dic[key] = [value, trnsl, gls, cmnt]
    return pivot_dic


def glossing(text):
    text = text.replace(' ', '\t')
    glossed_text = re.split(r'([a-z+])', text)
    return glossed_text


def main(file, informant, date, expe, others, theme):
    transc, transl, gloss, comment = elan_data(file)
    mapped_dic = mapping(transc, transl, gloss, comment)
    return to_word(mapped_dic, informant, date, expe, others, theme)


app = Flask(__name__)


@app.route('/')
def index():
    # files = os.listdir('./mysite')
    # for file in files:
    #     if file.endswith('.docx'):
    #         os.remove(f'/home/vantral/mysite/{file}')
    return render_template(
        'index.html'
    )

if __name__ == '__main__':
    app.run(debug=True)


@app.route('/results', methods = ['POST'])
def upload_route_summary():
    if request.method == 'POST':

        f = request.files['fileupload']

        fstring = f.read().decode('utf-8')
        FILE[0] = fstring

    return render_template('data.html')


@app.route('/itog', methods = ['GET'])
def create_file():
    informant = request.args.get('informant')
    date = request.args.get('date')
    expe = request.args.get('expe')
    others = request.args.get('others')
    theme = request.args.get('theme')
    name = main(FILE[0], informant, date, expe, others, theme)

    response = send_file(name, attachment_filename=name, as_attachment=True)
    response.headers["x-filename"] = name
    response.headers["Access-Control-Expose-Headers"] = 'x-filename'
    return response
