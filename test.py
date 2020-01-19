import re


def open_file(filename):
    text = open(filename, encoding='utf-8').read()
    return text


def elan(filename):
    elan = open_file(filename)
    elan = elan.splitlines()
    transc = []
    transl = []
    gloss = []
    for line in elan:
        tokens = line.split('\t')
        layer = tokens[0]
        time_start = tokens[2]
        time_finish = tokens[4]
        text = tokens[8]
        if layer == 'transcription':
            transc.append([text, time_start, time_finish])
        elif layer == 'translation':
            transl.append([text, time_start, time_finish])
        elif layer == 'gloss':
            gloss.append([text, time_start, time_finish])
    return transc, transl, gloss


def small_caps(text):
    new_text = re.sub('<.+?>', ' ', text)
    pattern = '[a-z-=]+'
    latins = re.findall(pattern, new_text)
    for latin in latins:
        text = text.replace(latin, '</w:t></w:r><w:r w:rsidRPr="00F6391B"><w:rPr><w:smallCaps/><w:lang w:val="en-US"/>'
                                   '</w:rPr><w:t>' + latin + '</w:t></w:r><w:r><w:t>')
    return text


def write_to_word(transc, transl, gloss):
    print(len(transl), len(transc))
    length = len(transc)
    informant = input('введите код информанта ')
    data = input('введите дату ')
    expe = input('введите свой код ')
    to_write = []
    for i in range(length):
        part = open_file('tag.txt')
        part = part.replace('informant', informant)
        part = part.replace('data', data)
        part = part.replace('expe', expe)
        part = part.replace('number', str(i + 1))
        if transc[i][1] == transl[i][1]:
            try:
                if gloss[i][1] == transl[i][1]:
                    part = part.replace('glossing',
                                        '</w:t></w:r><w:r><w:rPr><w:lang w:val="en-US"/></w:rPr><w:tab/><w:t>'.join(
                                            gloss[i][0].split()))
                    part = small_caps(part)
            except Exception:
                part = part.replace('glossing', '')
            part = part.replace('TEXT', '</w:t></w:r><w:r><w:rPr><w:lang w:val="en-US"/></w:rPr><w:tab/><w:t>'.join(
                transc[i][0].split()))
            part = part.replace('translation', '</w:t></w:r><w:r><w:t>' + transl[i][0])
        part = part.replace('optional', transl[i][1] + '—' + transl[i][2])
        to_write.append(part)
    docx = open_file('document1.xml')
    docx = docx.replace('PASTE_HERE', ''.join(to_write))
    with open('document.xml', 'w', encoding='utf-8') as f:
        f.write(docx)


def main():
    file = input('введите название илановского файла или назовите его 1.txt и нажмите Enter')
    if file == '':
        file = '1.txt'
    el = elan(file)
    transc = el[0]
    transl = el[1]
    gloss = el[2]
    write_to_word(transc, transl, gloss)


if __name__ == '__main__':
    main()
