import docx


def take_text(path):
    qwe = []
    for p in path.paragraphs:
        if '\n' in p.text:
            for j in p.text.split('\n'):
                qwe.append(j)
        else:
            qwe.append(p.text)
        p.clear()
    return qwe


doc_path = docx.Document('4.docx')

text_array = take_text(doc_path)
print(text_array)

alphabet = 'абвгдеёжзийклмнопрстуфхцчшщъыьэюя'


# четность слов в строке

yes_array1 = []
no_array1 = []
words_in_line = []

for i in text_array:
    words_counter = 0
    for j in i.split(' '):
        try:
            if j[0].lower() in alphabet and j[0].lower() != '':
                words_counter += 1
                continue
        except IndexError:
            continue
    if words_counter % 2 == 0:
        yes_array1.append([i, words_counter])
    elif words_counter % 2 == 1:
        no_array1.append([i, words_counter])
    words_in_line.append(words_counter)
print(words_in_line)
print(yes_array1)
print(len(yes_array1))
print(no_array1)
print(len(no_array1))
print('')
# наличия пробела в конце строки

yes_array2 = []
no_array2 = []
space_in_line = []

for i in yes_array1:
    space_flag = 0
    if i[0][::-1][0] == ' ':
        space_flag = 1
        yes_array2.append([i, '+'])
    else:
        no_array1.append([i, '-'])
print(yes_array2)
print(len(yes_array2))
print(no_array1)
print(len(no_array1))