"""Всего лабораторных работ три, четвертая для сдачи долгов.
Одно задание на три лабораторных работы. На второй лабораторной работе будет добавлено дополнительное задание к создаваемому приложению.

Автоматическая генерация писем в формате DOCX с использованием WPF и Open XML, язык C#

Цель работы
Познакомиться с библиотекой DocumentFormat.OpenXml и принципами работы с шаблонами DOCX для автоматической подстановки пользовательских данных в текстовые документы средствами WPF-приложения.

Задание
1. Создайте WPF-приложение с формой ввода данных (например, должность адресата, ФИО, тема письма и т. д.).
2. Подготовьте шаблон документа в формате .docx, в котором в нужных местах вставьте уникальные плейсхолдеры (например, {RECIPIENT_POST}, {RECIPIENT_NAME}, {LETTER_BODY} и т. д.).
3. Используя библиотеку DocumentFormat.OpenXml, реализуйте в программе функционал, который копирует исходный шаблон в новый файл, находит все плейсхолдеры в теле документа (а при необходимости — в колонтитулах) и заменяет их введёнными пользователем значениями.
4. Проверьте корректность результата, открыв итоговый документ в Microsoft Word или другом совместимом редакторе.
*Вместо языка C# можно использовать любой другой, на котором умеете программировать. Использование WPF не является обязательным условием, но форма ввода должна быть
"""

import tkinter as tk
import tkinter.ttk as ttk
import tkinter.filedialog as tkfd
import tkinter.messagebox as tkms
import docx
import shutil
import os

placeholders = [['{recipient_post}'], ['{recipient_name}'], ['{letter_body}'], ['{date}'], ['{sender_name}']]
def main():
    def generate(file_mask, file_doc):
        try:
            if file_mask == '':
                raise Exception('Необходимо выбрать шаблон')
            if file_doc != '': adress = file_doc
            else: adress = 'Сгенерированное письмо.docx'
            shutil.copy2(file_mask, adress)
            doc = docx.Document(adress)
            dict_ph = {}
            for ph in placeholders:
                key = ph[0]
                value = ph[1].get()
                dict_ph[key] = value
            for paragraph in doc.paragraphs:
                text = ''.join(run.text for run in paragraph.runs)
                if any(placeholder in text for placeholder in dict_ph.keys()):
                    run = paragraph.runs[0]
                    fmt = {
                        'bold': run.bold,
                        'italic': run.italic,
                        'underline': run.underline,
                        'font_name': run.font.name,
                        'font_size': run.font.size,
                        'font_color': run.font.color.rgb
                    }
                    for placeholder, value in dict_ph.items():
                        text = text.replace(placeholder, value)
                    paragraph.clear()
                    new_run = paragraph.add_run(text)
                    if fmt['bold']: new_run.bold = fmt['bold']
                    if fmt['italic']: new_run.italic = fmt['italic']
                    if fmt['underline']: new_run.underline = fmt['underline']
                    if fmt['font_name']: new_run.font.name = fmt['font_name']
                    if fmt['font_size']: new_run.font.size = fmt['font_size']
                    if fmt['font_color']: new_run.font.color.rgb = fmt['font_color']
                doc.save(adress)
            tkms.showinfo(title='Успех', message=f'Файл успешно сгенерирован по адресу:\n{adress}')
        except Exception as ex:
            print(ex.__repr__())
            tkms.showinfo(title='Ошибка', message=ex.__str__())
    root = tk.Tk()
    root.title("Генерация писем в формате DOCX")

    mask = tk.StringVar(value='')
    doc =  tk.StringVar(value='')

    frame_select = ttk.Frame(root, padding="10")
    frame_select.grid(row=0, column=0)

    ttk.Label(frame_select, text="Шаблон документа:").grid(row=0, column=0)
    mask_entry = ttk.Entry(frame_select, textvariable=mask, width=50)
    mask_entry.grid(row=0, column=1)
    but_mask = ttk.Button(frame_select, text='Выбрать', command=lambda: mask.set(value = tkfd.askopenfilename()))
    but_mask.grid(row=0, column=2)

    ttk.Label(frame_select, text="Исходный документ:").grid(row=1, column=0)
    doc_entry = ttk.Entry(frame_select, textvariable=doc, width=50)
    doc_entry.grid(row=1, column=1)
    but_doc = ttk.Button(frame_select, text='Выбрать', command=lambda: doc.set(value = tkfd.asksaveasfilename()))
    but_doc.grid(row=1, column=2)


    frame_write = ttk.Frame(root, padding="10")
    frame_write.grid(row=1, column=0)

    for row_i, text_label in enumerate(["Должность адресата (в род.пад.):",
                                "ФИО адресата (в род.пад.):", 
                                "Тело письма:", 
                                "Дата:",
                                "ФИО адресанта:"]):
        ttk.Label(frame_write, text=text_label).grid(row=row_i, column=0)
        entry = ttk.Entry(frame_write, width=40)
        entry.grid(row=row_i, column=1)
        placeholders[row_i].append(entry)

    placeholders[0][1].insert(0, 'Преподавателю')
    placeholders[1][1].insert(0 ,'Хандусенко В.О.')
    placeholders[2][1].insert(0, 'Прошу проверить мою лабораторную работу по курсу "ФиЛП" на тему "Генерация писем в формате DOCX с использованием WPF и Open XML". Лабораторная работа выполнена в полном объёме, все пункты задания соблюдены.')
    placeholders[3][1].insert(0, '18.05.2026')
    placeholders[4][1].insert(0, 'Гераськин А.А.')
    

    generate_button = ttk.Button(frame_write, text="Сгенерировать письмо", command=lambda: generate(mask.get(), doc.get()))
    generate_button.grid(row=5, column=0, columnspan=2, pady=[15, 5])

    root.mainloop()

main()