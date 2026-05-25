'''
Функциональное и логическое программирование, лабораторное занятие:

По поводу второй лабораторной:
1. Необходимо расширить функциональность существующего приложения, реализовав автоматическую подстановку данных о приложениях (их заголовки и тексты) в письмо.
2. Добавьте возможность указывать одно/несколько приложений к письму. Каждое приложение содержит:
* Слово "Приложение" справа, если приложение одно. Если приложений несколько, то "Приложение 1", "Приложение 2" и т. д.
* Заголовок приложения по центру.
* Текст приложения (произвольный контент без рисунков и картинок).
3. В теле письма до подписи или после основной текстовой части вставить перечень приложений с названиями. Также желательно, но не обязательно в тексте письма указать количество листов (страниц) каждого приложения.
    Например. 
    Приложение: 
1. Причина отмены лабораторной на 3 л. 
2. Почему всё плохо на 2 л
'''
import tkinter as tk
import tkinter.ttk as ttk
import tkinter.filedialog as tkfd
import tkinter.messagebox as tkms
import docx
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import shutil
import os

placeholders = [['{recipient_post}'], ['{recipient_name}'], ['{letter_body}'], ['{date}'], ['{sender_name}']]
adds = []
def main():
    def create_add(i):
        frame_add = ttk.Frame(notebook, padding="10")
        frame_add.columnconfigure(1, weight=1)
        frame_add.grid(row=0, column=0, sticky='nsew')
        notebook.add(frame_add, text=f'Приложение {i}')
        adds.append([])
        for row_i, text_label in enumerate(["Название приложения: ",
                                            "Кол-во страниц:",
                                            "Текст приложения:"]):
            ttk.Label(frame_add, text=text_label).grid(row=row_i, column=0, pady=[0, 3])
            entry = ttk.Entry(frame_add)
            entry.grid(row=row_i, column=1, sticky='nswe', pady=[0, 3])
            adds[i-1].append(entry)
    def generate(file_mask, file_doc):
        try:
            if file_mask == '':
                raise Exception('Необходимо выбрать шаблон')
            if file_doc != '': adress = file_doc
            else: adress = os.path.join(os.getcwd(), 'Сгенерированное письмо.docx')
            shutil.copy2(file_mask, adress)
            doc = docx.Document(adress)
            dict_ph = {}
            fmt_text = None
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
                        if placeholder == '{letter_body}': fmt_text = {'font_name': run.font.name, 'font_size': run.font.size}
                    paragraph.clear()
                    new_run = paragraph.add_run(text)
                    if fmt['bold']: new_run.bold = fmt['bold']
                    if fmt['italic']: new_run.italic = fmt['italic']
                    if fmt['underline']: new_run.underline = fmt['underline']
                    if fmt['font_name']: new_run.font.name = fmt['font_name']
                    if fmt['font_size']: new_run.font.size = fmt['font_size']
                    if fmt['font_color']: new_run.font.color.rgb = fmt['font_color']

                    if dict_ph['{letter_body}'] in text and len(adds):
                        run_add = paragraph.add_run('\n\tПриложения: \n')
                        run_add.font.size = fmt_text['font_size']
                        run_add.font.name = fmt_text['font_name']
                        for arr_add in adds:
                            run_add = paragraph.add_run(f' - \"{arr_add[0].get()}\" на {arr_add[1].get()} л.\n')
                            run_add.font.size = fmt_text['font_size']
                            run_add.font.name = fmt_text['font_name']
                doc.save(adress)
            for i, add in enumerate(adds): 
                paragraph_title1 = doc.add_paragraph(text='\n\n')
                paragraph_title1.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
                run = paragraph_title1.add_run(f'Приложение {i+1}' if len(adds) > 1 else 'Приложение')
                run.font.size = fmt_text['font_size']
                run.font.name = fmt_text['font_name']

                paragraph_title2 = doc.add_paragraph('\n')
                paragraph_title2.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                run = paragraph_title2.add_run(add[0].get())
                run.font.size = fmt_text['font_size']
                run.font.name = fmt_text['font_name']

                paragraph_title2 = doc.add_paragraph(text='\n')
                paragraph_title2.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
                run = paragraph_title2.add_run(f'\t{add[2].get()}')
                run.font.size = fmt_text['font_size']
                run.font.name = fmt_text['font_name']
            doc.save(adress)
            tkms.showinfo(title='Успех', message=f'Файл успешно сгенерирован по адресу:\n{adress}')
        except Exception as ex:
            print(ex.__repr__())
            tkms.showinfo(title='Ошибка', message=ex.__str__())
    root = tk.Tk()
    root.title("Генерация писем в формате DOCX")

    notebook = ttk.Notebook()
    notebook.grid(row=0, column=0, sticky='nsew')

    frame_root = ttk.Frame(notebook)
    frame_root.grid(row=0, column=0, sticky='nsew')
    notebook.add(frame_root, text='Письмо')

    mask = tk.StringVar()
    doc =  tk.StringVar()
    if os.path.exists(os.path.join(os.getcwd(), 'Шаблон ФиЛП.docx')):
        mask.set(os.path.join(os.getcwd(), 'Шаблон ФиЛП.docx'))

    frame_select = ttk.Frame(frame_root, padding="10")
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


    frame_write = ttk.Frame(frame_root, padding="10")
    frame_write.grid(row=1, column=0, sticky='nsew')
    frame_write.columnconfigure(1, weight=1)

    for row_i, text_label in enumerate(["Должность адресата (в род.пад.):",
                                "ФИО адресата (в род.пад.):", 
                                "Тело письма:", 
                                "Дата:",
                                "ФИО адресанта:"]):
        ttk.Label(frame_write, text=text_label).grid(row=row_i, column=0)
        entry = ttk.Entry(frame_write)
        entry.grid(row=row_i, column=1, sticky='nsew')
        placeholders[row_i].append(entry)

    placeholders[0][1].insert(0, 'Преподавателю')
    placeholders[1][1].insert(0 ,'Хандусенко В.О.')
    placeholders[2][1].insert(0, 'Прошу проверить мою 2 лабораторную работу по курсу "ФиЛП", темой которой было расширение функционала программы. Лабораторная работа выполнена в полном объёме, все пункты задания соблюдены.')
    placeholders[3][1].insert(0, '25.05.2026')
    placeholders[4][1].insert(0, 'Гераськин А.А.')
    
    ttk.Button(frame_write, text="Создать приложение", command=lambda: create_add(len(notebook.tabs()))).grid(row=5, column=0, columnspan=2, pady=[15, 5])

    generate_button = ttk.Button(frame_write, text="Сгенерировать письмо", command=lambda: generate(mask.get(), doc.get()))
    generate_button.grid(row=6, column=0, columnspan=2, pady=[5, 5])

    root.mainloop()

main()