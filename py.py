import os
import docx
from docx import Document

paths = []
folder = os.getcwd()
for root, dirs, files in os.walk(folder):
    for file in files:
        if file.endswith('docx') and not file.startswith('~'):
            paths.append(os.path.join(root, file))


counter = 0
print(f'Введите слово')
text = input()
text = text.lower()
print(f'Ищу - {text}')
mydoc = docx.Document()
for path in paths:  
    doc = docx.Document(path)
    properties = doc.core_properties
    # print('Автор документа:', properties.author)
    # print('Автор последней правки:', properties.last_modified_by)
    # print('Дата создания документа:', properties.created)
    # print('Дата последней правки:', properties.modified)
    # print('Дата последней печати:', properties.last_printed)
    # print('Количество сохранений:', properties.revision)
    for paragraph in doc.paragraphs:
        # index = paragraph.text.find("Сочетанная")
        # if index > -1:      
        if text in paragraph.text.lower():
            counter += 1    
            # filename = os.path.basename(path)
            # file_name = path.split(".")[0]
            file_name_with_extension = path.split("/")[-1]
            file_name = file_name_with_extension.split(".")[0]
            url = f'{file_name}_диклофенак.docx'  
            mydoc.add_paragraph(file_name)          
            # doc.save(url)  
            print(f'{file_name}_диклофенак')
            print(counter)
            break
mydoc.add_paragraph(f'Найдено совпадений - {counter}')       
mydoc.save("result.docx")


# from glob import glob
# import re
# import os
# import win32com.client as win32
# from win32com.client import constants

# # Create list of paths to .doc files
# paths = glob('C:\\Users\\Сергей\\Downloads\\pyth\\2022 — копия\\*.doc', recursive=True)

# def save_as_docx(path):
#     # Opening MS Word
#     word = win32.gencache.EnsureDispatch('Word.Application')
#     doc = word.Documents.Open(path)
#     doc.Activate ()

#     # Rename path with .docx
#     new_file_abs = os.path.abspath(path)
#     new_file_abs = re.sub(r'\.\w+$', '.docx', new_file_abs)

#     # Save and Close
#     word.ActiveDocument.SaveAs(
#         new_file_abs, FileFormat=constants.wdFormatXMLDocument
#     )
#     doc.Close(False)

# for path in paths:
#     save_as_docx(path)