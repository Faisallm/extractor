import docx
import openpyxl
import os

folder_path = os.getcwd()
print(f"folder path: {folder_path}" )

file_list = os.listdir(folder_path)
del file_list[0]
print(f"folder path: {file_list}" )





for file_index, word in enumerate(file_list):
    doc = docx.Document(word)

    # open excel workbook
    workbook = openpyxl.Workbook()

    sheet = workbook.active
    

    sheet = workbook.create_sheet(str(file_index))
    sheet.title = word

    alphabets = ["A", "B", "C", "D"]
    for row_index, paragraph in enumerate(doc.paragraphs):
        
        for index, item in enumerate(str(paragraph.text).split(',')):
            # I will now push it to the excel file
            sheet[f'{alphabets[index]}{row_index+1}'] = item

    print(f"Completed document {file_index+1}, file_name: {word}")
    workbook.save(f"{word}.xlsx")




