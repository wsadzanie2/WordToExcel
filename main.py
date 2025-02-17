from os.path import isfile
from docx.document import Document as Doc_type
from docx.api import Document
import xlsxwriter

PATH_SEP = "/"
input_file = ''
output_file = ''

TTY = False
if TTY:
    # HANDLE INPUT FILE
    while True:
        input_file = input("Podaj ścieżkę do pliku wejściowego (.docx): ")
        if not isfile(input_file):
            print("Ścieżka do pliku jest do kitu. Podaj ścieżkę prowadzącą DO PLIKU")
            continue
        if not input_file.endswith(".docx"):
            print("Program wspiera obecnie jedynie pliki docx. Jeśli odpalisz ten plik i coś się zepsuje to nie moja sprawa.")
            if input("Chcesz spróbować? (t/N)") in ['t', 'T', 'y', 'Y', 'tak', 'yes']:
                break
            continue
        break
    
    # HANDLE OUTPUT FILE
    while True:
        output_file = input("Podaj ścieżkę do pliku wyjściowego (.xlsx): ")
        if not output_file.endswith(".xlsx"):
            print("Program wspiera obecnie jedynie pliki xlsx. Jeśli odpalisz ten plik i coś się zepsuje to nie moja sprawa.")
            if input("Chcesz spróbować? (t/N)") in ['t', 'T', 'y', 'Y', 'tak', 'yes']:
                break
            continue
        if not isfile(input_file):
            with open(input_file, "a"):
                pass
        break

else:
    try:
        import easygui # type: ignore
    except Exception as e:
        print("Brakuje biblioteki easygui")
        print("na Windowsie trzeba wpisać pip install easygui, a na Ubuntu sudo apt-get install python3-easygui")
        raise e
    loop = True
    while loop:
        input_file = easygui.fileopenbox(filetypes="*.docx", default="*.docx", title="Wybierz plik (.docx)")
        if input_file is None:
            exit()
        if not (isfile(input_file) and input_file.endswith(".docx")): # type: ignore
            raise Exception("Plik nie jest plikiem .docx")

        output_file = easygui.filesavebox(filetypes="*.xlsx", default="*.xlsx", title="Wybierz plik Worda (.xlsx)")
        if output_file is None:
            exit()
        if not (isfile(output_file) and output_file.endswith(".xlsx")):
            raise Exception("Plik nie jest plikiem .xlsx")
        box = easygui.ccbox(msg=f"{input_file.split(PATH_SEP)[-1]} -> {output_file.split(PATH_SEP)[-1]}", choices=('Dobrze jest', 'Jeszcze raz')) # type: ignore
        if box:
            loop = False


print("Started working")
print(f"{input_file} -> {output_file}")

# doc: Doc_type = Document("documents/test2.docx")
doc: Doc_type = Document(input_file) # type: ignore
tables = doc.tables

# workbook = xlsxwriter.Workbook("output/output.xlsx")
workbook = xlsxwriter.Workbook(output_file)
worksheet = workbook.add_worksheet()

x = 0
y = 0

def getText(document):
    doc = document
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
        # fullText.append(para)
    # return fullText
    return '\n'.join(fullText)

# for paragraph in getText(doc):
#    for run in paragraph.runs:
#        print(run.text)

for thing in doc.iter_inner_content():
    try:
        thing.runs # type: ignore
        worksheet.write(y, x, thing.text) # type: ignore
        y += 1
        x = 0
    except AttributeError as e:
        # handle tables
        for row in thing.rows: # type: ignore
            for cell in row.cells:
                # print(cell.text, end='|')
                worksheet.write(y, x, cell.text)
                x += 1
            y += 1
            x = 0
            # print()


workbook.close()
