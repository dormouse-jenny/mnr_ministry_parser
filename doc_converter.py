import utils
import os

def doc_to_docx_linux(file_name):
    import subprocess

    print(f'Конвертация в docx файла {file_name}')
    #soffice --headless --convert-to docx --outdir doc_doc doc_doc/mnr_2009.doc

    subprocess.call([
        'soffice',
        '--headless',
        '--convert-to',
        'docx',
        '--outdir',
        utils.DOCX_DIR,
        file_name])

def doc_to_docx_win(file_name):
    #Не проверялось
    import win32com.client as win32
    from win32com.client import constants

    year_str = utils.year_from_filename(file_name)
    if not year_str:
        return

     # Opening MS Word

    word = win32.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(file_name)
    doc.Activate ()


     # Save and Close
    word.ActiveDocument.SaveAs(
        utils.get_docx_filename(year_str),
        FileFormat=constants.wdFormatXMLDocument
    )
    doc.Close(False)

def main():
    for doc_file in utils.get_all_doc():
        if os.name == "posix":
            doc_to_docx_linux(doc_file)
        elif os.name == "nt":
            doc_to_docx_win(doc_file)


if __name__=="__main__":
    main()
