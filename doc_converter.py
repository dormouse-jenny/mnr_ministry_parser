# -*- coding: utf-8 -*-
import os
import settings, utils


def doc_to_docx_linux(file_name):
    from subprocess import run, PIPE, STDOUT

    try:
        from subprocess import DEVNULL # py3k
    except ImportError:
        DEVNULL = open(os.devnull, 'wb')

    print(f'Конвертация в docx файла {file_name}')
    #soffice --headless --convert-to docx --outdir doc_doc doc_doc/mnr_2009.doc

    run([
        'soffice',
        '--headless',
        '--convert-to',
        'docx',
        '--outdir',
        settings.DOCX_DIR,
        file_name],
        stdin=PIPE,
        stdout=DEVNULL,
        stderr=STDOUT)

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
