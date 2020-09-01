import os
import re

DOCX_DIR = "docx_docs"
DOC_DIR = "doc_docs"
CSV_DIR = "csv_docs"
FILENAME_PATTERN = "mnr"
CSV_DELIMETR = ";"

def get_doc_filename(year_str):

    if not os.path.exists(DOC_DIR):
        os.mkdir(DOC_DIR)

    return DOC_DIR+"/"+FILENAME_PATTERN+"_"+year_str+".doc"

def get_docx_filename(year_str):

    if not os.path.exists(DOCX_DIR):
        os.mkdir(DOCX_DIR)

    return DOCX_DIR+"/"+FILENAME_PATTERN+"_"+year_str+".docx"

def get_csv_filename(year_str):

    if not os.path.exists(CSV_DIR):
        os.mkdir(CSV_DIR)

    return CSV_DIR+"/"+FILENAME_PATTERN+"_"+year_str+".csv"

def year_from_filename(file_name):
    year_list = re.findall(r'20\d\d',file_name)
    if year_list:
        return year_list[0]
    else:
        return ""

def get_all_docx():
    files = os.listdir(DOCX_DIR)
    for file in files:
        if file.endswith(".docx"):
            yield DOCX_DIR+"/"+file

def get_all_doc():
    files = os.listdir(DOC_DIR)
    for file in files:
        if file.endswith(".doc"):
            yield DOC_DIR+"/"+file
