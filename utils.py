# -*- coding: utf-8 -*-
import os
import re
import settings

def get_doc_filename(year_str):

    if not os.path.exists(settings.DOC_DIR):
        os.mkdir(settings.DOC_DIR)

    return settings.DOC_DIR+"/"+settings.FILENAME_PATTERN+"_"+year_str+".doc"

def get_docx_filename(year_str):

    if not os.path.exists(settings.DOCX_DIR):
        os.mkdir(settings.DOCX_DIR)

    return settings.DOCX_DIR+"/"+settings.FILENAME_PATTERN+"_"+year_str+".docx"

def get_csv_filename(year_str):

    if not os.path.exists(settings.CSV_DIR):
        os.mkdir(settings.CSV_DIR)

    return settings.CSV_DIR+"/"+settings.FILENAME_PATTERN+"_"+year_str+".csv"

def year_from_filename(file_name):
    year_list = re.findall(r'20\d\d',file_name)
    if year_list:
        return year_list[0]
    else:
        return ""

def get_all_docx():
    files = os.listdir(settings.DOCX_DIR)
    for file in files:
        if file.endswith(".docx"):
            yield settings.DOCX_DIR+"/"+file

def get_all_doc():
    files = os.listdir(settings.DOC_DIR)
    for file in files:
        if file.endswith(".doc"):
            yield settings.DOC_DIR+"/"+file
