# -*- coding: utf-8 -*-
import os
import settings
import utils
import scraper
import doc_converter
import doc_parser

if not os.path.exists(settings.DOC_DIR):
    os.mkdir(settings.DOC_DIR)

scraper.scrap_site()

if not os.path.exists(settings.DOCX_DIR):
    os.mkdir(settings.DOCX_DIR)

doc_converter.main()

if not os.path.exists(settings.CSV_DIR):
    os.mkdir(settings.CSV_DIR)

doc_parser.main()
