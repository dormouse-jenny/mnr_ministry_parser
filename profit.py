import os

import utils
import scraper
import doc_converter
import doc_parser

if not os.path.exists(utils.DOC_DIR):
    os.mkdir(utils.DOC_DIR)

#scraper.scrap_site()

if not os.path.exists(utils.DOCX_DIR):
    os.mkdir(utils.DOCX_DIR)

#doc_converter.main()

if not os.path.exists(utils.CSV_DIR):
    os.mkdir(utils.CSV_DIR)

doc_parser.main()
