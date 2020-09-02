# -*- coding: utf-8 -*-
from bs4 import BeautifulSoup
import requests
import re
from urllib.parse import urljoin
import utils, settings

def scrap_site():

    url = settings.MAIN_URL
    page = requests.get(url)

    if page.status_code != 200:
        #сообщить об ошибке
        pass

    print(page.status_code)
    soup = BeautifulSoup(page.text, "html.parser")
    doc_links = soup.find_all('a', {'class' : 'docs-full-item__link'})

    for doc_link in doc_links:
        #print(doc_link.get('href'))
        doc_caption = ''.join(doc_link.stripped_strings)

        #Подведомственные учреждения пропускаем
        if doc_caption.find('подведомственных')!=-1 or \
            doc_caption.find('руководителями федеральных')!=-1:
            continue

        doc_year = utils.year_from_filename(doc_caption)
        if not doc_year:
            print(f'Не удалось выделить год из заголовка : "{doc_caption}"')
            continue

        print(f'Скачивание файла за {doc_year} год')
        doc_url = urljoin(url, doc_link.get('href'))
        web_file = requests.get(doc_url)

        if web_file.status_code != 200:
            #сообщить об ошибке
            print("Скачивание не удалось")
            continue

        doc_file = open(utils.get_doc_filename(doc_year),"wb")
        doc_file.write(web_file.content)
        doc_file.close()
