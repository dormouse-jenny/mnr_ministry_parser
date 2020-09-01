# -*- coding: utf-8 -*-
from bs4 import BeautifulSoup
import requests
import re
from urllib.parse import urljoin
import utils

def scrap_site():
    url = ('http://www.mnr.gov.ru/open_ministry/anticorruption/'
        +'svedeniya_o_dokhodakh_raskhodakh_ob_imushchestve_i'
        +'_obyazatelstvakh_imushchestvennogo_kharaktera/')


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

        match = re.search(r'20\d\d',doc_caption)
        if match:
            doc_year = match[0]
        else:
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
