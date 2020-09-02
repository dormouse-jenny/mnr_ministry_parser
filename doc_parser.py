 # -*- coding: utf-8 -*-
import sys
import zipfile
import re
import locale
import itertools
import csv
from bs4 import BeautifulSoup
from decimal import Decimal

import settings, utils


class Item():
    def __init__(self,type,own_type = "Собственность"):
        self.type = type.strip()
        self.own_type = own_type.strip()
        self.proportion = None
        self.meters = None#Оставим строкой
        self.country = None
        self.car_type = None
        self.car_model = None

    @classmethod
    def new_flat(cls,type):
        #недвижимость по умолчанию создается в пользование
        #если есть вид собственности он добавится позднее
        return cls(type,"Пользование")

    @classmethod
    def new_car(cls,car_type="автомобиль",car_model=""):
        new_car = cls("транспорт")
        new_car.car_type = car_type
        new_car.car_model = car_model
        return new_car

    def add_owntype(self,own_type_str):
        #в own_type могут быть пропорции собственности
        #могут встречаться дроби вида ½ ¼ ¾
        own_type_str = re.sub("½","1/2",own_type_str)
        own_type_str = re.sub("¼","1/4",own_type_str)
        own_type_str = re.sub("¾","3/4",own_type_str)
        proportion = re.findall(r'\d+/\d+',own_type_str)
        if proportion:
            own_type_str = re.sub(proportion[-1],"",own_type_str)
            self.proportion = proportion[-1]

        self.own_type = own_type_str.strip()


    def add_meters(self, meters_str):
        #meters_str = 12 или 12,34 или 12.34
        meters_str = re.sub(",",".",meters_str)
        meters_dec = Decimal(meters_str)
        self.meters = locale.format("%.1f", meters_dec)

    def add_country(self, country_str):
        self.country = country_str.strip()

    def get_fields():
        fields = []
        fields.append("type")
        fields.append("own_type")
        fields.append("proportion")
        fields.append("meters")
        fields.append("country")
        fields.append("car_type")
        fields.append("car_model")
        return fields

    def get_entry(self):
        entry = []
        entry.append(self.type)
        entry.append(self.own_type)
        entry.append(self.proportion)
        entry.append(self.meters)
        entry.append(self.country)
        entry.append(self.car_type)
        entry.append(self.car_model)
        return entry

    def get_empty():
        entry = []
        entry.append(None)#type
        entry.append(None)#own_type
        entry.append(None)#proportion
        entry.append(None)#meters
        entry.append(None)#country
        entry.append(None)#car_type
        entry.append(None)#car_model
        return entry


class Person:
    clerk_iter = itertools.count(start = 1)
    #у человека может быть несколько детей. Чтоб различать их записи:
    #Первый будет Несовершеннолетний ребенок
    #Второй Несовершеннолетний ребенок 2
    child_iter = itertools.count(start = 2)

    def __init__(self,its_clerk):
        if its_clerk:
            self.id = str(next(Person.clerk_iter))
        else:
            self.id = ""

        self.its_clerk = its_clerk
        self.name = ""
        self.position = ""
        self.family = ""
        self.ownflats = []
        self.rentflats = []
        self.cars = []
        self.money = ""
        self.source = ""
        Person._previous_relative = ""

    @classmethod
    def newClerk(cls):
        Person.child_iter = itertools.count(start = 2)
        return cls(its_clerk = True)

    @classmethod
    def newRelative(cls,other_person):
        relative = cls(False)
        relative.name = other_person.name
        relative.position = other_person.position
        relative.id = other_person.id
        Person._previous_relative = re.sub(r" \d$","",other_person.family)

        return relative

    def add_description(self,description):
        #У чиновника в графе описание фио, у родственника родственное отношение
        if self.its_clerk:
            self.name = description
        else:
            if description == Person._previous_relative:
                self.family = description+" "+str(next(Person.child_iter))
            else:
                self.family = description

    def add_position(self,position):
        if self.its_clerk:
            if not position:
                print("ОШИБКА ИСХОДНИКА : "\
                    +f"У {self.name} не указана должность")
            else:
                self.position = position
        else:
            if position:
                print("ОШИБКА ИСХОДНИКА : "\
                    +f"У {self.family} {self.name} указана должность")

    def _add_items(self, kind_of_item, items):
        for item in items:
            if kind_of_item == "own_flat":
                self.ownflats.append(Item.new_flat(item))
            elif kind_of_item == "rent_flat":
                self.rentflats.append(Item.new_flat(item))

    def add_ownflat(self,items):
        self._add_items("own_flat", items)

    def add_rentflat(self,items):
        self._add_items("rent_flat", items)

    def add_car(self,items):
        #С автомобилями все непросто
        #Возможен вариант А/М <Enter> ЛэндРовер DEFENDER (два параграфа 1 машина)
        #А возможен Мотовездеход РМ 500 (1 параграф 1 машина)
        #А можно и вот так: 2 А/М <Enter> Хонда Цивик <Enter> Хундай Крета
        #И вот так: А/М <Enter> ГАЗ 21 <Enter> ВАЗ 21063
        #и как это разбирать?
        car_type_patern = re.compile(r'автомобиль|мотоцикл|мотороллер')
        car_type = ""
        for item in items:
            item = item.strip()
            item = re.sub("А/М|А / М|А /М|А/ М","Автомобиль",item)
            if car_type_patern.findall(item.lower()) \
                and not car_type_patern.match(item.lower()):
                print('ОШИБКА ИСХОДНИКА : '\
                    +f'У {self.family} {self.name} невозможно разобрать '\
                    +f'список транспортных средств. Приведите "{item}" '\
                    +'в читаемый вид (вид марка) вручную.')
                self.cars.clear()
                return

            car_type_match = car_type_patern.fullmatch(item.lower())

            if car_type_match and car_type:
                #предыдущая машина была машиной без названия
                #добавляем ее
                self.cars.append(Item.new_car(car_type))
                #и следующая запись тоже будет машиной
                car_type = car_type_match.group(0).lower()
            elif car_type:
                #предыдущая запись была А/М а это ее марка
                self.cars.append(Item.new_car(car_type,item))
                car_type = ""
            elif car_type_match:
                #следующая запись должна бы быть названием машины
                #а предыдущая им не была и была добавлена
                car_type = car_type_match.group(0).lower()
            else:
                #Это запись в одну строку
                line_break = re.findall(r'\w+',item)
                if line_break:
                    first_part = line_break[0]
                    second_part = item.replace(first_part,"",1).strip()
                    #если строка сосотоит из цифр возможно это ошибка разбора
                    #или первая часть начи6нается с цифры
                    if re.match(r'\d',first_part) or \
                        re.fullmatch(r'\d*',second_part):
                        print('ОШИБКА ИСХОДНИКА : '\
                            +f'У {self.family} {self.name} невозможно разобрать '\
                            +f'список транспортных средств. Приведите "{item}" '\
                            +'в читаемый вид (вид марка) вручную.')
                        self.cars.clear()
                        return
                    else:
                        self.cars.append(Item.new_car(first_part,second_part))
                else:
                    #скорее всего не добавлен вид транспортного средства в шаблон
                    print('ОШИБКА ИСХОДНИКА : '\
                        +f'У {self.family} {self.name} невозможно разобрать '\
                        +f'список транспортных средств. Приведите "{item}" '\
                        +'в читаемый вид (вид марка) вручную.')
                    self.cars.clear()
                    return

    def _add_propertyes(self,list_of_items, list_of_props,prop_name):

        if len(list_of_items) != len(list_of_props):
            print('ОШИБКА ИСХОДНИКА : '\
                +f'У {self.family} {self.name} не совпадает количество '\
                +f'для свойства "{prop_name.upper()}"'\
                +'. Вставьте ноль или пробел в нужное место вручную.')

        for num in range(min(len(list_of_items),len(list_of_props))):
            if prop_name == "вид собственности":
                list_of_items[num].add_owntype(list_of_props[num])

            elif prop_name == "площадь":
                meters_str = list_of_props[num].strip()

                if re.fullmatch(r'(\d+)|(\d+,\d+)|(\d+.\d+)',meters_str):
                    list_of_items[num].add_meters(meters_str)
                else:
                    print('ОШИБКА ИСХОДНИКА : '\
                        +f'У {self.family} {self.name} некорректно записаны '\
                        +f'метры площади "{meters_str}"')

            elif prop_name == "страна расположения":
                list_of_items[num].add_country(list_of_props[num])

    def add_ownflat_owntype(self, items):
        self._add_propertyes(self.ownflats, items, "вид собственности")

    def add_ownflat_meters(self, items):
        self._add_propertyes(self.ownflats, items, "площадь")

    def add_ownflat_country(self, items):
        self._add_propertyes(self.ownflats, items, "страна расположения")

    def add_rentflat_meters(self, items):
        self._add_propertyes(self.rentflats, items, "площадь")

    def add_rentflat_country(self, items):
        self._add_propertyes(self.rentflats, items, "страна расположения")

    def add_money(self,money_str):
        #Строка вида 1234567,89 (с запятой)
        #А иногда вида 1234567.89 (с точкой)
        money_str = re.sub(" ","",money_str)
        if not money_str:
            #бедняжечка
            pass
        elif re.fullmatch(r'(\d+)|(\d+,\d{1,2})|(\d+.\d{1,2})',money_str):
            money_str = re.sub(",",".",money_str)
            #print("money_str : "+money_str)
            money_dec = Decimal(money_str)
            self.money = locale.currency(money_dec,symbol="",grouping = True)
        else:
            print('ОШИБКА ИСХОДНИКА : '\
                +f'У {self.family} {self.name} некорректно записан '\
                +f'декларированный доход "{money_str}"')

    def add_source(self,source_str):
        if re.findall(r'\w+',source_str):
            print(f'У {self.family} {self.name} признался откуда взял деньги: '\
                +f'"{source_str}" Необходимо добавить разбор источников финансирования')

    def get_fields():
        fields = []
        fields.append("id")
        fields.append("name")
        fields.append("position")
        fields.append("family")
        fields.append("money")
        fields.append("source")
        fields+=Item.get_fields()
        return fields

    def get_entryes(self):
        entry =  []
        entry.append(self.id)
        entry.append(self.name)
        entry.append(self.position)
        entry.append(self.family)
        entry.append(self.money)
        entry.append(self.source)
        if not self.ownflats and not self.rentflats and not self.cars:
            # если чиновник без ничего
            yield entry+Item.get_empty()

        for item in (self.ownflats+self.rentflats+self.cars):
            yield entry+item.get_entry()


def docx_to_csv(file_name):

    year_str = utils.year_from_filename(file_name)
    if not year_str:
        print("Для перевода в csv передан не тот файл. Имя файла не содержит год.")
        return

    print(f'Конвертация в csv файла {file_name}')

    document = zipfile.ZipFile(file_name)
    xml_content = document.read('word/document.xml')
    document.close()


    soup = BeautifulSoup(xml_content,"html.parser")

    table = soup.find('w:tbl')

    #В таблице возможны два варианта структуры. с доходом в 3 колонке и в 12
    #Разберемся который вариант
    tab_row = table.find_all('w:tr')[0]
    tab_cells = tab_row.find_all('w:tc')
    cell_content = ' '.join(tab_cells[-2].stripped_strings)
    #print("len(tab_cell) : "+str(len(tab_cells))+"cell_content : "+cell_content)
    if len(tab_cells) == 8 and re.search(r'\sгодовой\sдоход\s',cell_content):
        new_variant = True
    else:
        new_variant = False
        print('Для старого формата разбор файла пока не написан.')
        return

    csv_file = open(utils.get_csv_filename(year_str),mode = 'w',newline='')
    csv_writer = csv.writer(csv_file,delimiter= settings.CSV_DELIMETR)
    csv_writer.writerow(Person.get_fields())

    #Структура разбираемой таблицы
    #Первые 2 строки - шапка

    for tab_row in table.find_all('w:tr')[2:]:

        for cell_num, tab_cell in enumerate(tab_row.find_all('w:tc'), start=1):
            cell_content = []

            #В ячейке несколько параграфов часть с информацией
            #Часть для визуального форматирования
            for tab_par in tab_cell.find_all('w:p'):
                text_content = ""

                for tab_text in tab_par.find_all('w:t'):
                    text_content = " ".join([text_content,]+list(tab_text.stripped_strings))

                if re.fullmatch(r'\s*-\s*',text_content):
                    #в некоторых файлах так указывается <ничего нет>
                    pass
                elif text_content:
                    cell_content.append(text_content)

            if new_variant:

                #01 номер ПП (только для самих чиновников)
                if cell_num == 1:
                    #Есть номер это новый чиновник
                    if cell_content:
                        rich_man = Person.newClerk()
                    else:
                        rich_man = Person.newRelative(rich_man)
                #02 ФИО или название родственника
                elif cell_num == 2:
                    rich_man.add_description("".join(cell_content))
                #03 - должность
                elif cell_num == 3:
                    rich_man.add_position("".join(cell_content))
                #04 - (собственность) вид объекта
                elif cell_num == 4:
                    rich_man.add_ownflat(cell_content)
                #05 - (собственность) вид собственности (индивидуальная, долевая)
                elif cell_num == 5:
                    rich_man.add_ownflat_owntype(cell_content)
                #06 - (собственность) площадь
                elif cell_num == 6:
                    rich_man.add_ownflat_meters(cell_content)
                #07 - (собственность) страна расположения
                elif cell_num == 7:
                    rich_man.add_ownflat_country(cell_content)
                #08 - (пользование) вид объекта
                elif cell_num == 8:
                    rich_man.add_rentflat(cell_content)
                #09 - (пользование) площадь
                elif cell_num == 9:
                    rich_man.add_rentflat_meters(cell_content)
                #10 - (пользование) страна расположения
                elif cell_num == 10:
                    rich_man.add_rentflat_country(cell_content)
                #11 - транспорт
                elif cell_num == 11:
                    rich_man.add_car(cell_content)
                #12 - годовой доход
                elif cell_num == 12:
                    rich_man.add_money("".join(cell_content))
                #13 - сведения об источниках
                elif cell_num == 13:
                    rich_man.add_source("".join(cell_content))
        else:#not new_variant:
            pass
        #все разобрали теперь выгружать
        for profit_entry in rich_man.get_entryes():
            csv_writer.writerow(profit_entry)

    csv_file.close()

def main():
    for docx_file in utils.get_all_docx():
        docx_to_csv(docx_file)


#Будем форматировать числа в красивый вид
locale.setlocale(locale.LC_ALL, "")

if __name__=="__main__":
    if len(sys.argv) == 2:
        #значит передано имя обрабатываемого файла
        docx_to_csv(sys.argv[1])
    else:
        #не передали значит грузим все
        main()
