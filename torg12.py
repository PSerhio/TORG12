import xlrd
import re


class TableHead:
    """Шапка таблицы"""
    # код колонки: [имя колонки, относительное смещение строки, *(номер колонки Excel)]
    def __init__(self):
        self.head = {1: ['номерпопорядку', 0],
                     2: ['наименование,характеристика,сорт,артикултовара', 1],
                     3: ['код', 1],
                     4: ['наименование', 1],
                     5: ['кодпоОКЕИ', 1],
                     6: ['видупаковки', 0],
                     7: ['водномместе', 1],
                     8: ['мест,штук', 1],
                     9: ['массабрутто', 0],
                     10: ['количество(массанетто)', 0],
                     11: ['цена,руб.коп.', 0],
                     12: ['суммабезучетаНДС,руб.коп.', 0],
                     13: ['ставка,%', 1],
                     14: ['сумма,руб.коп.', 1],
                     15: ['суммасучетомНДС,руб.коп.', 0]
                     }


class TableString:
    """Строка таблицы"""
    def __init__(self, number, tovar, code, unit, unit_code, kind_of_pack, quantity_per_packplace,
                 packplace_quantity, weight, quantity, price, summa_bez_nds, stavka_nds, summa_nds,
                 summa_incl_nds):
        self.number = number
        self.tovar = tovar
        self.code = code
        self.unit = unit
        self.unit_code = unit_code
        self.kind_of_pack = kind_of_pack
        self.quantity_per_packplace = quantity_per_packplace
        self.packplace_quantity = packplace_quantity
        self.weight = weight
        self.quantity = quantity
        self.price = price
        self.summa_bez_nds = summa_bez_nds
        self.stavka_nds = stavka_nds
        self.summa_nds = summa_nds
        self.summa_incl_nds = summa_incl_nds

    def __str__(self):
        return str(self.number)+'|'+self.tovar+'|'+str(self.code)+'|'+str(self.quantity)+'|'+str(self.price)+'|'\
               + str(self.summa_bez_nds)


class Torg12(TableHead, TableString):

    def __init__(self, file_name):
        self.name = file_name
        self.book = xlrd.open_workbook(self.name, formatting_info=True, on_demand=True, encoding_override="cp1251")
        self.sheet = self.book.sheet_by_index(0)
        self.valid = True
        self.number_document = None
        self.date_document = None
        self.values_document_row = None
        self.values_document = {}
        self.__nrows = self.sheet.nrows
        self.__ncols = self.sheet.ncols
        self.__pages = {}
        self.__head_table = TableHead()
        self.value_table = []

        if self.check_valid():
            self.__get_table_value()
            self.__check_document_value()

    def __str__(self):
        if self.valid:
            result = ''
            result += f'№ документа: {self.number_document} дата: {self.date_document}\n'
            for item in self.value_table:
                result += f'{item}\n'
            result += f'Всего по накладной кол-во: {self.values_document.get(10)} сумма: {self.values_document.get(15)}'
            return result
        else:
            return 'Документ не обработан'

    def __get_pages(self):
        i = 1
        for row in range(0, self.__nrows):
            for col in range(0, self.__ncols):
                if 'Страница' in str(self.sheet.cell_value(row, col)):
                    self.__pages[i] = row + 1
                    i += 1
                if 'Номер документа' in str(self.sheet.cell_value(row, col)):
                    number = self.sheet.cell_value(row+1, col)
                    if type(number) is float or type(number) is int:
                        self.number_document = str(int(number))
                if 'Дата составления' in str(self.sheet.cell_value(row, col)):
                    self.date_document = str(self.sheet.cell_value(row+1, col))
                if 'Всего по накладной' in str(self.sheet.cell_value(row, col)):
                    self.values_document_row = row
        # print(self.pages)
        return self.__pages

    def check_valid(self):
        self.__get_pages()
        row = self.__pages.get(1)
        # читаем стоку с цифрами и сверяем с соответствующими значениями текстовой информации (сверяем с образцом)
        # запоминаем колонку Excel для колонки таблицы
        for key in self.__head_table.head:
            for col in range(0, self.__ncols):
                if self.sheet.cell_value(row + 2, col) == key:
                    item = self.__head_table.head.get(key)
                    line = self.sheet.cell_value(row + item[1], col)
                    # убираем управляющие символы, пробелы, тире и переводим все в нижний регистр
                    line = re.sub(r'[\x00-\x20\x2d]+', '', line).lower()
                    if item[0].lower() == line:
                        self.__head_table.head.get(key).append(col)
                    else:
                        error = f'Ошибка соответствия колонок {key} и {self.sheet.cell_value(row + item[1], col)}'
                        print(error)
                        self.valid = False
                        break

        # читаем текстовые значения, и сверяем с цифровым значением с образцом (на наличие лишних столбцов)
        for col in range(0, self.__ncols):
            if self.sheet.cell_value(row, col):
                if not self.sheet.cell_value(row + 2, col):
                    error = f'Колонка "{self.sheet.cell_value(row, col)}" не предусмотренная стандартом ТОРГ12 ' \
                            f'№ {self.sheet.cell_value(row + 2, col)}'
                    print(error)
                    self.valid = False
                    break
            if self.sheet.cell_value(row + 1, col):
                if not self.sheet.cell_value(row + 2, col):
                    error = f'Колонка не предусмотренная стандартом ТОРГ12{self.sheet.cell_value(row + 2, col)} ' \
                            f'№ {self.sheet.cell_value(row + 2, col)}'
                    print(error)
                    self.valid = False
                    break
        # print(self.head_table.head)
        self.values_document = {8: self.sheet.cell_value(self.values_document_row, self.__head_table.head[8][2]),
                                9: self.sheet.cell_value(self.values_document_row, self.__head_table.head[9][2]),
                                10: self.sheet.cell_value(self.values_document_row, self.__head_table.head[10][2]),
                                12: self.sheet.cell_value(self.values_document_row, self.__head_table.head[12][2]),
                                14: self.sheet.cell_value(self.values_document_row, self.__head_table.head[14][2]),
                                15: self.sheet.cell_value(self.values_document_row, self.__head_table.head[15][2])}
        return self.valid

    def __get_table_value(self):
        for key in self.__pages:
            page_rows = self.__pages[key]
            for row in range(page_rows + 3, self.__nrows):
                if self.sheet.cell_value(row, self.__head_table.head[1][2]):
                    number = int(self.sheet.cell_value(row, self.__head_table.head[1][2]))
                    kod = self.sheet.cell_value(row, self.__head_table.head[3][2])
                    if type(kod) is float or type(kod) is int:
                        kod = str(int(kod))
                    table_string = TableString(number,
                                               self.sheet.cell_value(row, self.__head_table.head[2][2]),
                                               kod,
                                               self.sheet.cell_value(row, self.__head_table.head[4][2]),
                                               self.sheet.cell_value(row, self.__head_table.head[5][2]),
                                               self.sheet.cell_value(row, self.__head_table.head[6][2]),
                                               self.sheet.cell_value(row, self.__head_table.head[7][2]),
                                               self.sheet.cell_value(row, self.__head_table.head[8][2]),
                                               self.sheet.cell_value(row, self.__head_table.head[9][2]),
                                               self.sheet.cell_value(row, self.__head_table.head[10][2]),
                                               self.sheet.cell_value(row, self.__head_table.head[11][2]),
                                               self.sheet.cell_value(row, self.__head_table.head[12][2]),
                                               self.sheet.cell_value(row, self.__head_table.head[13][2]),
                                               self.sheet.cell_value(row, self.__head_table.head[14][2]),
                                               self.sheet.cell_value(row, self.__head_table.head[15][2])
                                               )
                    self.value_table.append(table_string)
                if not self.sheet.cell_value(row, self.__head_table.head[1][2]):
                    break

    def __check_document_value(self):
        mest = 0
        massa = 0
        kolvo = 0
        summa_bez_nds = 0
        summa_nds = 0
        summa_incl_nds = 0
        for item in self.value_table:
            if item.packplace_quantity:
                mest += item.packplace_quantity
            if item.weight:
                massa += item.weight
            if item.quantity:
                kolvo += item.quantity
            if item.summa_bez_nds:
                summa_bez_nds += item.summa_bez_nds
            if item.summa_nds:
                summa_nds += item.summa_nds
            if item.summa_incl_nds:
                summa_incl_nds += item.summa_incl_nds
        if mest == 0:
            mest = ''
        if massa == 0:
            massa = ''
        if kolvo == 0:
            kolvo = ''
        if summa_bez_nds == 0:
            summa_bez_nds = ''
        else:
            summa_bez_nds = round(summa_bez_nds, 2)
        if summa_nds == 0:
            summa_nds = ''
        else:
            summa_nds = round(summa_nds, 2)
        if summa_incl_nds == 0:
            summa_incl_nds = ''
        else:
            summa_incl_nds = round(summa_incl_nds, 2)
        current_document_value = {8: mest, 9: massa, 10: kolvo, 12: summa_bez_nds,
                                  14: summa_nds, 15: summa_incl_nds}
        # print(current_document_value)
        # print(self.values_document)
        if current_document_value != self.values_document:
            error = "Итоговые данные не бьются с табличной частью. \n" \
                    "Возможные причины: отсутствие части документа, неправильные расчеты, округление\n" \
                    "Если выше сказанное проверено, обратитесь к разработчику"
            print(error)
