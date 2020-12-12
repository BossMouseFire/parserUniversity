import json
import openpyxl
from openpyxl.styles import Alignment
from openpyxl.styles import Font, Fill

FILEname = "C:/Users/Дмитрий/Desktop/University education/OLIMP/Python/TOI Individual Task/Statistic information Test.xlsx"

try:
    wb = openpyxl.load_workbook(FILEname)
except Exception as error:
    wb = openpyxl.Workbook()

    # Удаление листа, создаваемого по умолчанию, при создании документа
    for sheet_name in wb.sheetnames:
        sheet = wb.get_sheet_by_name(sheet_name)
        wb.remove_sheet(sheet)
    # создание листов по умолчанию
    titlelist = wb.create_sheet('Title')
    firstlist = wb.create_sheet('Statistics')
    secondlist = wb.create_sheet('Full Statistics')
titlelist = wb.get_sheet_by_name('Title')
firstlist = wb.get_sheet_by_name('Statistics')
secondlist = wb.get_sheet_by_name('Full Statistics')

# оформление листа
class sheetdecoration():
    def preparetitlelist(self):
        titlelist['A1'] = "Анализ работы искусственного интеллекта сайта ResearchGate по определению ключевых слов на других источниках (Elibrary)"
        titlelist['A4'] = "Работу выполнили: Абанин Д. А., Ильбеков Д. С."
        titlelist['A7'] = "Примечание"
        titlelist['A8'] = """Коэффицент совпадения(k) - это процент сопоставимости ключевых слов из Researchgate с Elibrary. Коэффициент совпадения расчитывается следующим образом: (a+b+c)/n*100%, где:
a - количество полностью совпавших ключевых слов
b - количество ""хороших"" совпадений (""Хорошие совпадения"" - это такие два ключевые слова, которые совпадают по двум и более корням, например: 1 ключевое слово:['applied', 'artificial', 'intelligence'] 2 ключевое слово:['artificial', 'intelligence'], программа находит следующие корни: {'artificial': 0, 'intelligence': 0} - как видим, это хорошее совпадение, которое можно считать почти попаданием)
c - количество возможных совпадений (возможные совпадение - это такие два ключевых слова, у которых совпадает только 1 корень)
Показатели a, b, c относятся к одному человеку
n - количество ключевых слов на сайте research у конкретного человека.
"""
        titlelist['A10'] = "Вывод"
        titlelist['A11'] = """В результате проделанной работы мы пришли к определенным выводам. Т.к. средний коэффицент совпадения довольно низок, можно заключить, что искусственный интелект сайта ResearchGate плохо распознает ключевые слова по предложенным статьям. Но можно предположить, что часть данных на сайте Elibrary не соответствуют действительности, потому что на одно и тоже ключевое слово может приходиться до 200 различных понятий, которые сильно связаны по смыслу. 
"""
        titlelist.column_dimensions['A'].width = 150
        titlelist.row_dimensions[8].height = 100
        # Добавление возможности переноса строк, если закончилась ширина столбца
        wrap_alignment = Alignment(wrap_text=True)
        titlelist['A8'].alignment = wrap_alignment
        titlelist['A11'].alignment = wrap_alignment
        # Изменение размера текста
        titlelist['A1'].font = Font(30)
        titlelist['A4'].font = Font(30)
        titlelist['A7'].font = Font(30)
        titlelist['A10'].font = Font(30)

        row_cells = 1
        while row_cells < 8:
            titlelist.row_dimensions[row_cells].height = 30
            row_cells += 1

    def preparefirstlist(self):
        firstlist['A1'] = "Фамилия Имя ученого"
        firstlist['B1'] = "Количество ключевых слов на reseachgate"
        firstlist['C1'] = "Количество ключевых слов на elibrary"
        firstlist['D1'] = "Количество полных совпадений"
        firstlist['E1'] = "Количество хороших совпадений"
        firstlist['F1'] = "Количество возможных совпадений"
        firstlist['G1'] = "Вывод"
        row_cells = 1
        while row_cells < 169:
            firstlist.row_dimensions[row_cells].height = 15
            row_cells += 1
        column_cells = 65
        while column_cells < 75:
            firstlist.column_dimensions[chr(column_cells)].width = 40
            column_cells += 1

    def preparesecondlist(self):
        secondlist['A1'] = "Фамилия Имя ученого"
        secondlist['B1'] = "Список ключевых слов"
        secondlist['D1'] = "Количество совпадений"
        secondlist['E1'] = "Список совпавших слов"
        secondlist['F1'] = "Список возможных совпадений"
        secondlist['H1'] = "Количество возможных совпадений"

        secondlist['B2'] = "ResearchGate"
        secondlist['C2'] = "Elibrary"
        secondlist['F2'] = "ResearchGate"
        secondlist['G2'] = "Elibrary"

        secondlist.merge_cells('B1:C1')
        secondlist.merge_cells('F1:G1')
        secondlist.merge_cells('A1:A2')
        secondlist.merge_cells('D1:D2')
        secondlist.merge_cells('E1:E2')
        secondlist.merge_cells('H1:H2')

        column_cells = 65
        while column_cells < 75:
            secondlist.column_dimensions[chr(column_cells)].width = 30
            column_cells += 1
sheetdecoration().preparetitlelist()
sheetdecoration().preparefirstlist()
sheetdecoration().preparesecondlist()

# считывание данных парсера Researchgate и Elibrary
with open("profiles_researchgate.json", "r", encoding="utf8") as file:
    dataResearch = json.load(file)
with open("profiles_elibrary.json", "r", encoding="utf8") as file:
    dataElibrary = json.load(file)

# Переменные, необходимые для работа алгоритма и вывода полезной информации
Sovpadenia = 0  # количество полных совпадений ( a - (Коэффициент схожести = 1)
PossibleSovpadenia = 0  # количество возможных совпадений ( c - (Коэффициент схожести = 0,1)
PossibleSovpadeniaHuman = 0  # переменная, необходимая для просчета возможных совпадений на одного человека
NotFindNamesSum = 0  # Количество ненайденных имен
FindNamesSum = 0  # Количество найденных имен
row_number_name = 2  # счетчик строк для вывода в таблицу Excel
# Словари совпадений
amountoffullsovpad = {}  # Словарь {Имя человека : Количество полных совпадений этого человека}
goodsovpad = {}  # Словарь {Имя человека : Количество хороших совпадений этого человека}
anysovpad = {}  # Словарь {Имя человека : Количество возможных совпадений этого человека}
# группы
average = 0
one_group = 0
two_group = 0
three_group = 0
four_group = 0

# Алгоритм поиска слов
for NameResearch in dataResearch:  # начинаем поиск человека из Research в Elibrary
    NotFindName = True  # Переменная "Найден ли человек?"
    # Заполнение таблицы Excel
    firstlist.cell(row=row_number_name, column=1).value = NameResearch
    firstlist.cell(row=row_number_name, column=2).value = len(dataResearch[NameResearch])
    try:
        firstlist.cell(row=row_number_name, column=3).value = len(dataElibrary[NameResearch])
    except KeyError:
        firstlist.cell(row=row_number_name, column=3).value = "Имя не найдено"
    for NameElibrary in dataElibrary:  # ищем совпадающее имя на Elibrary
        if NameResearch == NameElibrary:  # имя нашлось
            NotFindName = False  # человек найден
            for keyWordResearch in dataResearch[NameResearch]:  # выбираем ключевое слово из имени Research
                INOTFIND = True  # Переменная "Найдено ли полное совпадение?
                for keyWordElibrary in dataElibrary[NameElibrary]:  # Пытаемся найти полное совпадение ключего слова
                    if keyWordResearch == keyWordElibrary:  # ключевое слово нашлось
                        Sovpadenia += 1
                        # Добавление в словарь полных совпадений
                        for namei in amountoffullsovpad.keys():
                            if namei == NameResearch:
                                amountoffullsovpad[namei] += 1
                                break
                        else:
                            amountoffullsovpad.update({NameResearch: 1})
                        # Удаление из базы данных найденного слова
                        Del = dataElibrary[NameElibrary].index(keyWordElibrary)
                        dataElibrary[NameElibrary].pop(Del)

                        Del = dataResearch[NameResearch].index(keyWordResearch)
                        dataResearch[NameResearch].pop(Del)
                        INOTFIND = False  # Ключевое слово нашлось
                        break
                if INOTFIND:  # если не было найдено ключевое слово, начинаем поиск совпадений
                    FINDPOSSIBLE = False
                    for keyWordElibrary in dataElibrary[NameElibrary]:  # Вновь начинаем искать слова на Elibrary
                        if len(keyWordResearch) > 5:  # Если длина слова имеет больше 5 букв (без местоимений, союзов)
                            # Делим ключевые слова на подслова (Ключевое слово может состоять из нескольких слов)
                            tempstr1 = str(keyWordResearch).split(' ')
                            tempstr2 = str(keyWordElibrary).split(' ')
                            FINDwithROOT = 0  # количество найденных корней у ключевого слова
                            ListofROOTS = {}  # все корни у ключевого слова

                            for keySplitResearch in tempstr1:  # Подслова из ключевого слова Research
                                TempDelletter = 0  # переменная для создания новых корней подслова Research
                                TempLen = len(keySplitResearch)  # Проверяем длину подслова
                                ROOTUSE = False  # Переменная "Использовался ли корень?"
                                # Урезаем слово, пока не найдем подхощий корень, корень слова должен состоять минимум из 6 букв
                                while TempLen + TempDelletter > 5:
                                    if ROOTUSE:
                                        break
                                    if TempDelletter == 0:
                                        RootStr = keySplitResearch
                                    else:
                                        RootStr = str(keySplitResearch)[0:TempDelletter]  # "КОРЕНЬ" слова
                                    DifferencefromROOT = 0

                                    # "Умное" сравнивание
                                    for keySplitElibrary in tempstr2:  # Подслова из ключевого слова Elibrary
                                        if ROOTUSE:
                                            break
                                        letter = 0  # Количество урезанных букв
                                        lengthWord = len(keySplitElibrary)  # Длина подслова из ключевого слова Elibrary
                                        # Проверяем на равенство "корня" и подслова из Elibrary
                                        while len(keySplitElibrary) + letter >= len(RootStr):
                                            # Обработка случая, если полученный корень сразу подходит к подслову
                                            if letter == 0:
                                                if keySplitElibrary == RootStr:
                                                    FINDPOSSIBLE = True
                                                    ROOTUSE = True
                                                    FINDwithROOT += 1
                                                    DifferencefromROOT = abs(letter)
                                                    ListofROOTS.update({RootStr: DifferencefromROOT})
                                                    break
                                            elif str(keySplitElibrary)[0:letter] == RootStr:  # вырезаем слово до корня
                                                FINDPOSSIBLE = True
                                                ROOTUSE = True
                                                FINDwithROOT += 1
                                                DifferencefromROOT = abs(letter)
                                                ListofROOTS.update({RootStr: DifferencefromROOT})
                                                break
                                            letter -= 1
                                    TempDelletter -= 1
                            if FINDwithROOT > 0:  # нашлись слова благодаря корню
                                if FINDwithROOT > 1:  # если корней больше двух
                                    for namei in goodsovpad.keys():
                                        if namei == NameResearch:
                                            goodsovpad[namei] += 1
                                            break
                                    else:
                                        goodsovpad.update({NameResearch: 1})
                                print(NameResearch)  # имя человека
                                print(tempstr1)  # строка research
                                print(tempstr2)  # строка elibrary
                                print(ListofROOTS)
                                print(FINDwithROOT)
                                print('\n')
                    if FINDPOSSIBLE:
                        PossibleSovpadenia += 1
    try:
        firstlist.cell(row=row_number_name, column=4).value = amountoffullsovpad[NameResearch]
    except KeyError:
        firstlist.cell(row=row_number_name, column=4).value = 0
        amountoffullsovpad.update({
            NameResearch: 0
        })
    try:
        firstlist.cell(row=row_number_name, column=5).value = goodsovpad[NameResearch]
    except KeyError:
        firstlist.cell(row=row_number_name, column=5).value = 0
        goodsovpad.update({
            NameResearch: 0
        })
    numAnySov = PossibleSovpadenia - PossibleSovpadeniaHuman
    anysovpad.update({
        NameResearch: numAnySov
    })
    PossibleSovpadeniaHuman = PossibleSovpadenia
    firstlist.cell(row=row_number_name, column=6).value = anysovpad[NameResearch]
    try:
        quality = (amountoffullsovpad[NameResearch] + goodsovpad[NameResearch] * 0.25 +
                   anysovpad[NameResearch] * 0.1) * 100 / \
                  (len(dataResearch[NameResearch]) + amountoffullsovpad[NameResearch])
        average += quality
        if quality > 50:
            one_group += 1
        elif quality > 40:
            two_group += 1
        elif quality > 25:
            three_group += 1
        else:
            four_group += 1
        firstlist.cell(row=row_number_name, column=7).value = f'Процент совпадения elibrary c ' \
            f'researchgate равен {round(quality, 2)}%.'
    except ZeroDivisionError:
        firstlist.cell(row=row_number_name, column=7).value = "Машинное обучение не определило ключевые слова"
    # Обработка ненайденных имен
    if NotFindName:
        NotFindNamesSum += 1
        print(NameResearch + "   NOT FIND NAME")
    else:
        FindNamesSum += 1
    row_number_name += 1

# Сохранение файла Excel
wb.save(FILEname)

# Вывод полученной информации
print('\n')
print('Количество ненайденных имен: ' + str(NotFindNamesSum))
print('Количество найденных имен: ' + str(FindNamesSum))
print('Полные совпадения слов: ' + str(Sovpadenia))
print('Возможные совпадения слов: ' + str(PossibleSovpadenia))
print('\n')
print("Количество полных совпадений у ученых: " + str(amountoffullsovpad))
print("Количество хороших совпадений у ученых (у ключевых слов обнаружено совпадение двух и более подслов): " + str(goodsovpad))
print("Количество возможных совпадений у ученых (у ключевых слов обнаружено совпадение одного подслова): " + str(anysovpad))
print()
print('Средний коэффицент совпадения: ', average / FindNamesSum)
print("Количество человек в 1 группе: " + str(one_group))
print("Количество человек в 2 группе: " + str(two_group))
print("Количество человек в 3 группе: " + str(three_group))
print("Количество человек в 4 группе: " + str(four_group))
