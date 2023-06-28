import requests
from docxtpl import DocxTemplate
import shutil


# ОБРАБОТКА ФИО

FIO_1 = str(input("Введите ФИО: "))
FIO_11 = FIO_1.split()
fam = FIO_11[0]
name = FIO_11[1]
otch = FIO_11[2]

# Запрашиваем ФИО в нужном нам падеже с сервиса Мофер
url = f'https://ws3.morpher.ru/russian/declension?s={fam}%20{name}%20{otch}'

# Маскируем запрос под браузерный,тк на сервисе стоит защита от бот-запросов
response = requests.get(url,headers={'User-Agent':
                            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                            "AppleWebKit/537.36 (KHTML, like Gecko) "
                            "Chrome/104.0.5112.79 "
                            "Safari/537.36"}, timeout=15).text

TPadej = (response[response.find('<Т>') + 3:response.find('</Т>')])


# Создание коротких подписей
famKor = TPadej.split()[0]
nameKor = f' {name[0]}.'
otchKor = f'{otch[0]}.'
FIO_mini_ImP = f'{fam}{nameKor}{otchKor}'
FIO_mini_TvP = f'{famKor}{nameKor}{otchKor}'

# Срздание ФИО в разных падежах
FIO_ImP = f'{fam} {name} {otch}'
FIO_TvP = TPadej

# Автоматическая смена окончаний в обращениях, в зависимости от пола

if otch[-1] == "а":
    gender = 'ая'
    gender_2 = "ка"
else:
    gender = 'ый'
    gender_2 = "ин"

# Ввод данных договора
print("Введите предмет договора: ")
predmet_data = []
a = str
while a != "*":
    a = input(" ")
    predmet_data.append(a)
else:
    predmet_data.remove("*")

predmet_data = (' '.join(map(str, predmet_data)))


predmetKOR = str(input("ВВедите короткий предмет договора (а именно создание) : "))
predmetKOR_2 = str(input("ВВедите короткий предмет договора ( в рамках исследовательской деятельности необходимо ): "))
svedenie = str(input("Функционал по созданию ?---------?  отсутствует в....... "))
format = str(input("ВВедите формат : "))
from_0 = str(input("ВВедите начало срока: "))
to = str(input( "ВВедите конец срока : "))

from_0 = from_0.split(".")
month_dict = {
    "01": "Января",
    "02": "Февраля",
    "03": "Марта",
    "04": "Апреля",
    "05": "Мая",
    "06": "Июня",
    "07": "Июля",
    "08": "Августа",
    "09": "Сентября",
    "10": "Октября",
    "11": "Ноября",
    "12": "Декабря"
}

month = str(month_dict.get(from_0[1]))

from_1 = from_0[0]

from_0 = f'{from_0[0]+"."+from_0[1]+"."+from_0[2]}'

pay = int(input("ВВедите зп: "))
pay_with_NDFL = (pay / 0.87)
print("при рассчете ЗП составляет" + " " + str(pay_with_NDFL) + "до скольки округлить ?:")
pay =input()

# Ввод Инициатора
print("По инициативе:")
Obosnov= str(input())


# Ввод компетенции Исполнителя

Kompit = str(input("Введите Компетенцию исполнителя : "))

# Перевод численного вида зарплаты в текстовый

response2 = requests.get(f'https://summa-propisyu.ru/?summ={pay}&vat=0&val=0&sep=0').text
pay_txt = (response2[response2.find('textarea')+ 98:response2.find('/textarea')])
pay_txt = (' '.join((pay_txt.split(' ')[:-3]))).capitalize()

# Ввод паспортных данных

print("Введите паспортные данные : ")
passport_data = []
for i in range(18):
    passport_data.append(input(" "))
print("Введите почту: ")
pochta = str(input())

# Выбор вида договора
typeof = str
print("Договор Авторский? 1-ДА 2-НЕТ")
Check = int(input())
if Check == 1:
    typeof = "Типовая. В связи с присутствием пунктов по отчуждению авторского права в типовом договоре."
if Check == 2:
    typeof = "Нетиповая. В связи с отсутствием пунктов по отчуждению авторского права в типовом договоре."

# ГПХ рендер

GPH = DocxTemplate(r"C:\Users\IGOR\Desktop\работа\шаблон\ДОГОВОР.docx")
context = {'FIO_ImP': FIO_ImP,
            'gender': gender,
            'predmet': predmet_data,
            'predmetKOR': predmetKOR,
            'format': format,
            'from_0': from_0,
            'to': to,
            'pay': pay,
            'pay_txt': pay_txt,
            'from_1': from_1,
            'month': month,
            'passport_data':
                       passport_data[0] + "\n" +
                       passport_data[1] + "\n" +
                       passport_data[2] + "\n" +
                       passport_data[3] + "\n" +
                       passport_data[4] + "\n" +
                       passport_data[5] + "\n" +
                       passport_data[6] + "\n" +
                       passport_data[7] + "\n" +
                       passport_data[8] + "\n" +
                       passport_data[9] + "\n" +
                       passport_data[10] + "\n" +
                       passport_data[11] + "\n" +
                       passport_data[12] + "\n" +
                       passport_data[13] + "\n" +
                       passport_data[14] + "\n" +
                       passport_data[15] + "\n" +
                       passport_data[16] + "\n" +
                       passport_data[17],
            'FIO_mini_ImP': FIO_mini_ImP}

GPH.render(context)
FileName = f'ДОГОВОР_{FIO_ImP}_{month}'
GPH.save(r"C:\Users\IGOR\Desktop\работа\Папка сохранения документа\Документ Microsoft Word.docx")
shutil.move(r"C:/Users/IGOR/Desktop/работа/Папка сохранения документа/Документ Microsoft Word.docx",
                f'C:/Users/IGOR/Desktop/работа/Папка сохранения документа/{FileName}.docx')

# Акт рендер

Akt = DocxTemplate(r"C:\Users\IGOR\Desktop\работа\шаблон\АКТ.docx")
context = {'FIO_ImP': FIO_ImP,
            'gender': gender,
            'predmet': predmet_data,
            'predmetKOR': predmetKOR,
            'format': format,
            'from_0': from_0,
            'to': to,
            'pay': pay,
            'pay_txt': pay_txt,
            'FIO_mini_ImP': FIO_mini_ImP}

Akt.render(context)
FileName2 = f'AKT_{FIO_ImP}_{month}'
Akt.save(r"C:\Users\IGOR\Desktop\работа\Папка сохранения документа\Документ Microsoft Word.docx")
shutil.move(r"C:/Users/IGOR/Desktop/работа/Папка сохранения документа/Документ Microsoft Word.docx",
                f'C:/Users/IGOR/Desktop/работа/Папка сохранения документа/{FileName2}.docx')

# СЗ ГПХ рендер

CZ_GPH = DocxTemplate(r"C:\Users\IGOR\Desktop\работа\шаблон\СЗ_ГПХ.docx")
context = {'FIO_mini_TvP': FIO_mini_TvP,
               'Obosnov': Obosnov,
               'predmet': predmet_data,
               'predmetKOR': predmetKOR,
               'predmetKOR_2': predmetKOR_2,
               'from_0': from_0,
               'to': to,
               'typeof': typeof,
               'pay': pay,
               'pay_txt': pay_txt,
               'FIO_TvP': FIO_TvP,
               'fam': fam}
CZ_GPH.render(context)
FileName3 = f'СЗ-ГПХ_{FIO_ImP}_{month}'
CZ_GPH.save(r"C:\Users\IGOR\Desktop\работа\Папка сохранения документа\Документ Microsoft Word.docx")
shutil.move(r"C:/Users/IGOR/Desktop/работа/Папка сохранения документа/Документ Microsoft Word.docx",
                f'C:/Users/IGOR/Desktop/работа/Папка сохранения документа/{FileName3}.docx')

# Обосвнование рендер

OBS = DocxTemplate(r"C:\Users\IGOR\Desktop\работа\шаблон\Обоснование.docx")
context = {
               'Kompit': Kompit,
                'Obosnov': Obosnov,
               'predmet': predmet_data,
               'predmetKOR': predmetKOR,
               'svedenie': svedenie,
               'format': format,
               'from_0': from_0,
               'to': to,
               'pay': pay,
               'pay_txt': pay_txt,
               'FIO_ImP': FIO_ImP
               }
OBS.render(context)
FileName4 = f'ОБОСНОВАНИЕ_{FIO_ImP}_{month}'
OBS.save(r"C:\Users\IGOR\Desktop\работа\Папка сохранения документа\Документ Microsoft Word.docx")
shutil.move(r"C:/Users/IGOR/Desktop/работа/Папка сохранения документа/Документ Microsoft Word.docx",
                f'C:/Users/IGOR/Desktop/работа/Папка сохранения документа/{FileName4}.docx')

# НДА рендер


NDA = DocxTemplate(r"C:\Users\IGOR\Desktop\работа\шаблон\НДА.docx")
context = {
               'FIO_ImP': FIO_ImP,
               'gender': gender,
               'gender_2': gender_2,
               'predmet': predmet_data,
               'passport_data':
                       passport_data[0] + "\n" +
                       passport_data[1] + "\n" +
                       passport_data[2] + "\n" +
                       passport_data[3] + "\n" +
                       passport_data[4] + "\n" +
                       passport_data[5] + "\n" +
                       passport_data[6] + "\n" +
                       passport_data[7] + "\n" +
                       passport_data[8] + "\n" +
                       passport_data[9] + "\n" +
                       passport_data[10] + "\n" +
                       passport_data[11] + "\n" +
                       passport_data[12] + "\n" +
                       passport_data[13] + "\n" +
                       passport_data[14] + "\n" +
                       passport_data[15] + "\n" +
                       passport_data[16] + "\n" +
                       passport_data[17],
               'FIO_mini_ImP': FIO_mini_ImP,
                'pochta': pochta,
               }
NDA.render(context)
FileName5 = f'НДА_{FIO_ImP}_{month}'
NDA.save(r"C:\Users\IGOR\Desktop\работа\Папка сохранения документа\Документ Microsoft Word.docx")
shutil.move(r"C:/Users/IGOR/Desktop/работа/Папка сохранения документа/Документ Microsoft Word.docx",
                f'C:/Users/IGOR/Desktop/работа/Папка сохранения документа/{FileName5}.docx')
