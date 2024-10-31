import xlwings as xw
import time
import calendar
import locale


def none_to_dash(value):
    if str(value) == "None":
        return '%10s' % '———'
    return '%10s' % str(value) or '%10s' % '———'


def get_source() -> str:
    from selenium import webdriver
    from selenium.webdriver.firefox.options import Options

    src: str = ''
    options = Options()
    options.add_argument('--headless')
    driver = webdriver.Firefox(options=options)
    link = 'http://rgkript.ru/raspisanie-zanyatiy/'
    print("Подключение модуля браузера...")
    driver.get(link)
    time.sleep(1)
    src = driver.page_source
    print('\033[F\033[K', end='', flush=True)
    print("Получение кода веб-страницы...")
    driver.quit()
    return src


def get_file_links(page_source) -> None:
    import urllib.request as req
    from bs4 import BeautifulSoup

    soup = BeautifulSoup(page_source, "html.parser")
    print('\033[F\033[K', end='', flush=True)
    print("Парсинг кода веб-страницы...")
    req.urlretrieve(soup.find(
        "img", class_="alignright wp-image-4749").parent["href"], "timetable_tmp.xls")
    print('\033[F\033[K', end='', flush=True)
    print("Скачивание расписания...")
    for strong in soup.findAll("strong"):
        if strong.text.__contains__("ЗАМЕНЫ"):
            req.urlretrieve(strong.parent["href"], "zameny.doc")
            print('\033[F\033[K', end='', flush=True)
            print("Скачивание замен...")


def get_replacements(group: str) -> list[list[str]]:
    from docx2python import docx2python
    import doc2docx

    print('\033[F\033[K', end='', flush=True)
    print("Анализ файла замен...")
    print('\033[F\033[K', end='', flush=True)
    doc2docx.convert("zameny.doc", "zameny.docx")
    doc = docx2python("zameny.docx")
    d = doc.text.split("\n")
    i = 10
    output_string = ""
    replacements: list[list[str]] = []
    for s in d:
        if i < 9:
            output_string += s.rstrip(" ") + " "
            if len(s) > 1:
                replacements[len(replacements)-1].append(s.rstrip(" "))
            i += 1
        if s == group:
            i = 1
            output_string = s.rstrip(" ") + " "
            replacements.append([s.rstrip(" ")])
    return replacements


locale.setlocale(locale.LC_TIME, 'ru_RU')

print("")

group = "ПО-33к"

for i in range(0, 6):
    print(f"{i+1} — {calendar.day_name[i]}")
weekday = int(input("Введите номер для недели:\t")) - 1

print('\033[F\033[K'*7, end='', flush=True)
print("День недели: ", calendar.day_name[weekday])
print("Группа: ", group, "\n")

get_file_links(get_source())

table: xw.Sheet = xw.Book("timetable_tmp.xls").sheets['Учебные группы']

print('\033[F\033[K', end='', flush=True)
print("Анализ файла расписания...")
i = 7
timetable = []
while True:
    if table[f'B{i}'].value == None:
        break

    temp_i = i+3
    if table[f'B{i}'].value == group:
        counter = 0
        while True:
            temp_string = ''
            timetable.append([])
            for cell in table.range(f'B{temp_i}:G{temp_i}').columns:
                temp_string += '%22s' % (none_to_dash(cell.value))
                timetable[-1].append(none_to_dash(str(cell.value).strip()))

            temp_string = ''
            for cell in table.range(f'B{temp_i+1}:G{temp_i+1}').columns:
                temp_string += '%22s' % (cell.value)
                timetable[-1].append(str(cell.value).strip())

            temp_string += '\n'
            temp_i += 4
            counter += 1
            if counter > 5:
                break
        break

    i += 36

replacements = get_replacements(group)

print('\033[F\033[K', end='', flush=True)

for i in range(0, 6):
    print("%10s" % timetable[i][weekday])
    print("%10s" % timetable[i][weekday+6])

for s in replacements:
    print(f"{s[1]} {s[2]}\n{" ".join(s[3:])}\n")

print('\033[F\033[K', end='', flush=True)
