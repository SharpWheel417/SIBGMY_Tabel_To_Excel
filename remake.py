import requests
from bs4 import BeautifulSoup as bs
from openpyxl import load_workbook
import re

URL_TEMPLATE = "https://ssmu.ru/sveden/employees/"
FILE_NAME = "Книга1.xlsx"


YEAR_REGEX = re.compile(r'(\d{4})\s*[г|Г]\.')
CODE_REGEX = re.compile(r'\d{2}\.\d{2}\.\d{2}')

html = requests.get(URL_TEMPLATE)
r = bs(html.text, "html.parser")

# def GetData():
#     r = requests.get(URL_TEMPLATE)
#     soup = bs(r.text, "html.parser")
#     return soup

## Получает имена с сайта по имени
## Массивы
def GetData(tag, soup):
    data= soup.find_all('td', itemprop=tag)
    return data

def GetDataTextByIndex(tag, soup, index):
    data= soup.find_all('td', itemprop=tag)
    return data[index].text

## Убирает пробелы
def ClearSpaces(string):
    return string.strip()


## Получаем фио по индексу
def GetFIO(num):
    fullName = GetData('fio', r)
    name = ClearSpaces(fullName[num].text)
    full = name.split()
    return full[0], full[1], full[2]

def GetCode(num):
    teachingDiscipline = r.find_all('td', itemprop='teachingDiscipline')
    code = re.findall(CODE_REGEX, teachingDiscipline[num].find_previous_sibling('td').text)
    code_string = ' '.join(code)
    return code_string





def main():

    # Ексель
    wb = load_workbook(FILE_NAME)
    sheet = wb.active

    ##Назанчаем индексы (какие строки брать)

    start = 0 # С какого препода начинаем
    interval = 1 # с каким индервалом идем
    much = 8 # Сколько надо
    end = (interval*much)+start # Конец

    numRow = 1

    for i in range(470, 500, interval):
    
    
        numRow+=1
        index_fio = i+8
        print(numRow)
        post = ClearSpaces(GetDataTextByIndex('teachingLevel', r, i)) # Должность

        if 'кафедры неврологии и нейрохирургии' in post:
            print('есть:', numRow)

            lastName, firstName, middleName = GetFIO(index_fio) # ФМО 
            dagree = ClearSpaces(GetDataTextByIndex('degree', r, i)) # Степень
            employee = ClearSpaces(GetDataTextByIndex('employeeQualification', r, i)) # Квалификация
            cvalification = ClearSpaces(GetDataTextByIndex('profDevelopment', r, i)) # Повышение квалификация
            allStash = ClearSpaces(GetDataTextByIndex('genExperience', r, i)) # Общее стаж
            specStash = ClearSpaces(GetDataTextByIndex('specExperience', r, i)) # Специальное ситаж
            desciplines = ClearSpaces(GetDataTextByIndex('teachingDiscipline', r, i)) # Дисциплины
            code = GetCode(i) # Код
            sheet["A"+str(numRow)].value = str(numRow-1)
            sheet["B"+str(numRow)].value = lastName
            sheet["C"+str(numRow)].value = firstName
            sheet["D"+str(numRow)].value = middleName
            sheet["E"+str(numRow)].value = post
            sheet["F"+str(numRow)].value = dagree
            sheet["G"+str(numRow)].value = employee
            sheet["H"+str(numRow)].value = cvalification
            sheet["I"+str(numRow)].value = allStash
            sheet["J"+str(numRow)].value = specStash
            sheet["L"+str(numRow)].value = desciplines
            sheet["K"+str(numRow)].value = code
# 
        # print(lastName)
        # print(firstName)
        # print(middleName)
        # print(post)
        # print(dagree)
        # print(employee)
        # print(cvalification)
        # print(allStash)
        # print(specStash)
        # print(desciplines)
        # print(code)
        else:
            continue

    wb.save(FILE_NAME)


main()