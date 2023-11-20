import requests
from bs4 import BeautifulSoup as bs
from openpyxl import load_workbook
import re

URL_TEMPLATE = "https://ssmu.ru/sveden/employees/"
FILE_NAME = "Книга1.xlsx"

yearRegex = re.compile(r'(\d{4})\s*[г|Г]\.')
codeRegex = re.compile(r'(\d{4})')

def main():

    url = URL_TEMPLATE
    result_list = {'href': [], 'title': [], 'about': []}
    r = requests.get(url)
    soup = bs(r.text, "html.parser")
    # print(soup)
    vacancies_names = soup.find_all('td', itemprop='fio')
    all_post = soup.find_all('td', itemprop='post')
    dagrees = soup.find_all('td', itemprop='degree')
    employes = soup.find_all('td', itemprop='employeeQualification')
    proffes = soup.find_all('td', itemprop='profDevelopment')
    genExperience = soup.find_all('td', itemprop='genExperience')
    specExperience = soup.find_all('td', itemprop='specExperience')

    code = soup.find_all('td')


    teachingDiscipline = soup.find_all('td', itemprop='teachingDiscipline')
    
    

    wb = load_workbook(FILE_NAME)
    sheet = wb.active

    x=2
    count =  10*25+20
    for i in range(20,count, 10):
        fio = vacancies_names[i].text.split(" ")
        post = all_post[i].text.split(" ")
        clear_post = post[0]+" "+post[1]+" "+post[2]

        # prof_parser = bs(proffes[i], "html.parser")
        # prof = prof_parser.find_all('p')


        if len(fio) < 3:
            print(f'Тут не все {fio}')
            continue
        
        lastName = fio[0]
        name = fio[1]
        middleName = fio[2]

        # print(lastName)
        # print(name)
        # print(middleName)

        code = re.findall(r'\d{2}\.\d{2}\.\d{2}', teachingDiscipline[i].find_previous_sibling('td').text)
        code_string = ' '.join(code)

        
        sheet["B"+str(x)].value = lastName
        sheet["C"+str(x)].value = name
        sheet["D"+str(x)].value = middleName
        sheet["E"+str(x)].value = clear_post
        sheet["F"+str(x)].value = dagrees[i].text
        sheet["G"+str(x)].value = employes[i].text
        sheet["H"+str(x)].value = ""
        sheet["I"+str(x)].value = genExperience[i].text
        sheet["J"+str(x)].value = specExperience[i].text
        sheet["L"+str(x)].value = teachingDiscipline[i].text
        sheet["K"+str(x)].value = code_string
                

        j=0

        for j in proffes[i].contents:
            sheet["H"+str(x)].value += j.text

            yearMatch = re.search(yearRegex, j.text)
            if yearMatch:
                yearStr = yearMatch.group(1)
                year = int(yearStr)
            else:
                print(f"год не найден: {j.text}")
                continue
            
            if year>2019:
                sheet["H"+str(x)].value += j.text
            else:
                print(f"год не подходит. Год: {year}")
                continue

        x+=1

        # print(vacancies_names)


    wb.save(FILE_NAME)

main()


 #def parse(url = URL_TEMPLATE):

#     writer = pd.ExcelWriter("file.xlsx", engine='xlsxwriter')
#     data = "10"
#     data.to_excel(writer, sheet_name='Тест')
#     writer.save()





#     # for name in vacancies_names:
#     # for name in vacancies_names:
#     #     result_list['href'].append('https://www.work.ua'+name.a['href'])
#     #     result_list['title'].append(name.a['title'])
#     # for info in vacancies_info:
#     #     result_list['about'].append(info.text)
#     # return result_list


# def excel():
#     # file = "file.xlsx"
#     # xl=pd.ExcelWriter(file)
#     # xl.sheets_names

#     # df1=xl.parse("Sheet1")

#     writer = pd.ExcelWriter("file.xlsx", engine='xlsxwriter')
#     data = "10"
#     data.to_excel(writer, sheet_name='Тест')
#     writer.save()

# # excel()





# # df = pd.DataFrame(data=parse())
# # df.to_csv(FILE_NAME)
