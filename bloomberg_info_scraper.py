import requests
from bs4 import BeautifulSoup
import re
import xlsxwriter

import time

start_time = time.time()

def auth_session(link):
    s = requests.Session()
    page0 = s.get(link)
    soup = BeautifulSoup(page0.text)
    personLink1 = "http://investing.businessweek.com/research/"
    personLink2 = soup.find('a',text="See Board Relationships")['href']
    personLink = personLink1 + personLink2[6:]
    return [personLink,s]

def main():
    link = input("Please, enter the webpage address: ")
    session = requests.Session()
    isPrivate = input("Is it one of those private pages that did not work in the old version?[y/n]: ")
    if isPrivate == "y":
        linkPlusSession = auth_session(link)
        link = linkPlusSession[0]
        session = linkPlusSession[1]
    page = session.post(link)
    soup = BeautifulSoup(page.text)
    names_raw = soup.find_all('div', class_="name")
    title_raw = soup.find_all('div', class_="title")
    large_detail_raw = soup.find_all('td', class_="largeDetail")
    companies_raw = soup.find_all('a')
    main_name = names_raw[0].get_text()#name of the main dude, name .xls file after it
    name_end_index = main_name.find('\xa0')#getting rid of that RETURN TO bs at the end of the name tag 
    main_name = main_name[:name_end_index-1]#
    main_title = title_raw[0].get_text()

    large_details = []
    for element in large_detail_raw:
        large_details.append(element.get_text())
    main_guy = {}

    main_guy["name"] = main_name
    main_guy["title"] = main_title
    main_guy["age"] = large_details[0]
    main_guy["annual_comp"] = large_details[1]

    names_raw.pop(0)
    names = []
    companies = []
    for i in names_raw:
        names.append(i.get_text())
    for i in companies_raw:
        companies.append(i.get_text())
    c = len(companies)#used as an index to cut off unnecessary stuff in the beginning
    
    for element in names:
        if ( c > companies.index(element)):
            c = companies.index(element)
    companies = companies[c:]        
    d = companies.index("\n\n")#used as an index to cut off unnecessary stuff at the end
    companies = companies[:d]
    peopleDict = {}
    counter = 0
    newPersonCounter = 0
    for element in companies:
        if element in names:
            peopleDict[element] = {}
            peopleDict[element]["company"] = []
            peopleDict[element]["affiliations"] = []
            currentPerson = element
            newPersonCounter = 0
        elif companies[newPersonCounter-1] in names:
            peopleDict[companies[counter-1]]["company"].append(element)
        elif ((newPersonCounter != 0) and (element != "Board Affiliations")):
            peopleDict[currentPerson]["affiliations"].append(element)
        newPersonCounter += 1
        counter += 1

    workbook = xlsxwriter.Workbook(main_guy["name"] + ".xlsx")
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': True})
    worksheet.write(0,0,"name",bold)
    worksheet.write(0,1,"title",bold)
    worksheet.write(0,2,"age",bold)
    worksheet.write(0,3,"annual compensation",bold)
    worksheet.write(1,0,main_guy["name"])
    worksheet.write(1,1,main_guy["title"])
    worksheet.write(1,2,main_guy["age"])
    worksheet.write(1,3,main_guy["annual_comp"])
    worksheet.write(3,0,"name",bold)
    worksheet.write(3,1,"company",bold)
    worksheet.write(3,2,"board affiliations",bold)
    row = 4
    old_row = row
    for name,info in peopleDict.items():
        worksheet.write(row,0,name)
        col = 1
        for k,v in info.items():
            if len(v) > 1:
                old_row = 0 + row
                for i in v:
                    worksheet.write(row,col,i)
                    row+= 1
            else:
                worksheet.write(row,col,v[0])
            col += 1
        #row = old_row
        row += 1
        #print(info)

    #print(peopleDict)
    #print(main_guy)
    workbook.close()

main()

print("--- %s seconds ---" % (time.time() - start_time))