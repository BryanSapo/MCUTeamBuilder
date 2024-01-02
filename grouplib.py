import bs4
import requests
import re
from random import shuffle
from Student import Student
import xlsxwriter

def group(original_list,n):
    grouped_list = [original_list[i:i+n] for i in range(0, len(original_list), n)]
    return grouped_list

def webSearch(class_,subject,n):
    url=f'https://www.mcu.edu.tw/student/new-query/sel-6-4.asp?text0={class_}&text3={subject}&ch=1'

    students=[]
    final_result={}
    data=requests.get(url)
    data.encoding=data.apparent_encoding
    # print(data.text)

    soup = bs4.BeautifulSoup(data.content, 'html5lib')
    tables = soup.find_all('table')

    for table in tables:
        trs = table.find_all('tr')
        for td in trs:
            # print(td.get_text())
            text=td.get_text()
            id = re.findall(r'\d{8}', text)
            name = re.findall(r'([^\d]+)選修', text)
            # print(id,name)
            if len(name)==len(id):
                for index in range(len(name)):
                    s=Student(id[index],name[index])
                    students.append(s)
    total=len(students)

    shuffle(students)
    after=group(students,n)
    groupNumber=1
    for i in after:
        id=[]
        name=[]
        for j in i:
            id.append(j.id)
            name.append(j.name)
        # print(j.id,j.name,end=' ',sep='')
            no={'no':groupNumber,"data":{'Name':name,'Id':id}}
        final_result.update({groupNumber:no})
        groupNumber+=1
        
    return final_result

def logExcel(final_result):
    workbook = xlsxwriter.Workbook('result.xlsx')
 
    # The workbook object is then used to add new 
    # worksheet via the add_worksheet() method.
    worksheet = workbook.add_worksheet()
    
    # Use the worksheet object to write
    # data via the write() method.
    worksheet.write('A1', '組別')
    worksheet.write('B1', '學號')
    worksheet.write('C1', '姓名')
    row=2
    for i in final_result:
        # print(final_result[i])
        k=0
        for j in range(len(final_result[i]['data']['Id'])):
            worksheet.write(f'A{row}', final_result[i]['no'])
            worksheet.write(f'B{row}', final_result[i]['data']['Id'][j])
            worksheet.write(f'C{row}', final_result[i]['data']['Name'][j])
            row+=1
    workbook.close()


if __name__=='__main__':
    subject=input("請輸入科目代號: ")
    class_=input("請輸入班及代號: ")
    n=int(input("請輸入一組多少人: "))
    final_result=webSearch(class_,subject,n)
    logExcel(final_result)