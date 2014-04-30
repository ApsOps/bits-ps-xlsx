import xlsxwriter
import requests
from bs4 import BeautifulSoup
import time

timeout = 60*10  #sec
global row
global worksheet
global workbook

def scan():
    global workbook
    url = "http://psd/problembankisem1415_Disp2.asp"
    cookies = dict(COOKIE_NAME='COOKIE_VALUE')  # Put your ASP.net Session cookie here
    payload = {'Discipline': 'ALL'}
    while True:
        try:
            r = requests.post(url, cookies=cookies, data=payload)
            break
        except:
            print "Network Error"
            time.sleep(10)
            
    html = BeautifulSoup(r.text)
    count = 0
    if len(html.find_all('a')) > 2:
        create_workbook()
        for a in html.find_all('a'):
            url = a.get('href')
            name = a.text
            if count >= 4:
                get_data(name, "http://psd/" + url)
            count += 1
        workbook.close()
    print "no. of stations = " + str(count-4)


def create_workbook():
    global workbook
    global worksheet
    workbook = xlsxwriter.Workbook('./PS_list[SIDS].xlsx')
    worksheet = workbook.add_worksheet()

    bold = workbook.add_format({'bold':1})

    worksheet.set_column('A:D', 30)
    worksheet.set_column('E:I', 20)

    worksheet.write('A1', 'Station Name', bold)
    worksheet.write('B1', 'Address', bold)
    worksheet.write('C1', 'Stipend(Rupees per Month)', bold)
    worksheet.write('D1', '**Preferred Discipline(s)', bold)
    worksheet.write('E1', 'Accomodation', bold)
    worksheet.write('F1', 'Food', bold)
    worksheet.write('G1', 'Transport', bold)
    worksheet.write('H1', 'Website', bold)
    worksheet.write('I1', 'Other Benefits', bold)
    
    global row
    row = 2

    print "file created"


def get_data(name, url):
    while True:
        try:
            r = requests.get(url)
            break
        except:
            print "Network Error"
            time.sleep(10)
    
    html = BeautifulSoup(r.text)
    tables = html.find_all('table')
    details = [name]
    try:
        if tables[1].find('td').text.strip() == "1":
            write_to_file(details)
            return
        for e in tables[1].find_all('td'):
            details.append(' '.join(e.text.split()))
        write_to_file(details)
    except:
        write_to_file(details)


def write_to_file(details):
    global row
    global worksheet
    for (column, e) in enumerate(details):
        worksheet.write(row, column, e)
    row += 1


def create_html():
    content = """
<html>
<head>
<title>PS Stations Excel file - SIDS</title>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<style>
body {
width: 700px;
margin: 50px auto;
text-align: center;
}
</style>
</head>
<body>

<h1>PS 2 Stations list in XLSX format</h1>

<div>

<p>Here is the link to the excel file for PS 2 stations with the details.</p>
<h2><a href="PS_list[SIDS].xlsx">PS_list[SIDS].xlsx</a></h2>

<p>This file is auto updated every 10 min. Last updated: 
    """
    f = open("./index.html", "w")
    f.write(content)
    f.write(time.strftime('%I:%M %p, %b %d, %Y'))
    f.write("<br>For any problems/suggestions, mail me @ <a href='mailto:aps.sids@gmail.com'>aps.sids@gmail.com</a></p><p>Interested folks can find the source code on <a href='https://github.com/aps-sids/bits-ps-xlsx'>Github</a> and follow me on <a href='https://twitter.com/aps_sids'>Twitter</a>.</p></div></body></html>")
    f.close()


while True:

    print "starting scan.."
    scan()

    create_html()
    print "html created"

    time.sleep(timeout)
