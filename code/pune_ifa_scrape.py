from urllib.request import Request, urlopen
from bs4 import BeautifulSoup
import openpyxl
import os

def deconstruct_td_data(data):
    lst = []
    for info in data.findAll("td"):
        lst.append(info.text)
    return lst

def create_excel_sheet(data):
    wb = openpyxl.Workbook()
    ws = wb.active
    for  row_data in data:
        for col_idx, col_val in enumerate(data[row_data]):
           ws.cell(row=row_data + 1,column=col_idx + 1,value=col_val)
    script_dir = os.path.dirname(os.path.dirname(os.path.realpath(__file__)))
    output_dir = os.path.join(script_dir,"output")
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    output_file = os.path.join(output_dir,"Pune_IFA_Data.xlsx")
    wb.save(output_file)
    

def retrieve_data_from_scraped_obj(scrapedObj):
    tr_arr = scrapedObj.find("div",{'data-id':"2458664"}).find_next('div').find('table').find('tbody').findAll('tr')[1:]
    data_list = {}
    for i,data in enumerate(tr_arr):
        data_list[i] = deconstruct_td_data(data)
    return data_list

def main():
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'}
    url = "https://pifaa.org/pifaa-members/"
    req = Request(url,headers=headers)
    html = urlopen(req)
    bsObj = BeautifulSoup(html,'lxml')
    retrieved_data = retrieve_data_from_scraped_obj(bsObj)
    create_excel_sheet(retrieved_data)
    
    
if __name__ == "__main__":
    main()