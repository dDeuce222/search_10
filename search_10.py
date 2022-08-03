from datetime import datetime
from playwright.sync_api import sync_playwright
import time
import pandas as pd

def search(keywords , link):
    save_path = input("input save_file_path : ")
    start = True
    start_img = True
    with sync_playwright() as p:
        browser = p.chromium.launch()
        context = browser.new_context()
        page = context.new_page()
        engine = 'xlsxwriter'
        mode = 'w'
        sheet_exists = 'error'
        for keyword in keywords:
            keyword = keyword.strip("\"")
            for i in range(1,10):
                results =[]
                search_link = link.format(KeyWord = keyword ,Type='general', PageNo = i)
                page.goto(search_link,timeout=0)
                time.sleep(2)
                articles = page.query_selector_all('article')
                for article in articles:
                    h3 = article.query_selector('h3')
                    a_tag = h3.query_selector('a')
                    result_name = a_tag.inner_text()
                    result_link = a_tag.get_attribute('href')
                    result = {'Keyword' : keyword ,'Title' : result_name ,'Url' : result_link}
                    results.append(result)
                if(mode == 'w'):
                    writer = pd.ExcelWriter(save_path, engine=engine,mode=mode)
                else:
                    writer = pd.ExcelWriter(save_path, engine=engine,mode=mode,if_sheet_exists=sheet_exists)
                df = pd.DataFrame(results)
                if(start):
                    need_header = True
                    start = False
                    row = 0
                else:
                    need_header = False
                    row = writer.sheets['General'].max_row
                df.to_excel(writer,sheet_name='General',startrow=row,header=need_header)
                writer.save()
                search_link = link.format(KeyWord = keyword , PageNo = i , Type = "images")
                page.goto(search_link,timeout=0)
                time.sleep(2)
                results = []
                articles = page.query_selector_all('article')
                for article in articles:
                    a_tag = article.query_selector('a')
                    result_name = a_tag.inner_text()
                    result_link = a_tag.get_attribute('href')
                    result = {'Keyword' : keyword ,'Title' : result_name ,'Url': result_link}
                    results.append(result)
                engine = 'openpyxl'
                mode = 'a'
                sheet_exists = 'overlay'
                writer = pd.ExcelWriter(save_path, engine=engine,mode=mode,if_sheet_exists = sheet_exists)
                df = pd.DataFrame(results)
                if(start_img):
                    need_header = True
                    start_img = False
                    row = 0
                else:
                    need_header = False
                    row = writer.sheets['Images'].max_row
                df.to_excel(writer,sheet_name='Images',startrow=row,index=False,header=need_header)
                writer.save()
                
            
            
def upload(file_name):
    keywords = []
    with open(file_name) as f:
        lines = f.readlines()
        for line in lines:
            keywords.append(line.strip('\n'))
        f.close()
    return keywords

def get_input():
    keywords = []
    while(True):
        keyword = input('insert keyword : ')
        if(keyword.lower() == 'quit'):
            break
        else:
            keywords.append(keyword)
    return keywords

def main(link):
    upload_type = input("Please select method of keywords input \n 1 : input manually \n 2: upload from file \n")
    if(upload_type == '1'):
        print('Input Quit to finish inputing\n')
        keywords = get_input()
    elif(upload_type == '2'):
        file_name = input("Please input file path : ")
        keywords = upload(file_name)
    else:
        print('Please insert valid selection')
        main(link)
    results = search(keywords,link)
    return results
    

link = """https://irsearch.live/search?q={KeyWord}&category_{Type}=1&pageno={PageNo}&language=en-US&time_range=None&safesearch=0&theme=simple"""
start_time = datetime.now()
main(link)