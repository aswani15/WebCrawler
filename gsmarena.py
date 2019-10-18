import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys
import pandas as pd
import sys,urllib2
from BeautifulSoup import BeautifulSoup 
import numpy as np
import random
import xlsxwriter
import collections 
import json

def getbrandurl(browser1,brands_df):
    rel_indx_all = 1
    response2=browser1.open('https://www.gsmarena.com/',timeout=(10.0))
    sleep_time1 = random.randint(2,20)
    time.sleep(sleep_time1) 
    html2 = response2.read()
    bsObj2 = BeautifulSoup(html2)  
    try:
            brands = BeautifulSoup(str(bsObj2.findAll("div",attrs={"class":"brandmenu-v2 light l-box clearfix"})))
            brands_list = brands.findAll('li')
            for brands in brands_list:   
                brands_df.at[rel_indx_all,'Brand Name'] = brands.findAll('a')[0].text
                brands_df.at[rel_indx_all,'url'] = 'https://www.gsmarena.com/'+ brands.findAll('a')[0]['href']
                rel_indx_all=rel_indx_all+1  
    except Exception as e:
            error_type, error_obj, error_info = sys.exc_info()
            print 'error in getting brand URL'
            print error_type, 'Line:', error_info.tb_lineno
            continue

    return brands_df

def getphoneurl(phones_df,browser,brands_df):
     browser1=urllib2.build_opener()
     browser1.addheaders=[('User-agent', 'Mozilla/5.0')]
     i=1
     sleep_time1 = random.randint(10,12)
     for url_idx in xrange(12,len(brands_df)+1): 
         browser.get(brands_df.loc[url_idx,'url'])
         while True:
            time.sleep(sleep_time1) 
            print browser.current_url
            response2 = browser1.open(browser.current_url)
            
            html2 = response2.read()
            bsObj2 = BeautifulSoup(html2)
            phones = BeautifulSoup(str(bsObj2.findAll("div",attrs={"id":"review-body"})))
            phones_list = phones.findAll('li')
            for phs in phones_list:   
                phones_df.at[i,'Brand Name'] = brands_df.loc[url_idx,'Brand Name'] 
                phones_df.at[i,'Phone Name'] = phs.findAll('a')[0].text
                phones_df.at[i,'ph_url'] = 'https://www.gsmarena.com/'+ phs.findAll('a')[0]['href']
                i=i+1  
            try:
                elm = browser.find_element_by_xpath('//*[contains(concat( " ", @class, " " ), concat( " ", "pages-next", " " ))]')            
                if 'disabled pages-next' in elm.get_attribute('class'):
                    break;
                elm.click()
            except Exception as e:
                break;     
     return  phones_df

def nested_dict():
    return collections.defaultdict(nested_dict)

brands_df = pd.DataFrame(columns = ['Brand Name','url'])
phones_df = pd.DataFrame(columns = ['Brand Name','Phone Name','ph_url'])

browser1=urllib2.build_opener()
browser1.addheaders=[('User-agent', 'Mozilla/5.0')]

#commenting orginal code################################################################


brands_df = getbrandurl(browser1,brands_df)

browser = webdriver.Firefox(executable_path='A:\geckodriver-v0.19.0-win64\geckodriver.exe')

phones_df = getphoneurl()

#
writer_all_Data = pd.ExcelWriter('phones1.xlsx', engine='xlsxwriter')
phones_df.to_excel(writer_all_Data,'Sheet1')
writer_all_Data.save()

writer_all_Data1 = pd.ExcelWriter('Brand1.xlsx', engine='xlsxwriter')
brands_df.to_excel(writer_all_Data1,'Sheet1')
writer_all_Data1.save()

##################################################################
##run from here###
## get files from email change the path
phones_df = pd.read_excel('phones1.xlsx', sheetname='Sheet1')
brands_df = pd.read_excel('Brand.xlsx', sheetname='Sheet1')
df= pd.read_csv('data.csv')
df.columns
phones_specs_dict = nested_dict()
innerdict = nested_dict()

for p in xrange(1,len(phones_df)+1):
     response1 = browser1.open(phones_df.loc[p,'ph_url'])
     brand_name = phones_df.loc[p,'Brand Name']
     ph_name = phones_df.loc[p,'Phone Name']
     html1 = response1.read()
     bsObj1 = BeautifulSoup(html1)
     table = bsObj1.findAll("table")
     for header in table:
         header_val = header.find("th").text
         tr_all = header.findAll('tr') 
         #,attrs={"class":"ttl"})
         for tr_each in tr_all:
             if len(tr_each)> 3:
                td_key=tr_each.find('td',attrs={"class":"ttl"}).text
                td_val=BeautifulSoup(str(tr_each)).find('td',attrs={"class":"nfo"}).text
                phones_specs_dict[ph_name] = innerdict 
                innerdict[header_val][td_key] = td_val

with open('data.json', 'w') as outfile:
    json.dump(phones_specs_dict, outfile)

