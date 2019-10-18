# -*- coding: utf-8 -*-
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

import pandas as pd
import sys,urllib2
from BeautifulSoup import BeautifulSoup 
import numpy as np
import random
import xlsxwriter
import re
import uuid


#def my_random_string(string_length=10):
#    """Returns a random string of length string_length."""
#    random = str(uuid.uuid4()) # Convert UUID format to a Python string.
#    random = random.upper() # Make all characters uppercase.
#    random = random.replace("-","") # Remove the UUID '-'.
#    return random[0:string_length]
#
#class BitDeal:
#    meta = 'tvg'
#    def __inPartial_Industrials__(self, deal_ID,hotel_name, hotel_ota, hotel_price, has_breakfast, tvg_hotel_id, hotel_category):
#        self.dealID = deal_ID
#        self.hname = hotel_name
#        self.hota = hotel_ota
#        self.hprice = hotel_price
#        self.htvgid = tvg_hotel_id
#        self.hbreakfast = has_breakfast
#        self.hcat = hotel_category

browser = webdriver.Firefox(executable_path='A:\geckodriver-v0.19.0-win64\geckodriver.exe')
time.sleep(15)
mouse = webdriver.ActionChains(browser)

data_pd_manta = pd.DataFrame( columns = ['Company','Type','UnqId','Missing'])
url_df_manta = pd.DataFrame( columns = ['Company','URL','UnqId','Missing'])
url_df_mod_manta = pd.DataFrame( columns = ['Company','URL','UnqId','Missing'])
final_url_manta_df = pd.DataFrame(columns = ['URL', 'Company','UnqId','Type'])
companies_manta_df = pd.DataFrame( columns = ['Company','URL','UnqId','Type','Missing'])
companies_manta_mod_df = pd.DataFrame( columns = ['Company','URL','UnqId','Type','Missing'])
people_manta_df = pd.DataFrame( columns = ['Person_Name','Title','UnqId','Company_Name','Person_Id'])

company_attrib_manta_df = pd.DataFrame( columns = ['Company','Revenue_Text','Revenue',
                            'Employee_Text','Employee_Number','UnqId','Industry','Type','Manta_Name'])


#company_attrib_manta_df = pd.read_excel('A:/Capstone_Code/files/Manta_Partial_Industrials_Company_Attribs_Data.xlsx', sheetname='Sheet1')
companies_manta_df = pd.read_excel('A:/Capstone_Code/files/Manta_Industrail_Partial_Company_Urls.xlsx', sheetname='Sheet1')
company_attrib_manta_df= pd.read_excel('A:/Capstone_Code/files/PartialDataDelete.xlsx', sheetname='Sheet1')

browser1=urllib2.build_opener()
browser1.addheaders=[('User-agent', 'Mozilla/5.0')]
comp_indx = 1
#companies_manta_df = companies_manta_df[companies_manta_df.Type <> 'Buyer']
companies_manta_df.reset_index()
companies_manta_df.index = np.arange(1, len(companies_manta_df) + 1)
comp_indx=1

#companies_manta_df.loc[companies_manta_df['Company']=='Foundation Partial_Industrials, Inc.'].index.values.astype(int)[0]

for comp_indx in xrange(1,len(companies_manta_df)+1):
 sleep_time = random.randint(20,60)
 random_number = random.randint(5,12)
 time.sleep(5)
 try:    
     time.sleep(random_number+10)     
     url_manta = companies_manta_df.loc[comp_indx,'URL']    
     browser.get(url_manta)
     time.sleep(sleep_time)
     wait = WebDriverWait(browser, 30)     
    #browser.get('https://www.manta.com/c/mm811bq/fitz-and-floyd-enterprises-llc')
     if bool(random.getrandbits(1)):
         if comp_indx%random_number==0:
          box_text = wait.until( EC.presence_of_element_located(
                           (By.CSS_SELECTOR, ".text-primary")))        
          point = browser.find_element_by_css_selector(".text-primary") 
          time.sleep(5) 
          mouse.move_to_element(point).perform()
         if comp_indx%random_number==3:
          box_text = wait.until( EC.presence_of_element_located(
                           (By.CSS_SELECTOR, ".icon-twitter")))           
          browser.find_element_by_css_selector(".icon-twitter").click()               
     page_source_web_manta = browser.page_source
     time.sleep(10)
     print 'inside try'
     bsObj_manta = BeautifulSoup(page_source_web_manta)
 except Exception as e:
     error_type, error_obj, error_info = sys.exc_info()
     print 'ERROR FOR URL:',url_manta
     print error_type, 'Line:', error_info.tb_lineno
     writer_all_manta_data = pd.ExcelWriter('A:/Capstone_Code/files/Manta_Partial_Industrials_Company_Attribs_Data.xlsx', engine='xlsxwriter')
     if len(company_attrib_manta_df)>0:
      company_attrib_manta_df.to_excel(writer_all_manta_data,'Sheet1')
     writer_all_manta_data.save()
     continue   
     
 revenue_indx = page_source_web_manta.find('annual revenue of')
 if revenue_indx > 1:
         revenue_value_uncleaned = page_source_web_manta[revenue_indx+17:revenue_indx+30]
         company_attrib_manta_df.at[comp_indx,'Revenue_Text'] = revenue_value_uncleaned
         revenue_cleaned = (re.findall(r'\d+', revenue_value_uncleaned) if len(revenue_value_uncleaned)>0 else '')
         if len(revenue_cleaned) > 1:
            company_attrib_manta_df.at[comp_indx,'Revenue'] = (float(revenue_cleaned[0])+float(revenue_cleaned[1]))/2
         else:
             if len(revenue_cleaned)==1:
                company_attrib_manta_df.at[comp_indx,'Revenue'] = revenue_cleaned[0]
 else:
     if str(bsObj_manta.findAll("td",{"rel":"annualRevenue"}))<> 'None':
        revenue_cleaned = bsObj_manta.findAll("td",{"rel":"annualRevenue"}) 
        if len(revenue_cleaned) > 1:
            company_attrib_manta_df.at[comp_indx,'Revenue'] = (float(revenue_cleaned[0].text)+float(revenue_cleaned[1].text))/2
        else:
             if len(revenue_cleaned)==1:
                company_attrib_manta_df.at[comp_indx,'Revenue'] = revenue_cleaned[0].text

              
 employee_indx= page_source_web_manta.find('employs a staff of approximately')
 if employee_indx > 1:
         employee_value_uncleaned = page_source_web_manta[employee_indx+19:employee_indx+50]
         company_attrib_manta_df.at[comp_indx,'Employee_Text'] = employee_value_uncleaned
         employee_cleaned = (re.findall(r'\d+', employee_value_uncleaned) if len(employee_value_uncleaned)>0 else '')
         if len(employee_cleaned) > 1:
            company_attrib_manta_df.at[comp_indx,'Employee_Number'] = (float(employee_cleaned[0])+float(employee_cleaned[1]))/2
         else:
             if len(employee_cleaned)==1:
                company_attrib_manta_df.at[comp_indx,'employee_cleaned'] = employee_cleaned[0]
 company_attrib_manta_df.at[comp_indx,'Company'] = companies_manta_df.loc[comp_indx,'Company']
 company_attrib_manta_df.at[comp_indx,'UnqId'] = companies_manta_df.loc[comp_indx,'UnqId']
 company_attrib_manta_df.at[comp_indx,'Type'] = companies_manta_df.loc[comp_indx,'Type'] 
 manta_company = bsObj_manta.find("span",attrs={"class":"text-primary visible-xs h2"})
 company_attrib_manta_df.at[comp_indx,'Manta_Name'] =(manta_company.text if str(manta_company) <> 'None' else '') 
 print companies_manta_df.loc[comp_indx,'Company']
 print companies_manta_df.loc[comp_indx,'UnqId']
 unq_id = companies_manta_df.loc[comp_indx,'UnqId']
 total_sleep = random.randint(200,400)
 total_sleep_1 = random.randint(100,200)

 if unq_id==450 or unq_id==480 or unq_id==530 or unq_id==550 or unq_id==590 or unq_id==620:
    time.sleep(total_sleep)
    browser.get("https://in.yahoo.com/")
 if unq_id==680 or unq_id==635 or unq_id==655 or unq_id==689 or unq_id==700 or unq_id==720:
    time.sleep(total_sleep_1) 
    browser.get("https://www.google.com")            
         
 try:
       if sleep_time ==30 or sleep_time==46 or sleep_time==62:
         browser.find_element_by_css_selector(".icon-facebook").click()
         #time.sleep(sleep_time)
         #child_window = browser.current_window_handle         
         #browser.find_element_by_tag_name('body').send_keys(Keys.CONTROL + 'w')
  
 except Exception as e:
       error_type, error_obj, error_info = sys.exc_info()
       print 'Random Click error'
       print error_type, 'Line:', error_info.tb_lineno
       writer_all_manta_data = pd.ExcelWriter('A:/Capstone_Code/files/Manta_Partial_Industrials_Company_Attribs_Data.xlsx', engine='xlsxwriter')
       if len(company_attrib_manta_df)>1:
        company_attrib_manta_df.to_excel(writer_all_manta_data,'Sheet1')
        writer_all_manta_data.save()
       continue 
 browser.find_element_by_css_selector(".text-primary")      
 
writer_all_manta_data = pd.ExcelWriter('A:/Capstone_Code/files/Manta_Partial_Industrials_Company_Attribs_Data.xlsx', engine='xlsxwriter')
company_attrib_manta_df.to_excel(writer_all_manta_data,'Sheet1')
writer_all_manta_data.save()
       
browser.quit()

