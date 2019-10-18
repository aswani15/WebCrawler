# -*- coding: utf-8 -*-
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys
import pandas as pd
#import sys,urllib2
#from BeautifulSoup import BeautifulSoup 
import numpy as np
import random
import xlsxwriter
import re

browser = webdriver.Firefox(executable_path='A:\geckodriver-v0.19.0-win64\geckodriver.exe')
time.sleep(15)

##crunchbase
data_pd_crunchbase = pd.DataFrame( columns = ['Company','Type','UnqId','Missing'])
url_df_crunchbase = pd.DataFrame( columns = ['Company','URL','UnqId','Missing'])
url_df_mod_crunchbase = pd.DataFrame( columns = ['Company','URL','UnqId','Missing'])
final_url_crunchbase_df = pd.DataFrame(columns = ['URL', 'Company','UnqId','Type'])
companies_crunchbase_df = pd.DataFrame( columns = ['Company','URL','UnqId','Type','Missing'])

##for manta
data_pd_manta = pd.DataFrame( columns = ['Company','Type','UnqId','Missing'])
url_df_manta = pd.DataFrame( columns = ['Company','URL','UnqId','Missing'])
url_df_mod_manta = pd.DataFrame( columns = ['Company','URL','UnqId','Missing'])
final_url_manta_df = pd.DataFrame(columns = ['URL', 'Company','UnqId','Type'])
companies_manta_df = pd.DataFrame( columns = ['Company','URL','UnqId','Type','Missing'])



df_excel = pd.read_excel('A:/Capstone_Code/Consumer Discretionary.xls', sheetname='Screening',skiprows=7)
comptrs_df = pd.read_excel('A:/Capstone_Code/files/Bloomberg_CD_CompanyAttributes.xlsx', sheetname='Sheet1')

index_data =1
for i in xrange(0,len(df_excel)):
    
    company_name = df_excel.loc[i,'Target/Issuer']
    index= company_name.find(':')
    if index >0:
       data_pd_crunchbase.at[index_data,'Company'] = company_name[0:index-1]
       data_pd_manta.at[index_data,'Company'] = company_name[0:index-1]
    else:
       data_pd_crunchbase.at[index_data,'Company'] = df_excel.loc[i,'Target/Issuer']
       #data_pd_manta.at[index_data,'Company'] = df_excel.loc[i,'Target Company']
    data_pd_crunchbase.at[index_data,'Type'] = 'Target'
    #data_pd_manta.at[index_data,'Type'] = 'Target'    
    index_data=index_data+1
    
i_indx =len(data_pd_crunchbase)+1

for j in xrange(1,len(comptrs_df)+1):  
  if comptrs_df.loc[j,'Type'] <> 'Buyer':     
    data_pd_crunchbase.at[i_indx,'Company']= comptrs_df.loc[j,'Company']
    data_pd_crunchbase.at[i_indx,'Type'] = comptrs_df.loc[j,'Type']
    #data_pd_manta.at[i_indx,'Company']= comptrs_df.loc[j,'Company']
    #data_pd_manta.at[i_indx,'Type'] = comptrs_df.loc[j,'Type']
    i_indx=i_indx+1
#drop duplicates
data_final_crunchbase_pd = data_pd_crunchbase.drop_duplicates()
data_final_manta_pd = data_pd_manta.drop_duplicates()


##reset index after removing duplicates
data_final_crunchbase_pd.reset_index()
data_final_crunchbase_pd.index = np.arange(1, len(data_final_crunchbase_pd) + 1)

data_final_manta_pd.reset_index()
data_final_manta_pd.index = np.arange(1, len(data_final_manta_pd) + 1)


browser.get('http://www.yahoo.com')
comp_names_indx=1
p_indx_url_exists = 1

for comp_names_indx in xrange(1,len(data_final_crunchbase_pd)+1): 
 sleep_time = random.randint(10,20)
 company_name = data_final_crunchbase_pd.loc[comp_names_indx,'Company']
 comp_type = data_final_crunchbase_pd.loc[comp_names_indx,'Type']
 ####manta scrape block
# try:
#    wait = WebDriverWait(browser, 30)
#    box = wait.until(EC.presence_of_element_located(
#            (By.NAME, "p"))) 
#    elem = browser.find_element_by_name('p')  # Find the search box
#    elem.clear() 
#    elem = browser.find_element_by_name('p')   
#    browser.wait = WebDriverWait(browser, 10)
#    elem.send_keys(company_name+'+manta' + Keys.RETURN)   
#    time.sleep(sleep_time)          
# except TimeoutException:
#    print("Yahoo not loaded")
#    writer_manata_all = pd.ExcelWriter('A:/Capstone_Code/files/Manta_CD_All_Urls.xlsx', engine='xlsxwriter')
#    final_url_manta_df.to_excel(writer_manata_all,'Sheet1')
#    writer_manata_all.save()
#
#    writer_manata_company_url = pd.ExcelWriter('A:/Capstone_Code/files/Manta_CD_Company_Urls.xlsx', engine='xlsxwriter')
#    companies_manta_df.to_excel(writer_manata_company_url,'Sheet1')
#    writer_manata_company_url.save()
#    continue

#assert 'Yahoo' in browser.title

 
 i=1
 j_all_urls=1
 url_df_manta = url_df_manta[0:0]
 url_df_mod_manta = url_df_mod_manta[0:0]
 website_desc = browser.find_elements_by_class_name("td-u")
 search_list=[]
 #for url in browser.find_elements_by_class_name("wr-bw"):    
 #   if "www.manta.com" in url.text: 
 #       company_name_search = company_name.strip()
 #       #company_name_search = (company_name.lower().replace('llc','') if company_name.find('llc') else company_name)
 #       #company_name_search = (company_name.lower().replace('inc','') if company_name.find('inc') else company_name)
 #       company_name_search = (company_name_search.lower().replace(',','') if company_name_search.find(',') else company_name_search)
 #       search_list = company_name_search.lower().split(' ')
 #       #compidentify_regex.search(website_desc[j_all_urls].text)
 #       description_url = website_desc[j_all_urls-1].text
 #       if len(search_list)>1:
 #        if len(search_list)>0 and description_url.find(search_list[0]) and description_url.find(search_list[1]):
 #          url_df_manta.at[i,'URL']=  url.text
 #          url_df_manta.at[i,'UnqId']=  comp_names_indx
 #          url_df_manta.at[i,'Company']=  company_name        
 #          i=i+1
 #       else:
 #         if len(search_list)==1 and description_url.find(search_list[0]):
 #             url_df_manta.at[i,'URL']=  url.text
 #             url_df_manta.at[i,'UnqId']=  comp_names_indx
 #             url_df_manta.at[i,'Company']=  company_name  
 #             i=i+1
 #   url_df_mod_manta.at[j_all_urls,'URL']=  url.text
 #   url_df_mod_manta.at[j_all_urls,'UnqId']=  comp_names_indx
 #   url_df_mod_manta.at[j_all_urls,'Company']=  company_name 
 #   j_all_urls = j_all_urls+1
 #           
 #if len(url_df_manta) >0:       
 #   companies_manta_df.at[p_indx_url_exists,'URL']=  "https://"+url_df_manta.loc[1,'URL']
 #   companies_manta_df.at[p_indx_url_exists,'Company']=  url_df_manta.loc[1,'Company']
 #   companies_manta_df.at[p_indx_url_exists,'UnqId']=  url_df_manta.loc[1,'UnqId']
 #   companies_manta_df.at[p_indx_url_exists,'Type']= comp_type
 #   data_final_manta_pd.at[comp_names_indx,'Missing']='N' 
 #   p_indx_url_exists= p_indx_url_exists+1
 #else :
 #   data_final_manta_pd.at[comp_names_indx,'Missing']='Y'
 #final_url_manta_df = final_url_manta_df.append(url_df_mod_manta)
 #
###crunchbase scrape block
 try:
    wait = WebDriverWait(browser, 30)
    box = wait.until(EC.presence_of_element_located(
            (By.NAME, "p"))) 
    elem = browser.find_element_by_name('p')  # Find the search box
    elem.clear() 
    elem = browser.find_element_by_name('p')  
    browser.wait = WebDriverWait(browser, 10)
    elem.send_keys(company_name+'+crunchbase' + Keys.RETURN)   
    time.sleep(sleep_time)
            
 except Exception as e:
    print("Yahoo not loaded")
    print e
    writer_crunchbase_all = pd.ExcelWriter('A:/Capstone_Code/files/Crunchbase_CD_All_Urls.xlsx', engine='xlsxwriter')
    final_url_crunchbase_df.to_excel(writer_crunchbase_all,'Sheet1')
    writer_crunchbase_all.save()

    writer_crunchbase_company_url = pd.ExcelWriter('A:/Capstone_Code/files/Crunchbase_CD_Company_Urls.xlsx', engine='xlsxwriter')
    companies_crunchbase_df.to_excel(writer_crunchbase_company_url,'Sheet1')
    writer_crunchbase_company_url.save()
    continue

#assert 'Yahoo' in browser.title

 i=1
 j_all_urls=1
 url_df_crunchbase = url_df_crunchbase[0:0]
 url_df_mod_crunchbase = url_df_mod_crunchbase[0:0]
 website_desc = browser.find_elements_by_class_name("td-u")
 search_list=[]
 for url in browser.find_elements_by_class_name("wr-bw"):    
    if "www.crunchbase.com" in url.text: 
        company_name_search = company_name.strip()
        #company_name_search = (company_name.lower().replace('llc','') if company_name.find('llc') else company_name)
        #company_name_search = (company_name.lower().replace('inc','') if company_name.find('inc') else company_name)
        company_name_search = (company_name_search.lower().replace(',','') if company_name_search.find(',') else company_name_search)
        search_list = company_name_search.lower().split(' ')
        #compidentify_regex.search(website_desc[j_all_urls].text)
        description_url = (website_desc[j_all_urls-1].text if len(website_desc)>=j_all_urls else '')
        if len(search_list)>1:
         if len(search_list)>0 and description_url.find(search_list[0]) and description_url.find(search_list[1]):
           url_df_crunchbase.at[i,'URL']=  url.text
           url_df_crunchbase.at[i,'UnqId']=  comp_names_indx
           url_df_crunchbase.at[i,'Company']=  company_name        
           i=i+1
        else:
          if len(search_list)==1 and description_url.find(search_list[0]):
              url_df_crunchbase.at[i,'URL']=  url.text
              url_df_crunchbase.at[i,'UnqId']=  comp_names_indx
              url_df_crunchbase.at[i,'Company']=  company_name  
              i=i+1
    url_df_mod_crunchbase.at[j_all_urls,'URL']=  url.text
    url_df_mod_crunchbase.at[j_all_urls,'UnqId']=  comp_names_indx
    url_df_mod_crunchbase.at[j_all_urls,'Company']=  company_name 
    j_all_urls = j_all_urls+1
            
 if len(url_df_crunchbase) >0:       
    companies_crunchbase_df.at[p_indx_url_exists,'URL']=  "https://"+url_df_crunchbase.loc[1,'URL']
    companies_crunchbase_df.at[p_indx_url_exists,'Company']=  url_df_crunchbase.loc[1,'Company']
    companies_crunchbase_df.at[p_indx_url_exists,'UnqId']=  url_df_crunchbase.loc[1,'UnqId']
    companies_crunchbase_df.at[p_indx_url_exists,'Type']= comp_type
    data_final_crunchbase_pd.at[comp_names_indx,'Missing']='N' 
    p_indx_url_exists= p_indx_url_exists+1
 else :
    data_final_crunchbase_pd.at[comp_names_indx,'Missing']='Y'
 final_url_crunchbase_df = final_url_crunchbase_df.append(url_df_mod_crunchbase)
 print company_name
 
companies_crunchbase_df.reset_index()
companies_crunchbase_df.index = np.arange(1, len(companies_crunchbase_df) + 1)
final_url_crunchbase_df.reset_index()
final_url_crunchbase_df.index = np.arange(1, len(final_url_crunchbase_df) + 1)
#final_url_manta_df.reset_index()
#final_url_manta_df.index = np.arange(1, len(final_url_manta_df) + 1)
#companies_manta_df.reset_index()
#companies_manta_df.index = np.arange(1, len(companies_manta_df) + 1) 
   
writer_manata_all = pd.ExcelWriter('A:/Capstone_Code/files/Crunchbase_CD_All_Urls.xlsx', engine='xlsxwriter')
final_url_crunchbase_df.to_excel(writer_manata_all,'Sheet1')
writer_manata_all.save()

writer_manata_company_url = pd.ExcelWriter('A:/Capstone_Code/files/Crunchbase_CD_Company_Urls.xlsx', engine='xlsxwriter')
companies_crunchbase_df.to_excel(writer_manata_company_url,'Sheet1')
writer_manata_company_url.save()

#writer_manata_all = pd.ExcelWriter('A:/Capstone_Code/files/Manta_CD_All_Urls.xlsx', engine='xlsxwriter')
#final_url_manta_df.to_excel(writer_manata_all,'Sheet1')
#writer_manata_all.save()
#
#writer_manata_company_url = pd.ExcelWriter('A:/Capstone_Code/files/Manta_CD_Company_Urls.xlsx', engine='xlsxwriter')
#companies_manta_df.to_excel(writer_manata_company_url,'Sheet1')
#writer_manata_company_url.save()

browser.quit()
