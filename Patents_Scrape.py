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
import glob
import os
import sys
reload(sys)

sys.setdefaultencoding('utf8')


def savefiles(pd_company_patent_info,fwd_citation_df,bckwd_citation_df):
    writer_patents = pd.ExcelWriter('A:/Capstone_Code/files/Patents_HealthCare_Data.xlsx', engine='xlsxwriter')
    pd_company_patent_info.to_excel(writer_patents,'Sheet1')
    writer_patents.save()

    writer_fwd_data = pd.ExcelWriter('A:/Capstone_Code/files/Forward_Citations_HealthCare.xlsx', engine='xlsxwriter')
    fwd_citation_df.to_excel(writer_fwd_data,'Sheet1')
    writer_fwd_data.save()

    writer_bckwd_data = pd.ExcelWriter('A:/Capstone_Code/files/Backward_Citations_HealthCare.xlsx', engine='xlsxwriter')
    bckwd_citation_df.to_excel(writer_bckwd_data,'Sheet1')
    writer_bckwd_data.save()


browser_patent_sl = webdriver.Firefox(executable_path='A:\geckodriver-v0.19.0-win64\geckodriver.exe')
time.sleep(15)
mouse = webdriver.ActionChains(browser_patent_sl)

data_final_company_pd = pd.DataFrame( columns = ['Company','Type','Patents'])
data_pd_comp = pd.DataFrame( columns = ['Company','Type'])
pd_company_patent_info = pd.DataFrame( columns = ['Company','UnqId','Patent_Number','Type','Num_Back_Citations','Num_Fwd_Citations','Priority_Date','Publication_Date','Assignee','Title','Researchers','Status','URL'])
bckwd_citation_df = pd.DataFrame( columns = ['Company','UnqId','Patent_Number','Backward_Citation','Examiner_Cited','Priority_Date','Publication_Date','Assignee','Title'])
fwd_citation_df = pd.DataFrame( columns = ['Company','UnqId','Patent_Number','Fwd_Citation','Examiner_Cited','Priority_Date','Publication_Date','Assignee','Title'])
companies_patent_df = pd.DataFrame(columns = ['id','title','assignee','inventor/author','priority date','filing/creation date','publication date','grant date','result link'])
pd_company_patent_info_final = pd.DataFrame( columns = ['Company','UnqId','Patent_Number','Type','Num_Back_Citations','Num_Fwd_Citations','Priority_Date','Publication_Date','Assignee','Title','Researchers','Status','URL'])
bckwd_citation_df_final = pd.DataFrame( columns = ['Company','UnqId','Patent_Number','Backward_Citation','Examiner_Cited','Priority_Date','Publication_Date','Assignee','Title'])
fwd_citation_df_final = pd.DataFrame( columns = ['Company','UnqId','Patent_Number','Fwd_Citation','Examiner_Cited','Priority_Date','Publication_Date','Assignee','Title'])

browser_patent=urllib2.build_opener()
browser_patent.addheaders=[('User-agent', 'Mozilla/5.0')]


df_excel = pd.read_excel('A:/Capstone_Code/Healthcare_20171016.xls')#, sheetname='Screening',skiprows=7)
comptrs_df = pd.read_excel('A:/Capstone_Code/files/Bloomberg_HealthCare_CompanyAttributes.xlsx', sheetname='Sheet1')
#fwd_citation_df= pd.read_excel('A:/Capstone_Code/files/Forward_Citations_HealthCare.xlsx', sheetname='Sheet1')
#bckwd_citation_df= pd.read_excel('A:/Capstone_Code/files/Backward_Citations_HealthCare.xlsx', sheetname='Sheet1')
#pd_company_patent_info= pd.read_excel('A:/Capstone_Code/files/Patents_HealthCare_Data.xlsx', sheetname='Sheet1')

index_data =1
for i in xrange(0,len(df_excel)):    
    company_name = df_excel.loc[i,'Target Company']
    data_pd_comp.at[index_data,'Company'] = company_name
    data_pd_comp.at[index_data,'Type'] = 'Target'
    index_data=index_data+1
    
#for i in xrange(0,len(df_excel)):    
#    company_name = df_excel.loc[i,'Bidder Company']
#    data_pd_comp.at[index_data,'Company'] = company_name
#    data_pd_comp.at[index_data,'Type'] = 'Buyer'
#    index_data=index_data+1    
    
i_indx =len(data_pd_comp)+1

for j in xrange(1,len(comptrs_df)+1):  
 if comptrs_df.loc[j,'Type'] <> 'Buyer':      
    data_pd_comp.at[i_indx,'Company']= comptrs_df.loc[j,'Company']
    data_pd_comp.at[i_indx,'Type'] = comptrs_df.loc[j,'Type']
    i_indx=i_indx+1
#drop duplicates
data_final_company_pd = data_pd_comp.drop_duplicates()


##reset index after removing duplicates
data_final_company_pd.reset_index()
data_final_company_pd.index = np.arange(1, len(data_final_company_pd) + 1)

repls = {',' : '', 'inc' : '','l.p':'','llc':'','corp':'','amp;':'','.': '','-': '','&': '','co.':'','company':''}

i_all_ptnt_inx = 1
i_comp_indx_patent=1

#data_final_company_pd.loc[data_final_company_pd['Company']=='PLx Pharma Inc.'].index.values.astype(int)[0]

for i_comp_indx_patent in xrange(664,len(data_final_company_pd)+1):
 sleep_time = random.randint(6,15)
 random_number = random.randint(5,12) 
 #reset pd df
 #pd_company_patent_info[0:0]
 #fwd_citation_df[0:0]
 #bckwd_citation_df[0:0] 

 company_name  =  data_final_company_pd.loc[i_comp_indx_patent,'Company'] 
 company_name_clean = reduce(lambda a, kv: a.replace(*kv), repls.iteritems(), company_name.lower().strip()).strip()
 company_name_clean = company_name_clean.replace(' ','+')
 url_assign = 'https://patents.google.com/?assignee='+company_name_clean

 browser_patent_sl.get(url_assign)
 time.sleep(sleep_time)
 path = "A:/patents/"

  ##delete file
 filelist_del = glob.glob(os.path.join(path, "*.csv"))
 for f_del in filelist_del:
    os.remove(f_del)

 try:
     wait = WebDriverWait(browser_patent_sl, 15)
     box = wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, "#count a")))
     browser_patent_sl.find_element_by_css_selector("#count a").click()  
       
 except Exception:
    data_final_company_pd.at[i_comp_indx_patent,'Patents'] = 'N' 
    print("Patent not found")
    savefiles(pd_company_patent_info,fwd_citation_df,bckwd_citation_df)
    continue           
 #if random_number ==6:
 #  browser_patent_sl.get('https://www.google.co.in')
 #if random_number ==9:
 #  browser_patent_sl.get('https://in.search.yahoo.com/')
 #if random_number ==11:
 #  browser_patent_sl.get('https://twitter.com/')
 #   
 companies_patent_df[0:0]
 time.sleep(random_number)
 all_files = glob.glob(os.path.join(path, "*.csv")) #make list of paths
 i_patent_files = 1
 #read files
 for file in all_files:
   if i_patent_files   ==1:
    companies_patent_df = pd.read_csv(file,skiprows=1)
   i_patent_files=i_patent_files+1 


 companies_patent_df.reset_index()
 companies_patent_df.index = np.arange(1, len(companies_patent_df) + 1)

 i_patent_info =1

 for i_patent_info in xrange(1,len(companies_patent_df)+1):
    fetch_sleep = random.randint(5,12)   
    fwd_cit_indx = len(fwd_citation_df)+1
    bckwd_cit_ind = len(bckwd_citation_df)+1
    i_all_ptnt_inx = len(pd_company_patent_info)+1
    url_patents =''
    pd_company_patent_info.at[i_all_ptnt_inx,'Patent_Number'] = companies_patent_df.loc[i_patent_info,'id']

    pd_company_patent_info.at[i_all_ptnt_inx,'Type'] = data_final_company_pd.loc[i_comp_indx_patent,'Type']
    pd_company_patent_info.at[i_all_ptnt_inx,'UnqId'] = i_comp_indx_patent
    pd_company_patent_info.at[i_all_ptnt_inx,'Company'] = company_name
    pd_company_patent_info.at[i_all_ptnt_inx,'Assignee'] = companies_patent_df.loc[i_patent_info,'assignee'] 
    pd_company_patent_info.at[i_all_ptnt_inx,'Publication_Date'] = companies_patent_df.loc[i_patent_info,'publication date']
    pd_company_patent_info.at[i_all_ptnt_inx,'Title'] = companies_patent_df.loc[i_patent_info,'title'].encode('utf8')
    pd_company_patent_info.at[i_all_ptnt_inx,'Researchers'] = companies_patent_df.loc[i_patent_info,'inventor/author']
    pd_company_patent_info.at[i_all_ptnt_inx,'Priority_Date'] = companies_patent_df.loc[i_patent_info,'priority date']  
    pd_company_patent_info.at[i_all_ptnt_inx,'URL'] = companies_patent_df.loc[i_patent_info,'result link']          

    #print companies_patent_df.loc[i_patent_info,'assignee'] 
    #fwd_citation_df[0:0]
    #bckwd_citation_df[0:0]
    try:
      url_patents = companies_patent_df.loc[i_patent_info,'result link']
      #url_patents='https://patents.google.com/patent/US20080050738A1/en'
      response_patent=browser_patent.open(url_patents,timeout=(5.0))
      time.sleep(fetch_sleep)
      #Initializing Beautifulsoup object
      html_patent = response_patent.read()
      bsObj_patent = BeautifulSoup(html_patent)
    except Exception as e:
     error_type, error_obj, error_info = sys.exc_info()
     print 'ERROR FOR URL:',url_patents
     print error_type, 'Line:', error_info.tb_lineno
     savefiles(pd_company_patent_info,fwd_citation_df,bckwd_citation_df)
     continue
    obj_legal_status = bsObj_patent.find("dd", { "itemprop" : "applicationNumber" })
    obj_inventor = bsObj_patent.findAll("dd", { "itemprop" : "inventor" })
    
    pd_company_patent_info.at[i_all_ptnt_inx,'Status'] = (bsObj_patent.find("span", { "itemprop" : "status" }).text
                                                              if str(bsObj_patent.find("span", { "itemprop" : "status" })) <> 'None' else '')                                                      
    
    #['Company','UnqId','Patent_Number','Type','Num_Back_Citations','Num_Fwd_Citations','Publication_Date','Assignee','Title','Researchers'])
   

    table_obj = bsObj_patent.findAll("tr", attrs={ "itemprop" : "backwardReferences" })
    i_ptnt_bckws_indx = 0
    reset_indx = 1
    if str(table_obj) <> 'None':
     for row in table_obj:         
         cells = row.findAll("td")
         if  len(cells) >1 :
           bckwd_citation_df.at[bckwd_cit_ind,'Company'] = company_name
           bckwd_citation_df.at[bckwd_cit_ind,'UnqId'] = i_patent_info
           bckwd_citation_df.at[bckwd_cit_ind,'Patent_Number'] = companies_patent_df.loc[i_patent_info,'id']
           bckwd_citation_df.at[bckwd_cit_ind,'Backward_Citation'] = cells[0].find("span", { "itemprop" : "publicationNumber" }).text
           bckwd_citation_df.at[bckwd_cit_ind,'Examiner_Cited'] = ('Y' if str(cells[0].find("span", { "itemprop" : "examinerCited" })) <> 'None' and 
                                                                    cells[0].find("span", { "itemprop" : "examinerCited" }).text == '*' else 'N')
           bckwd_citation_df.at[bckwd_cit_ind,'Priority_Date'] = cells[1].text
           bckwd_citation_df.at[bckwd_cit_ind,'Publication_Date'] = cells[2].text
           bckwd_citation_df.at[bckwd_cit_ind,'Assignee'] = cells[3].text
           bckwd_citation_df.at[bckwd_cit_ind,'Title'] = cells[4].text
           bckwd_cit_ind=bckwd_cit_ind+1
           i_ptnt_bckws_indx=i_ptnt_bckws_indx+1

    table_obj_fwd = bsObj_patent.findAll("tr", attrs={ "itemprop" : "forwardReferences" })
    i_ptnt_fwd_indx = 0
    
    if str(table_obj) <> 'None':
     for row in table_obj:         
         cells = row.findAll("td")
         if len(cells) >1 :
           fwd_citation_df.at[fwd_cit_indx,'Company'] = company_name
           fwd_citation_df.at[fwd_cit_indx,'UnqId'] = i_patent_info
           fwd_citation_df.at[fwd_cit_indx,'Patent_Number'] = companies_patent_df.loc[i_patent_info,'id']
           fwd_citation_df.at[fwd_cit_indx,'Fwd_Citation'] = cells[0].find("span", { "itemprop" : "publicationNumber" }).text
           fwd_citation_df.at[fwd_cit_indx,'Examiner_Cited'] = ('Y' if str(cells[0].find("span", { "itemprop" : "examinerCited" })) <> 'None' and 
                                                                     cells[0].find("span", { "itemprop" : "examinerCited" }).text == '*' else 'N')
           fwd_citation_df.at[fwd_cit_indx,'Priority_Date'] = cells[1].text
           fwd_citation_df.at[fwd_cit_indx,'Publication_Date'] = cells[2].text
           fwd_citation_df.at[fwd_cit_indx,'Assignee'] = cells[3].text
           fwd_citation_df.at[fwd_cit_indx,'Title'] = cells[4].text
           fwd_cit_indx=fwd_cit_indx+1
           i_ptnt_fwd_indx=i_ptnt_fwd_indx+1 
    pd_company_patent_info.at[i_all_ptnt_inx,'Num_Back_Citations'] = i_ptnt_bckws_indx
    pd_company_patent_info.at[i_all_ptnt_inx,'Num_Fwd_Citations'] = i_ptnt_fwd_indx
    
    i_all_ptnt_inx=i_all_ptnt_inx+1
    #print i_all_ptnt_inx

    
 print('Company name'+data_final_company_pd.loc[i_comp_indx_patent,'Company'])
 print i_all_ptnt_inx

fwd_citation_df_final = fwd_citation_df_final.append(fwd_citation_df)
bckwd_citation_df_final = bckwd_citation_df_final.append(bckwd_citation_df)
pd_company_patent_info_final = pd_company_patent_info_final.append(pd_company_patent_info)


pd_company_patent_info_final.reset_index()
pd_company_patent_info_final.index = np.arange(1, len(pd_company_patent_info_final) + 1)

fwd_citation_df_final.reset_index()
fwd_citation_df_final.index = np.arange(1, len(fwd_citation_df_final) + 1)

bckwd_citation_df_final.reset_index()
bckwd_citation_df_final.index = np.arange(1, len(bckwd_citation_df_final) + 1) 
   
savefiles(pd_company_patent_info_final,fwd_citation_df_final,bckwd_citation_df_final)

browser_patent_sl.quit()