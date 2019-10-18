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

def getCompanyAttributes(company_name,UnqId,bsObj,company_type,target_company):  
    company_attrib_df_mod = pd.DataFrame( columns = ['Company','Desc','Address','Ph','website','Founded_year','twitterhandle','fbhandle','revenue','UnqId','Industry','Competetior','Type','Bloomberg_Name'])
    competetiors_url_df_mod = pd.DataFrame( columns = ['Competetior_Name','URL','UnqId','Region','Target_Company'])
    competetiors_list=[]
    try:     
       comp_attrib_list = bsObj.find("p", { "id" : "bDesc" })     
       company_attrib_df_mod.at[1,'Company'] = company_name
       company_attrib_df_mod.at[1,'Type'] = company_type
       company_attrib_df_mod.at[1,'UnqId'] = UnqId
       if  company_type == 'Competetior':
           company_attrib_df_mod.at[1,'Competetior'] = target_company
       bloomberg_list = bsObj.find("span", { "itemprop" : "name" })
       company_attrib_df_mod.at[1,'Bloomberg_Name'] = (bloomberg_list.text if str(bloomberg_list) <> 'None' else '')
            
       company_attrib_df_mod.at[1,'Desc'] = (comp_attrib_list.text if str(comp_attrib_list) <> 'None' else '')
       address_list = bsObj.find("div", { "class" : "detailsDataContainerLt" }).find("div", { "itemprop" : "address" }) 
       company_attrib_df_mod.at[1,'Address'] = (address_list.text if str(address_list) <> 'None' else '')       
       ph_list =  bsObj.find("p", attrs={"itemprop" : "telephone"})
       company_attrib_df_mod.at[1,'Ph'] = (ph_list.text if str(ph_list) <> 'None' else '')
       competetiors_list_Obj = bsObj.findAll("table", attrs={ "class" : "table" })
       competetiors_list = (competetiors_list_Obj[0] if str(competetiors_list_Obj) <> 'None' else '')
       founded_list = bsObj.find("div", { "class" : "detailsDataContainerLt" }).find("span", { "itemprop" : "foundingDate" })
       company_attrib_df_mod.at[1,'Founded_year'] = (founded_list.text  
                                                               if str(founded_list) <> 'None' else '')    
       industry_list = bsObj.find( attrs={"name" : "industry"})
       company_attrib_df_mod.at[1,'Industry'] = (industry_list['content'] if str(industry_list) <> 'None' else '')
       website_list = bsObj.find("a", { "itemprop" : "url" })
       company_attrib_df_mod.at[1,'website'] = (website_list['href'] if str(website_list) <> 'None' else '')
       
    except Exception as e:
            error_type, error_obj, error_info = sys.exc_info()
            print 'ERROR in getting company attributes:',company_name
            print error_type, 'Line:', error_info.tb_lineno
            pass 
    print company_name        
    if len(competetiors_list)> 0 and company_type == 'Target':    
          competetiors_url_df_mod = getCompetetiorsURL(competetiors_list,bsObj,company_name,UnqId,competetiors_url_df_mod)      
    return company_attrib_df_mod,competetiors_url_df_mod

def getCompetetiorsURL(competetiors_list,bsObj,company_name,UnqId,competetiors_url_df_mod):    
    i_comp = 1
    reset_indx = 1
    competetiors_url_df_mod = competetiors_url_df_mod[0:0]
    #print "inside competetiors"
    #print company_name
    for row in competetiors_list.findAll("tr"):         
         cells = row.findAll("td")
         if  reset_indx>1 and len(cells) >1 :
             competetiors_url_df_mod.at[i_comp,'Target_Company'] = company_name
             competetiors_url_df_mod.at[i_comp,'UnqId'] = UnqId
             competetiors_url_df_mod.at[i_comp,'Competetior_Name'] = cells[0].text
             competetiors_url_df_mod.at[i_comp,'URL'] = 'https://www.bloomberg.com/research/stocks/private/'+cells[0].find("a")['href']
             competetiors_url_df_mod.at[i_comp,'Region'] = cells[1].text
             i_comp=i_comp+1  
                                             
         reset_indx =reset_indx+1

    
    return  competetiors_url_df_mod  

def getpeopleinfo(bsObj1,company_name,UnqId):
    people_df_mod = pd.DataFrame( columns = ['Person_Name','Title','UnqId','Age','Number_Relations','Company_Name','Person_Id','URL','Primary_Company'])
    table_exec = bsObj1.find("table", attrs={ "id" : "keyExecs" })
    table_board = bsObj1.findAll("table", attrs={ "class" : "table" })
    i_ppl_indx = 1
    reset_indx = 1
    if str(table_exec) <> 'None':
     for row in table_exec.findAll("tr"):         
         cells = row.findAll("td")
         if reset_indx>1 and len(cells) >1 :
           people_df_mod.at[i_ppl_indx,'Company_Name'] = company_name
           people_df_mod.at[i_ppl_indx,'UnqId'] = UnqId
           people_df_mod.at[i_ppl_indx,'Person_Id'] = i_ppl_indx
           if 'No Relationships' not in cells[1].text:
            people_df_mod.at[i_ppl_indx,'Number_Relations'] = cells[1].find("strong").text
            people_df_mod.at[i_ppl_indx,'URL'] = cells[1].find("a")['href'].replace("../..","https://www.bloomberg.com/research")
           people_df_mod.at[i_ppl_indx,'Person_Name'] = cells[0].text
           people_df_mod.at[i_ppl_indx,'Title'] = cells[2].text
           people_df_mod.at[i_ppl_indx,'Age'] = cells[3].text           
           i_ppl_indx = i_ppl_indx+1
         reset_indx =reset_indx+1
    reset_indx = 1 
    if str(table_board) <> 'None' and len(table_board)>0:    
     for row in table_board[1].findAll("tr"):         
         cells = row.findAll("td")
         if reset_indx>1 and len(cells) >1 :
           people_df_mod.at[i_ppl_indx,'Company_Name'] = company_name
           people_df_mod.at[i_ppl_indx,'UnqId'] = UnqId
           people_df_mod.at[i_ppl_indx,'Person_Id'] = i_ppl_indx
           if 'No Relationships' not in cells[1].text:
            people_df_mod.at[i_ppl_indx,'Number_Relations'] = cells[1].find("strong").text
            people_df_mod.at[i_ppl_indx,'URL'] = cells[1].find("a")['href'].replace("../..","https://www.bloomberg.com/research")
           people_df_mod.at[i_ppl_indx,'Person_Name'] = cells[0].text
           people_df_mod.at[i_ppl_indx,'Primary_Company'] = cells[2].text
           people_df_mod.at[i_ppl_indx,'Age'] = cells[3].text  
           people_df_mod.at[i_ppl_indx,'Title']  = 'Board Member'       
           i_ppl_indx = i_ppl_indx+1
         reset_indx =reset_indx+1

    return people_df_mod

def getrelations(people_df):
    relation_df_mod = pd.DataFrame( columns = ['Person_Id','Person_Name','UnqId','Relation_Name','Relation_Company'])
    rel_indx_all = 1
    relation_list=[]
    for i_rel_indx in xrange(1,len(people_df)+1):
     if str(people_df.loc[i_rel_indx,'URL']) <> 'nan':
         try:
            response2=browser1.open(people_df.loc[i_rel_indx,'URL'],timeout=(10.0))
            sleep_time1 = random.randint(2,20)
            time.sleep(sleep_time1) 
            html2 = response2.read()
            bsObj2 = BeautifulSoup(html2)  
            relation_list = bsObj2.findAll("div",attrs={"class":"relationBox"})
         except Exception as e:
            error_type, error_obj, error_info = sys.exc_info()
            print 'ERROR FOR URL:',people_df.loc[i_rel_indx,'URL']
            print error_type, 'Line:', error_info.tb_lineno
            continue
#Initializing Beautifulsoup object

     #relations = relation_list[0]
         for relations in relation_list:   
            relation_df_mod.at[rel_indx_all,'Person_Id'] = people_df.loc[i_rel_indx,'Person_Id']
            relation_df_mod.at[rel_indx_all,'Person_Name'] = people_df.loc[i_rel_indx,'Person_Name']
            relation_df_mod.at[rel_indx_all,'UnqId'] = people_df.loc[i_rel_indx,'UnqId']
            relation_df_mod.at[rel_indx_all,'Relation_Name'] = relations.findAll("a")[0].text 
            relation_df_mod.at[rel_indx_all,'Relation_Company'] = relations.findAll("a")[1].text 
            rel_indx_all=rel_indx_all+1
    
    return relation_df_mod

def savefiles(company_attrib_df,people_df,relation_df,competetiors_url_df):
 if len(company_attrib_df)>1:
    writer_comp = pd.ExcelWriter('A:/Capstone_Code/files/Bloomberg_CD_CompanyAttributes.xlsx', engine='xlsxwriter')
    company_attrib_df.to_excel(writer_comp,'Sheet1')
    writer_comp.save()
    
 if len(people_df)>1:
    writer_ppl = pd.ExcelWriter('A:/Capstone_Code/files/Bloomberg_CD_People.xlsx', engine='xlsxwriter')
    people_df.to_excel(writer_ppl,'Sheet1')
    writer_ppl.save()
    
 if len(relation_df)>1:
    writer_rel = pd.ExcelWriter('A:/Capstone_Code/files/Bloomberg_CD_Relations.xlsx', engine='xlsxwriter')
    relation_df.to_excel(writer_rel,'Sheet1')
    writer_rel.save()

 if len(competetiors_url_df)>1:
    writer_all_Competetiors = pd.ExcelWriter('A:/Capstone_Code/files/Bloomberg_CD_Competetiors_Data.xlsx', engine='xlsxwriter')
    competetiors_url_df.to_excel(writer_all_Competetiors,'Sheet1')
    writer_all_Competetiors.save()


                                    

company_attrib_df = pd.DataFrame( columns = ['Company','Desc','Address','Ph','website','Founded_year','twitterhandle','fbhandle','revenue','UnqId','Industry','Competetior','Type','Bloomberg_Name'])
competetiors_url_df = pd.DataFrame( columns = ['Competetior_Name','URL','UnqId','Region','Target_Company'])
relation_df = pd.DataFrame( columns = ['Person_Id','Person_Name','UnqId','Relation_Name','Relation_Company'])
people_df = pd.DataFrame( columns = ['Person_Name','Title','UnqId','Age','Number_Relations','Company_Name','Person_Id','URL','Primary_Company'])
company_attrib_df_mod = pd.DataFrame( columns = ['Company','Desc','Address','Ph','website','Founded_year','twitterhandle','fbhandle','revenue','UnqId','Industry','Competetior','Type','Bloomberg_Name'])
competetiors_url_df_mod = pd.DataFrame( columns = ['Competetior_Name','URL','UnqId','Region','Target_Company'])
relation_df_mod = pd.DataFrame( columns = ['Person_Id','Person_Name','UnqId','Relation_Name','Relation_Company'])
people_df_mod = pd.DataFrame( columns = ['Person_Name','Title','UnqId','Age','Number_Relations','Company_Name','Person_Id','URL','Primary_Company'])   

url_df = pd.DataFrame( columns = ['URL', 'Company','UnqId'])
url_df_final = pd.DataFrame( columns = ['URL', 'Company','UnqId'])
url_df_mod =  pd.DataFrame( columns = ['URL', 'Company','UnqId'])
final_url_df = pd.DataFrame(columns = ['URL', 'Company','UnqId','Type'])
data_pd = pd.DataFrame( columns = ['Company','Type','Missing'])

browser = webdriver.Firefox(executable_path='A:\geckodriver-v0.19.0-win64\geckodriver.exe')
time.sleep(5)


df_excel = pd.read_excel('A:/Capstone_Code/files/Modelling_File_CD.xlsx')#, sheetname='Screening',skiprows=7)
df_excel = df_excel.loc[df_excel['Bloomberg_Name'].isnull()]
df_excel = df_excel.loc[ df_excel['Type'].isnull()]

df_excel.reset_index()
df_excel.index = np.arange(1, len(df_excel) + 1)


for j in xrange(0,len(df_excel)):
   
  indx_clean =  df_excel.loc[j,'Company'].find(',')
    data_pd.at[j,'Company']
   data_pd['Company'] = df_excel['Company']
   data_pd['Type'] = 'Target'
data_pd_final= data_pd.drop_duplicates()

data_pd_final.reset_index()
data_pd_final.index = np.arange(1, len(data_pd_final) + 1)

browser.get('http://www.yahoo.com')
comp_names_indx=1
p_indx_url_exists = 1
for comp_names_indx in xrange(1,len(data_pd_final)+1): 
 sleep_time = random.randint(6,20)
 company_name = data_pd_final.loc[comp_names_indx,'Company']
 comp_type = data_pd_final.loc[comp_names_indx,'Type']
 #indx_reset_comp=indx_reset_comp+1
 try:
    wait = WebDriverWait(browser, 30)
    box = wait.until(EC.presence_of_element_located(
            (By.NAME, "p")))       
 except TimeoutException:
    print("Yahoo not loaded")
    writer_all_Urls = pd.ExcelWriter('A:/Capstone_Code/files/IT_AllUrls.xlsx', engine='xlsxwriter')
    url_df_final.to_excel(writer_all_Urls,'Sheet1')
    writer_all_Urls.save()

    writer_cmp_Urls = pd.ExcelWriter('A:/Capstone_Code/files/IT_CompanyUrls.xlsx', engine='xlsxwriter')
    final_url_df.to_excel(writer_cmp_Urls,'Sheet1')
    writer_cmp_Urls.save()



    writer_all_Data = pd.ExcelWriter('A:/Capstone_Code/files/IT_AllData.xlsx', engine='xlsxwriter')
    data_pd_final.to_excel(writer_all_Data,'Sheet1')
    writer_all_Data.save()

    continue

#assert 'Yahoo' in browser.title
 elem = browser.find_element_by_name('p')  # Find the search box
 elem.clear()
 browser.wait = WebDriverWait(browser, 4)
 elem.send_keys(company_name+'+bloomberg' + Keys.RETURN)   
 time.sleep(sleep_time)
 i=1
 j_all_urls=1
 url_df = url_df[0:0]
 url_df_mod = url_df_mod[0:0]
 for url in browser.find_elements_by_class_name("wr-bw"):    
    if "privcapId" in url.text: 
        url_df.at[i,'URL']=  url.text
        url_df.at[i,'UnqId']=  comp_names_indx
        url_df.at[i,'Company']=  company_name        
        i=i+1
    
    url_df_mod.at[j_all_urls,'URL']=  url.text
    url_df_mod.at[j_all_urls,'UnqId']=  comp_names_indx
    url_df_mod.at[j_all_urls,'Company']=  company_name 
    j_all_urls = j_all_urls+1
            
 if len(url_df) >0:       
    final_url_df.at[p_indx_url_exists,'URL']=  "https://"+url_df.loc[1,'URL']
    final_url_df.at[p_indx_url_exists,'Company']=  url_df.loc[1,'Company']
    final_url_df.at[p_indx_url_exists,'UnqId']=  url_df.loc[1,'UnqId']
    final_url_df.at[p_indx_url_exists,'Type']= comp_type
    data_pd_final.at[comp_names_indx,'Missing']='N' 
    p_indx_url_exists= p_indx_url_exists+1
 else :
    data_pd_final.at[comp_names_indx,'Missing']='Y'
 url_df_final = url_df_final.append(url_df_mod)
 #if sleep_time == 4 or sleep_time ==6:
     #url.click()

browser.quit()
writer_all_Urls = pd.ExcelWriter('A:/Capstone_Code/files/IT_AllUrls.xlsx', engine='xlsxwriter')
url_df_final.to_excel(writer_all_Urls,'Sheet1')
writer_all_Urls.save()

writer_cmp_Urls = pd.ExcelWriter('A:/Capstone_Code/files/IT_CompanyUrls.xlsx', engine='xlsxwriter')
final_url_df.to_excel(writer_cmp_Urls,'Sheet1')
writer_cmp_Urls.save()



writer_all_Data = pd.ExcelWriter('A:/Capstone_Code/files/IT_AllData.xlsx', engine='xlsxwriter')
data_pd_final.to_excel(writer_all_Data,'Sheet1')
writer_all_Data.save()


browser1=urllib2.build_opener()
browser1.addheaders=[('User-agent', 'Mozilla/5.0')]
comp_indx = 1

final_url_df= pd.read_excel('A:/Capstone_Code/files/Bloomberg_CD_CompanyUrls.xlsx')
competetiors_url_df = pd.read_excel('A:/Capstone_Code/files/Bloomberg_CD_Competetiors_Data.xlsx')
company_attrib_df= pd.read_excel('A:/Capstone_Code/files/Bloomberg_CD_CompanyAttributes.xlsx')
people_df= pd.read_excel('A:/Capstone_Code/files/Bloomberg_CD_People.xlsx')
relation_df= pd.read_excel('A:/Capstone_Code/files/Bloomberg_CD_Relations.xlsx')

for comp_indx in xrange(1,len(final_url_df)+1):
 people_df_mod=people_df_mod[0:0]
 competetiors_url_df_mod=competetiors_url_df_mod[0:0]
 relation_df_mod=relation_df_mod[0:0]
 company_attrib_df_mod=company_attrib_df_mod[0:0]
 sleep_time = random.randint(5,40)
 time.sleep(sleep_time)
 try:
     url_bloomberg = final_url_df.loc[comp_indx,'URL']
     response=browser1.open(url_bloomberg,timeout=(5.0))
     #Initializing Beautifulsoup object
     html = response.read()
     bsObj = BeautifulSoup(html)
 except Exception as e:
     error_type, error_obj, error_info = sys.exc_info()
     print 'ERROR FOR URL:',url_bloomberg
     print error_type, 'Line:', error_info.tb_lineno
     savefiles(company_attrib_df,people_df,relation_df,competetiors_url_df)
     continue

 #class method to fetch company attributes
 #if comp_indx == 1:
 #   company_attrib_df,competetiors_url_df = getCompanyAttributes(final_url_df.loc[comp_indx,'Company'],final_url_df.loc[comp_indx,'UnqId'],bsObj,final_url_df.loc[comp_indx,'Type'])
# else:
 company_attrib_df_mod,competetiors_url_df_mod = getCompanyAttributes(final_url_df.loc[comp_indx,'Company'],final_url_df.loc[comp_indx,'UnqId'],bsObj,final_url_df.loc[comp_indx,'Type'],'')
 company_attrib_df= company_attrib_df.append(company_attrib_df_mod)
 competetiors_url_df= competetiors_url_df.append(competetiors_url_df_mod)
 
 print company_attrib_df_mod   
 notselobj_ppl = bsObj.find("div",attrs={"class" : "fLeft tabPeople"}).find("a",attrs={"class" : "notSelected"})
 selobj_ppl = bsObj.find("div",attrs={"class" : "fLeft tabPeople"}).find("a",attrs={"class" : "selected"})
 try:
     people_tab_select_url = (notselobj_ppl['href'] if str(notselobj_ppl) <> 'None' else selobj_ppl['href']).replace("../","https://www.bloomberg.com/research/stocks/")
 except Exception as e:
     error_type, error_obj, error_info = sys.exc_info()
     print 'ERROR People tab not found:'
     print error_type, 'Line:', error_info.tb_lineno
     savefiles(company_attrib_df,people_df,relation_df,competetiors_url_df)
     continue    
 try:
     response1=browser1.open(people_tab_select_url,timeout=(5.0))
     sleep_time_ppl = random.randint(1,5)
     time.sleep(sleep_time_ppl) 
     html1 = response1.read()
     bsObj1 = BeautifulSoup(html1)
 except Exception as e:
     error_type, error_obj, error_info = sys.exc_info()
     print 'ERROR FOR URL:',people_tab_select_url
     print error_type, 'Line:', error_info.tb_lineno
     savefiles(company_attrib_df,people_df,relation_df,competetiors_url_df)
     continue
#Initializing Beautifulsoup object
 
 if comp_indx == 1:
    people_df = getpeopleinfo(bsObj1,final_url_df.loc[comp_indx,'Company'],final_url_df.loc[comp_indx,'UnqId'])
    relation_df = getrelations(people_df)
 else:
    people_df_mod = getpeopleinfo(bsObj1,final_url_df.loc[comp_indx,'Company'],final_url_df.loc[comp_indx,'UnqId'])
    relation_df_mod = getrelations(people_df_mod)    
    people_df = people_df.append(people_df_mod)
    relation_df= relation_df.append(relation_df_mod)
    
competetiors_url_df.reset_index()
competetiors_url_df.index = np.arange(1, len(competetiors_url_df) + 1)   
company_attrib_df.reset_index()
company_attrib_df.index = np.arange(1, len(company_attrib_df) + 1)   
company_attrib_df.reset_index()
company_attrib_df.index = np.arange(1, len(company_attrib_df) + 1)   
people_df.reset_index()
people_df.index = np.arange(1, len(people_df) + 1) 
relation_df.reset_index()
relation_df.index = np.arange(1, len(relation_df) + 1)

savefiles(company_attrib_df,people_df,relation_df,competetiors_url_df)

icompt_attrib_indx = max(company_attrib_df['UnqId'])+1
uniqueUrlsList = pd.unique(competetiors_url_df.URL.ravel())
#uniqueUrlsList = pd.DataFrame(uniqueUrlsList)


browser1=urllib2.build_opener()
browser1.addheaders=[('User-agent', 'Mozilla/5.0')]

#'Rakhi Properties & Leasing Pvt Ltd

for i in xrange(1,len(uniqueUrlsList)):
 url_compt = uniqueUrlsList[i] 
 region=  competetiors_url_df[competetiors_url_df.URL==url_compt].Region.unique()
 competetior = competetiors_url_df[competetiors_url_df.URL==url_compt].Competetior_Name.unique()
 if region[0] == "United States" or region[0]=='Americas':
   try:
     response2=browser1.open(url_compt,timeout=(5.0))
     sleep_time_compt = random.randint(4,20)
     time.sleep(sleep_time_compt) 
     #Initializing Beautifulsoup object
     html2 = response2.read()
     bsObj2 = BeautifulSoup(html2) 
   except Exception as e:
     error_type, error_obj, error_info = sys.exc_info()
     print 'ERROR FOR URL:',url_compt
     print error_type, 'Line:', error_info.tb_lineno
     savefiles(company_attrib_df,people_df,relation_df,[])
     continue

   company_attrib_df_mod = company_attrib_df_mod[0:0]
   competetiors_url_df_mod = competetiors_url_df_mod[0:0]
   company_attrib_df_mod,competetiors_url_df_mod = getCompanyAttributes(competetior[0],icompt_attrib_indx,bsObj2,'Competetior','')
   company_attrib_df= company_attrib_df.append(company_attrib_df_mod)
   #competetiors_url_df= competetiors_url_df.append(competetiors_url_df_mod)
   print company_attrib_df_mod    
   
   try:
     notselobj_ppl = bsObj2.find("div",attrs={"class" : "fLeft tabPeople"}).find("a",attrs={"class" : "notSelected"})
     selobj_ppl = bsObj2.find("div",attrs={"class" : "fLeft tabPeople"}).find("a",attrs={"class" : "selected"})
     people_tab_select_url = (notselobj_ppl['href'] if str(notselobj_ppl) <> 'None' else selobj_ppl['href']).replace("../","https://www.bloomberg.com/research/stocks/")
   except Exception as e:
     error_type, error_obj, error_info = sys.exc_info()
     print 'ERROR People tab not found:'
     print error_type, 'Line:', error_info.tb_lineno
     savefiles(company_attrib_df,people_df,relation_df,[])
     continue    
      
   try:
     response3=browser1.open(people_tab_select_url,timeout=(5.0))
     sleep_time_ppl1 = random.randint(5,15)
     time.sleep(sleep_time_ppl1) 
     #Initializing Beautifulsoup object
     html3 = response3.read()
     bsObj3 = BeautifulSoup(html3)
   except Exception as e:
     error_type, error_obj, error_info = sys.exc_info()
     print 'ERROR FOR URL:',people_tab_select_url
     print error_type, 'Line:', error_info.tb_lineno
     savefiles(company_attrib_df,people_df,relation_df,[])
     continue

   people_df_mod = getpeopleinfo(bsObj3,competetior[0],icompt_attrib_indx)
   time.sleep(sleep_time_compt)  
   relation_df_mod = getrelations(people_df_mod)    
   people_df = people_df.append(people_df_mod)
   relation_df= relation_df.append(relation_df_mod)
   icompt_attrib_indx=icompt_attrib_indx+1

company_attrib_df.reset_index()
company_attrib_df.index = np.arange(1, len(company_attrib_df) + 1)   
company_attrib_df.reset_index()
company_attrib_df.index = np.arange(1, len(company_attrib_df) + 1)   
people_df.reset_index()
people_df.index = np.arange(1, len(people_df) + 1) 
relation_df.reset_index()
relation_df.index = np.arange(1, len(relation_df) + 1)
   

savefiles(company_attrib_df,people_df,relation_df,[])




#final_url_df[(final_url_df.UnqId==629)]
#relation_df[(relation_df.UnqId==629)]
#max(relation_df["UnqId"])

#relation_df = pd.read_excel('A:/Capstone_Code/files/Bloomberg_IT_Relations.xlsx', sheetname='Sheet1')
#final_url_df = pd.read_excel('A:/Capstone_Code/files/Bloomberg_IT_CompanyUrls.xlsx', sheetname='Sheet1')
#company_attrib_df=pd.read_excel('A:/Capstone_Code/files/IT_CompanyAttributes.xlsx', sheetname='Sheet1')
#people_df=pd.read_excel('A:/Capstone_Code/files/Bloomberg_IT_People.xlsx', sheetname='Sheet1')
#competetiors_url_df = pd.read_excel('A:/Capstone_Code/files/Bloomberg_IT_Competetiors_Data.xlsx', sheetname='Sheet1')
#
#
#competetiors_url_df[(competetiors_url_df.URL=='https://www.bloomberg.com/research/stocks/private/snapshot.asp?privcapId=6527758')].Region