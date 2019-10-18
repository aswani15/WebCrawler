import time

import pandas as pd
import sys,urllib2
from BeautifulSoup import BeautifulSoup 
import numpy as np
import random
#from yahoo_finance import Share
import xlsxwriter
#import xlsxwriter

def fetchfinvalues(table_IncomeStmnt,company,comp_type):
   df_pd = pd.DataFrame( columns = ['Company','Type','UnqId','Year','Description','Value','Missing'])
   i_fin_indx=1
   reset_indx=1
   year1=''
   year2=''
   year3=''
   #row = table_IncomeStmnt.findAll("tr")[0]
   if str(table_IncomeStmnt) <> 'None':
    for row in table_IncomeStmnt.findAll("tr"):         
         cells = row.findAll("td")
         if reset_indx == 1 and len(cells) >1:
           year1= (cells[1].text[-4:]  if str(cells[1].text) <> 'None' and len(cells[1].text)>6 else '')
           year2= (cells[2].text[-4:] if len(cells)>2 else '')
           year3= (cells[3].text[-4:] if len(cells)>3 else '')     
         if reset_indx>1 and len(cells) >1 :             
           df_pd = (assignvalues(df_pd,i_fin_indx,year1,1,
                                      data_final_pd.loc[com_tick_ind,'Company'],
                                       data_final_pd.loc[com_tick_ind,'Type'],cells) if len(year1)>0 else df_pd)
           df_pd = (assignvalues(df_pd,i_fin_indx+1,year2,2,
                                      data_final_pd.loc[com_tick_ind,'Company'],
                                       data_final_pd.loc[com_tick_ind,'Type'],cells) if len(year2)>0 else df_pd)
           df_pd = (assignvalues(df_pd,i_fin_indx+2,year3,3,
                                      data_final_pd.loc[com_tick_ind,'Company'],
                                       data_final_pd.loc[com_tick_ind,'Type'],cells) if len(year3)>0 else df_pd)
         
           i_fin_indx = i_fin_indx+3
         reset_indx =reset_indx+1    
   return df_pd

def assignvalues(df_pd,i_fin_indx,year1,index_i,company,comp_type,cells):
     df_pd.at[i_fin_indx,'Company'] = company
     df_pd.at[i_fin_indx,'Type'] = comp_type
     df_pd.at[i_fin_indx,'Year'] = year1
     df_pd.at[i_fin_indx,'Description'] = (cells[0].text if len(cells[0])>0 else '')
     df_pd.at[i_fin_indx,'Value'] = (cells[index_i].text.replace(',','') if len(cells[index_i])>0 else '')  
     df_pd.at[i_fin_indx,'Missing'] = 'N'
     return df_pd

data_pd = pd.DataFrame( columns = ['Company','Type','Ticker','UnqId','Missing'])
finance_pd = pd.DataFrame( columns = ['Company','Type','UnqId','Year','Description','Value','Missing'])
finance_pd_mod = pd.DataFrame( columns = ['Company','Type','UnqId','Year','Description','Value','Missing'])

#'Tot_Revenue','Profit','RnD_Exp','Tot_OPEX','OPS_Inc_Loss','EBITA','Net_Income',
#'Tot_Cur_Assets','Tot_Cur_Liab','Net_Tang_Assets','Tot_Assets','Tot_Liab','Tot_Cash_Flow_Ops',
#'Tot_Cash_Flow_Invst','Net_Borrowings','Tot_Cash_Flow_FinAct','Change_In_CashFlow','Year','Missing'])

df_excel = pd.read_excel('A:/Capstone_Code/Industrials_Old.xls', sheetname='Screening',skiprows=7)

index_data =1
for i in xrange(0,len(df_excel)):
    data_pd.at[index_data,'Company'] = df_excel.loc[i,'Target/Issuer']
    company_name = (df_excel.loc[i,'Target/Issuer'])
    index= company_name.find(':')
    if index >0:
       data_pd.at[index_data,'Ticker'] = company_name[index+1:len(company_name)-1]
    data_pd.at[index_data,'Type'] = 'Target'
    index_data=index_data+1
    
i_indx =len(data_pd)+1
for j in xrange(0,len(df_excel)):
    data_pd.at[i_indx,'Company'] = df_excel.loc[j,'Buyers/Investors']
    company_name_buy = (df_excel.loc[j,'Buyers/Investors'])
    index_buy= company_name_buy.find(':')
    if index_buy >0:
       data_pd.at[i_indx,'Ticker'] = company_name_buy[index_buy+1:len(company_name_buy)-1]
    data_pd.at[i_indx,'Type'] = 'Buyer'
    i_indx=i_indx+1
#drop duplicates
data_final_pd = data_pd.drop_duplicates()

##reset index after removing duplicates
data_final_pd.reset_index()
data_final_pd.index = np.arange(1, len(data_final_pd) + 1)


browser=urllib2.build_opener()
browser.addheaders=[('User-agent', 'Mozilla/5.0')]

for com_tick_ind in xrange(1,len(data_final_pd)+1):
 finance_pd_mod=finance_pd_mod[0:0]
 url_yf = 'https://finance.yahoo.com/quote/'

 if str(data_final_pd.loc[com_tick_ind,'Ticker']) <> 'nan':
  try:
     ticker = data_final_pd.loc[com_tick_ind,'Ticker']
     url_yf_ins = url_yf+str(ticker)+'/financials'  
     #yahoo = Share(str(ticker))   
     response=browser.open(url_yf_ins,timeout=(5.0))
     sleep_time_yahoo = random.randint(10,40)
     time.sleep(sleep_time_yahoo)
     #Initializing Beautifulsoup object
     html = response.read()
     bsObj = BeautifulSoup(html)
     table_IncomeStmnt = bsObj.find("table", attrs={ "class" : "Lh(1.7) W(100%) M(0)" })
  except Exception as e:
     error_type, error_obj, error_info = sys.exc_info()
     print 'ERROR FOR URL:',url_yf_ins
     print error_type, 'Line:', error_info.tb_lineno
     data_final_pd.at[com_tick_ind,'Missing']='Y'
     writer_fin = pd.ExcelWriter('A:/Capstone_Code/files/Industrials_Financial_Data_Data.xlsx', engine='xlsxwriter')
     finance_pd.to_excel(writer_fin,'Sheet1')
     writer_fin.save()

     writer_fin_data = pd.ExcelWriter('A:/Capstone_Code/files/Industrials_All_Data_Fin.xlsx', engine='xlsxwriter')
    
     data_final_pd.to_excel(writer_fin_data,'Sheet1')
     writer_fin_data.save()

     continue 
  data_final_pd.at[com_tick_ind,'Missing']='N'     
  finance_pd_mod = fetchfinvalues(table_IncomeStmnt,
                               data_final_pd.loc[com_tick_ind,'Company'],
                                       data_final_pd.loc[com_tick_ind,'Type']) 
  finance_pd=finance_pd.append(finance_pd_mod)
  finance_pd_mod=finance_pd_mod[0:0]
  try:
     url_yf_bs = url_yf+str(ticker)+'/balance-sheet'
     response_bs=browser.open(url_yf_bs,timeout=(5.0))
     html_bs = response_bs.read()
     bsObj_bs = BeautifulSoup(html_bs)
     table_blncsheet = bsObj_bs.find("table", attrs={ "class" : "Lh(1.7) W(100%) M(0)" })     
  except Exception as e:
     error_type, error_obj, error_info = sys.exc_info()
     print 'ERROR FOR URL:',url_yf
     print error_type, 'Line:', error_info.tb_lineno
     data_final_pd.at[com_tick_ind,'Missing']='Y'
     continue 
  finance_pd_mod = fetchfinvalues(table_blncsheet,
                               data_final_pd.loc[com_tick_ind,'Company'],
                                       data_final_pd.loc[com_tick_ind,'Type']) 
  finance_pd=finance_pd.append(finance_pd_mod)
  finance_pd_mod=finance_pd_mod[0:0]
  try:
     url_yf_cf = url_yf+str(ticker)+'/cash-flow'
     response_cf=browser.open(url_yf_cf,timeout=(5.0))
     html_cf = response_cf.read()
     bsObj_cf = BeautifulSoup(html_cf)
     table_cashflow = bsObj_cf.find("table", attrs={ "class" : "Lh(1.7) W(100%) M(0)" })     
  except Exception as e:
     error_type, error_obj, error_info = sys.exc_info()
     print 'ERROR FOR URL:',url_yf_cf
     print error_type, 'Line:', error_info.tb_lineno
     writer_fin = pd.ExcelWriter('A:/Capstone_Code/files/Industrials_Financial_Data_Data.xlsx', engine='xlsxwriter')
     finance_pd.to_excel(writer_fin,'Sheet1')
     writer_fin.save()

     writer_fin_data = pd.ExcelWriter('A:/Capstone_Code/files/Industrials_All_Data_Fin.xlsx', engine='xlsxwriter')
     data_final_pd.to_excel(writer_fin_data,'Sheet1')
     writer_fin_data.save()

     data_final_pd.at[com_tick_ind,'Missing']='Y'
     continue 
  finance_pd_mod = fetchfinvalues(table_cashflow,
                               data_final_pd.loc[com_tick_ind,'Company'],
                                       data_final_pd.loc[com_tick_ind,'Type']) 
  finance_pd=finance_pd.append(finance_pd_mod)
  finance_pd_mod=finance_pd_mod[0:0]
 
  finance_pd.reset_index()
  finance_pd.index = np.arange(1, len(finance_pd) + 1)   
  print  data_final_pd.loc[com_tick_ind,'Company']
 #size_finance_pd=len(finance_pd)
 #finance_pd.at[size_finance_pd+1,'Description'] = 'PE Ratio'
 #finance_pd.at[size_finance_pd+1,'Company'] = data_final_pd.loc[com_tick_ind,'Company']
 #finance_pd.at[size_finance_pd+1,'Company'] = data_final_pd.loc[com_tick_ind,'Type']
 #finance_pd.at[size_finance_pd+1,'Value'] = yahoo.get_price_earnings_ratio()
  if sleep_time_yahoo==10 or sleep_time_yahoo==20 or sleep_time_yahoo==15:
     try:
         url_yf_random = url_yf+str(ticker)+'/key-statistics'
         response_random=browser.open(url_yf_random,timeout=(5.0))
     except Exception as e:
         continue
  if sleep_time_yahoo==4 or sleep_time_yahoo==8 or sleep_time_yahoo==16:
     try:
         url_yf_random = url_yf+str(ticker)+'/analysts'
         response_random=browser.open(url_yf_random,timeout=(5.0))
     except Exception as e:
         continue
     
writer_fin = pd.ExcelWriter('A:/Capstone_Code/files/Industrials_Financial_Data_Data.xlsx', engine='xlsxwriter')
finance_pd.to_excel(writer_fin,'Sheet1')
writer_fin.save()

writer_fin_data = pd.ExcelWriter('A:/Capstone_Code/files/Industrials_All_Data_Fin.xlsx', engine='xlsxwriter')
data_final_pd.to_excel(writer_fin_data,'Sheet1')
writer_fin_data.save()

