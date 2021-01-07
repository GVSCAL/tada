# -*- coding: utf-8 -*-
"""
Created on Mon May 25 16:08:38 2020

@author: chenwang
THC_summary_V2 :compare to V1, No need export csv from Lv0 report website.
Now script scraping information from website based runid.

input = runids from a txt file 'runidlist.txt' , including runs we want extract information.
output = One excel with all information from THC. ==> 
this goes to PowerBI for visualisation.
  
careful : 
    error if run have no Lv0 report.
    error if run have no THC table.
    better to create some kind of checking function.
    
Update from V2.1 : output method upgraded to xlsxwriter, doing 'format as table' for PowerBI
Update from V2.2 : remove duplicated run ids, add unit for criterias.
"""

import requests
from bs4 import BeautifulSoup
import pandas as pd
import numpy as np
from functools import reduce
import re
from datetime import datetime
import sys,os

from configparser import ConfigParser


class FooException(Exception):
    def __init__(self, run):
        self.runID = run


class DataGrasper():

    def __init__(self):
        self.runIDs = []
        self.data_frames = []

        cfg = ConfigParser()
        cfg.read('../config.ini')
        self.direc_path = cfg.get('user_setting','directory_path')

    def open_txt_file(self, path):
        with open(path) as f:
            txt_orig = f.read().upper()                                   ## get runid from a txt file
            self.runid_list = re.findall(r'\w\w[12]\d{5}\d?',txt_orig)           ## find runids into a list according to seach pattern
            self.runid_list = list(dict.fromkeys(self.runid_list))                      # remove duplicated runs.
            # self.data_frames = []
        return self.runid_list

    def search_online_by_runID(self, runid_list):
        self.data_frames = []
        print('runid:',runid_list)
        try:
            for runs in runid_list:
                self.data_frames.append(self.get_df_from_runid(runs))       #put all runs df together.
        except AttributeError:
            raise FooException(runs)

    def get_df_from_runid(self,RunID):
        print('get_df_from_runid:',RunID)
    ##this function returns df we need by scraping infor from Lv0 website based on RunID
        url = 'http://frbriunil007.bri.fr.corp/dashboard/MIT_reports.php?FEA_ref='+RunID
        data = requests.get(url)
        soup = BeautifulSoup(data.text, 'html.parser')      #GET All Lv0 report web information(html) into 'soup'

        div = soup.find('div',{'class' : 'col-lg-4'})
        h3 = div.find_all('h3')
        project_name__ = h3[0].text                           #get project name and OEM name
        project_name = project_name__.split(' /')[1]          #get project name only
        OEM = project_name__.split(' /')[0]                   #get OEM name
        loadcase__ = h3[1].text
        loadcase = loadcase__.split(' /')[0]                               #get load case

        div_2 = soup.find('div',{'class' : 'col-lg-4','style':'vertical-align: middle;'})
        h4 = div_2.find_all('h4')
        design_loop__ = h4[2].text                           #get design loop infor
        design_loop = design_loop__.split(':')[-1]           #remove prefix for design loop infor

        div_3 = soup.find_all('div',{'class' : 'col-lg-12'})
        pre = div_3[5].find('pre')
        position = pre.text.split('\n')[2]                 #get seat adjustment position infor
        pulse__ = pre.text.split('\n')[1]                    #get pulse infor
        dummy__ = pre.text.split('\n')[0]                    #get dummy infor
        pulse = pulse__.split(':')[-1]                     #remove prefix for pulse infor
        dummy = dummy__.split(':')[-1]                      #remove prefix for dummy infor

        div_4 = soup.find_all('div',{'class' : 'col-md-6'})
        status1 = div_4[0].find('h4')
        status2 = div_4[1].find('h4')
        integrity = status1.text                        #get integrity infor(OK/NOK)
        specs = status2.text                            #get specs infor(OK/NOK)
        
        div_5 = soup.find('div',{'class' : 'nav-tabs-custom'})
        seat__ = div_5.text.upper()
        seat = re.search(r'\d\d?W[MP]P?',seat__)           ##get seat version infor
        if seat:
            seatversion = seat.group()                       #get seat version according to pattern number*(1or2)+W+P or M
        else: seatversion = 'unknown'                     #return unknown if found nothing.                     
        
        criteria = []
        for tr in soup.find_all('tr'):
            values_criteria = [td.text for td in tr.find_all('td')]
            criteria.append(values_criteria)                #Get criteria and its value into a list 'criteria' 

        

        df = pd.DataFrame(criteria)
    #############Seperate position into HA and TRK position##
        position__ = position.split(' : ')[1]
        if position__: ##if function continue if position__ are empty or not.
            xx = position__.split(' = ')[1]
            self.TRK_position = xx.split(',')[0]
            self.HA_position = xx.split(',')[1]
        else:
            self.TRK_position = "Unknown"
            self.HA_position = "Unknown"
    ##########################################################
        def pd_cleanup(df_orig):
            
            df_orig[0].replace(regex=[r'^\d+$',r'^\s+$'], value = np.nan, inplace=True)  #replace criteria name 'empty' or 'only numbers' to np.nan
            df_clean = df_orig[df_orig[0].notnull()]                                      #remove missing data row if criteria have no name
            df_short = df_clean.loc[:, 0:1]                                   #get data for criteria and values only.
            df_short.loc[:,0] = df_short.loc[:,0].astype(str).str.replace(u'\xa0', '')          # remove strange white space in header
            df_short.loc[:,1] = df_short.loc[:,1].astype(str).str.replace(' ', '')            #clean up numbering format, eg: 1 000
            df_short.loc[:,1] = df_short.loc[:,1].astype(float).abs()                           #get absulute data value
            df_short.reset_index(drop=True,inplace=True)                                              #re_index
            return df_short
    ##########################################################
        def rename_labels(df):
            df.columns = ['Items','Values']
            return df                                                    #rename lables names
    ##########################################################
        df_cleaned = pd_cleanup(df)
        df_renamed = rename_labels(df_cleaned)                           #use functions defined above to get df format we want.
        
        df_renamed.drop_duplicates(subset='Items', keep='first',inplace = True)  ## remove duplicated line(duplicated criterias)

        new_row = pd.DataFrame({'Items':['RunID','OEM','project_name','seatversion','loadcase','dummy','design_loop','TRK_position','HA_position','pulse','integrity','specs'],
                                'Values':[RunID,OEM,project_name,seatversion,loadcase,dummy,design_loop,self.TRK_position,self.HA_position,pulse,integrity,specs]})   ##create a new df with MIT run infor
        merged_df = pd.concat([new_row,df_renamed],ignore_index = True)    #Merge two data frame. reindex from 0.
        return merged_df

    ##***********************************************************************************##
    

    def merge_on_items(self,left,right):                                                      #define a function merge to dataframe according to 'criteria'
        merged_on_items = pd.merge(left,right,on=['Items'],how='outer')
        return merged_on_items

    def generate_xml(self):
        df_merged_all = reduce(self.merge_on_items,self.data_frames)               #loop over all runs, using reduce function.
        df_merged_all.columns = df_merged_all.iloc[0]                    #REINDEX column lables

        ###########################################################################
        df_temp = df_merged_all.copy()            ##make a copy of df
        df_temp.drop_duplicates(inplace = True)  ## remove full duplicated line(duplicated criterias)

        df_T = df_temp.T  # do transport of df

        df_T.columns = df_T.iloc[0] 
        df_T = df_T[1:]             # set 1st row as header

        df_T.columns = pd.io.parsers.ParserBase({'names':df_T.columns})._maybe_dedup_names(df_T.columns)  ##rename duplicated lable name if exist
        df_T[df_T.eq(0)] = np.nan  # set 0 valus to NAN
        df_T.dropna(axis='columns',how='all',inplace=True)          # remove column if only have value=0 or empty

        for col in df_T.columns[12:]:
            if len(df_T[col].unique()) == 1:                         # unique=2 means one value + others = empty
                df_T.drop(col,inplace=True,axis=1)                    # drop columns which only have one single value.
                
        # datetime object containing current date and time
        now = datetime.now()
        # dd/mm/YY H:M:S
        dt_string = now.strftime("%d-%m-%Y_%H%M%S")

        ################# Using xlsxwriter to output excel with 'format as table'#############
        # declare directory where we want export xlsx file.
        filepath_ = 'THC_summary_raw_data_' + dt_string + '.xlsx'
        filepath = os.path.join(self.direc_path, filepath_)

        # os.makedirs(os.path.dirname(filepath),exist_ok=True)
        writer = pd.ExcelWriter(filepath, engine='xlsxwriter') 
        df_T.to_excel(writer, sheet_name='FEA', index=False,header = True)  # output THC summary.T excel to the working folder.

        # worksheet is instance of Excel sheet "FEA" - used for inserting the table
        worksheet = writer.sheets['FEA']
        # workbook is instance of whole book - used i.e. for cell format assignment 
        workbook = writer.book

        header_cell_format = workbook.add_format()
        header_cell_format.set_rotation(90)
        header_cell_format.set_align('center')
        header_cell_format.set_align('vcenter')

        # create list of dicts for header names
        #  (columns property accepts {'header': value} as header name)
        col_names = [{'header': col_name} for col_name in df_T.columns]

        # add table with coordinates: first row, first col, last row, last col; 
        #  header names or formating can be inserted into dict 
        worksheet.add_table(0, 0, df_T.shape[0], df_T.shape[1] - 1, {
            'columns': col_names,
            # 'style' = option Format as table value and is case sensitive 
            # (look at the exact name into Excel)
            'style': 'Table Style Medium 10',
            'name': 'FEA'  # name table as 'FEA' for powerBI
        })

        writer.save()   ## export the excel file !
        return filepath
