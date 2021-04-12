import re
import pandas as pd 
import numpy as np
import os 
from datetime import datetime
import sys
import xlsxwriter
#from MakePivotTable import *

from configparser import ConfigParser
from pathlib import Path


class DataTransferer():
    """This class analyse all column names in the raw excel file and generates a standardized excel with standard columns
    """
    def __init__(self,raw_file_name = r'df2.xlsx', exist_df= pd.DataFrame()):
        cfg = ConfigParser()
        cfg.read('../config.ini')
        
        self.direc_path = cfg.get('user_setting','directory_path')
        database_path = cfg.get('debug','database_path')
        Path(database_path).mkdir(parents=True, exist_ok=True)

        db_excelname = 'MIT_database.xlsx'
        self.db_path = os.path.join(database_path, db_excelname)
        if not os.path.isfile(self.db_path):
            workbook = xlsxwriter.Workbook(self.db_path)
            workbook.add_worksheet('FEA')
            workbook.add_worksheet('History__')
            self.df_History__ = pd.DataFrame(columns=['time', 'runID'])

            header_cell_format = workbook.add_format()
            header_cell_format.set_rotation(90)
            header_cell_format.set_align('center')
            header_cell_format.set_align('vcenter')

            workbook.close()
        else:
            self.df_History__ = pd.read_excel(self.db_path, sheet_name='History__')

        self.df_database = pd.read_excel(self.db_path, sheet_name='FEA')

        # contains names of all basic columns
        self.basic_info_list = ['RunID','OEM','project_name','seatversion','loadcase',
                            'dummy','design_loop','TRK_position','HA_position','pulse','integrity','specs']
        
        # raw_file_name is the user input, template_filename is the template df.
        # the goal is to send raw_file_name information to template_filename(template with standard criteria names)
        self.df1 = pd.DataFrame(columns = self.basic_info_list)
        self.df1.columns = self.df1.columns.str.strip()  #remove white space in each column nameprint(df1)

        self.exist_df = exist_df

        # filename = r'df2.xlsx'
        self.df2 = pd.read_excel(raw_file_name)
        self.df2.columns = self.df2.columns.str.strip()  #remove white space in each column nameprint(df2)

        # All keywords are in this dictionary. 
        # For exemple, ['latch','DS'] means a column should match 2 keywords 'latch' and 'DS' at the same time
        ds_list = ['DS', 'Door', 'outer', 'outter', 'external', 'outboard']
        ts_list = ['TS','tunnel','inner','internal','inboard']
        self.dict_keywords = {'Latch force DS': self.create_regex_dict_keywords_two(['latch','HDM'],ds_list),
                       'Latch force TS': self.create_regex_dict_keywords_two(['latch','HDM'],ts_list),
                       'Recliner torque DS': self.create_regex_dict_keywords_three(['recliner'],['torque'],ds_list),
                       'Recliner torque TS': self.create_regex_dict_keywords_three(['recliner'],['torque'],ts_list),
                       'Recliner axial force DS': self.create_regex_dict_keywords_three(['recliner'],['axial','axis','axial','axis'],ds_list),
                       'Recliner axial force TS': self.create_regex_dict_keywords_three(['recliner'],['axial','axis','axial','axis'],ts_list),
                       'Belt displacement DS': self.create_regex_dict_keywords_three(['Belt','anchor'],['displacement','dis','disp'],ds_list),
                       'Belt displacement TS': self.create_regex_dict_keywords_three(['Belt','anchor'],['displacement','dis','disp'],ts_list),
                       'Rear bracket force DS': self.create_regex_dict_keywords_three(['rear','re'],['pivot','bracket'],ds_list),  
                       'Rear bracket force TS': self.create_regex_dict_keywords_three(['rear','re'],['pivot','bracket'],ts_list), 
                       'Front bracket force DS': self.create_regex_dict_keywords_three(['front','fr'],['pivot','bracket'],ds_list),
                       'Front bracket force TS': self.create_regex_dict_keywords_three(['front','fr'],['pivot','bracket'],ts_list),
                       'Belt bracket force': [['Belt', 'bracket','Force'],['BBB']], # need to verify
                       'Lap bracket force': [['Lap', 'bracket','Force'],['Lap', 'belt','Force']],  # need to verify
                       'HA torque': self.create_regex_dict_keywords_two(['HA','Nano','Epump','E-pump'],['torque']),
                       'Backrest dynamic angle DS': self.create_regex_dict_keywords_four(['Backrest'],['angle'],['dynamic','dyna','dyn'],ds_list),
                       'Backrest dynamic angle TS': self.create_regex_dict_keywords_four(['Backrest'],['angle'],['dynamic','dyna','dyn'],ts_list),
                       'Backrest static angle DS': self.create_regex_dict_keywords_four(['Backrest'],['angle'],['static','stat'],ds_list),
                       'Backrest static angle TS': self.create_regex_dict_keywords_four(['Backrest'],['angle'],['static','stat'],ts_list),
                       'Backrest deflection DS': self.create_regex_dict_keywords_three(['Backrest'],['deflection'],ds_list),
                       'Backrest deflection TS': self.create_regex_dict_keywords_three(['Backrest'],['deflection'],ts_list),
                       'Backrest x displ' : self.create_regex_dict_keywords_three(['backrest'],['displacement','dis','disp'],['x']),# need to verify
                       'PELVIS DX': [['PELVIS', 'x', 'dis'],['PELVIS', 'Dx']],  # need to verify
                       'PELVIS DZ': [['PELVIS', 'z', 'dis'],['PELVIS', 'Dz']],  # need to verify
                       'Tilt Axial Force': [['Tilt', 'Axial'],['Tilt', 'axis']], # need to verify
                       'Tilt Shear Force': [['Tilt', 'Shear'],['Tilt', 'Share']], # need to verify
                       'Track sliding DS': self.create_regex_dict_keywords_two(['sliding'],ds_list),
                       'Track sliding TS': self.create_regex_dict_keywords_two(['sliding'],ts_list),
                       'Upper profile section force TS': self.create_regex_dict_keywords_two(['PUPP section force','profile section force','profile_section_force'],ts_list),
                       'Upper profile section force DS': self.create_regex_dict_keywords_two(['PUPP section force','profile section force','profile_section_force'],ds_list),
                       'Jack pulling force - TRK sub': self.create_regex_dict_keywords_two(['pulling','jack'],['force','beam']),
                       'Pullig displacement - TRK sub': self.create_regex_dict_keywords_two(['pulling','jack'],['dis','displacement']),
                       'Ultimate torque - HA-Rec sub': self.create_regex_dict_keywords_two(['ultimate','ulti'],['torque']),
                       'Pinion torque': self.create_regex_dict_keywords_two(['pinion'],['torque']),
                       'OrbitingGear force - HA sub': self.create_regex_dict_keywords_two(['OrbitingGear'],['force']),
                       'Leadscrew force axial - e-tilt sub': self.create_regex_dict_keywords_three(['Leadscrew'],['force'],['axial','axis']),
                       'Leadscrew force radial - e-tilt sub': self.create_regex_dict_keywords_three(['Leadscrew'],['force'],['radi','radial']),
                       'Res_F C_ROLLER_1-C_CAM': self.create_regex_dict_keywords_three(['Resultant contact force'],['COMMAND_ROLLER_1'],['COMMAND_CAM']),
                       'Res_F C_ROLLER_2-C_CAM': self.create_regex_dict_keywords_three(['Resultant contact force'],['COMMAND_ROLLER_2'],['COMMAND_CAM']),
                       'Res_F C_ROLLER_3-C_CAM': self.create_regex_dict_keywords_three(['Resultant contact force'],['COMMAND_ROLLER_3'],['COMMAND_CAM']),
                       'Res_F C_ROLLER_4-C_CAM': self.create_regex_dict_keywords_three(['Resultant contact force'],['COMMAND_ROLLER_4'],['COMMAND_CAM']),
                       'Res_F C_ROLLER_5-C_CAM': self.create_regex_dict_keywords_three(['Resultant contact force'],['COMMAND_ROLLER_5'],['COMMAND_CAM']),
                       'Res_F C_ROLLER_6-C_CAM': self.create_regex_dict_keywords_three(['Resultant contact force'],['COMMAND_ROLLER_6'],['COMMAND_CAM']),
                       'Res_F L_ROLLER_10-PINION': self.create_regex_dict_keywords_three(['Resultant contact force'],['LOCK_ROLLER_10'],['PINION']),
                       'Res_F L_ROLLER_9-PINION': self.create_regex_dict_keywords_three(['Resultant contact force'],['LOCK_ROLLER_9'],['PINION']),
                       'Res_F L_ROLLER_8-PINION': self.create_regex_dict_keywords_three(['Resultant contact force'],['LOCK_ROLLER_8'],['PINION']),
                       'Res_F L_ROLLER_7-PINION': self.create_regex_dict_keywords_three(['Resultant contact force'],['LOCK_ROLLER_7'],['PINION']),
                       'Res_F L_ROLLER_6-PINION': self.create_regex_dict_keywords_three(['Resultant contact force'],['LOCK_ROLLER_6'],['PINION']),
                       'Res_F L_ROLLER_5-PINION': self.create_regex_dict_keywords_three(['Resultant contact force'],['LOCK_ROLLER_5'],['PINION']),
                       'Res_F L_ROLLER_4-PINION': self.create_regex_dict_keywords_three(['Resultant contact force'],['LOCK_ROLLER_4'],['PINION']),
                       'Res_F L_ROLLER_3-PINION': self.create_regex_dict_keywords_three(['Resultant contact force'],['LOCK_ROLLER_3'],['PINION']),
                       'Res_F L_ROLLER_2-PINION': self.create_regex_dict_keywords_three(['Resultant contact force'],['LOCK_ROLLER_2'],['PINION']),
                       'Res_F L_ROLLER_1-PINION': self.create_regex_dict_keywords_three(['Resultant contact force'],['LOCK_ROLLER_1'],['PINION']),
                       'Res_F L_ROLLER_10-L_RING': self.create_regex_dict_keywords_three(['Resultant contact force'],['LOCK_ROLLER_10'],['LOCKER_RING']),
                       'Res_F L_ROLLER_9-L_RING': self.create_regex_dict_keywords_three(['Resultant contact force'],['LOCK_ROLLER_9'],['LOCKER_RING']),
                       'Res_F L_ROLLER_8-L_RING': self.create_regex_dict_keywords_three(['Resultant contact force'],['LOCK_ROLLER_8'],['LOCKER_RING']),
                       'Res_F L_ROLLER_7-L_RING': self.create_regex_dict_keywords_three(['Resultant contact force'],['LOCK_ROLLER_7'],['LOCKER_RING']),
                       'Res_F L_ROLLER_6-L_RING': self.create_regex_dict_keywords_three(['Resultant contact force'],['LOCK_ROLLER_6'],['LOCKER_RING']),
                       'Res_F L_ROLLER_5-L_RING': self.create_regex_dict_keywords_three(['Resultant contact force'],['LOCK_ROLLER_5'],['LOCKER_RING']),
                       'Res_F L_ROLLER_4-L_RING': self.create_regex_dict_keywords_three(['Resultant contact force'],['LOCK_ROLLER_4'],['LOCKER_RING']),
                       'Res_F L_ROLLER_3-L_RING': self.create_regex_dict_keywords_three(['Resultant contact force'],['LOCK_ROLLER_3'],['LOCKER_RING']),
                       'Res_F L_ROLLER_2-L_RING': self.create_regex_dict_keywords_three(['Resultant contact force'],['LOCK_ROLLER_2'],['LOCKER_RING']),
                       'Res_F L_ROLLER_1-L_RING': self.create_regex_dict_keywords_three(['Resultant contact force'],['LOCK_ROLLER_1'],['LOCKER_RING']),
                       'Res_F Pinion-Collar-1': self.create_regex_dict_keywords_three(['Resultant contact force'],['Pinion'],['Collar-1']),
                       'Res_F Pinion-Collar-2': self.create_regex_dict_keywords_three(['Resultant contact force'],['Pinion'],['Collar-2']),
                       'Torque MX on pinions section': self.create_regex_dict_keywords_two(['Torque MX'],['pinions section']),
                       'pinion_ends_1_MZ': self.create_regex_dict_keywords_two(['MZ'],['pinion_ends_1']),
                       'pinion_ends_2_MZ': self.create_regex_dict_keywords_two(['MZ'],['pinion_ends_2']),
                       'pinion_ends_3_MZ': self.create_regex_dict_keywords_two(['MZ'],['pinion_ends_3']),
                       'pinion_ends_4_MZ': self.create_regex_dict_keywords_two(['MZ'],['pinion_ends_4']),
                       'pinion_ends_5_MZ': self.create_regex_dict_keywords_two(['MZ'],['pinion_ends_5']),
                       'pinion_ends_6_MZ': self.create_regex_dict_keywords_two(['MZ'],['pinion_ends_6']),
                       'pinion_teeth_1_MZ': self.create_regex_dict_keywords_two(['MZ'],['pinion_teeth_1']),
                       'pinion_teeth_2_MZ': self.create_regex_dict_keywords_two(['MZ'],['pinion_teeth_2']),
                       'pinion_teeth_3_MZ': self.create_regex_dict_keywords_two(['MZ'],['pinion_teeth_3']),
                       'pinion_teeth_4_MZ': self.create_regex_dict_keywords_two(['MZ'],['pinion_teeth_4']),
                       'pinion_teeth_5_MZ': self.create_regex_dict_keywords_two(['MZ'],['pinion_teeth_5']),
                       'pinion_teeth_6_MZ': self.create_regex_dict_keywords_two(['MZ'],['pinion_teeth_6']),
                       'pinion_teeth_7_MZ': self.create_regex_dict_keywords_two(['MZ'],['pinion_teeth_7']),
                       'pinion_teeth_8_MZ': self.create_regex_dict_keywords_two(['MZ'],['pinion_teeth_8']),
                       'pinion_teeth_9_MZ': self.create_regex_dict_keywords_two(['MZ'],['pinion_teeth_9']),
                       'pinion_teeth_10_MZ': self.create_regex_dict_keywords_two(['MZ'],['pinion_teeth_10']),
                       'pinion_teeth_11_MZ': self.create_regex_dict_keywords_two(['MZ'],['pinion_teeth_11']),
                       'pinion_OG_teeth_1_MZ': self.create_regex_dict_keywords_two(['MZ'],['pinion_OG_teeth_1']),
                       'pinion_OG_teeth_2_MZ': self.create_regex_dict_keywords_two(['MZ'],['pinion_OG_teeth_2']),
                       'pinion_OG_teeth_3_MZ': self.create_regex_dict_keywords_two(['MZ'],['pinion_OG_teeth_3']),
                       'pinion_OG_teeth_4_MZ': self.create_regex_dict_keywords_two(['MZ'],['pinion_OG_teeth_4']),
                       
                       }

        # dictionary of regex
        self.dict_regex = {}
        # write all keyword in a regex form. 
        # For exemple, r"((?=.*latch)(?=.*DS))|((?=.*latch)(?=.*Door))|..."
        for key, value in self.dict_keywords.items():
            regex = r''
            for item in value:
                regex += r'('
                for word in item:
                    regex += r'(?=.*'+word+r')'
                regex += r')'
                regex += r'|'
            regex = regex.rstrip('|')
            self.dict_regex[key] = regex

        self.common_criteria = ['Latch force DS','Latch force TS','recliner torque TS']


    def create_regex_dict_keywords_two(self,keywords1,keywords2):
        """This class creates a list of possible keyword combination. It select one elements from each argument and combine them into a new list, then appends the list to keywords_list

        :param list keywords1: First keyword list
        :param list keywords2: Second keyword list
        :type keywords2: [type]
        :return: List which contains lists of keyword combination
        :rtype: list
        """
        keywords_list = []
        for keyword1 in keywords1 :
            for keyword2 in keywords2:
                keywords_list.append([keyword1,keyword2])
        return keywords_list

    def create_regex_dict_keywords_three(self,keywords1,keywords2,keywords3):
        """See function create_regex_dict_keywords_two
        """
        keywords_list = []
        for keyword1 in keywords1 :
            for keyword2 in keywords2:
                for keyword3 in keywords3:
                    keywords_list.append([keyword1,keyword2,keyword3])
        return keywords_list

    def create_regex_dict_keywords_four(self,keywords1,keywords2,keywords3,keywords4):
        """See function create_regex_dict_keywords_two
        """
        keywords_list = []
        for keyword1 in keywords1 :
            for keyword2 in keywords2:
                for keyword3 in keywords3:
                    for keyword4 in keywords4:
                        keywords_list.append([keyword1,keyword2,keyword3,keyword4])
        return keywords_list

    # send data from df2 to df1, according to column name
    def send_data(self,df1_column_name,df2_column_name):  
        """This function send data from raw dataframe to regular dataframe, by providing their correspond column names

        :param str df1_column_name: Column name of the regular dataframe
        :param str df2_column_name: Column name of the raw dataframe
        """
        print(self.df1)
        print('columns:',self.df1.columns)
        if df1_column_name not in self.df1.columns:
            # if column name does not exist, create an empty column
            self.df1[df1_column_name] = np.nan
            print('create empty column:',df1_column_name)
            print('col1:', df1_column_name)
            print('col2:', df2_column_name)

        self.df1[df1_column_name] = self.df1[df1_column_name].combine_first(self.df2[df2_column_name])  # combine 2 columns together
            
    def concatenate_to_db(self, runID_list):   
        """Concatenate search information to database after each search is finished

        :param list runID_list: List of runID searched
        """
        # Concatenate result at the front of DB
        self.df_database = pd.concat([self.df1, self.df_database], axis=0, ignore_index=True)
        # If duplicates, keep the first (newer) occurance
        self.df_database = self.df_database.drop_duplicates(subset=['RunID'], keep='first')
        writer = pd.ExcelWriter(self.db_path, engine='xlsxwriter') 
        self.df_database.to_excel(writer, sheet_name='FEA', index=False,header = True)

        now = datetime.now()
        dt_string = now.strftime("%d-%m-%Y %H:%M:%S")
        new_History__ = {'time': dt_string, 'runID': ','.join(str(e) for e in runID_list).rstrip(',')}

        self.df_History__ = self.df_History__.append(new_History__, ignore_index=True)
        self.df_History__.to_excel(writer, sheet_name='History__', index=False,header = True)

        worksheet1 = writer.sheets['FEA']
        col_names = [{'header': col_name} for col_name in self.df_database.columns]
        worksheet1.add_table(0, 0, self.df_database.shape[0], self.df_database.shape[1] - 1, {
            'columns': col_names,
            # 'style' = option Format as table value and is case sensitive 
            # (look at the exact name into Excel)
            'style': 'Table Style Medium 9',
            'name': 'FEA'  
        })

        worksheet2 = writer.sheets['History__']
        col_names2 = [{'header': col_name} for col_name in self.df_History__.columns]
        worksheet2.add_table(0, 0, self.df_History__.shape[0], self.df_History__.shape[1] - 1, {
            'columns': col_names2,
            # 'style' = option Format as table value and is case sensitive 
            # (look at the exact name into Excel)
            'style': 'Table Style Medium 9',
            'name': 'History__'  
        })
        worksheet2.set_column(0,0,18)
        worksheet2.set_column(1,1,80)
        writer.save()
        print('concatenated')

    # match column according to regex
    def update_df1_according_to_match(self):
        """This function match column using regular expression, then send matched content from raw dataframe to regular ddataframe
        
        :return: Updated regular dataframe
        :rtype: pandas.Dataframe
        :return: List of messages generated during match process
        :rtype: list
        """
        msg_list = []
        for key in self.dict_regex:
            print("Searching column:",key,"...")
            regex = self.dict_regex[key]
            print('\t>regex:', regex)
            matched = False
            for df2_column_name in self.df2.columns:
                if re.match(regex, df2_column_name, re.I):   
                    print("\t> Column matched!")
                    print("\t> Column founded: ",df2_column_name)
                    self.send_data(key,df2_column_name)
                    matched = True
                    
            if not matched:
                msg =  key +" not found!"
                print("\t=> ",msg)
                msg_list.append(msg)    
        
        return self.df1, msg_list

    def send_basic_info(self):
        """This function send basic information from raw dataframe to regular dataframe
        
        :return: Updated regular dataframe
        :rtype: pandas.Dataframe
        """
        for basic_info in self.basic_info_list:
            self.send_data(basic_info,basic_info)
        return self.df1

    def getAllCriterias(self):
        """This function returns all columns except basic columns from the regular dataframe
        
        :return: List which contains column names
        :rtype: list
        """
        # +1 because we should consider 'load_case_short_name' column that is added
        all_criteria = self.df1.columns[len(self.basic_info_list) + 1:].to_list()
        return all_criteria

    def getUncommonCriterias(self,all_criteria):
        """This function returns uncommon column names which is not stored in our column name dictionary
        
        :param list all_criteria: List which contains all column names except basic column names
        :return: List which contains all uncommon columns
        :rtype: list
        """
        #this returns uncommon criterias
        uncommon_criterias = list(set(all_criteria) - set(self.common_criteria))
        return uncommon_criterias

    # this returns two list of criterias. one is for common criterias. 2nd is the other criterias.
    def getInfo(self):
        """This function calls update_df1_according_to_match
        
        :Returns: 
            - all_criteria: List of all column names except basic column name
            - uncommon_criterias: List of all uncommon column names which is not stored in our column
            
        :rtype: list, list
        """
        
        all_criteria = self.getAllCriterias()
        uncommon_criterias = self.getUncommonCriterias(all_criteria)
        uncommon_criterias = sorted(uncommon_criterias)
        return all_criteria, uncommon_criterias

    def generate_reg_excel(self):
        """This function generates regular excel

            :return: Path of regular excel generated
            :rtype: str
        """
        from GraphGenerator import dfToDict
        import random 
        now = datetime.now()
        print('now:', now)
        # dd/mm/YY H:M:S
        dt_string = now.strftime("%d-%m-%Y_%H%M%S")
        print('dt_string:', dt_string)
        
        filepath_ = 'THC_summary_regular_excel_' + dt_string + '.xlsx'
        filepath = os.path.join(self.direc_path,filepath_)
        print('filepath:', filepath)

        self.df1, msg_list = self.update_df1_according_to_match()
        self.df1 = self.send_basic_info()

        
        
        def find_key_for(input_dict, value):    
            matched = '_'
            for k in input_dict.keys():
                if k == value:
                    matched = input_dict[k]
            return matched

        def rename_loadcase(loadcase):
            loadcase_dict = {
                'Luggage crash': 'LUG',
                'Rear Crash': 'RC',
                'ECE14': 'ECE14',
                'Front Crash': 'FC',
                'FMVSS202a': 'FMVSS202',
                'Lateral Crash': 'LC',
                'Whiplash': 'Whiplash',
                'Z Crash': 'Z Crash',
                'ECE17': 'ECE17',
                'ECE21': 'ECE21',
                'IFX Trans -': 'IFX',
                'IFX Trans +': 'IFX',
                'TopTether': 'IFX',
                'RR pulling (Mech)': 'RR pull',
                'FR pulling (Mech)': 'FR pull',
                }
            loadcase_short = find_key_for(loadcase_dict,loadcase)
            return loadcase_short

        def rename_dummy(dummy):
            dummy_dict = {
                ' D95': '95',
                ' D50': '50',
                ' D05' : '05',
                ' E14': ' ',
                ' IFX' : 'IFX',
                ' BRD' : 'BRD',
                }
            dummy_short = find_key_for(dummy_dict,dummy)
            return dummy_short

        def rename_trkposition(trkposition):
            trkposition_dict = {
                'Tracks:rear most - 1 notch': 'Rm-1n',
                'Tracks:rear most': 'Rm',
                'Tracks:middle': 'Mp',
                'Tracks:front most': 'Fm',
                'Tracks:front most - 1 notch': 'Fm-1n' }
            trkposition_short = find_key_for(trkposition_dict,trkposition)
            return trkposition_short

        def rename_HAposition(HAposition):
            HAposition_dict = {
                ' HA:lower most': 'Dm',
                ' HA:middle': 'Mp',
                ' HA:upper most': 'Um' ,
                ' HA:no_adjm': '' ,
                'Unknown': '' ,}
            HAposition_short = find_key_for(HAposition_dict,HAposition)
            return HAposition_short

        def getLoadcase_full_name(row):
            loadcase_short = rename_loadcase(row['loadcase'])
            dummy_short = rename_dummy(row['dummy'])
            trkposition_short = rename_trkposition(row['TRK_position'])
            HAposition_short = rename_HAposition(row['HA_position'])
            seat_version = row['seatversion']
            loadcase_name_full_short = loadcase_short + dummy_short + ' ' + trkposition_short + HAposition_short
            return loadcase_name_full_short
        
        last_basic_info_column_id = 12


        print(self.df1)


        if not 'loadcase_short_name' in self.df1:
            self.df1.insert(last_basic_info_column_id ,'loadcase_short_name',value='null')
        self.df1['loadcase_short_name'] = self.df1.apply(getLoadcase_full_name, axis=1)

        last_basic_info_column_id += 1  # as we added new column(full loadcase name)
        

        # if in add RunID mode
        if not self.exist_df.empty:
            print('self.exist_df:', self.exist_df)
            #self.df1 = pd.concat([self.exist_df, self.df1], axis=0, ignore_index=True, join='inner')
            common_columns = self.exist_df.columns.intersection(self.df1.columns)
            self.df1 = pd.concat([self.exist_df,self.df1[common_columns]])
            print('merged:\n', self.df1)
          

        os.makedirs(os.path.dirname(filepath),exist_ok=True)

        # Create a Pandas Excel writer using XlsxWriter as the engine.
        writer = pd.ExcelWriter(filepath, engine='xlsxwriter') 
        self.df1.to_excel(writer, sheet_name='FEA', index=False,header = True)
        
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
        col_names = [{'header': col_name} for col_name in self.df1.columns]

        # add table with coordinates: first row, first col, last row, last col; 
        #  header names or formating can be inserted into dict 
        worksheet.add_table(0, 0, self.df1.shape[0], self.df1.shape[1] - 1, {
            'columns': col_names,
            # 'style' = option Format as table value and is case sensitive 
            # (look at the exact name into Excel)
            'style': 'Table Style Medium 10',
            'name': 'FEA'  # name table as 'FEA' for powerBI
        })

        ## creatfve excel sheet for each column
        colors = ['#E41A1C', '#377EB8', '#4DAF4A', '#984EA3', '#FF7F00']
        count_row = self.df1.shape[0]
        RunID_column_id = self.df1.columns.get_loc("RunID")
        all_columns = self.df1.columns.values.tolist()
        selected_columns = all_columns[last_basic_info_column_id:]

        
        for column in selected_columns :
            
            df_basic_info = self.df1.iloc[:,:last_basic_info_column_id].copy()
            # df_basic_info['loadcase_short_name'] = df_basic_info.apply(getLoadcase_full_name, axis=1)
            df_select_column = self.df1[column].copy()
            df_sheet = pd.concat([df_basic_info, df_select_column], axis=1, sort=False)
            
            if "loadcase_short_name" in df_sheet.columns:
                loadcase_short_name_column_id = df_sheet.columns.get_loc("loadcase_short_name")

            sheet_name = column

            xs, ys = dfToDict(df_sheet,'dummy',column)
            length_xs = len(xs)

            d = { 'dummy_mean': xs, column + ' mean' : ys }
            df_mean = pd.DataFrame(data=d)
            df_sheet_2 = pd.concat([df_sheet, df_mean], axis=1, sort=False)
            df_sheet_2.to_excel(writer, sheet_name=sheet_name,index=False,header = True)

            # Access the XlsxWriter workbook and worksheet objects from the dataframe.
            workbook  = writer.book
            worksheet = writer.sheets[sheet_name]

           

            # Configure the series of the chart from the dataframe data.
            if "loadcase_short_name" in df_sheet.columns:
                 # Create a chart object.
                chart = workbook.add_chart({'type': 'column'})
                chart.add_series({
                    'name':       [sheet_name, 0, last_basic_info_column_id],
                    'categories': [sheet_name, 1, loadcase_short_name_column_id, count_row, loadcase_short_name_column_id],
                    'values':     [sheet_name, 1, last_basic_info_column_id, count_row, last_basic_info_column_id],
                    'fill':       {'color':  random.choice(colors)},
                    'overlap':    -5,
                })

                # Configure the chart axes.
                x_axis_name = 'Load case'
                y_axis_name = column
                chart.set_x_axis({'name': x_axis_name})
                chart.set_y_axis({'name': y_axis_name , 'major_gridlines': {'visible': False}})

                # Insert the chart into the worksheet.
                worksheet.insert_chart('O8', chart)

    ##########################################################################33
            # Configure the series of the chart from the dataframe data.
            if "loadcase_short_name" in df_sheet.columns:
                ## Create a 2nd chart object.
                chart = workbook.add_chart({'type': 'column'})
                chart.add_series({
                    'name':       [sheet_name, 0, last_basic_info_column_id],
                    'categories': [sheet_name, 1, last_basic_info_column_id + 1, length_xs, last_basic_info_column_id + 1],
                    'values':     [sheet_name, 1, last_basic_info_column_id + 2, length_xs, last_basic_info_column_id + 2],
                    'fill':       {'color':  random.choice(colors)},
                    'overlap':    -5,
                })

                # Configure the chart axes.
                x_axis_name2 = 'dummy'
                y_axis_name2 = column
                chart.set_x_axis({'name': x_axis_name2})
                chart.set_y_axis({'name': y_axis_name2 , 'major_gridlines': {'visible': False}})

                # Insert the chart into the worksheet.
                worksheet.insert_chart('O24', chart)

        # Close the Pandas Excel writer and output the Excel file.
        writer.save()

        print('reg excel saved at path :',filepath)
        
#        run_excel(filepath, 'FEA')

        return filepath

