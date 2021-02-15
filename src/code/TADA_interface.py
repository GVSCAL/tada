# --------------------------------------------------------------------------------------
# Important thing to know about Streamlit:                                              |
#   Any time something must be updated on the screen, Streamlit just reruns your entire |
#   Python script from top to bottom                                                    |
# --------------------------------------------------------------------------------------

import streamlit as st
import numpy as np
import matplotlib.pyplot as plt
import os
import io
from datetime import datetime

from DataGrasper import *
from DataTransferer import *
from GraphGenerator import *
import xlsxwriter
import SessionState
import requests
import xlsxwriter.exceptions
import random

import pandas as pd
from configparser import ConfigParser

# streamlit components
from pandas_profiling import ProfileReport
from streamlit_pandas_profiling import st_profile_report


class StInterface():
    """This class displays Streamlit user interface
    """
    def __init__(self):
        self.multi_runIDs = []
        self.all_runIDs = []
        self.file_runid_list = []
    
        self.cb_loop = []
        self.cb_graph_type = []
        self.cb_uncommon_criteria = []

        self.loop_list = []
        self.uncommon_criteria = []

        self.fig_list = []

        self._searched = False
        self._generated = False
        self.canGenerate = False
        self.cb_view_table = False
        self.multiSelectKey = 0
        self.multiSelectKey2 = 0

        self.uploaderKey = 0
        self.uploaderKey2 = 0
        self.uploaderKey3 = 0
        self.tmp_excel_path = None
        self.regular_excel_path = None
        self.set_working_dir()

        
    
    def set_working_dir(self):
        """This function sets working directory to the location where source code file locates
        """
        file_path = r'%s' % os.path.abspath(__file__)
        working_dir_path = os.path.dirname(file_path)
        os.chdir(working_dir_path)
        print('working_dir_path : ', os.getcwd())

    def setSearched(self, state):
        """This function sets search state. 

        :param bool state: A boolean value.
        """
        self._searched = state
    def getSearchedState(self):
        """This function returns search state, an instance variable which represents if 'Search' button has been clicked or not

        :return: A boolean value
        :rtype: bool
        """
        return self._searched


    def setGenerated(self, state):
        """This function sets generate state. If button 'Generate graph' is clicked, then the generate state will be set to true.

        :param bool state: Receive a boolean value.
        """
        self._generated = state
    def getGeneratedState(self):
        """This function returns generate state, an instance variable which represents if graphs has been generated or not

        Returns:
            bool: returns a boolean value
        """
        return self._generated

    def interface_mainPage(self):
        """This function displays the basic elements to the interface
        """
        import streamlit.components.v1 as components

        cfg = ConfigParser()
        cfg.read('../config.ini')
        
        direc_path = cfg.get('user_setting','directory_path')
        direc_path = direc_path.replace('\\', '/')

        if self.page == 'Main Page':
            st.title('THC Automated Display Analysis')
        else:
            st.title(self.page)

        with st.beta_expander("How to use?"):
            if self.page == 'TADA Based on Excel':
                st.markdown(f"- Select a existing Excel file with desired columns, then select RunIDs to add. Click **'Search Online'**. It will generate new Excel which contains added runIDs.")
                st.markdown(f'- Excels will be created in **{direc_path}**') 
            else:
                st.write('This tool can make result summary for a bunch of RunIDs based on results searched from [MIT report website](http://frbriunil007.bri.fr.corp/dashboard/MIT_reports.php).')
                st.markdown(f"- Select your txt file (Runids), then click **'Search Online'**.\nYou can also click **Generate PDF** to get quick charts.")
                st.markdown(f'- Excels and PDFs will be created in **{direc_path}**')          
            components.html(
                """
                <button id="btn" onclick="
                    var dummy = document.createElement('textarea');
                    document.body.appendChild(dummy);
                    dummy.value = '%s';
                    dummy.select();
                    dummy.setSelectionRange(0, 99999);
                    document.execCommand('copy');
                    alert('Path copied: %s');
                    document.body.removeChild(dummy);">Copy path</button>
                """ % (direc_path, direc_path),
                height=30
            )

    def interface_profilingPage(self):
        st.title(self.page)
        with st.beta_expander("Introduction"):
            st.write('This profiling tool can make quick data analysis of a existing dataframe')
            st.write('For each column the following statistics - if relevant for the column type - are presented in an interactive HTML report:')
            st.markdown('- **Type inference**: detect the types of columns in a dataframe.')
            st.markdown('- **Essentials**: type, unique values, missing values')
            st.markdown('- **Quantile statistics** like minimum value, Q1, median, Q3, maximum, range, interquartile range')
            st.markdown('- **Descriptive statistics** like mean, mode, standard deviation, sum, median absolute deviation, coefficient of variation, kurtosis, skewness')
            st.markdown('- **Most frequent values**')
            st.markdown('- **Histograms**')
            st.markdown('- **Correlations** highlighting of highly correlated variables, Spearman, Pearson and Kendall matrices')
            st.markdown('- **Missing values** matrix, count, heatmap and dendrogram of missing values')
            st.markdown('- **Duplicate rows** Lists the most occurring duplicate rows')
            st.markdown('- **Text analysis** learn about categories (Uppercase, Space), scripts (Latin, Cyrillic) and blocks (ASCII) of text data')


            st.write('For further information, visit [this page](https://pandas-profiling.github.io/pandas-profiling/docs/master/rtd/).')

        st.write('')
        with st.beta_expander('How to use?'):
            st.write("Select your XLSX file, then click **'Generate Profiling'**")
        st.subheader('Choose a XLSX file')


    def display_sidebar_widget(self):
        """This function displays sidebar widgets
        """
        st.sidebar.image('../pic/logo_gvs - cut.jpg', width=250)
        self.page = st.sidebar.selectbox('Page',options=['Main Page','Compare RunIDs','Quick Data Analysis','TADA Based on Excel'])
        
        st.sidebar.header('Options')
        self.nb_per_page = st.sidebar.select_slider('Number of graphs per page', options = list(np.arange(6)+1), value=6)
        self.ph_cbViewTable = st.sidebar.empty()

        st.sidebar.header('Useful links')
        st.sidebar.write('<a href="https://www.faurecia.com" target="_blank"><dir style="background-color:#ffffff; padding:10px 10px"><img src="https://www.faurecia.com/sites/groupe/files/logo%402x.png" width="60%"></dir></a>', unsafe_allow_html = True)
        st.sidebar.write('<a href="http://frbriunil007.bri.fr.corp/dashboard/MIT_reports.php" target="_blank" style="color: white; font-size:25px; text-decoration: none;"><dir style="background-color:#D73925; padding:10px 10px; "><b>MIT Report</b></dir></a>', unsafe_allow_html = True)
        st.sidebar.write('<a href="https://faurus.ww.faurecia.com/community/thehub" target="_blank"><dir style="background-color:#003684; padding:10px 10px"><img src="https://faurus.ww.faurecia.com/9.0.1.1597bde/resources/images/palette-1022/customNavLogoImage-1570090458285-faurus-logos.png" width="50%"></dir></a>', unsafe_allow_html = True)

        

    def display_profiling(self):
        uploaded_file = st.file_uploader('Choose a XLSX file which contains data to analyse', type=['xlsx'],accept_multiple_files=False, key = self.uploaderKey2)
        if uploaded_file:
            df = pd.read_excel(uploaded_file)
            st.dataframe(df)
        
        if st.button('Generate profiling'):
            if not uploaded_file:
                st.error('Please upload a file')
                st.stop()
            pr = ProfileReport(df, explorative=True)
            st_profile_report(pr)

        
        

    def get_uploaded(self, old_upload_len, old_id_list, last_current_runIDs_value):
        """This function gets selected runIDs and adjust the display of related widgets

        :param old_upload_len: Previous uploaded file length
        :type old_upload_len: int
        :param old_id_list: Previous runID list in multi-select widget
        :type old_id_list: list
        :param last_current_runIDs_value: Temporary variable to store previous runID list. When a runID is removed, this variable will be sent to the widget
        :type last_current_runIDs_value: list
        :Returns:
            - len(uploaded_file): Number of ujploaded file 
            - self.multi_runIDs: List of selected runIDs in the multi-select widget 
            - current_runIDs: Temporary list variable which contains runIDs.
        :rtype: int, list, list
        """
        print(self.page)
        if  self.page == 'Main Page' or self.page == 'Compare RunIDs':
            print('Refreshing uploaded files...')
            print('old_id_list\n', old_id_list)
            multiple_files = True
            st.subheader('Choose a txt file')
            print('self.uploaderKey:',self.uploaderKey)
            uploaded_file = st.file_uploader("Choose a txt file which contains RunIDs", type=['txt'], accept_multiple_files=multiple_files, key = self.uploaderKey)
            file_runid_list = self.file_runid_list
            input_add_list = []

            if uploaded_file:
                string_data = ""
                for up_file in uploaded_file:
                    # To convert to a string based IO
                    stringio = io.StringIO(up_file.read().decode("utf-8"))

                    # To read file as string
                    string_data += stringio.read()
                    
                file_runid_list = self.toRunidList(string_data)
            
            print('old file:',old_upload_len,'now file:',len(uploaded_file))

            

        st.subheader('Add other runIDs here if needed')
        text_input = st.text_area("RunIDs")


        if st.button('Add to list'):
            if text_input!='':
                input_add_list = self.toRunidList(text_input)
            if not input_add_list:
                st.error('No valid runID')
                st.stop()


        print('file_runid_list:\n',file_runid_list)
        print('input add list:\n',input_add_list)

        
        # One file added
        if old_upload_len < len(uploaded_file):
            current_runIDs = sorted(list(set().union(old_id_list, file_runid_list)))
            
        # One file deleted
        elif old_upload_len > len(uploaded_file):
            current_runIDs = sorted(list(set().union(input_add_list, file_runid_list)))
        # File unchanged
        else:
            # Button 'add to list' clicked
            if (input_add_list):
                self.multiSelectKey += 1    # if there are new items added by input text, increment the multiselect widget key
                current_runIDs = sorted(list(set().union(self.multi_runIDs, input_add_list)))
            else:
                print("other button or 'x' clicked")
                current_runIDs = last_current_runIDs_value      # we don't change the value of current_runIDs(which will be sent as default value for multiselect) so that multiselect widget will ignore this variable


        self.all_runIDs = sorted(list(set().union(current_runIDs, self.all_runIDs)))
        
        
        # If new items added, we should increment the multiselect widget key, in order to delete the previous widget
        if (len(last_current_runIDs_value) < len(current_runIDs)):
            self.multiSelectKey += 1

        print('all runids:\n',self.all_runIDs)
        print('last_current_runIDs_value:\n',last_current_runIDs_value)
        print('current runids:\n',current_runIDs)
        print('self.multi_runIDs before:\n', self.multi_runIDs)
        self.multi_runIDs = st.multiselect('Selected runIDs:\n', self.all_runIDs, current_runIDs,key=self.multiSelectKey)    # Here we pass the current key number
        print('self.multi_runIDs after:\n', self.multi_runIDs)
        
        st.text(f'Total: {len(self.multi_runIDs)} RunIDs')

        if len(uploaded_file)!=old_upload_len or len(self.multi_runIDs)!=len(old_id_list):
            self._searched = False

        return len(uploaded_file), self.multi_runIDs, current_runIDs



    def add_new_runID(self, old_id_list):
        input_runIDs = []
        st.subheader('Choose a XLSX file')
        uploaded_file = st.file_uploader('Choose a XLSX file which contains data to analyse', type=['xlsx'],accept_multiple_files=False, key = self.uploaderKey3)
        if uploaded_file:
            self.exist_df = pd.read_excel(uploaded_file)
            st.dataframe(self.exist_df)

        st.subheader('Select runIDs here')
        text_input = st.text_area("RunIDs")

        if st.button('Add to list'):
            if text_input!='':
                input_runIDs = self.toRunidList(text_input)
            if not input_runIDs:
                st.error('No valid runID')
                st.stop()
        print('input runids:',input_runIDs)
        print('old id list:', old_id_list)

        if (input_runIDs):
            self.multiSelectKey2 += 1    # if there are new items added by input text, increment the multiselect widget key
        
        current_runIDs = sorted(list(set().union(input_runIDs,old_id_list)))
        self.multi_runIDs = st.multiselect('Selected runIDs:\n',current_runIDs,current_runIDs, key=self.multiSelectKey2)    # Here we pass the current key number
        print('self.multiSelect:', self.multi_runIDs)

        return self.multi_runIDs


    def toRunidList(self,runid_string):
        """This function converts a string to a list of RunIDs
        
        :param str runid_string: Input string
        :return: A list of runIDs
        :rtype: list
        """
        txt_orig = runid_string.upper()                                   ## convert to upper case
        runid_list = re.findall(r'\w\w[12]\d{5}\d?',txt_orig)           ## find runids into a list according to search pattern
        runid_list = list(dict.fromkeys(runid_list))                      # remove duplicated runs.
        return runid_list


    # @st.cache(suppress_st_warning=True)
    def search_online(self):
        """This function handles online search interface
        """
        if not self.multi_runIDs:
            st.error('No runID selected')
            return
        if self.page == 'TADA Based on Excel' and not hasattr(self, 'exist_df'):
            st.error('Please select a existing Excel file.')
            st.stop()
        self.setSearched(True)
        grasper = DataGrasper()
        try:
            grasper.search_online_by_runID(self.multi_runIDs)
        except requests.ConnectionError:
            st.error('Connection error: Please check your network environment')
            st.stop()
        except FooException as e:
            st.error(f'Can not find runID: {e.runID}')
            st.stop()
        
        else:
            st.success('Searching successful, generating Excels...')
        
        with st.spinner('Generating temporary excel file...'):
            tmp_excel_path = grasper.generate_xml()
        if not self.page == 'TADA Based on Excel':
            transferer = DataTransferer(raw_file_name = tmp_excel_path)
        else:
            transferer = DataTransferer(raw_file_name = tmp_excel_path, exist_df=self.exist_df)
        
        with st.spinner('Generating regular excel file...'):
            regular_excel_path = transferer.generate_reg_excel()
            all_criteria, uncommon_criteria = transferer.getInfo()

            self.all_criteria = all_criteria
            self.uncommon_criteria = uncommon_criteria
        
        try:
            with st.spinner('Updating database...'):
                transferer.concatenate_to_db(self.multi_runIDs)
        except xlsxwriter.exceptions.FileCreateError as e:
            st.warning(f'{e}\n\nTo save history, please make sure the database excel is closed.')
        except PermissionError as e:
            st.warning(f'{e}\n\nTo save history, please make sure the database excel is closed.')
        else:
            st.success('Excel generated successfully')
        self.tmp_excel_path, self.regular_excel_path = tmp_excel_path, regular_excel_path

    def display_excel_path(self):
        """This function displays the path of generated excel
        """
        st.info(f'Temporary excel path: {self.tmp_excel_path}')
        st.info(f'Regular excel path: {self.regular_excel_path}')


    def display_default_loop(self, path):
        """This function displays run loops as checkbox widgets
        
        :param str path: Regular excel path
        """
        try:
            with st.beta_expander("Select loops"):
                c1, c2 = st.beta_columns((1,10))
                
                self.loop_list = self.get_all_loop(path)   
                self.cb_loop.clear()
                for i in range(len(self.loop_list)):
                    self.cb_loop.append(c2.checkbox(self.loop_list[i], True, key=i))
        except FileNotFoundError:  
            st.error(f'Can not find file {path}.\nPlease make sure the file exists or recheck your connection.')     
            st.stop()


    def get_all_loop(self, path):
        """This function returns a list of all loops presented in the excel

        :param str path: Path of the regular excel
        :return: list of all loops
        :rtype: list
        """
        # create generator object
        self.gen = GraphGenerator(path)
        df_selected = self.gen.df_origin[['OEM','project_name','design_loop']]
        dic_loop = df_selected.groupby(['design_loop']).apply(list).to_dict()
        
        # self.clear_all_loop()
        loop_list = list(dic_loop.keys())
        return loop_list

    def display_common_creterias(self):
        """This function displays common creterias as checkbox widgets
        """
        with st.beta_expander("Select graph types"):
            c1, c2 = st.beta_columns((1,10))
            graph_types = ['Status','Belt bracket on track','Longitudinal load','Recliner torque','Front bracket load','Rear brackets load']
            self.cb_graph_type.clear()
            for i in range(len(graph_types)):
                self.cb_graph_type.append(c2.checkbox(graph_types[i], True, key=i))

    def display_uncommon_criterias(self):
        """This function displays uncommon creterias as checkbox widgets
        """
        with st.beta_expander("Select other columns"):
            c1, c2 = st.beta_columns((1,10))
            print(self.uncommon_criteria)
            self.cb_uncommon_criteria.clear()
            uc_list = self.uncommon_criteria
            for i in range(len(uc_list)):
                self.cb_uncommon_criteria.append(c2.checkbox(uc_list[i], True, key=i))


    def verifyCanGenerate(self):
        """This function checks if it's possible to generate charts

        :return: Returns a boolean
        :rtype: boolean
        """
        result = True
        if (sum(self.cb_loop) == 0):
            st.warning('Please select at least one loop.')
            result = False

        if (sum(self.cb_graph_type) == 0):
            st.warning('Please select at least one graph.')
            result = False
        
        return result

    def show_excel_data(self, page):
        """This function print the regularized dataframe to the interface

        :param page: Current page name
        :type page: string
        """

        self.cb_view_table = self.ph_cbViewTable.checkbox('View Excel table')
        if self.cb_view_table:
            with st.beta_expander("Regular excel table", expanded=True):
                with st.spinner('Displaying dataframe...'):
                    if page=='Main Page':
                        st.write(self.gen.df_origin) 
                    elif page=='Compare RunIDs':
                        df = self.gen.df_origin
                        df = df.set_index('RunID',inplace=False)
                        df = df.T
                        st.write(df)
            
        


    def generate_charts(self, page):
        """This function generates a PDF which contains all generated graphs
        """
        uc_list = self.uncommon_criteria
        max_per_page = self.nb_per_page

        otheritems_list = []
        selected_loop = []


        for i in range(len(self.cb_loop)):
            if self.cb_loop[i]:
                selected_loop.append(self.loop_list[i])


        for i in range(len(self.cb_uncommon_criteria)):
            if self.cb_uncommon_criteria[i]:
                otheritems_list.append(uc_list[i])   

        # for debug

        # st.write('self.cb_graph_type:',self.cb_graph_type)
        # st.write('selected_loop:',selected_loop)
        # st.write('otheritems_list:',otheritems_list)
        # st.write('max_per_page:',max_per_page)
        # st.write(self.gen.df_origin)



        try:
            print("max per page", max_per_page)
            self.fig_list, save_path, _ = self.gen.generate_pdf(self.cb_graph_type, selected_loop, otheritems_list, page, max_per_page = max_per_page)
            self._generated = True

        except PermissionError as e:
            st.warning(f"PermissionError: {e}")
            return
        

        self.mode = self.gen.mode
        st.success("PDF File generated successfully")
        st.info(f'PDF path: {save_path}')

    def plot_graphs(self):
        """This function draws charts to interface
        """

        st.info(f'Total: {len(self.fig_list)} graphs, mode: {self.mode}')
        with st.beta_expander('Graphs', expanded=True): 
            for graph in self.fig_list:
                st.pyplot(graph)



    def incrementUploader(self):
        """This function increment the inner key field for file uploader widget
        """
        self.uploaderKey += 1

    def incrementUploader2(self):
        """This function increment the inner key field for file uploader widget
        """
        self.uploaderKey2 += 1

    def initialize(self):
        """This function initialize the Streamlilt interface
        """
        tmp_multiSelectKey = self.multiSelectKey     # Store old keys
        tmp_uploaderKey = self.uploaderKey
        self.__init__()                             # Initialize all class variables
        self.multiSelectKey = tmp_multiSelectKey   # Restore old keys
        self.uploaderKey = tmp_uploaderKey 
        print('Interface initialized!')



class DataStorage():
    """This class stores global variables, which we don't want lose them during every user action
    """
    def __init__(self):
        self.upload_length = 0
        self.last_current_runIDs = []
        self.id_list = []
        self.last_id2=[]
        self.page = 'Main Page'

    def initialize(self):
        print('initialize dataStorage')
        self.__init__()                 # Initialize all fields
        




if __name__ == "__main__":
    MainWindow = StInterface()
    dataStore = DataStorage()
    state = SessionState.get(count = random.random(), interface =  MainWindow, data = dataStore)

    print('Session state: ', state.count)

    interface = state.interface             # Get current state object
    dataStore = state.data                  # Get current state object
    
    interface.display_sidebar_widget()

    print('dataStore.page:',dataStore.page)
    print('interface.page:',interface.page)


    share_data_page = ['Main Page', 'Compare RunIDs']

    if st.sidebar.button('Restart') or not ((interface.page in share_data_page and dataStore.page in share_data_page) or (interface.page == dataStore.page)):
        st.balloons()
        dataStore.initialize()
        interface.initialize()
        interface.incrementUploader()
        interface.incrementUploader2()

    if interface.page != dataStore.page:
        interface.setGenerated(False)

    if (interface.page == 'Main Page' or interface.page == 'Compare RunIDs' or interface.page == 'TADA Based on Excel'):
        interface.interface_mainPage()

        if not interface.page == 'TADA Based on Excel':
            upload_len, curr_id_list, current_runIDs= interface.get_uploaded(dataStore.upload_length, dataStore.id_list, dataStore.last_current_runIDs)
            # update dataStore
            dataStore.upload_length = upload_len
            dataStore.id_list = curr_id_list
            dataStore.last_current_runIDs = current_runIDs
            print('dataStore.id_list', dataStore.id_list)
        else:
            current_id = interface.add_new_runID(dataStore.last_id2)
            dataStore.last_id2 = current_id

        # bottons for click
        if st.button("Search"):
            interface.setGenerated(False)
            with st.spinner('Searching runIDs online ...'):     
                interface.search_online()

        
        if interface.getSearchedState():
            interface.display_excel_path()
            if not interface.page == 'TADA Based on Excel':
                interface.display_default_loop(interface.regular_excel_path) 
                interface.display_common_creterias()
                interface.display_uncommon_criterias()
                
                if interface.verifyCanGenerate():
                    interface.show_excel_data(interface.page)
                    if st.button('Generate Graphs'):
                        with st.spinner('Generating graphs...'):       
                            interface.generate_charts(interface.page)
            
        if interface.getGeneratedState():
            interface.plot_graphs()

        
    elif (interface.page == 'Quick Data Analysis'):
        interface.interface_profilingPage()
        interface.display_profiling()


    print('****************')
    dataStore.page = interface.page
    state.data = dataStore      # Update state
    state.interface = interface     # Update state
    state.count += 1
