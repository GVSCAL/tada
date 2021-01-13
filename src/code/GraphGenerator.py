import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
import numpy as np
import textwrap
import os
from typing import Tuple
from matplotlib.backends.backend_pdf import PdfPages
from datetime import datetime
import sys
sys.path.insert(1, os.path.abspath(os.path.join(os.getcwd(),'..')))
from pdfrw import PdfReader, PdfWriter, PageMerge
from configparser import ConfigParser
plt.style.use('ggplot')

# x/y label names
RT_DS = 'Recliner torque - DS [Nm]'
RT_TS = 'Recliner torque - TS [Nm]'

FBF_DS = 'DS Front bracket force [kN]'
FBF_TS = 'TS Front bracket force [kN]'

LAP_FORCE = 'Lap bracket force'

BFD_DS = 'Belt Fixation DS displacement [mm]'

# all graph keywords
KW_LOADCASE = 'loadcase_short_name'
KW_RUNID = 'RunID'
KW_latch_DS = 'Latch force DS'
KW_latch_TS = 'Latch force TS'  

KW_Belt_disp_DS = 'Belt displacement DS'       
KW_Belt_disp_TS = 'Belt displacement TS' 
KW_recliner_torque_DS = 'Recliner torque DS'      
KW_recliner_torque_TS = 'Recliner torque TS'      

KW_Front_Bracket_Force_DS = 'Front bracket force DS'
KW_Front_Bracket_Force_TS = 'Front bracket force TS'

KW_Belt_Bracket_Force = 'Belt bracket force'

KW_Rear_Bracket_Force_DS = 'Rear bracket force DS'
KW_Rear_Bracket_Force_TS = 'Rear bracket force TS'



class GraphGenerator():
    """This class contains functions to draw various types of graphs
    """
    def __init__(self,filepath,design_loop = ''):
        print("PATH:",filepath)

        cfg = ConfigParser()
        cfg.read('../config.ini')
        
        self.direc_path = cfg.get('user_setting','directory_path')

        
        self.df_origin = pd.read_excel(filepath)
        self.df_origin.columns = self.df_origin.columns.str.strip()  #remove white space in each column name
        self.df_origin = self.df_origin.apply(lambda x: x.str.strip() if x.dtype == "object" else x)


    def two_pie_chart(self,column_name1 = 'integrity', column_name2 = 'specs'):
        """Function which returns a fig object of a two pie chart
        
        :param column_name1: Column name of 'integrity', defaults to 'integrity'
        :type column_name1: str, optional
        :param column_name2: Column name of 'specification', defaults to 'specs'
        :type column_name2: str, optional
        :return: Generated figure object
        :rtype: matplotlib.pyplot.figure
        """
        fig = plt.figure()
        fig.add_subplot(121)
        pie_chart(self.df,column_name1)
        fig.add_subplot(122)
        pie_chart(self.df,column_name2)

        oem, pjt_name, _ = self.basic_info()
        oem_str, pjt_name_str = 'Oem: ','Project: '
        for s in oem:
            oem_str+=s+', '
        for s in pjt_name:
            pjt_name_str+=s+', '
        oem_str = oem_str.strip(', ')
        pjt_name_str = pjt_name_str.strip(', ')
        plt.text(0.70,0.10,oem_str, transform=fig.transFigure, size=10)
        plt.text(0.70,0.05,pjt_name_str, transform=fig.transFigure, size=10)
        plt.subplots_adjust(bottom=0.30)
        return fig

        
    def belt_bracket(self,dataframe,column_type='loadcase_short_name', column_bar=KW_Belt_Bracket_Force,column_line=BFD_DS, loadcase_name = ""):
        """This function plots graph for 'belt bracket'

        :param dataframe: Regular dataframe
        :type dataframe: pandas.Dataframe
        :param column_type: Column name for the x axis, defaults to 'loadcase_short_name'
        :type column_type: str, optional
        :param column_bar: Column name for the bar chart., defaults to KW_Belt_Bracket_Force
        :type column_bar: str, optional
        :param column_line: Column name for the line chart., defaults to BFD_DS
        :type column_line: str, optional
        :param loadcase_name: Load case name (only for multi loop mode, defaults to ""
        :type loadcase_name: str, optional
        :return: Generated figure object
        :rtype: matplotlib.pyplot.figure
        """
        xs,ys = dfToDict(dataframe,column_type,column_bar)
        if not self.mode=='multiple loop':
            title = "Belt Bracket on track"
            # fig = plt.figure()

            # plt.bar(xs,ys,alpha=.7,color='dodgerblue',edgecolor = "k", label=column_bar)
            # plt.legend(prop={'size':8},loc='upper right')
            # plt.ylabel("Load[kN]")
           
            # plt.ylim(bottom = 0)    # set coordinate limits
            # plt.title(title,fontsize= 'large' , pad = 0)

            # x_axis = range(len(xs))
            # plt.xticks(x_axis, [textwrap.fill(label, 8) for label in xs], 
            # rotation = 90, fontsize=8, horizontalalignment="center")
            # plt.tight_layout(pad=1.0)           # makes space on the figure canvas for the labels
            # plt.tick_params(axis='x', pad=6)
            fig = self.single_bar_chart(dataframe, column_type, column_bar, title)
            return fig
        else:
            title = "Belt Bracket on track ("+ loadcase_name + ")"
            fig = self.single_line_chart(dataframe, column_type, column_bar, title)
            return fig



    # def belt_bracket_compare(self,dataframe,column_type='loadcase_short_name',column_line=BFD_DS, title = "Belt Bracket on track"):
    #     xs_line,ys_line = dfToDict(dataframe,column_type,column_line)

    #     fig = plt.figure()

    #     # ax1 = fig.add_subplot(111)
    #     plt.plot(xs_line,ys_line,alpha=.7,color='dodgerblue',edgecolor = "k", label=column_line)
    #     plt.legend(prop={'size':8},loc='upper right')
    #     plt.ylabel("Load[kN]")
     
    #     plt.ylim(bottom = 0)    # set coordinate limits
    #     plt.title(title,fontsize= 'large' , pad = 0)
        
    #     x_axis = range(len(xs_line))
    #     plt.xticks(x_axis, [textwrap.fill(label, 8) for label in xs_line], 
    #     rotation = 90, fontsize=8, horizontalalignment="center")
    #     plt.tight_layout(pad=1.0)           # makes space on the figure canvas for the labels
    #     plt.tick_params(axis='x', pad=6)
        
    #     return fig

    def longitudinal_load(self,dataframe:pd.core.frame.DataFrame,abs_column_name='loadcase_short_name', doorside_column_name='Latch Outer force', tunnelside_column_name='Latch Inner force', loadcase_name = '') -> matplotlib.pyplot.figure:
        """This function plots graph for 'longitudinal load'

        :param dataframe: Regular dataframe
        :type dataframe: pandas.Dataframe
        :param abs_column_name: Column name for x axis, defaults to 'loadcase_short_name'
        :type abs_column_name: str, optional
        :param doorside_column_name: Column name for first bar, defaults to 'Latch Outer force'
        :type doorside_column_name: str, optional
        :param tunnelside_column_name: Column name for second bar, defaults to 'Latch Inner force'
        :type tunnelside_column_name: str, optional
        :param loadcase_name: Load case name (only for multi loop mode), defaults to ''
        :type loadcase_name: str, optional
        :return: Generated figure object
        :rtype: matplotlib.pyplot.figure
        """

        title = "Longitudinal Load" if not self.mode=='multiple loop' else "Longitudinal Load ("+ loadcase_name + ")"
        if not self.mode=='multiple loop':
            title = "Longitudinal Load" 
            chart_type = 'bar'
        else:
            title = "Longitudinal Load ("+ loadcase_name + ")"
            chart_type = 'line'
        fig = self.double_value_chart(dataframe,abs_column_name, doorside_column_name, tunnelside_column_name,title,abs_column_name,'Load[kN]', chart_type=chart_type)
        # plt.axhline(y=20, color='red', linestyle='--')
        return fig


    def recliner_torque(self,dataframe:pd.core.frame.DataFrame,abs_column_name='loadcase_short_name', doorside_column_name='recliner torque DS', tunnelside_column_name='recliner torque TS', loadcase_name = '')-> matplotlib.pyplot.figure:    
        """This function plots graph for 'recliner torque'

        :param dataframe: Regular dataframe
        :type dataframe: pd.core.frame.DataFrame
        :param abs_column_name: Column name for x axis, defaults to 'loadcase_short_name'
        :type abs_column_name: str, optional
        :param doorside_column_name: Column name for first bar, defaults to 'recliner torque DS'
        :type doorside_column_name: str, optional
        :param tunnelside_column_name: Column name for second bar, defaults to 'recliner torque TS'
        :type tunnelside_column_name: str, optional
        :param loadcase_name: Load case name (only for multi loop mode), defaults to ''
        :type loadcase_name: str, optional
        :return: Generated figure object
        :rtype: matplotlib.pyplot.figure
        """
        if not self.mode=='multiple loop':
            title = "Recliner Torque" 
            chart_type = 'bar'
        else:
            title = "Longitudinal Load ("+ loadcase_name + ")"
            chart_type = 'line'
        fig = self.double_value_chart(dataframe,abs_column_name, doorside_column_name, tunnelside_column_name,title,abs_column_name,'Torque[N.m]', chart_type=chart_type)
            
        # plt.axhline(y=2000, color='red', linestyle='--')
        return fig

    def front_brackets_load(self,dataframe:pd.core.frame.DataFrame, abs_column_name = 'loadcase_short_name', doorside_column_name = FBF_DS , tunnelside_column_name= FBF_TS, loadcase_name = '')-> matplotlib.pyplot.figure:
        """This function plots graph for 'front brackets load'

        :param dataframe: Regular dataframe
        :type dataframe: pd.core.frame.DataFrame
        :param abs_column_name: Column name for x axis, defaults to 'loadcase_short_name'
        :type abs_column_name: str, optional
        :param doorside_column_name: Column name for first bar, defaults to FBF_DS
        :type doorside_column_name: str, optional
        :param tunnelside_column_name: Column name for second bar, defaults to FBF_TS
        :type tunnelside_column_name: str, optional
        :param loadcase_name: Load case name (only for multi loop mode), defaults to ''
        :type loadcase_name: str, optional
        :return: Generated figure object
        :rtype: matplotlib.pyplot.figure
        """

        if not self.mode=='multiple loop':
            title = "Front Brackets Load" 
            chart_type = 'bar'
        else:
            title = "Front Brackets Load ("+ loadcase_name + ")"
            chart_type = 'line'
        fig = self.double_value_chart(dataframe,abs_column_name, doorside_column_name, tunnelside_column_name, title, abs_column_name,'Load[kN]', chart_type=chart_type)
        return fig
      
    def rear_brackets_load(self, dataframe:pd.core.frame.DataFrame, abs_column_name = 'loadcase_short_name', doorside_column_name = KW_Rear_Bracket_Force_DS, tunnelside_column_name = KW_Rear_Bracket_Force_TS, loadcase_name = '')-> matplotlib.pyplot.figure:
        """This function plots graph for 'rear brackets load'

        :param dataframe: Regular dataframe
        :type dataframe: pd.core.frame.DataFrame
        :param abs_column_name: Column name for x axis, defaults to 'loadcase_short_name'
        :type abs_column_name: str, optional
        :param doorside_column_name: Column name for first bar, defaults to KW_Rear_Bracket_Force_DS
        :type doorside_column_name: str, optional
        :param tunnelside_column_name: Column name for second bar, defaults to KW_Rear_Bracket_Force_TS
        :type tunnelside_column_name: str, optional
        :param loadcase_name: Load case name (only for multi loop mode), defaults to ''
        :type loadcase_name: str, optional
        :return: Generated figure object
        :rtype: matplotlib.pyplot.figure
        """
        if not self.mode=='multiple loop':
            title = "Rear Brackets Load" 
            chart_type = 'bar'
        else:
            title = "Rear Brackets Load ("+ loadcase_name + ")"
            chart_type = 'line'

        fig = self.double_value_chart(dataframe,abs_column_name, doorside_column_name, tunnelside_column_name, title, abs_column_name,'Load[kN]', chart_type=chart_type)
        return fig

    def double_value_chart(self,dataframe:pd.core.frame.DataFrame,abs_column_name:str, vert_column_name1:str, vert_column_name2:str,title:str,xlabel:str,ylabel:str, chart_type:str)-> matplotlib.pyplot.figure:
        """This function generate a double value (bar/line) chart

        :param dataframe: Regular dataframe
        :type dataframe: pd.core.frame.DataFrame
        :param abs_column_name: Column name for x axis
        :type abs_column_name: str
        :param vert_column_name1: Column name for first y axis
        :type vert_column_name1: str
        :param vert_column_name2: Column name for second y axis
        :type vert_column_name2: str
        :param title: Graph title name
        :type title: str
        :param xlabel: x label name
        :type xlabel: str
        :param ylabel: y label name
        :type ylabel: str
        :param chart_type: Type of the chart (bar/line)
        :type chart_type: str
        :return: Generated figure object
        :rtype: matplotlib.pyplot.figure
        """
        fig, xs = raw_double_value_chart(dataframe,abs_column_name, vert_column_name1, vert_column_name2, chart_type=chart_type)
        if not fig:
            return False
        # plt.xlabel(xlabel)
        plt.ylabel(ylabel)
        # plt.legend(prop={'size':8},loc = 'best')
        # plt.xticks([x for x in range(len(xs))],xs)  #set x labels and locations
        x_axis = range(len(xs))
        plt.xticks(x_axis, [textwrap.fill(label, 8) for label in xs], 
        rotation = 90, fontsize='small', horizontalalignment="center")
        plt.tight_layout(pad=1.0)           # makes space on the figure canvas for the labels
        plt.tick_params(axis='x', pad=6)
        plt.title(title,pad=0)
        return fig

    def single_bar_chart(self, dataframe:pd.core.frame.DataFrame,abs_column_name, vert_column_name, title = '')-> matplotlib.pyplot.figure:
        """This function generates a one bar chart

        :param dataframe: Regular dataframe
        :type dataframe: pd.core.frame.DataFrame
        :param abs_column_name: Column name for x axis
        :type abs_column_name: str
        :param vert_column_name: Column name for y axis
        :type vert_column_name: str
        :param title: Graph title name, defaults to ''
        :type title: str, optional
        :return: Generated figure object
        :rtype: matplotlib.pyplot.figure
        """
        xs_bar,ys_bar = dfToDict(dataframe,abs_column_name,vert_column_name)
        print('xs bar',xs_bar)
        print('ys bar',ys_bar)


        for i in range(len(xs_bar) - 1, -1, -1):
                if np.isnan(ys_bar[i]):
                    xs_bar.pop(i)
                    ys_bar.pop(i) 


        if not self.mode=='multiple loop':
            if len(xs_bar) <= 1:
                return False 
        else:
            if len(xs_bar) <= 1:
                return False
                
        fig = plt.figure()
        plt.bar(xs_bar,ys_bar,alpha=.8,color='dodgerblue',edgecolor = "k",label = vert_column_name)
        plt.legend(prop={'size':8},loc=3)
        # plt.xlabel(abs_column_name)
        plt.ylabel(vert_column_name)
        x_axis = range(len(xs_bar))
        plt.xticks(x_axis, [textwrap.fill(label, 8) for label in xs_bar], 
           rotation = 90, fontsize='small', horizontalalignment="center")
        # plt.tight_layout(pad=1.0)           # makes space on the figure canvas for the labels
        plt.tick_params(axis='x', pad=4)
        
        if title:
            plt.title(title,pad=0)
        else:
            plt.title(vert_column_name,pad=0)

        print('xs_bar: ', xs_bar)
        print('ys_bar: ', ys_bar)

        max_value, min_value = 0,0
        if len(ys_bar)>0:
            max_value = max(ys_bar)
            min_value = min(ys_bar)
        
        plt.text(0.05,0.95,'MAX', transform=fig.transFigure, size=10)
        plt.text(0.05,0.90,max_value, transform=fig.transFigure, size=10)
        plt.text(0.15,0.95,'MIN', transform=fig.transFigure, size=10)
        plt.text(0.15,0.90,min_value, transform=fig.transFigure, size=10)
        return fig

  




    def single_line_chart(self, dataframe:pd.core.frame.DataFrame, abs_column_name:str, vert_column_name:str, title = '' )-> matplotlib.pyplot.figure:
        """This function generates a one line chart

        :param dataframe: Regular dataframe
        :type dataframe: pd.core.frame.DataFrame
        :param abs_column_name: Column name for x axis
        :type abs_column_name: str
        :param vert_column_name: Column name for y axis
        :type vert_column_name: str
        :param title: Graph title name, defaults to ''
        :type title: str, optional
        :return: Generated figure object
        :rtype: matplotlib.pyplot.figure
        """
        xs_line,ys_line = dfToDict(dataframe,abs_column_name,vert_column_name)
        for i in range(len(xs_line) - 1, -1, -1):
                if np.isnan(ys_line[i]):
                    xs_line.pop(i)
                    ys_line.pop(i) 


        if not self.mode=='multiple loop':
            if len(xs_line) == 0:
                return False 
        else:
            if len(xs_line) <= 1:
                return False


        fig = plt.figure()
        
        plt.plot(xs_line,ys_line,'o-',color = 'mediumblue',lw = 3,label = vert_column_name)
        plt.legend(prop={'size':8},loc=3)
        # plt.xlabel(abs_column_name)
        plt.ylabel(vert_column_name)
        x_axis = range(len(xs_line))
        plt.xticks(x_axis, [textwrap.fill(label, 8) for label in xs_line], 
           rotation = 90, fontsize='small', horizontalalignment="center")
        # plt.tight_layout(pad=1.0)           # makes space on the figure canvas for the labels
        plt.tick_params(axis='x', pad=4)


        # scale the graph
        scale = max(ys_line) - min(ys_line)
        factor = 0.2
        plt.ylim(bottom = min(ys_line)-factor*scale, top = max(ys_line)+factor*scale)
        
        if title:
            plt.title(title,pad=0)
        else:
            plt.title(vert_column_name,pad=0)

        max_value, min_value = 0,0
        if len(ys_line)>0:
            max_value = max(ys_line)
            min_value = min(ys_line)
        
        plt.text(0.05,0.95,'MAX', transform=fig.transFigure, size=10)
        plt.text(0.05,0.90,max_value, transform=fig.transFigure, size=10)
        plt.text(0.15,0.95,'MIN', transform=fig.transFigure, size=10)
        plt.text(0.15,0.90,min_value, transform=fig.transFigure, size=10)
        return fig

    # return basic information of the component
    def basic_info(self):
        """This function return basic information of the runID set (OEM, Project name, Loadcase)

        :Returns:
            - oem_list: OEM list
            - pjtname_list: Project name list
            - loadcase_list: Loadcase list
        :rtype: list, list, list
        """
        x_axis_lable = 'loadcase_short_name'
        df_selected = self.df[['OEM','project_name',x_axis_lable]]
        dic_oem = df_selected.groupby(['OEM']).apply(list).to_dict()
        dic_pjt_name = df_selected.groupby(['project_name']).apply(list).to_dict()
        dic_loadcase = df_selected.groupby([x_axis_lable]).apply(list).to_dict()

        oem_list = [k for k in dic_oem.keys()]
        oem_list = list(filter(None, oem_list))
        pjtname_list = [k for k in dic_pjt_name.keys()]
        pjtname_list = list(filter(None, pjtname_list))
        loadcase_list = [k for k in dic_loadcase.keys()]
        loadcase_list = list(filter(None, loadcase_list))
        print('OEM list: ', oem_list)
        print('Project name list: ', pjtname_list)
        print('Loadcase list: ', loadcase_list)
        return oem_list, pjtname_list, loadcase_list

    # function to combine multiple PDF pages into one page
    def combine_pages(self,srcpages):
        """This function combines multiple pages of a PDF file in to one page 
        """
        SCALE = 0.5
        srcpages = PageMerge() + srcpages
        print(srcpages.xobj_box[2:])
        x_increment, y_increment = (SCALE * i for i in srcpages.xobj_box[2:])

        nb_page = len(srcpages)
        for i, page in enumerate(srcpages):
            page.scale(SCALE)
            page.x = x_increment if i & 1 else 0
            page.y = y_increment*((nb_page-1-i) // 2)
        return srcpages.render()

    # function to generate PDF file
    def generate_pdf(self,cb_selected:list,design_loop:list,otheritems:list, page, max_per_page = 6):
        """This function generates PDF file

        :param cb_selected: List of integers which contains checkbox index of selected graphs types
        :type cb_selected: list
        :param design_loop: List of strings which contains selected design loop
        :type design_loop: list
        :param otheritems: List of strings which contains uncommon column names
        :type otheritems: list
        :param page: Current page selected in the menu
        :param max_per_page: Maximum number of chart in one PDF page, defaults to 6
        :type max_per_page: int, optional
        :Returns: 
            - fig_list: List of matplotlib.pyplot.figure
            - savepath: Path of PDF generated
            - msg_list: List of strings which contains message
        :rtype: list, str, list
        """


        now = datetime.now()
        # dd/mm/YY H:M:S
        dt_string = now.strftime("%d-%m-%Y_%H%M%S")
        
        filename = 'THC_Summary_Report_' + dt_string + '.pdf'
        savepath = os.path.join(self.direc_path, filename)

        # filter design loop
        self.df = self.df_origin[self.df_origin.design_loop.isin(design_loop)]
        print(self.df)
        _,_,self.loadcase_short_name = self.basic_info()
        if len(design_loop)>1 and page == 'Main Page':
            self.mode = 'multiple loop'
        elif len(design_loop)==1 and page == 'Main Page':
            self.mode = 'single loop'
        elif page == 'Compare RunIDs':
            self.mode = 'compare RunIDs'
        

        # empty warning message list
        msg_list = []
        fig_list = []
        with PdfPages(savepath) as pdf:  # create a PDF file
            if page == 'Main Page':
                loadcase_column_name = get_found_column(self.df,KW_LOADCASE)
            elif page == 'Compare RunIDs':
                loadcase_column_name = get_found_column(self.df, KW_RUNID)
            print('page:', page)
            print('loadcase_column_name', loadcase_column_name)

            # add figure to the PDF file
            # Status selected
            if(cb_selected[0]):
                fig1 = self.two_pie_chart('integrity', 'specs')
                if fig1:
                    pdf.savefig(fig1)
                    fig_list.append(fig1)
                    plt.close(fig1)
                
            if self.mode == 'single loop' or self.mode == 'compare RunIDs':
                # Belt bracket on track selected
                if(cb_selected[1]):
                    # draw status
                    fbf_ds_column_name = get_found_column(self.df,KW_Belt_Bracket_Force)
                    bfd_ds_column_name = get_found_column(self.df,KW_Belt_disp_DS)
                    if(fbf_ds_column_name and bfd_ds_column_name):
                        fig2 = self.belt_bracket(self.df,loadcase_column_name,fbf_ds_column_name)
                        if fig2:
                            pdf.savefig(fig2)
                            fig_list.append(fig2)
                            plt.close(fig2)
                        
                # Longitudinal load selected
                if(cb_selected[2]):  
                    # draw Belt bracket on DS track
                    latch_column_name1 = get_found_column(self.df,KW_latch_DS)
                    latch_column_name2 = get_found_column(self.df,KW_latch_TS)
                    if(latch_column_name1 and latch_column_name2):
                        fig3 = self.longitudinal_load(self.df, loadcase_column_name, latch_column_name1, latch_column_name2)
                        if fig3:
                            pdf.savefig(fig3)
                            fig_list.append(fig3)
                            plt.close(fig3)
                        
                # Recliner torque selected
                if(cb_selected[3]):
                    recliner_column_name1 = get_found_column(self.df,KW_recliner_torque_DS)
                    recliner_column_name2 = get_found_column(self.df,KW_recliner_torque_TS)
                    if(recliner_column_name1 and recliner_column_name2):
                        fig4 = self.recliner_torque(self.df, loadcase_column_name, recliner_column_name1, recliner_column_name2)
                        if fig4:
                            pdf.savefig(fig4)
                            fig_list.append(fig4)
                            plt.close(fig4)
                        
                # Front bracket load selected
                if(cb_selected[4]):
                    fbf_ds_column_name = get_found_column(self.df,KW_Front_Bracket_Force_DS)
                    fbf_ts_column_name = get_found_column(self.df,KW_Front_Bracket_Force_TS)
                    if(fbf_ds_column_name and fbf_ts_column_name):
                        fig5 = self.front_brackets_load(self.df, loadcase_column_name,fbf_ds_column_name,fbf_ts_column_name)
                        if fig5:
                            pdf.savefig(fig5)
                            fig_list.append(fig5)
                            plt.close(fig5)

                        
                # Rear brackets load selected        
                if(cb_selected[5]):
                    rbf_ds_column_name = get_found_column(self.df,KW_Rear_Bracket_Force_DS)
                    rbf_ts_column_name = get_found_column(self.df,KW_Rear_Bracket_Force_TS)
                    if(rbf_ds_column_name and rbf_ts_column_name):
                        fig6 = self.rear_brackets_load(self.df, loadcase_column_name,rbf_ds_column_name,rbf_ts_column_name)
                        if fig6:
                            pdf.savefig(fig6)
                            fig_list.append(fig6)
                            plt.close(fig6)
                        

                # generate other graphs selected
                for item in otheritems:
                    fig = self.single_bar_chart(self.df, loadcase_column_name,item)
                    if fig:
                        pdf.savefig(fig)
                        fig_list.append(fig)
                        plt.close(fig)
                    
                

            # if in multi loop mode (multiple design loops selected)
            elif self.mode == 'multiple loops':
                # Belt bracket on track selected
                if(cb_selected[1]):
                    fbf_ds_column_name = get_found_column(self.df,KW_Belt_Bracket_Force)   
                    bfd_ds_column_name = get_found_column(self.df,KW_Belt_disp_DS)
                    
                    if(fbf_ds_column_name and bfd_ds_column_name):
                        for loadcase_short_name in self.loadcase_short_name:
                            dataframe = self.df[self.df.loadcase_short_name == loadcase_short_name]
                            title = "Belt Bracket on track (" + loadcase_short_name + ")"
                            fig = self.belt_bracket(dataframe,column_type='design_loop',column_bar = fbf_ds_column_name, column_line = bfd_ds_column_name, loadcase_name=loadcase_short_name)
                            if fig:
                                pdf.savefig(fig)
                                fig_list.append(fig)
                                plt.close(fig)
                                
                # Longitudinal load selected
                if(cb_selected[2]):
                    latch_column_name1 = get_found_column(self.df,KW_latch_DS)
                    latch_column_name2 = get_found_column(self.df,KW_latch_TS)
                    
                    if(latch_column_name1 and latch_column_name2):
                        for loadcase_short_name in self.loadcase_short_name:
                            dataframe = self.df[self.df.loadcase_short_name == loadcase_short_name]
                            fig = self.longitudinal_load(dataframe, 'design_loop', latch_column_name1, latch_column_name2, loadcase_name = loadcase_short_name)
                            if fig:
                                pdf.savefig(fig)
                                fig_list.append(fig)
                                plt.close(fig)
                # Recliner torque selected
                if(cb_selected[3]): 
                    recliner_column_name1 = get_found_column(self.df,KW_recliner_torque_DS)
                    recliner_column_name2 = get_found_column(self.df,KW_recliner_torque_TS)
                    
                    if(recliner_column_name1 and recliner_column_name2):
                        for loadcase_short_name in self.loadcase_short_name:
                            dataframe = self.df[self.df.loadcase_short_name == loadcase_short_name]
                            fig = self.recliner_torque(dataframe, 'design_loop', recliner_column_name1, recliner_column_name2, loadcase_short_name)
                            if fig:
                                pdf.savefig(fig)
                                fig_list.append(fig)
                                plt.close(fig)
                # Front bracket load selected
                if(cb_selected[4]): 
                    fbf_ds_column_name = get_found_column(self.df,KW_Front_Bracket_Force_DS)
                    fbf_ts_column_name = get_found_column(self.df,KW_Front_Bracket_Force_TS)
                    if(fbf_ds_column_name and fbf_ts_column_name):
                        for loadcase_short_name in self.loadcase_short_name:
                            dataframe = self.df[self.df.loadcase_short_name == loadcase_short_name]
                            title  = "Front Brackets Load ("+ loadcase_short_name + ")"
                            
                            fig = self.front_brackets_load(self.df, 'design_loop',fbf_ds_column_name,fbf_ts_column_name, loadcase_short_name)

                            if fig:
                                pdf.savefig(fig)
                                fig_list.append(fig)
                                plt.close(fig)
                # Rear brackets load selected
                if(cb_selected[5]): 
                    rbf_ds_column_name = get_found_column(self.df,KW_Rear_Bracket_Force_DS)
                    rbf_ts_column_name = get_found_column(self.df,KW_Rear_Bracket_Force_TS)
                    if(rbf_ds_column_name and rbf_ts_column_name):          
                        for loadcase_short_name in self.loadcase_short_name:
                            dataframe = self.df[self.df.loadcase_short_name == loadcase_short_name]
                            title  = "Rear Brackets Load ("+ loadcase_short_name + ")"
                            fig = self.rear_brackets_load(dataframe, 'design_loop',rbf_ds_column_name,rbf_ts_column_name, loadcase_short_name)
                            if fig:
                                pdf.savefig(fig)
                                fig_list.append(fig)
                                plt.close(fig)

                for item in otheritems:
                    for loadcase_short_name in self.loadcase_short_name:
                        dataframe = self.df[self.df.loadcase_short_name == loadcase_short_name]
                        title  = item +" ("+ loadcase_short_name + ")"
                        fig = self.single_line_chart(dataframe, 'design_loop', item, title)
                        if fig:
                            pdf.savefig(fig)
                            fig_list.append(fig)
                            plt.close(fig)

        



        # original multi-pages PDF file generated      

        # read the original multi-pages DF file
        pages = PdfReader(savepath).pages

        # overwrite the original PDF file by single page with multiple graphs
        writer = PdfWriter(savepath)
        for index in range(0, len(pages), max_per_page):
            writer.addpage(self.combine_pages(pages[index:index + max_per_page]))
        writer.write()

        return fig_list, savepath, msg_list



# draw a pie chart
def pie_chart(dataframe:pd.core.frame.DataFrame,column:str):
    """This function plot a pie chart

    :param dataframe: Regular dataframe
    :type dataframe: pd.core.frame.DataFrame
    :param column: Column name for the pie chart
    :type column: str
    """
    
    column = column.strip()
    color_dict = {"OK":'limegreen',"OK Limit":'yellow',"NOK Limit":'orange',"NOK":'r','None':'w'}
    df = dataframe[['RunID',column]]    #select columns and generate a dataframe
    dic = df.groupby([column])['RunID'].apply(list).to_dict()  #convert the dataframe to a dictionary

    group_key = []
    group_value = []
    colors = []
    for key in dic:
        group_key.append(key)
        group_value.append(len(dic[key]))
        colors.append(color_dict[key])
    group_value = np.asarray(group_value)
    percent = 100.*group_value/group_value.sum()

    def my_autopct(pct):    # function which allows to display the actual amount on the chart
        total = sum(group_value)
        val = int(round(pct*total/100.0))
        return '{v:d}'.format(v=val)

    patches, texts, _ = plt.pie(group_value, labels = group_key, colors = colors, shadow=True, textprops = {'fontsize':9, 'color':'k'}, startangle=90, autopct=my_autopct, wedgeprops = {'linewidth': 1, 'edgecolor':'k'})  #draw a pie chart
    labels = ['{0} - {1:1.2f} %'.format(i,j) for i,j in zip(group_key, percent)]
    sort_legend = True
    if sort_legend:
        patches, labels, dummy =  zip(*sorted(zip(patches, labels, group_value),key=lambda x: x[2],reverse=True))

    plt.legend(patches, labels, loc='best',bbox_to_anchor=(0.8, 0.1), fontsize=8)
    plt.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
    plt.title(column)
    

# draw a double bar chart
def raw_double_value_chart(dataframe:pd.core.frame.DataFrame,column_type:str, vert_column_name1:str, vert_column_name2:str, chart_type = 'bar')-> Tuple[matplotlib.pyplot.figure, list]:
    """Generates raw double value (bar/line) chart

    :param dataframe: Regular dataframe
    :type dataframe: pd.core.frame.DataFrame
    :param column_type: Column name for x axis
    :type column_type: str
    :param vert_column_name1: Column name for first y axis
    :type vert_column_name1: str
    :param vert_column_name2: Column name for second y axis
    :type vert_column_name2: str
    :param chart_type: Type of the chart, defaults to 'bar'
    :type chart_type: str, optional
    :Returns: 
        -fig: Generated figure object
        -xs: List of x axis value
    :rtype: matplotlib.pyplot.figure, list
    """
    
    fig = plt.figure()  #create a empty figure
    xs,ys_1 = dfToDict(dataframe,column_type,vert_column_name1)   #generate the dictionary
    xs2,ys_2 = dfToDict(dataframe,column_type,vert_column_name2)
    print('xs1:', xs)
    print('ys1:', ys_1)
    print('xs2:', xs2)
    print('ys2:', ys_2)
    print(chart_type)
    # Remove NaN values
    for i in range(len(xs) - 1, -1, -1):
            if np.isnan(ys_1[i]) or np.isnan(ys_2[i]):
                xs.pop(i)
                ys_1.pop(i) 
                ys_2.pop(i) 

    if chart_type == 'bar':
        # if no bars can be generated, return False
        if len(xs) == 0:
            return False, False
        # define x coordinate
        x = np.arange(len(xs)) 
        total_width, n = 0.9 , 2
        width = total_width / n
        x = x - (total_width - width) / 2
        plt.bar(x, ys_1, color = "dodgerblue",edgecolor = "k",width=width,label=vert_column_name1)    #draw bar chart
        plt.bar(x + width, ys_2, color = "mediumblue",edgecolor = "k",width=width,label=vert_column_name2)
        
    elif chart_type == 'line':
        # if less or equal than one line point can be generated, return False
        if len(xs) <= 1:
            return False, False

        plt.plot(xs,ys_1,'o-',color = 'dodgerblue',lw = 3,label = vert_column_name1)
        plt.plot(xs,ys_2,'o-',color = 'mediumblue',lw = 3,label = vert_column_name2)
    
    plt.legend(prop={'size': 8}, loc='best')
    # display the maximum
    max1, max2 = 0,0
    min1, min2 = 0,0
    if len(ys_1)>0:
        max1 = max(ys_1)
        min1 = min(ys_1)
    if len(ys_2)>0:
        max2 = max(ys_2)
        min2 = min(ys_2)
    plt.text(0.05,0.95,'MAX DS', transform=fig.transFigure, size=10)
    plt.text(0.05,0.90,max1, transform=fig.transFigure, size=10)
    plt.text(0.15,0.95,'MAX TS', transform=fig.transFigure, size=10)
    plt.text(0.15,0.90,max2, transform=fig.transFigure, size=10)
    if chart_type == 'line':
        scale = max(max1, max2) - min(min1, min2)
        factor = 0.2
        plt.ylim(bottom =  min(min1, min2)-factor*scale, top = max(max1, max2)+factor*scale)
    return fig, xs

# convert a data frame to a dictionary which contains the selected column
def dfToDict(dataframe:pd.core.frame.DataFrame,column_key:str,column_value:str):
    """This function convert a data frame to a dictionary which contains the selected column, then returns keys and values of the dictionary


    :param dataframe: Regular dataframe
    :type dataframe: pd.core.frame.DataFrame
    :param column_key: The column name which constitutes dictionary key
    :type column_key: str
    :param column_value: The column name which constituted dictionary value
    :type column_value: str
    :returns: 
        - xs - List of dictionary keys
        - ys - List of dictionary values
    :rtype: list, list
    """

    column_key = column_key.strip()
    column_value = column_value.strip()
    df = dataframe[[column_key,column_value]]       # temporary df which contains only 2 columns
    dic = df.groupby([column_key])[column_value].apply(list).to_dict()  # generate a dictionary whose key is the column name, and the value are all values in this column
    dic = {k.strip():v for k,v in dic.items()}  # remove blank space
    dic = {k:v for k,v in dic.items() if k != ''} # remove empty column name
    dic = {k:np.around(np.nanmean(v),2) for k,v in dic.items() if k != ''}  # calculate the average for each key
    xs = [k for k in dic.keys()]    # extract keys and values into 2 lists
    ys = [v for v in dic.values()]

    return xs,ys


def get_found_column(dataframe:pd.core.frame.DataFrame, colName: str):
    """This function search the column by keyword and return the column name. If no column name is found, then return False.

    :param dataframe: Regular dataframe
    :type dataframe: pd.core.frame.DataFrame
    :param colName: The column name that we want to search
    :type colName: str
    :return: The column name found, false if not found
    :rtype: str or bool
    """
    
    
    if colName in dataframe.columns:
        print('Find colName: ',colName)
        return colName
    
    else:
        print('Cannot find colName: ', colName)
        return False
