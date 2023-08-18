#! python 3
# -*- coding: utf-8 -*-
"""
Description of this script.
"""
# --- MODULE IMPORTS ---------------------------------------------------------#
# Example: import numpy as np

import os
import xlsxwriter
import datetime
#%%--------Input Ignored Minmaxes----------------------#
Ignore_Minimax=[]





#%%
Text_ign = ''
for mm in Ignore_Minimax:
    Text_ign= Text_ign + ' '+mm

#%% --- FUNCTION DEFINITIONS ---------------------------------------------------#
def list_shmr_files(directory):
# List csmr file in current directory
    shmr_files = []
    res_files = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith('.shmr1'):
                case_fs=file.split('mm')[1].replace('sh.res.shmr1','')
                case_fs=case_fs.split('.')[0]
                ege=0
                ege_switch = True
                for x in Ignore_Minimax:
                    ege_float = Ignore_Minimax[ege]
                    if case_fs==ege_float:
                        ege_switch=False
                        continue
                    ege=ege+1
                if ege_switch:
                    shmr = root + str("/") + file
                    res = root + str("/") + file[:len(file)-5]
                    shmr_files.append(shmr)
                    res_files.append(res)

    return shmr_files,res_files


def envelope_ULS_ALS(limit_state):
    Envelope_result_dict[dg]['env'][limit_state]['element'] = {}
    Envelope_result_dict[dg]['env'][limit_state]['node'] = {}
    Envelope_result_dict[dg]['env'][limit_state]['LC2'] = {}
    Envelope_result_dict[dg]['env'][limit_state]['angle'] = {}
    Envelope_result_dict[dg]['env'][limit_state]['Nf'] = {}
    Envelope_result_dict[dg]['env'][limit_state]['Mf'] = {}
    Envelope_result_dict[dg]['env'][limit_state]['Vf'] = {}
    Envelope_result_dict[dg]['env'][limit_state]['Vccd'] = {}
    Envelope_result_dict[dg]['env'][limit_state]['Vcd'] = {}
    Envelope_result_dict[dg]['env'][limit_state]['asv'] = str(-1)
    Envelope_result_dict[dg]['env'][limit_state]['asv_zero'] = {} 
    Envelope_result_dict[dg]['env'][limit_state]['case'] = {} 
    
def width_columns(worksheet):
    std_width = 9
    worksheet.set_column('A:A', std_width-2)
    worksheet.set_column('B:B', std_width-2)
    worksheet.set_column('C:C', std_width)
    worksheet.set_column('D:D', std_width+1)
    worksheet.set_column('E:E', std_width-2)
    worksheet.set_column('F:F', std_width+1)
    worksheet.set_column('G:G', std_width-2)
    worksheet.set_column('H:H', std_width+4)
    worksheet.set_column('I:I', std_width-2)
    worksheet.set_column('J:J', std_width)    
    worksheet.set_column('K:K', std_width)    
    worksheet.set_column('L:L', std_width)    
    worksheet.set_column('M:M', std_width)    
    worksheet.set_column('N:N', std_width)    
    worksheet.set_column('O:O', std_width)    
    
def check_envelope(dg,result,limit_state,bracket,case_dict,case):
    if bracket == 'larger':
        #If asv is larger - update envelope dictionary
        if float(case_dict[result]) > float(Envelope_result_dict[dg]['env'][limit_state][result]):
            Envelope_result_dict[dg]['env'][limit_state][result] = case_dict[result]

            Envelope_result_dict[dg]['env'][limit_state]['element'] = case_dict['element']
            Envelope_result_dict[dg]['env'][limit_state]['lc_thermal'] = case_dict['lc_thermal']
            Envelope_result_dict[dg]['env'][limit_state]['Mf'] = case_dict['Mf']
            Envelope_result_dict[dg]['env'][limit_state]['Nf'] = case_dict['Nf']
            Envelope_result_dict[dg]['env'][limit_state]['Vf'] = case_dict['Vf']
            Envelope_result_dict[dg]['env'][limit_state]['Vccd'] = case_dict['Vccd']
            Envelope_result_dict[dg]['env'][limit_state]['Vcd'] = case_dict['Vcd']
            Envelope_result_dict[dg]['env'][limit_state]['LC2'] = case_dict['LC2']
            Envelope_result_dict[dg]['env'][limit_state]['node'] = case_dict['node']
            Envelope_result_dict[dg]['env'][limit_state]['angle'] = case_dict['angle']
            Envelope_result_dict[dg]['env'][limit_state]['angle'] = case_dict['angle']
            Envelope_result_dict[dg]['env'][limit_state]['asv_min'] = case_dict['asv_min']
            Envelope_result_dict[dg]['env'][limit_state]['case'] = case
        
        #If asv is equal - update envelope dictionary if Vf/Vcd is larger.
        elif float(case_dict[result]) == float(Envelope_result_dict[dg]['env'][limit_state][result]):
            
            if case_dict['Vcd'] == 0:
                case_ratio = 1
            else:
                case_ratio = abs(case_dict['Vf'])/case_dict['Vcd']
            if Envelope_result_dict[dg]['env'][limit_state]['Vcd'] == 0:
                env_ratio = 1
            else:
                env_ratio = abs(Envelope_result_dict[dg]['env'][limit_state]['Vf'])/Envelope_result_dict[dg]['env'][limit_state]['Vcd']
            if case_ratio > env_ratio:
                Envelope_result_dict[dg]['env'][limit_state][result] = case_dict[result]                    
                
                Envelope_result_dict[dg]['env'][limit_state]['element'] = case_dict['element']
                Envelope_result_dict[dg]['env'][limit_state]['lc_thermal'] = case_dict['lc_thermal']
                Envelope_result_dict[dg]['env'][limit_state]['Mf'] = case_dict['Mf']
                Envelope_result_dict[dg]['env'][limit_state]['Nf'] = case_dict['Nf']
                Envelope_result_dict[dg]['env'][limit_state]['Vf'] = case_dict['Vf']
                Envelope_result_dict[dg]['env'][limit_state]['Vccd'] = case_dict['Vccd']
                Envelope_result_dict[dg]['env'][limit_state]['Vcd'] = case_dict['Vcd']
                Envelope_result_dict[dg]['env'][limit_state]['LC2'] = case_dict['LC2']
                Envelope_result_dict[dg]['env'][limit_state]['node'] = case_dict['node']
                Envelope_result_dict[dg]['env'][limit_state]['angle'] = case_dict['angle']
                Envelope_result_dict[dg]['env'][limit_state]['case'] = case

        if float(case_dict[result]) > float(Envelope_result_dict[dg]['env']['all'][result]):
            Envelope_result_dict[dg]['env']['all'][result] = case_dict[result]

            Envelope_result_dict[dg]['env']['all']['element'] = case_dict['element']
            Envelope_result_dict[dg]['env']['all']['lc_thermal'] = case_dict['lc_thermal']
            Envelope_result_dict[dg]['env']['all']['Mf'] = case_dict['Mf']
            Envelope_result_dict[dg]['env']['all']['Nf'] = case_dict['Nf']
            Envelope_result_dict[dg]['env']['all']['Vf'] = case_dict['Vf']
            Envelope_result_dict[dg]['env']['all']['Vccd'] = case_dict['Vccd']
            Envelope_result_dict[dg]['env']['all']['Vcd'] = case_dict['Vcd']
            Envelope_result_dict[dg]['env']['all']['LC2'] = case_dict['LC2']
            Envelope_result_dict[dg]['env']['all']['node'] = case_dict['node']
            Envelope_result_dict[dg]['env']['all']['angle'] = case_dict['angle']
            Envelope_result_dict[dg]['env']['all']['angle'] = case_dict['angle']
            Envelope_result_dict[dg]['env']['all']['asv_min'] = case_dict['asv_min']
            Envelope_result_dict[dg]['env']['all']['case'] = case
            Envelope_result_dict[dg]['env']['all']['limit_state'] = case_dict['limit_state']
        
        #If asv is equal - update envelope dictionary if Vf/Vcd is larger.
        elif float(case_dict[result]) == float(Envelope_result_dict[dg]['env']['all'][result]):
            if case_dict['Vcd'] == 0:
                case_ratio = 1
            else:
                case_ratio = abs(case_dict['Vf'])/case_dict['Vcd']
            if Envelope_result_dict[dg]['env']['all']['Vcd'] == 0:
                env_ratio = 1
            else:
                env_ratio = abs(Envelope_result_dict[dg]['env']['all']['Vf'])/Envelope_result_dict[dg]['env']['all']['Vcd']
            if case_ratio > env_ratio:
                Envelope_result_dict[dg]['env']['all'][result] = case_dict[result]                    
                
                Envelope_result_dict[dg]['env']['all']['element'] = case_dict['element']
                Envelope_result_dict[dg]['env']['all']['lc_thermal'] = case_dict['lc_thermal']
                Envelope_result_dict[dg]['env']['all']['Mf'] = case_dict['Mf']
                Envelope_result_dict[dg]['env']['all']['Nf'] = case_dict['Nf']
                Envelope_result_dict[dg]['env']['all']['Vf'] = case_dict['Vf']
                Envelope_result_dict[dg]['env']['all']['Vccd'] = case_dict['Vccd']
                Envelope_result_dict[dg]['env']['all']['Vcd'] = case_dict['Vcd']
                Envelope_result_dict[dg]['env']['all']['LC2'] = case_dict['LC2']
                Envelope_result_dict[dg]['env']['all']['node'] = case_dict['node']
                Envelope_result_dict[dg]['env']['all']['angle'] = case_dict['angle']
                Envelope_result_dict[dg]['env']['all']['case'] = case
                Envelope_result_dict[dg]['env']['all']['limit_state'] = case_dict['limit_state']


def set_table_blue_table_uls_als(counter_rows,worksheet,workbook):
    table = 'B2:' + 'O' + str(counter_rows)
    italic = workbook.add_format({'italic': True,'font_color': 'red'})
                                           
    worksheet.add_table(table,
                        {'columns': [{'header': 'DG','format': italic},
									{'header': 'Case'},
									{'header': 'Thrm.'},
									{'header': 'LS'},
									{'header': 'Element'},
									{'header': 'Point'},
									{'header': 'LC'},
									{'header': 'angle'},
									{'header': 'Nf'},
									{'header': 'Mf'},
									{'header': 'Vf'},
									{'header': 'Vccd'},
									{'header': 'Vcd'},
									{'header': 'Asv'}],'style': 'Table Style Medium 2','header_row': 1})
    
    subscript= workbook.add_format({'font_script':2,'font_color': 'white'})
    standard= workbook.add_format({'font_color': 'white','size':8,'bold': True})
    
    worksheet.write_rich_string('J2',standard,'N',subscript, 'f')        
    worksheet.write_rich_string('K2',standard,'M',subscript, 'f')        
    worksheet.write_rich_string('L2',standard,'V',subscript, 'f')        
    worksheet.write_rich_string('M2',standard,'V',subscript, 'ccd')        
    worksheet.write_rich_string('N2',standard,'V',subscript, 'cd')        
    worksheet.write_rich_string('O2',standard,'A',subscript, 'sv')         
    
def set_table_blue_table_uls_als_envelope(counter_rows,worksheet,workbook):
    table = 'B2:' + 'O' + str(counter_rows)
    italic = workbook.add_format({'italic': True,'font_color': 'red'})
    
    worksheet.add_table(table,
                        {'columns': [{'header': 'DG','format': italic},
									{'header': 'Gov.Case'},
									{'header': 'Gov.Thrm.'},
									{'header': 'LS'},
									{'header': 'Element'},
									{'header': 'Point'},
									{'header': 'LC'},
									{'header': 'angle'},
									{'header': 'Nf'},
									{'header': 'Mf'},
									{'header': 'Vf'},
									{'header': 'Vccd'},
									{'header': 'Vcd'},
									{'header': 'Asv'}],'style': 'Table Style Medium 2','header_row': 1})    
    
    subscript= workbook.add_format({'font_script':2,'font_color': 'white'})
    standard= workbook.add_format({'font_color': 'white','size':8,'bold': True})
    
    worksheet.write_rich_string('J2',standard,'N',subscript, 'f')        
    worksheet.write_rich_string('K2',standard,'M',subscript, 'f')        
    worksheet.write_rich_string('L2',standard,'V',subscript, 'f')        
    worksheet.write_rich_string('M2',standard,'V',subscript, 'ccd')        
    worksheet.write_rich_string('N2',standard,'V',subscript, 'cd')        
    worksheet.write_rich_string('O2',standard,'A',subscript, 'sv')         



def summary_blue_table_all_phases_envelope(Envelope_result_dict,workbook):    
    # Create a new workbook and add a worksheet

    worksheet = workbook.add_worksheet('ULS_ALS_envelope')
    row = 2
    header = 'The table contains values read from .shmr-files. Results must be verified with .emf-tables.'+' Ignored Minimaxes: '+Text_ign
    worksheet.write(0,1,header)     
    num_format = workbook.add_format({'num_format': '#,###0','size':8})
        
#
    for dg, dg_dict in Envelope_result_dict.items():
        worksheet.write(row,1,float(dg))
        worksheet.write(row,2,float(Envelope_result_dict[dg]['env']['all']["case"]))
        worksheet.write(row,3,float(Envelope_result_dict[dg]['env']['all']["lc_thermal"]))

        worksheet.write(row,4,Envelope_result_dict[dg]['env']['all']["limit_state"])
        worksheet.write(row,5,Envelope_result_dict[dg]['env']['all']["element"])
        worksheet.write(row,6,round(float(Envelope_result_dict[dg]['env']['all']["node"]),1))
        worksheet.write(row,7,Envelope_result_dict[dg]['env']['all']["LC2"])
        worksheet.write(row,8,round(float(Envelope_result_dict[dg]['env']['all']["angle"]),1))
        worksheet.write(row,9,int(Envelope_result_dict[dg]['env']['all']["Nf"]),num_format)
        worksheet.write(row,10,int(Envelope_result_dict[dg]['env']['all']["Mf"]),num_format)
        worksheet.write(row,11,int(Envelope_result_dict[dg]['env']['all']["Vf"]),num_format)
        worksheet.write(row,12,int(Envelope_result_dict[dg]['env']['all']["Vccd"]),num_format)
        worksheet.write(row,13,int(Envelope_result_dict[dg]['env']['all']["Vcd"]),num_format)
        worksheet.write(row,14,round(float(Envelope_result_dict[dg]['env']['all']["asv"]),1),num_format)
              
        row = row +1
    if row ==2:
        row = row+1                            
    
    set_table_blue_table_uls_als_envelope(row,worksheet,workbook)    
    width_columns(worksheet)
    workbook.formats[0].set_font_size(8)
    worksheet.set_zoom(130)
      
  
def summary_blue_table_all_phases(All_result_dict,workbook):    
    # Create a new workbook and add a worksheet
    worksheet = workbook.add_worksheet('ULS_ALS_long')
    row = 2
    header = 'The table contains values read from .shmr-files. Results must be verified with .emf-tables.'+' Ignored Minimaxes: '+Text_ign
    worksheet.write(0,1,header)     
    
    num_format = workbook.add_format({'num_format': '#,###0','size':8})
    

    for dg, dg_dict in All_result_dict.items():
        for thermal_case,thermal_dict in dg_dict.items():
            if thermal_case == 'req_link':
                continue
            for case, case_dict in thermal_dict.items():
                worksheet.write(row,1,float(dg))
                worksheet.write(row,2,float(case))
                worksheet.write(row,3,float(case_dict["lc_thermal"]))
                worksheet.write(row,4,case_dict["limit_state"])
                worksheet.write(row,5,case_dict["element"])
                worksheet.write(row,6,round(float(case_dict["node"]),1))
                worksheet.write(row,7,case_dict["LC2"])
                worksheet.write(row,8,round(float(case_dict["angle"]),1))
                worksheet.write(row,9,int(case_dict["Nf"]),num_format)
                worksheet.write(row,10,int(case_dict["Mf"]),num_format)
                worksheet.write(row,11,int(case_dict["Vf"]),num_format)
                worksheet.write(row,12,int(case_dict["Vccd"]),num_format)
                worksheet.write(row,13,int(case_dict["Vcd"]),num_format)
                worksheet.write(row,14,int(float(case_dict["asv"])),num_format)
                                           
                row = row +1
    if row ==2:
        row = row+1                            
    
    set_table_blue_table_uls_als(row,worksheet,workbook)    
    width_columns(worksheet)
    workbook.formats[0].set_font_size(8)
#    workbook.formats[0].set_num_format('# ###0')
    worksheet.set_zoom(130)

# --- Main script ---
    
directory = os.getcwd()
shmr_files,res_files= list_shmr_files(directory)

found = False

counter_rows = 1

#Open all .shmr files
All_result_dict = {}
for file_no, shmr_file in enumerate (shmr_files):
    with open(shmr_file) as f:

        Vcd_ratio = 0
        for line_no, line in enumerate(f):
            if 'Shear CHECK FOR DG:' in line:
                dg = line.split()[4]
                case = line.split()[6] + '.'  + line.split()[9]
                lc_thermal = line.split()[len(line.split())-1]
                
                if '27' in case or '31' in case or '33' in case:
                    limit_state = 'ALS'
                else:
                    limit_state = 'ULS'        
                
                try:
                    test = All_result_dict[dg]
                except:
                    All_result_dict[dg] = {}
                    
                try:
                    test = All_result_dict[dg][lc_thermal]
                except:
                    All_result_dict[dg][lc_thermal] = {}
                    
                All_result_dict[dg][lc_thermal][case] = {}                    
                All_result_dict[dg][lc_thermal][case]['limit_state'] = limit_state
                All_result_dict[dg][lc_thermal][case]['lc_thermal'] = lc_thermal

            if "DATA_BASE VERSION" in line:
                All_result_dict[dg][lc_thermal][case]['Data_base_version'] = line.split()[3]
            
            ## Read summary -shmr file
            asv_zero = False
            if '   max  ' in line:
                counter = 0
                asv_min = line.split()[len(line.split())-1]
                if line.split()[2][0] == '0':
                    asv_zero = True
                    
                else:
                    asv = line.split()[2]
                    LC = line.split()[3]
                    
                with open(res_files[file_no]) as res_file_3:
                    lines_3 = res_file_3.readlines()                    
                    
                ## Read .res file    
                with open(res_files[file_no]) as res_f:
                    
                    found = False

                    ## Search for asv result
                    for line_no_res1, line_res1 in enumerate(res_f):
                    
                        #Find max Ved/Vcd if no shear reinforcement is required
                        if asv_zero == True:
                            if len(line_res1.split()) > 7:
                                if len(line_res1.split()) == 10 and line_res1.split()[7] == '0.':
                                    angle_test = float(line_res1.split()[0])

                                    Ved_test = float(line_res1.split()[3])
                                    Vcd_test = float(line_res1.split()[5])
                                    
                                    if Vcd_test  == 0:
                                        Vcd_ratio_test = 1
                                    else:
                                        Vcd_ratio_test = abs(Ved_test/Vcd_test)
                                    
                                    if Vcd_ratio_test > Vcd_ratio:
                                        Vcd_ratio = Vcd_ratio_test
                                    
                                        angle = float(line_res1.split()[0])
                                        
                                        number = int(16 + (angle/10))
                                        
                                        Nf = float(line_res1.split()[1])
                                        Mf = float(line_res1.split()[2])
                                        Vf = float(line_res1.split()[3])
                                        Vccd = float(line_res1.split()[4])
                                        Vcd = float(line_res1.split()[5])
                                        asv = float(line_res1.split()[7])

                                        try:
                                            #Water pressure is on
                                            LC2 = (lines_3[line_no_res1-number-4].split()[4]) 
                                            element = (lines_3[line_no_res1-number-2].split()[3])                                     
                                            node = (lines_3[line_no_res1-number-2].split()[6])      
                                        except:
                                            #Water pressure is off
                                            LC2 = (lines_3[line_no_res1-number-2].split()[4])                                     
                                            element = (lines_3[line_no_res1-number].split()[3])                                     
                                            node = (lines_3[line_no_res1-number].split()[6])        
                        if found == True:                           
                            break
                        
                        if asv_zero == False:                        
                            if asv in line_res1 and '** Maximum required stirrup area is' in line_res1 and asv_zero == False:
                                ## Found asv result. Re-Loop trough .res to find additional results such as concrete shear capacity.                            
                                    
                                found = False
                                
                                try:
                                    #Water pressure is on
                                    LC2 = (lines_3[line_no_res1-40].split()[4]) 
                                    element = (lines_3[line_no_res1-38].split()[3])                                     
                                    node = (lines_3[line_no_res1-38].split()[6])                                     
                                except:
                                    #Water pressure is off
                                    LC2 = (lines_3[line_no_res1-38].split()[4])                                     
                                    element = (lines_3[line_no_res1-36].split()[3])                                     
                                    node = (lines_3[line_no_res1-36].split()[6])                                     
                                
                                for line_no_loop in range(line_no_res1-23,line_no_res1):
                                    if LC == LC2:
                                        if asv + '  ' in lines_3[line_no_loop]:
                                            if case == '27.1' or case == '31.1':
                                                limit_state = 'ALS'
                                            else:
                                                limit_state = 'ULS'
                                                
                                                
                                            line_list = lines_3[line_no_loop].split()
                                            
                                            angle = float(line_list[0])
                                            Nf = float(line_list[1])
                                            Mf = float(line_list[2])
                                            Vf = float(line_list[3])
                                            Vccd = float(line_list[4])
                                            Vcd = float(line_list[5])
                                            asv = float(line_list[7])
#                                            
                                            All_result_dict[dg][lc_thermal][case]['element'] = element
                                            All_result_dict[dg][lc_thermal][case]['node'] = node
                                            All_result_dict[dg][lc_thermal][case]['LC2'] = LC2
                                            All_result_dict[dg][lc_thermal][case]['angle'] = angle
                                            All_result_dict[dg][lc_thermal][case]['Nf'] = Nf
                                            All_result_dict[dg][lc_thermal][case]['Mf'] = Mf
                                            All_result_dict[dg][lc_thermal][case]['Vf'] = Vf
                                            All_result_dict[dg][lc_thermal][case]['Vccd'] = Vccd
                                            All_result_dict[dg][lc_thermal][case]['Vcd'] = Vcd
                                            All_result_dict[dg][lc_thermal][case]['asv'] = asv
                                            All_result_dict[dg][lc_thermal][case]['asv_zero'] = asv_zero
                                            All_result_dict[dg][lc_thermal][case]['asv_min'] = asv_min
                                            
                                            
                                            counter_rows = counter_rows + 1
                                            found = True
                                    if found == True:                               
                                        break
                                
        if asv_zero == True:
            
            All_result_dict[dg][lc_thermal][case]['element'] = element
            All_result_dict[dg][lc_thermal][case]['node'] = node
            All_result_dict[dg][lc_thermal][case]['LC2'] = LC2
            All_result_dict[dg][lc_thermal][case]['angle'] = angle
            All_result_dict[dg][lc_thermal][case]['Nf'] = Nf
            All_result_dict[dg][lc_thermal][case]['Mf'] = Mf
            All_result_dict[dg][lc_thermal][case]['Vf'] = Vf
            All_result_dict[dg][lc_thermal][case]['Vccd'] = Vccd
            All_result_dict[dg][lc_thermal][case]['Vcd'] = Vcd
            All_result_dict[dg][lc_thermal][case]['asv'] = asv
            All_result_dict[dg][lc_thermal][case]['asv_zero'] = asv_zero
            All_result_dict[dg][lc_thermal][case]['asv_min'] = asv_min
            
            counter_rows = counter_rows + 1        
                                

## Create envelope dict            
Envelope_result_dict = {}                
for dg, dg_dict in All_result_dict.items():
    Envelope_result_dict[dg] = {}                
    Envelope_result_dict[dg]['env'] = {}                
    Envelope_result_dict[dg]['env']['ALS'] = {}                
    Envelope_result_dict[dg]['env']['ULS'] = {}                
    Envelope_result_dict[dg]['env']['all'] = {}                
    
    envelope_ULS_ALS('ULS')    
    envelope_ULS_ALS('ALS')    
    envelope_ULS_ALS('all')    
    
    for thermal_case,thermal_dict in dg_dict.items():    
        if thermal_case == 'req_link':
            continue
        for case, case_dict in thermal_dict.items():  
            limit_state = case_dict["limit_state"]
            if limit_state == 'ULS' or limit_state == 'ALS':
                
                check_envelope(dg,"asv",limit_state,"larger",case_dict,case)
          

# --- Write tables to excel document
file_root='result_tables_shear_m_'
date=str(datetime.date.today())
filename=file_root+date+'.xlsx'             
wb = xlsxwriter.Workbook(filename)                


summary_blue_table_all_phases_envelope(Envelope_result_dict,wb)
summary_blue_table_all_phases(All_result_dict,wb)     

wb.close()  
print('Your results are ready in '+directory)