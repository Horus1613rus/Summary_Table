#! python 3
# -*- coding: utf-8 -*-
"""
Created on Thu Feb 21 18:41:04 2019
@author: nx134
Modified on Tue Nov 12 14:00:00 2019
@co-author: Antl

Description of this script.

# _*_ coding: utf-8
"""
#%% --- MODULE S ---------------------------------------------------------#
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


def createFolder(directory):
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
    except OSError:
        print ('Error: Creating directory. ' +  directory)
        
            
def list_csmr_files(directory):
# List csmr file in current directory
    csmr_files = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith('.csmr'):
                case_fs=file.split('mm')[1].replace('.res.csmr','')
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
                    csmr = root + str("/") + file
                    csmr_files.append(csmr)
                    end = root.index("dg")
                    rooten = root[:end]

    req_link = r".\requirements\RequFile_09.41d_rev2.txt"
    folder_req = "requi_09.41D/"
    return csmr_files, req_link,folder_req

def list_shmr_files(directory):
# List csmr file in current directory
    shmr_files = []
    res_files = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith('.shmr'):
                shmr = root + str("/") + file
                res = root + str("/") + file[:len(file)-5]
                shmr_files.append(shmr)
                res_files.append(res)
                end = root.index("dg")
                rooten = root[:end]

    return shmr_files,res_files


def read_csmr_all(line,dg,case, All_result_dict,LC,line_no,thermal_effects):
    excluded_points  = 'False'
    if 'EXCLUDED' in line:
        excluded_points  = 'True'
        All_result_dict[dg][lc_thermal][case]['excluded_points'] = 'True'
        return excluded_points,thermal_effects
    elif "Thermal effects ..." in line:
        thermal_effects = 'True'
        return excluded_points,thermal_effects
    elif len(line[34:].split()) > 4 or len(line[34:].split()) == 4:
        All_result_dict[dg][lc_thermal][case]['excluded_points'] = 'False'
        text = line[5:32]

        if thermal_effects == 'True':
            text = text + '_thermal'

        All_result_dict[dg][lc_thermal][case][text] = {}
        Value = (line[34:].split()[0])

        if len(line[34:].split()) == 4:
            LC = (line[34:].split()[3])
        elif (line[34:].split()[3][0]) == "Y" or (line[34:].split()[3][0]) == "T" or (line[34:].split()[3][0]) == "X" or (line[34:].split()[3][0]) == "W":
            LC = (line[34:].split()[3]) + (line[34:].split()[4])
        else:
            LC = (line[34:].split()[3])

        if 'utilisation' in text:
            All_result_dict[dg][lc_thermal][case][text]['value']  = Value[:len(Value)-2]
        else:
            All_result_dict[dg][lc_thermal][case][text]['value']  = Value[:len(Value)]

        All_result_dict[dg][lc_thermal][case][text]['part'] = (line[34:].split()[1])
        All_result_dict[dg][lc_thermal][case][text]['node'] = (line[34:].split()[2])
        All_result_dict[dg][lc_thermal][case][text]['LC'] = LC

    elif len(line[34:].split()) > 0:
        text = line[5:32]
        if line_no > 48:
            text = text + '_thermal'
        All_result_dict[dg][lc_thermal][case][text] = {}
        Value = (line[34:].split()[0])
        if 'utilisation' in text:
            All_result_dict[dg][lc_thermal][case][text]['value']  = Value[:len(Value)-2]
        else:
            All_result_dict[dg][lc_thermal][case][text]['value']  = Value[:len(Value)]
    return(excluded_points,thermal_effects)    

    
def links_and_requirement(All_result_dict,dg,case,file,limit_state,Data_base_version,temp_case,folder_req,soil_series):
    if limit_state == "SLS":
        subtract = 12
    elif limit_state == "ULS" or limit_state == "ALS":
        subtract = 13
    else:
        subtract = 0    
    
    emf_link = file[:len(file)-subtract] + "_" + soil_series + limit_state + "_" + Data_base_version + "_1_" + temp_case + ".emf"
    pro_link = file[:file.index(temp_case)+5] + "\consecn" + limit_state[0]  + "_1." + str(dg)
    csmr_link = file
    res_link = file[:len(file)-5]
    f_link = file[:len(file)-9] + '.f'
    
    All_result_dict[dg][lc_thermal][case]['emf_link'] = emf_link
    All_result_dict[dg][lc_thermal][case]['pro_link'] = pro_link
    All_result_dict[dg][lc_thermal][case]['csmr_link'] = csmr_link
    All_result_dict[dg][lc_thermal][case]['res_link'] = res_link
    All_result_dict[dg][lc_thermal][case]['f_link'] = f_link
    
    
    
    
    #Requirements
    with open(req_link) as req_file_summary:
        for line_no,line in enumerate(req_file_summary):
            if line_no == 0:
                if line.split()[0] == 'reui_file':	
                    reui_no = 0
                elif line.split()[1] == 'reui_file':	
                    reui_no = 1
                elif line.split()[2] == 'reui_file':	
                    reui_no = 2
                elif line.split()[3] == 'reui_file':	
                    reui_no = 3
                    
            if line.split()[0] == dg:
                req_link_dg_temp = line.split()[reui_no]
                end = req_link.index("requirements")
                req_link_dg_temp = req_link[:end+13] + folder_req  + req_link_dg_temp                
                break
            else:
                req_link_dg_temp = "error"
    
    
    All_result_dict[dg]['req_link'] = req_link_dg_temp
    
    with open(pro_link) as pro_file:
        pro_file_lines = pro_file.readlines()
        thickness = pro_file_lines[0].split()[6]
        All_result_dict[dg][lc_thermal][case]['thickness'] = thickness
        
        
    
    with open(req_link_dg_temp) as req_file_dg:
        for line_no,line in enumerate(req_file_dg):
            line_no_requisss = 7
#            if line_no ==7:
                
            if len(line.split()) == 20 and line.split()[0] == case[:len(case)-2]:
                Requirements_dict[dg][case] = {}
                Requirements_dict[dg][case]['prop'] = line.split()[1]
                Requirements_dict[dg][case]['ls'] = line.split()[2]
                Requirements_dict[dg][case]['cepa'] = line.split()[3]
                Requirements_dict[dg][case]['csig'] = line.split()[4]
                Requirements_dict[dg][case]['sepa'] = line.split()[5]
                Requirements_dict[dg][case]['ssig'] = line.split()[6]
                Requirements_dict[dg][case]['peps'] = line.split()[7]
                Requirements_dict[dg][case]['psig'] = line.split()[8]
                Requirements_dict[dg][case]['C'] = line.split()[9]
                Requirements_dict[dg][case]['msig'] = line.split()[10]
                Requirements_dict[dg][case]['wk+'] = line.split()[11]
                Requirements_dict[dg][case]['wk-'] = line.split()[12]
                Requirements_dict[dg][case]['wk1'] = line.split()[13]
                Requirements_dict[dg][case]['wk2'] = line.split()[14]
                Requirements_dict[dg][case]['wke'] = line.split()[15]
                Requirements_dict[dg][case]['wkto'] = line.split()[16]
                Requirements_dict[dg][case]['wk2+'] = line.split()[17]
                Requirements_dict[dg][case]['Wk2-'] = line.split()[18]
                Requirements_dict[dg][case]['csigp'] = line.split()[19]                
            
        
def temp_case_function(file):
    if 'lower' in file:
        temp_case = 'lower'
    elif 'upper' in file:
        temp_case= 'upper'
    elif 'lower2' in file:
        temp_case= 'lower2'
    else:
        temp_case = 'unknown'
    return temp_case



def set_table_blue_table_uls_als(counter_rows,worksheet,workbook):
    table = 'B2:' + 'AQ' + str(counter_rows)
    bold   = workbook.add_format({'bold': True})
    italic = workbook.add_format({'italic': True,'font_color': 'red'})
    
    currency_format = workbook.add_format({'num_format': '$#,##0'})
    percentage_format = workbook.add_format({'num_format': ' ##\%;[Red](0.00##\%)','size':8})
                                           
    worksheet.add_table(table,
                        {'columns': [{'header': 'DG','format': italic},
									{'header': 'Case'},
									{'header': 'Thrm.'},
									{'header': 'LS'},
									{'header': 'σc,bf'},
									{'header': 'σc,tf'},
                                    {'header': 'ɛc,bf'},
									{'header': 'ɛc,tf'},
									{'header': 'σs,1'},
									{'header': 'σs,2'},
									{'header': 'σs,3'},
									{'header': 'σs,4'},
									{'header': 'σs,5'},
									{'header': 'σs,6'},
									{'header': 'σs,7'},
									{'header': 'σs,8'},
									{'header': 'σps,ad,1'},
									{'header': 'σps,ad,2'},
									{'header': 'UFs1'},
									{'header': 'UFs2'},
									{'header': 'UFs3'},
									{'header': 'UFs4'},
									{'header': 'UFs5'},
									{'header': 'UFs6'},
									{'header': 'UFs7'},
									{'header': 'UFs8'},
                                    {'header': 'ɛs1,min'},
									{'header': 'ɛs1,max'},
									{'header': 'ɛs2,min'},
									{'header': 'ɛs2,max'},
									{'header': 'ɛs3,min'},
									{'header': 'ɛs3,max'},
									{'header': 'ɛs4,min'},
                                    {'header': 'ɛs4,max'},
                                    {'header': 'ɛs5,min'},
									{'header': 'ɛs5,max'},
									{'header': 'ɛs6,min'},
									{'header': 'ɛs6,max'},
									{'header': 'ɛs7,min'},
									{'header': 'ɛs7,max'},
									{'header': 'ɛs8,min'},
                                    {'header': 'ɛs8,max'}],'style': 'Table Style Medium 2','header_row': 1})
    
    worksheet.set_column('T:AA', None, percentage_format)

    subscript= workbook.add_format({'font_script':2,'font_color': 'white'})
    standard= workbook.add_format({'font_color': 'white','size':8,'bold': True})
    
    worksheet.write_rich_string('F2','σ',subscript, 'c,bf')
    worksheet.write_rich_string('G2','σ',subscript, 'c,tf')
    worksheet.write_rich_string('H2','ɛ',subscript, 'c,bf')
    worksheet.write_rich_string('I2','ɛ',subscript, 'c,tf')    
    worksheet.write_rich_string('J2','σ',subscript, 's,1')    
    worksheet.write_rich_string('K2','σ',subscript, 's,2')    
    worksheet.write_rich_string('L2','σ',subscript, 's,3')    
    worksheet.write_rich_string('M2','σ',subscript, 's,4')    
    worksheet.write_rich_string('N2','σ',subscript, 's,5')    
    worksheet.write_rich_string('O2','σ',subscript, 's,6')    
    worksheet.write_rich_string('P2','σ',subscript, 's,7')    
    worksheet.write_rich_string('Q2','σ',subscript, 's,8')    
    worksheet.write_rich_string('R2','σ',subscript, 'ps,ad,1')    
    worksheet.write_rich_string('S2','σ',subscript, 'ps,ad,2')    
    worksheet.write_rich_string('T2',standard,'UF',subscript, 's1')    
    worksheet.write_rich_string('U2',standard,'UF',subscript, 's2')    
    worksheet.write_rich_string('V2',standard,'UF',subscript, 's3')    
    worksheet.write_rich_string('W2',standard,'UF',subscript, 's4')    
    worksheet.write_rich_string('X2',standard,'UF',subscript, 's5')    
    worksheet.write_rich_string('Y2',standard,'UF',subscript, 's6')    
    worksheet.write_rich_string('Z2',standard,'UF',subscript, 's7')    
    worksheet.write_rich_string('AA2',standard,'UF',subscript, 's8')
    worksheet.write_rich_string('AB2', 'ɛ',subscript, 's1,min')    
    worksheet.write_rich_string('AC2', 'ɛ',subscript, 's1,max')    
    worksheet.write_rich_string('AD2', 'ɛ',subscript, 's2,min')    
    worksheet.write_rich_string('AE2', 'ɛ',subscript, 's2,max')    
    worksheet.write_rich_string('AF2', 'ɛ',subscript, 's3,min')    
    worksheet.write_rich_string('AG2', 'ɛ',subscript, 's3,max')    
    worksheet.write_rich_string('AH2', 'ɛ',subscript, 's4,min')    
    worksheet.write_rich_string('AI2', 'ɛ',subscript, 's4,max')
    worksheet.write_rich_string('AJ2', 'ɛ',subscript, 's5,min')    
    worksheet.write_rich_string('AK2', 'ɛ',subscript, 's5,max')    
    worksheet.write_rich_string('AL2', 'ɛ',subscript, 's6,min')    
    worksheet.write_rich_string('AM2', 'ɛ',subscript, 's6,max')    
    worksheet.write_rich_string('AN2', 'ɛ',subscript, 's7,min')    
    worksheet.write_rich_string('AO2', 'ɛ',subscript, 's7,max')    
    worksheet.write_rich_string('AP2', 'ɛ',subscript, 's8,min')    
    worksheet.write_rich_string('AQ2', 'ɛ',subscript,'s8,max')    
    

def set_table_blue_table_sls(counter_rows,worksheet,workbook):
    table = 'B2:' + 'AC' + str(counter_rows)
    bold   = workbook.add_format({'bold': True})
    italic = workbook.add_format({'italic': True,'font_color': 'red'})
    
    currency_format = workbook.add_format({'num_format': '$#,##0'})
    percentage_format = workbook.add_format({'num_format': ' ##\%;[Red](0.00##\%)','size':8})
    

    worksheet.add_table(table,
                        {'columns': [{'header': 'DG','format': italic},
									{'header': 'Case'},
									{'header': 'Thrm.'},
									{'header': 'LS'},
									{'header': 'σc,bf'},
									{'header': 'σc,tf'},
									{'header': 'σs,1'},
									{'header': 'σs,2'},
									{'header': 'σs,3'},
									{'header': 'σs,4'},
									{'header': 'σs,5'},
									{'header': 'σs,6'},
									{'header': 'σs,7'},
									{'header': 'σs,8'},
									{'header': 'σps,ad,1'},
									{'header': 'σps,ad,2'},
									{'header': 'cu'},
									{'header': 'wk'},
									{'header': 'cu,th.'},
									{'header': 'wk,th.'}],'style': 'Table Style Medium 2','header_row': 1})
    

    subscript= workbook.add_format({'font_script':2,'font_color': 'white'})
    standard= workbook.add_format({'font_color': 'white','size':8,'bold': True})
    
    worksheet.write_rich_string('F2','σ',subscript, 'c,bf')    
    worksheet.write_rich_string('G2','σ',subscript, 'c,tf')    
    worksheet.write_rich_string('H2','σ',subscript, 's,1')    
    worksheet.write_rich_string('I2','σ',subscript, 's,2')    
    worksheet.write_rich_string('J2','σ',subscript, 's,3')    
    worksheet.write_rich_string('K2','σ',subscript, 's,4')    
    worksheet.write_rich_string('L2','σ',subscript, 's,5')    
    worksheet.write_rich_string('M2','σ',subscript, 's,6')    
    worksheet.write_rich_string('N2','σ',subscript, 's,7')    
    worksheet.write_rich_string('O2','σ',subscript, 's,8')    
    worksheet.write_rich_string('P2','σ',subscript, 'ps,ad,1')    
    worksheet.write_rich_string('Q2','σ',subscript, 'ps,ad,2')    
    worksheet.write_rich_string('R2','c',subscript, 'u')    
    worksheet.write_rich_string('S2','w',subscript, 'k')    
    worksheet.write_rich_string('T2','c',subscript, 'u,th.')    
    worksheet.write_rich_string('U2','w',subscript, 'k,th.')    
     
    
   
def width_columns_blue_table(worksheet):
    std_width = 10.5
    w_uf = 5
    w_es = 4
    w_std = 4
    worksheet.set_column('A:A', std_width-w_std)
    worksheet.set_column('B:B', std_width-w_std-1)
    worksheet.set_column('C:C', std_width-w_std-2)
    worksheet.set_column('D:D', std_width-w_std)
    worksheet.set_column('E:E', std_width-w_std-2)
    worksheet.set_column('F:F', std_width-w_es-1)
    worksheet.set_column('G:G', std_width-w_es-1)
    worksheet.set_column('H:H', std_width-w_es)
    worksheet.set_column('I:I', std_width-w_es)
    worksheet.set_column('J:J', std_width-w_es)    
    worksheet.set_column('K:K', std_width-w_es)    
    worksheet.set_column('L:L', std_width-w_es)    
    worksheet.set_column('M:M', std_width-w_es) 
    worksheet.set_column('N:N', std_width-w_es) 
    worksheet.set_column('O:O', std_width-w_es) 
    worksheet.set_column('P:P', std_width-3) 
    worksheet.set_column('Q:Q', std_width-3) 
    worksheet.set_column('R:R', std_width-w_uf) 
    worksheet.set_column('S:S', std_width-w_uf) 
    worksheet.set_column('T:T', std_width-w_uf) 
    worksheet.set_column('U:U', std_width-w_uf) 
    worksheet.set_column('V:V', std_width-w_uf) 
    worksheet.set_column('W:W', std_width-w_uf) 
    worksheet.set_column('X:X', std_width-w_uf) 
    worksheet.set_column('Y:Y', std_width-w_uf)
    worksheet.set_column('Z:AG', std_width-w_std)

   

    
def summary_blue_table_uls_als(All_result_dict,workbook,limit_state):    
    # Create a new workbook and add a worksheet
    worksheet = workbook.add_worksheet(limit_state + '_long')
    row = 2
    header = limit_state + ' result table. The table contains values read from .csmr-files. Results must be verified with .emf-tables.'+' Ignored Minimaxes: '+Text_ign
    worksheet.write(0,1,header)  
    
    show_l5 = False
    show_l6 = False
    show_l7 = False
    show_l8 = False
    show_p1 = False
    show_p2 = False
#
    for dg, dg_dict in All_result_dict.items():
        for thermal_case,thermal_dict in dg_dict.items():
            if thermal_case == 'req_link':
                continue
            for case, case_dict in thermal_dict.items():
                if case_dict["limit_state"] == limit_state:
                    
                    worksheet.write(row,1,float(dg))
                    worksheet.write(row,2,float(case))
                    worksheet.write(row,3,float(case_dict["lc_thermal"]))
                    worksheet.write(row,4,case_dict["limit_state"])
                    worksheet.write(row,5,round(float(case_dict["Min conc. stress bot face :"]["value"]),1))
                    worksheet.write_comment(row,5,case_dict["Min conc. stress bot face :"]["part"]+" "+case_dict["Min conc. stress bot face :"]["node"]+'\n'+case_dict["Min conc. stress bot face :"]["LC"])
                    
                    worksheet.write(row,6,round(float(case_dict["Min conc. stress top face :"]["value"]),1))
                    worksheet.write_comment(row,6,case_dict["Min conc. stress top face :"]["part"]+' '+case_dict["Min conc. stress top face :"]["node"]+"\n"+case_dict["Min conc. stress top face :"]["LC"])
                    
                    worksheet.write(row,7,round(float(case_dict["Min conc. strain bot face :"]["value"]),1))
                    worksheet.write_comment(row,7,case_dict["Min conc. strain bot face :"]["part"]+" "+case_dict["Min conc. strain bot face :"]["node"]+'\n'+case_dict["Min conc. strain bot face :"]["LC"])
                    
                    worksheet.write(row,8,round(float(case_dict["Min conc. strain top face :"]["value"]),1))
                    worksheet.write_comment(row,8,case_dict["Min conc. strain top face :"]["part"]+' '+case_dict["Min conc. strain top face :"]["node"]+"\n"+case_dict["Min conc. strain top face :"]["LC"])
                    
                    if max(abs(float(case_dict["Max steel stress layer  1 :"]["value"])),abs(float(case_dict["Min steel stress layer  1 :"]["value"])))<=abs(float(case_dict["Min steel stress layer  1 :"]["value"])):
                        worksheet.write(row,9, round(float(case_dict["Min steel stress layer  1 :"]["value"]),0))
                    else:
                        worksheet.write(row,9, round(float(case_dict["Max steel stress layer  1 :"]["value"]),0))
                    #worksheet.write_comment(row,7,case_dict["Max steel stress layer  1 :"]["part"]+' '+case_dict["Max steel stress layer  1 :"]["node"]+"\n"+case_dict["Max steel stress layer  1 :"]["LC"][-12:])
                    
                    if max(abs(float(case_dict["Max steel stress layer  2 :"]["value"])),abs(float(case_dict["Min steel stress layer  2 :"]["value"])))<=abs(float(case_dict["Min steel stress layer  2 :"]["value"])):
                        worksheet.write(row,10, round(float(case_dict["Min steel stress layer  2 :"]["value"]),0))
                    else:
                        worksheet.write(row,10, round(float(case_dict["Max steel stress layer  2 :"]["value"]),0))
                    #worksheet.write_comment(row,8,case_dict["Max steel stress layer  2 :"]["part"]+' '+case_dict["Max steel stress layer  2 :"]["node"]+"\n"+case_dict["Max steel stress layer  2 :"]["LC"][-12:])
                    
                    if max(abs(float(case_dict["Max steel stress layer  3 :"]["value"])),abs(float(case_dict["Min steel stress layer  3 :"]["value"])))<=abs(float(case_dict["Min steel stress layer  3 :"]["value"])):
                        worksheet.write(row,11, round(float(case_dict["Min steel stress layer  3 :"]["value"]),0))
                    else:
                        worksheet.write(row,11, round(float(case_dict["Max steel stress layer  3 :"]["value"]),0))
                    #worksheet.write_comment(row,9,case_dict["Max steel stress layer  3 :"]["part"]+' '+case_dict["Max steel stress layer  3 :"]["node"]+"\n"+case_dict["Max steel stress layer  3 :"]["LC"][-12:])
                    
                    if max(abs(float(case_dict["Max steel stress layer  4 :"]["value"])),abs(float(case_dict["Min steel stress layer  4 :"]["value"])))<=abs(float(case_dict["Min steel stress layer  4 :"]["value"])):
                        worksheet.write(row,12, round(float(case_dict["Min steel stress layer  4 :"]["value"]),0))
                    else:
                        worksheet.write(row,12, round(float(case_dict["Max steel stress layer  4 :"]["value"]),0))
                    #worksheet.write_comment(row,10,case_dict["Max steel stress layer  4 :"]["part"]+' '+case_dict["Max steel stress layer  4 :"]["node"]+"\n"+case_dict["Max steel stress layer  4 :"]["LC"][-12:])
                    
                    #if "Max steel stress layer  5 :" in case_dict:
                        #worksheet.write(row,11,round(float(case_dict["Max steel stress layer  5 :"]["value"])))
                        #worksheet.write_comment(row,11,case_dict["Max steel stress layer  5 :"]["part"]+' '+case_dict["Max steel stress layer  5 :"]["node"]+"\n"+case_dict["Max steel stress layer  5 :"]["LC"][-12:])
                    #else:
                    show_l5 = worksheet_write(worksheet,case_dict,"Max steel stress layer  5 :",row,13,show_l5)
                    
                    
                    #if "Max steel stress layer  6 :" in case_dict:
                        #worksheet.write(row,12,round(float(case_dict["Max steel stress layer  6 :"]["value"])))
                       # worksheet.write_comment(row,12,case_dict["Max steel stress layer  6 :"]["part"]+' '+case_dict["Max steel stress layer  6 :"]["node"]+"\n"+case_dict["Max steel stress layer  6 :"]["LC"][-12:])
                    #else:
                    show_l6 = worksheet_write(worksheet,case_dict,"Max steel stress layer  6 :",row,14,show_l6)
                    

                    #if "Max steel stress layer  7 :" in case_dict:
                        #worksheet.write(row,13,round(float(case_dict["Max steel stress layer  7 :"]["value"])))
                        #worksheet.write_comment(row,13,case_dict["Max steel stress layer  7 :"]["part"]+' '+case_dict["Max steel stress layer  7 :"]["node"]+"\n"+case_dict["Max steel stress layer  7 :"]["LC"][-12:])
                    #else:
                    show_l7 = worksheet_write(worksheet,case_dict,"Max steel stress layer  7 :",row,15,show_l7)
                    
                         
                    #if "Max steel stress layer  8 :" in case_dict:
                        #worksheet.write(row,14,round(float(case_dict["Max steel stress layer  8 :"]["value"])))
                        #worksheet.write_comment(row,14,case_dict["Max steel stress layer  8 :"]["part"]+' '+case_dict["Max steel stress layer  8 :"]["node"]+"\n"+case_dict["Max steel stress layer  8 :"]["LC"][-12:])
                    #else:
                    show_l8 = worksheet_write(worksheet,case_dict,"Max steel stress layer  8 :",row,16,show_l8)
                    
                                       
                    show_p1 = worksheet_write(worksheet,case_dict,"Max prestress adstress  1 :",row,17,show_p1)
                                        
                    show_p2 = worksheet_write(worksheet,case_dict,"Max prestress adstress  2 :",row,18,show_p2)
                                        
                    worksheet.write(row,19,float(case_dict["Steel utilisation layer  1:"]["value"]))
                    #worksheet.write_comment(row,17,case_dict["Steel utilisation layer  1:"]["part"]+' '+case_dict["Steel utilisation layer  1:"]["node"]+"\n"+case_dict["Steel utilisation layer  1:"]["LC"][-12:])
                    
                    worksheet.write(row,20,float(case_dict["Steel utilisation layer  2:"]["value"]))
                    #worksheet.write_comment(row,18,case_dict["Steel utilisation layer  2:"]["part"]+' '+case_dict["Steel utilisation layer  2:"]["node"]+"\n"+case_dict["Steel utilisation layer  2:"]["LC"][-12:])
                    
                    worksheet.write(row,21,float(case_dict["Steel utilisation layer  3:"]["value"]))
                    #worksheet.write_comment(row,19,case_dict["Steel utilisation layer  3:"]["part"]+' '+case_dict["Steel utilisation layer  3:"]["node"]+"\n"+case_dict["Steel utilisation layer  3:"]["LC"][-12:])
                    
                    worksheet.write(row,22,float(case_dict["Steel utilisation layer  4:"]["value"]))
                    #worksheet.write_comment(row,20,case_dict["Steel utilisation layer  4:"]["part"]+' '+case_dict["Steel utilisation layer  4:"]["node"]+"\n"+case_dict["Steel utilisation layer  4:"]["LC"][-12:])

                    #if "Steel utilisation layer  5:" in case_dict:
                     #   worksheet.write(row,21,round(float(case_dict["Steel utilisation layer  5:"]["value"])))
                     #   worksheet.write_comment(row,21,case_dict["Steel utilisation layer  5:"]["part"]+' '+case_dict["Steel utilisation layer  5:"]["node"]+"\n"+case_dict["Steel utilisation layer  5:"]["LC"][-12:])
                      #  show_l5 =1
                    #else:
                    show_l5 = worksheet_write(worksheet,case_dict,"Steel utilisation layer  5:",row,23,show_l5)
                    
                    #if "Steel utilisation layer  6:" in case_dict:
                     #   worksheet.write(row,22,round(float(case_dict["Steel utilisation layer  6:"]["value"])))
                      #  worksheet.write_comment(row,22,case_dict["Steel utilisation layer  6:"]["part"]+' '+case_dict["Steel utilisation layer  6:"]["node"]+"\n"+case_dict["Steel utilisation layer  6:"]["LC"][-12:])
                       # show_l6 =1
                    #else:
                    show_l6 = worksheet_write(worksheet,case_dict,"Steel utilisation layer  6:",row,24,show_l6)
                        
                    #if "Steel utilisation layer  7:" in case_dict:
                     #   worksheet.write(row,23,round(float(case_dict["Steel utilisation layer  7:"]["value"])))
                      #  worksheet.write_comment(row,23,case_dict["Steel utilisation layer  7:"]["part"]+' '+case_dict["Steel utilisation layer  7:"]["node"]+"\n"+case_dict["Steel utilisation layer  7:"]["LC"][-12:])
                      #  show_l7 =1
                    #else:
                    show_l7 = worksheet_write(worksheet,case_dict,"Steel utilisation layer  7:",row,25,show_l7)
                        
                    #if "Steel utilisation layer  8:" in case_dict:
                     #   worksheet.write(row,24,round(float(case_dict["Steel utilisation layer  8:"]["value"])))
                      #  worksheet.write_comment(row,24,case_dict["Steel utilisation layer  8:"]["part"]+' '+case_dict["Steel utilisation layer  8:"]["node"]+"\n"+case_dict["Steel utilisation layer  8:"]["LC"][-12:])
                       # show_l8 =1
                   # else:
                    show_l8 = worksheet_write(worksheet,case_dict,"Steel utilisation layer  8:",row,26,show_l8)
                    
                    worksheet.write(row,27, float(case_dict["Min steel strain layer  1 :"]["value"]))
                    worksheet.write_comment(row,27,case_dict["Min steel strain layer  1 :"]["part"]+' '+case_dict["Min steel strain layer  1 :"]["node"]+"\n"+case_dict["Min steel strain layer  1 :"]["LC"])
                    
                    worksheet.write(row,28, float(case_dict["Max steel strain layer  1 :"]["value"]))
                    worksheet.write_comment(row,28,case_dict["Max steel strain layer  1 :"]["part"]+' '+case_dict["Max steel strain layer  1 :"]["node"]+"\n"+case_dict["Max steel strain layer  1 :"]["LC"])
                    
                    worksheet.write(row,29, float(case_dict["Min steel strain layer  2 :"]["value"]))
                    worksheet.write_comment(row,29,case_dict["Min steel strain layer  2 :"]["part"]+' '+case_dict["Min steel strain layer  2 :"]["node"]+"\n"+case_dict["Min steel strain layer  2 :"]["LC"])
                    
                    worksheet.write(row,30, float(case_dict["Max steel strain layer  2 :"]["value"]))
                    worksheet.write_comment(row,30,case_dict["Max steel strain layer  2 :"]["part"]+' '+case_dict["Max steel strain layer  2 :"]["node"]+"\n"+case_dict["Max steel strain layer  2 :"]["LC"])
                    
                    worksheet.write(row,31, float(case_dict["Min steel strain layer  3 :"]["value"]))
                    worksheet.write_comment(row,31,case_dict["Min steel strain layer  3 :"]["part"]+' '+case_dict["Min steel strain layer  3 :"]["node"]+"\n"+case_dict["Min steel strain layer  3 :"]["LC"])
                    
                    worksheet.write(row,32, float(case_dict["Max steel strain layer  3 :"]["value"]))
                    worksheet.write_comment(row,32,case_dict["Max steel strain layer  3 :"]["part"]+' '+case_dict["Max steel strain layer  3 :"]["node"]+"\n"+case_dict["Max steel strain layer  3 :"]["LC"])
                    
                    worksheet.write(row,33, float(case_dict["Min steel strain layer  4 :"]["value"]))
                    worksheet.write_comment(row,33,case_dict["Min steel strain layer  4 :"]["part"]+' '+case_dict["Max steel strain layer  4 :"]["node"]+"\n"+case_dict["Max steel strain layer  4 :"]["LC"])
                    
                    worksheet.write(row,34, float(case_dict["Max steel strain layer  4 :"]["value"]))
                    worksheet.write_comment(row,34,case_dict["Max steel strain layer  4 :"]["part"]+' '+case_dict["Max steel strain layer  4 :"]["node"]+"\n"+case_dict["Max steel strain layer  4 :"]["LC"])
                    
                    if "Min steel strain layer  5 :" in case_dict:
                        worksheet.write(row,35,float(case_dict["Min steel strain layer  5 :"]["value"]))
                        worksheet.write_comment(row,35,case_dict["Min steel strain layer  5 :"]["part"]+' '+case_dict["Min steel strain layer  5 :"]["node"]+"\n"+case_dict["Min steel strain layer  5 :"]["LC"])
                    else:
                        show_l5 = worksheet_write(worksheet,case_dict,"Min steel strain layer  5 :",row,35,show_l5)
                    
                    
                    if "Max steel strain layer  5 :" in case_dict:
                        worksheet.write(row,36,float(case_dict["Max steel strain layer  5 :"]["value"]))
                        worksheet.write_comment(row,36,case_dict["Max steel strain layer  5 :"]["part"]+' '+case_dict["Max steel strain layer  5 :"]["node"]+"\n"+case_dict["Max steel strain layer  5 :"]["LC"])
                    else:
                        show_l5 = worksheet_write(worksheet,case_dict,"Max steel stress layer  5 :",row,36,show_l5)
                        
                    
                    if "Min steel strain layer  6 :" in case_dict:
                        worksheet.write(row,37,float(case_dict["Min steel strain layer  6 :"]["value"]))
                        worksheet.write_comment(row,37,case_dict["Min steel strain layer  6 :"]["part"]+' '+case_dict["Min steel strain layer  6 :"]["node"]+"\n"+case_dict["Min steel strain layer  6 :"]["LC"])
                    else:
                        show_l6 = worksheet_write(worksheet,case_dict,"Min steel strain layer  7 :",row,37,show_l6)
                    
                    
                    if "Max steel strain layer  6 :" in case_dict:
                        worksheet.write(row,38,float(case_dict["Max steel strain layer  6 :"]["value"]))
                        worksheet.write_comment(row,38,case_dict["Max steel strain layer  6 :"]["part"]+' '+case_dict["Max steel strain layer  6 :"]["node"]+"\n"+case_dict["Max steel strain layer  6 :"]["LC"])
                    else:
                        show_l6 = worksheet_write(worksheet,case_dict,"Max steel stress layer  6 :",row,38,show_l6)
                        
                        
                    if "Min steel strain layer  7 :" in case_dict:
                        worksheet.write(row,39,float(case_dict["Min steel strain layer  7 :"]["value"]))
                        worksheet.write_comment(row,39,case_dict["Min steel strain layer  7 :"]["part"]+' '+case_dict["Min steel strain layer  7 :"]["node"]+"\n"+case_dict["Min steel strain layer  7 :"]["LC"])
                    else:
                        show_l7 = worksheet_write(worksheet,case_dict,"Min steel strain layer  7 :",row,39,show_l7)
                    
                    
                    if "Max steel strain layer  7 :" in case_dict:
                        worksheet.write(row,40,float(case_dict["Max steel strain layer  7 :"]["value"]))
                        worksheet.write_comment(row,40,case_dict["Max steel strain layer  7 :"]["part"]+' '+case_dict["Max steel strain layer  7 :"]["node"]+"\n"+case_dict["Max steel strain layer  7 :"]["LC"])
                    else:
                        show_l7 = worksheet_write(worksheet,case_dict,"Max steel stress layer  7 :",row,40,show_l7)
                    
                    
                    if "Min steel strain layer  8 :" in case_dict:
                        worksheet.write(row,41,float(case_dict["Min steel strain layer  8 :"]["value"]))
                        worksheet.write_comment(row,41,case_dict["Min steel strain layer  8 :"]["part"]+' '+case_dict["Min steel strain layer  8 :"]["node"]+"\n"+case_dict["Min steel strain layer  8 :"]["LC"])
                    else:
                        show_l8 = worksheet_write(worksheet,case_dict,"Min steel strain layer  8 :",row,41,show_l8)
                    
                    
                    if "Max steel strain layer  8 :" in case_dict:
                        worksheet.write(row,42,float(case_dict["Max steel strain layer  8 :"]["value"]))
                        worksheet.write_comment(row,42,case_dict["Max steel strain layer  8 :"]["part"]+' '+case_dict["Max steel strain layer  8 :"]["node"]+"\n"+case_dict["Max steel strain layer  8 :"]["LC"])
                    else:
                        show_l8 = worksheet_write(worksheet,case_dict,"Max steel stress layer  8 :",row,42,show_l8)

                    
                    row = row +1
    if row ==2:
        row = row+1                           
    
    set_table_blue_table_uls_als(row,worksheet,workbook)    
    width_columns_blue_table(worksheet)
    workbook.formats[0].set_font_size(8)
    worksheet.set_zoom(130)

    ##    Hide columns with all 'NA'
    if show_l5 == False:
        worksheet.set_column(13,13, None, None, {'hidden': 1})
        worksheet.set_column(23,23, None, None, {'hidden': 1})
        worksheet.set_column(35,35, None, None, {'hidden': 1})
        worksheet.set_column(36,36, None, None, {'hidden': 1})
    if show_l6 == False:
        worksheet.set_column(14,14, None, None, {'hidden': 1})
        worksheet.set_column(24,24, None, None, {'hidden': 1})
        worksheet.set_column(37,37, None, None, {'hidden': 1})
        worksheet.set_column(38,38, None, None, {'hidden': 1})
    if show_l7 == False:
        worksheet.set_column(15,15, None, None, {'hidden': 1})
        worksheet.set_column(25,25, None, None, {'hidden': 1})
        worksheet.set_column(39,39, None, None, {'hidden': 1})
        worksheet.set_column(40,40, None, None, {'hidden': 1})
    if show_l8 == False:
        worksheet.set_column(16,16, None, None, {'hidden': 1})
        worksheet.set_column(26,26, None, None, {'hidden': 1})
        worksheet.set_column(41,41, None, None, {'hidden': 1})
        worksheet.set_column(42,42, None, None, {'hidden': 1})
    if show_p1 == False:
        worksheet.set_column(17,17, None, None, {'hidden': 1})
    if show_p2 == False:
        worksheet.set_column(18,18, None, None, {'hidden': 1})
    

def summary_blue_table_uls_als_envelope(Envelope_result_dict,workbook,limit_state):    
    # Create a new workbook and add a worksheet
    worksheet = workbook.add_worksheet(limit_state + '_envelope')
    row = 2
    header = limit_state + ' result table. The table contains values read from .csmr-files. Results must be verified with .emf-tables.'+' Ignored Minimaxes: '+Text_ign
    worksheet.write(0,1,header)     
    
    show_l5 = False
    show_l6 = False
    show_l7 = False
    show_l8 = False
    show_p1 = False
    show_p2 = False
#
    for dg, dg_dict in Envelope_result_dict.items():
        
        worksheet.write(row,1,float(dg))
        worksheet.write(row,2,'env.')
        worksheet.write(row,4,limit_state)

        worksheet.write(row,5,round(float(Envelope_result_dict[dg]["env"][limit_state]["Min conc. stress bot face :"]["value"]),1))
        worksheet.write(row,6,round(float(Envelope_result_dict[dg]["env"][limit_state]["Min conc. stress top face :"]["value"]),1))
        worksheet.write(row,7,round(float(Envelope_result_dict[dg]["env"][limit_state]["Min conc. strain bot face :"]["value"]),1))
        worksheet.write(row,8,round(float(Envelope_result_dict[dg]["env"][limit_state]["Min conc. strain top face :"]["value"]),1))
        
        if max(abs(float(Envelope_result_dict[dg]["env"][limit_state]["Max steel stress layer  1 :"]["value"])),abs(float(Envelope_result_dict[dg]["env"][limit_state]["Min steel stress layer  1 :"]["value"])))<=abs(float(Envelope_result_dict[dg]["env"][limit_state]["Min steel stress layer  1 :"]["value"])):
            worksheet.write(row,9, round(float(Envelope_result_dict[dg]["env"][limit_state]["Min steel stress layer  1 :"]["value"]),0))
        else:
            worksheet.write(row,9, round(float(Envelope_result_dict[dg]["env"][limit_state]["Max steel stress layer  1 :"]["value"]),0))
            #worksheet.write_comment(row,7,Envelope_result_dict[dg]["env"][limit_state]["Max steel stress layer  1 :"]["part"]+' '+Envelope_result_dict[dg]["env"][limit_state]["Max steel stress layer  1 :"]["node"]+"\n"+Envelope_result_dict[dg]["env"][limit_state]["Max steel stress layer  1 :"]["LC"][-12:])
                    
        if max(abs(float(Envelope_result_dict[dg]["env"][limit_state]["Max steel stress layer  2 :"]["value"])),abs(float(Envelope_result_dict[dg]["env"][limit_state]["Min steel stress layer  2 :"]["value"])))<=abs(float(Envelope_result_dict[dg]["env"][limit_state]["Min steel stress layer  2 :"]["value"])):
            worksheet.write(row,10, round(float(Envelope_result_dict[dg]["env"][limit_state]["Min steel stress layer  2 :"]["value"]),0))
        else:
            worksheet.write(row,10, round(float(Envelope_result_dict[dg]["env"][limit_state]["Max steel stress layer  2 :"]["value"]),0))
        #worksheet.write_comment(row,8,Envelope_result_dict[dg]["env"][limit_state]["Max steel stress layer  2 :"]["part"]+' '+Envelope_result_dict[dg]["env"][limit_state]["Max steel stress layer  2 :"]["node"]+"\n"+Envelope_result_dict[dg]["env"][limit_state]["Max steel stress layer  2 :"]["LC"][-12:])
                    
        if max(abs(float(Envelope_result_dict[dg]["env"][limit_state]["Max steel stress layer  3 :"]["value"])),abs(float(Envelope_result_dict[dg]["env"][limit_state]["Min steel stress layer  3 :"]["value"])))<=abs(float(Envelope_result_dict[dg]["env"][limit_state]["Min steel stress layer  3 :"]["value"])):
            worksheet.write(row,11, round(float(Envelope_result_dict[dg]["env"][limit_state]["Min steel stress layer  3 :"]["value"]),0))
        else:
            worksheet.write(row,11, round(float(Envelope_result_dict[dg]["env"][limit_state]["Max steel stress layer  3 :"]["value"]),0))
                    #worksheet.write_comment(row,9,Envelope_result_dict[dg]["env"][limit_state]["Max steel stress layer  3 :"]["part"]+' '+Envelope_result_dict[dg]["env"][limit_state]["Max steel stress layer  3 :"]["node"]+"\n"+Envelope_result_dict[dg]["env"][limit_state]["Max steel stress layer  3 :"]["LC"][-12:])
                    
        if max(abs(float(Envelope_result_dict[dg]["env"][limit_state]["Max steel stress layer  4 :"]["value"])),abs(float(Envelope_result_dict[dg]["env"][limit_state]["Min steel stress layer  4 :"]["value"])))<=abs(float(Envelope_result_dict[dg]["env"][limit_state]["Min steel stress layer  4 :"]["value"])):
            worksheet.write(row,12, round(float(Envelope_result_dict[dg]["env"][limit_state]["Min steel stress layer  4 :"]["value"]),0))
        else:
            worksheet.write(row,12, round(float(Envelope_result_dict[dg]["env"][limit_state]["Max steel stress layer  4 :"]["value"]),0))

            
        show_l5 = worksheet_write_envelope(worksheet,dg,"Max steel stress layer  5 :",limit_state,13,row,show_l5)
        show_l6 = worksheet_write_envelope(worksheet,dg,"Max steel stress layer  6 :",limit_state,14,row,show_l6)
        show_l7 = worksheet_write_envelope(worksheet,dg,"Max steel stress layer  7 :",limit_state,15,row,show_l7)
        show_l8 = worksheet_write_envelope(worksheet,dg,"Max steel stress layer  8 :",limit_state,16,row,show_l8)
        
        show_p1 = worksheet_write_envelope(worksheet,dg,"Max prestress adstress  1 :",limit_state,17,row,show_p1)
        show_p2 = worksheet_write_envelope(worksheet,dg,"Max prestress adstress  2 :",limit_state,18,row,show_p2)        
                     
        worksheet.write(row,19,float(Envelope_result_dict[dg]["env"][limit_state]["Steel utilisation layer  1:"]["value"]))
        worksheet.write(row,20,float(Envelope_result_dict[dg]["env"][limit_state]["Steel utilisation layer  2:"]["value"]))
        worksheet.write(row,21,float(Envelope_result_dict[dg]["env"][limit_state]["Steel utilisation layer  3:"]["value"]))
        worksheet.write(row,22,float(Envelope_result_dict[dg]["env"][limit_state]["Steel utilisation layer  4:"]["value"]))                
        
        show_l5 = worksheet_write_envelope(worksheet,dg,"Steel utilisation layer  5:",limit_state,23,row,show_l5)
        show_l6 = worksheet_write_envelope(worksheet,dg,"Steel utilisation layer  6:",limit_state,24,row,show_l6)
        show_l7 = worksheet_write_envelope(worksheet,dg,"Steel utilisation layer  7:",limit_state,25,row,show_l7)
        show_l8 = worksheet_write_envelope(worksheet,dg,"Steel utilisation layer  8:",limit_state,26,row,show_l8)        
            
        row = row +1
    if row ==2:
        row = row+1                          
    
    
    set_table_blue_table_uls_als(row,worksheet,workbook)    
    width_columns_blue_table(worksheet)
    workbook.formats[0].set_font_size(8)
    worksheet.set_zoom(130)

    ##    Hide columns with all 'NA'
    if show_l5 == False:
        worksheet.set_column(13,13, None, None, {'hidden': 1})
        worksheet.set_column(23,23, None, None, {'hidden': 1})
    if show_l6 == False:
        worksheet.set_column(14,14, None, None, {'hidden': 1})
        worksheet.set_column(24,24, None, None, {'hidden': 1})
    if show_l7 == False:
        worksheet.set_column(15,15, None, None, {'hidden': 1})
        worksheet.set_column(25,25, None, None, {'hidden': 1})
    if show_l8 == False:
        worksheet.set_column(16,16, None, None, {'hidden': 1})
        worksheet.set_column(26,26, None, None, {'hidden': 1})
    
    if show_p1 == False:
        worksheet.set_column(17,17, None, None, {'hidden': 1})
    if show_p2 == False:
        worksheet.set_column(18,18, None, None, {'hidden': 1})
    
    worksheet.set_column(3,3, None, None, {'hidden': 1})
        
def summary_blue_table_sls_envelope(Envelope_result_dict,workbook,limit_state):    
    # Create a new workbook and add a worksheet
    worksheet = workbook.add_worksheet(limit_state + '_envelope')
    row = 2
    header = limit_state + ' result table. The table contains values read from .csmr-files. Results must be verified with .emf-tables.'+' Ignored Minimaxes: '+Text_ign
    worksheet.write(0,1,header)     
    
    show_l5 = False
    show_l6 = False
    show_l7 = False
    show_l8 = False
    show_p1 = False
    show_p2 = False
#
    for dg, dg_dict in Envelope_result_dict.items():
        
        worksheet.write(row,1,float(dg))
        worksheet.write(row,2,'env.')
        worksheet.write(row,4,limit_state)

        worksheet.write(row,5,round(float(Envelope_result_dict[dg]["env"][limit_state]["Min conc. stress bot face :"]["value"]),1))
        worksheet.write(row,6,round(float(Envelope_result_dict[dg]["env"][limit_state]["Min conc. stress top face :"]["value"]),1))
        
        if max(abs(float(Envelope_result_dict[dg]["env"][limit_state]["Max steel stress layer  1 :"]["value"])),abs(float(Envelope_result_dict[dg]["env"][limit_state]["Min steel stress layer  1 :"]["value"])))<=abs(float(Envelope_result_dict[dg]["env"][limit_state]["Min steel stress layer  1 :"]["value"])):
            worksheet.write(row,7, round(float(Envelope_result_dict[dg]["env"][limit_state]["Min steel stress layer  1 :"]["value"]),0))
        else:
            worksheet.write(row,7, round(float(Envelope_result_dict[dg]["env"][limit_state]["Max steel stress layer  1 :"]["value"]),0))
            #worksheet.write_comment(row,7,Envelope_result_dict[dg]["env"][limit_state]["Max steel stress layer  1 :"]["part"]+' '+Envelope_result_dict[dg]["env"][limit_state]["Max steel stress layer  1 :"]["node"]+"\n"+Envelope_result_dict[dg]["env"][limit_state]["Max steel stress layer  1 :"]["LC"][-12:])
                    
        if max(abs(float(Envelope_result_dict[dg]["env"][limit_state]["Max steel stress layer  2 :"]["value"])),abs(float(Envelope_result_dict[dg]["env"][limit_state]["Min steel stress layer  2 :"]["value"])))<=abs(float(Envelope_result_dict[dg]["env"][limit_state]["Min steel stress layer  2 :"]["value"])):
            worksheet.write(row,8, round(float(Envelope_result_dict[dg]["env"][limit_state]["Min steel stress layer  2 :"]["value"]),0))
        else:
            worksheet.write(row,8, round(float(Envelope_result_dict[dg]["env"][limit_state]["Max steel stress layer  2 :"]["value"]),0))
        #worksheet.write_comment(row,8,Envelope_result_dict[dg]["env"][limit_state]["Max steel stress layer  2 :"]["part"]+' '+Envelope_result_dict[dg]["env"][limit_state]["Max steel stress layer  2 :"]["node"]+"\n"+Envelope_result_dict[dg]["env"][limit_state]["Max steel stress layer  2 :"]["LC"][-12:])
                    
        if max(abs(float(Envelope_result_dict[dg]["env"][limit_state]["Max steel stress layer  3 :"]["value"])),abs(float(Envelope_result_dict[dg]["env"][limit_state]["Min steel stress layer  3 :"]["value"])))<=abs(float(Envelope_result_dict[dg]["env"][limit_state]["Min steel stress layer  3 :"]["value"])):
            worksheet.write(row,9, round(float(Envelope_result_dict[dg]["env"][limit_state]["Min steel stress layer  3 :"]["value"]),0))
        else:
            worksheet.write(row,9, round(float(Envelope_result_dict[dg]["env"][limit_state]["Max steel stress layer  3 :"]["value"]),0))
                    #worksheet.write_comment(row,9,Envelope_result_dict[dg]["env"][limit_state]["Max steel stress layer  3 :"]["part"]+' '+Envelope_result_dict[dg]["env"][limit_state]["Max steel stress layer  3 :"]["node"]+"\n"+Envelope_result_dict[dg]["env"][limit_state]["Max steel stress layer  3 :"]["LC"][-12:])
                    
        if max(abs(float(Envelope_result_dict[dg]["env"][limit_state]["Max steel stress layer  4 :"]["value"])),abs(float(Envelope_result_dict[dg]["env"][limit_state]["Min steel stress layer  4 :"]["value"])))<=abs(float(Envelope_result_dict[dg]["env"][limit_state]["Min steel stress layer  4 :"]["value"])):
            worksheet.write(row,10, round(float(Envelope_result_dict[dg]["env"][limit_state]["Min steel stress layer  4 :"]["value"]),0))
        else:
            worksheet.write(row,10, round(float(Envelope_result_dict[dg]["env"][limit_state]["Max steel stress layer  4 :"]["value"]),0))
            
        show_l5 = worksheet_write_envelope(worksheet,dg,"Max steel stress layer  5 :",limit_state,11,row,show_l5)
        show_l6 = worksheet_write_envelope(worksheet,dg,"Max steel stress layer  6 :",limit_state,12,row,show_l6)
        show_l7 = worksheet_write_envelope(worksheet,dg,"Max steel stress layer  7 :",limit_state,13,row,show_l7)
        show_l8 = worksheet_write_envelope(worksheet,dg,"Max steel stress layer  8 :",limit_state,14,row,show_l8)
        
        show_p1 = worksheet_write_envelope(worksheet,dg,"Max prestress adstress  1 :",limit_state,15,row,show_p1)
        show_p2 = worksheet_write_envelope(worksheet,dg,"Max prestress adstress  2 :",limit_state,16,row,show_p2)        

        worksheet.write(row,18, round(float(Envelope_result_dict[dg]["env"][limit_state]["wk_static"]["value"]),2))
        worksheet.write(row,25, float(Envelope_result_dict[dg]["env"][limit_state]["wk_thermal"]["req"]))
        
        worksheet.write(row,20, round(float(Envelope_result_dict[dg]["env"][limit_state]["wk_thermal"]["value"]),2))
        worksheet.write(row,26, float(Envelope_result_dict[dg]["env"][limit_state]["wk_thermal"]["req"]))
                              
        
        ##Compression depth
        if float(Envelope_result_dict[dg]["env"][limit_state]["Min Therm comp  depth (mm):_thermal"]["value_with_req"]) < 1000:
            worksheet.write(row,17, round(float(Envelope_result_dict[dg]["env"][limit_state]["Min compression depth (mm):"]["value_with_req"]),1))
        else:
            worksheet.write(row,17,round(float(Envelope_result_dict[dg]["env"][limit_state]["Min compression depth (mm):"]["value_without_req"]),1))                 
        worksheet.write(row,27,round(float(Envelope_result_dict[dg]['env'][limit_state]["Min compression depth (mm):"]["req"]),1))                 

        if float(Envelope_result_dict[dg]["env"][limit_state]["Min Therm comp  depth (mm):_thermal"]["value_with_req"]) < 1000:
            worksheet.write(row,19,round(float(Envelope_result_dict[dg]["env"][limit_state]["Min Therm comp  depth (mm):_thermal"]["value_with_req"]),1))                 
        else:
            worksheet.write(row,19,round(float(Envelope_result_dict[dg]["env"][limit_state]["Min Therm comp  depth (mm):_thermal"]["value_without_req"]),1))                 
        worksheet.write(row,28,round(float(Envelope_result_dict[dg]['env'][limit_state]["Min Therm comp  depth (mm):_thermal"]["req"]),1))                 

        row = row +1
        
    if row ==2:
        row = row+1                           
    
    
    set_table_blue_table_sls(row,worksheet,workbook)    
    width_columns_blue_table(worksheet)
    workbook.formats[0].set_font_size(8)
    worksheet.set_zoom(130)

    ##    Hide columns with all 'NA'
    if show_l5 == False:
        worksheet.set_column(11,11, None, None, {'hidden': 1})
    if show_l6 == False:
        worksheet.set_column(12,12, None, None, {'hidden': 1})
    if show_l7 == False:
        worksheet.set_column(13,13, None, None, {'hidden': 1})
    if show_l8 == False:
        worksheet.set_column(14,14, None, None, {'hidden': 1})
    if show_p1 == False:
        worksheet.set_column(15,15, None, None, {'hidden': 1})
    if show_p2 == False:
        worksheet.set_column(16,16, None, None, {'hidden': 1})

    worksheet.set_column(3,3, None, None, {'hidden': 1})        

    red = workbook.add_format({'bg_color': '#F6B7B7'})           
        
    worksheet.conditional_format('S3:S'+str(row), {'type': 'cell',
                                             'criteria': '>',
                                             'value': 'Z3:Z'+str(row),
                                             'format': red})
        
    worksheet.conditional_format('U3:U'+str(row), {'type': 'cell',
                                             'criteria': '>',
                                             'value': 'AA3:AA'+str(row),
                                             'format': red})    

    
    worksheet.conditional_format('R3:R'+str(row), {'type': 'cell',
                                             'criteria': '<',
                                             'value': 'AB3:AB'+str(row),
                                             'format': red})
        
    worksheet.conditional_format('T3:T'+str(row), {'type': 'cell',
                                             'criteria': '<',
                                             'value': 'AC3:AC'+str(row),
                                             'format': red})    


    

def summary_blue_table_sls(All_result_dict,workbook):    
    # Create a new workbook and add a worksheet
    worksheet = workbook.add_worksheet('SLS_long')
    row = 2
    header = 'SLS result table. The table contains values read from .csmr-files. Results must be verified with .emf-tables.'+' Ignored Minimaxes: '+Text_ign
    worksheet.write(0,1,header)     
    limit_state = 'SLS'
    
    show_l5 = False
    show_l6 = False
    show_l7 = False
    show_l8 = False
    show_p1 = False
    show_p2 = False    
#
    for dg, dg_dict in All_result_dict.items():
        for thermal_case,thermal_dict in dg_dict.items():
            if thermal_case == 'req_link':
                continue        
            for case, case_dict in thermal_dict.items():
                if case == 'req_link':
                    continue
                if case_dict["limit_state"] == 'SLS':
                    
                    worksheet.write(row,1,float(dg))
                    worksheet.write(row,2,float(case))
                    worksheet.write(row,3,float(case_dict["lc_thermal"]))
                    worksheet.write(row,4,case_dict["limit_state"])
                    worksheet.write(row,5,round(float(case_dict["Min conc. stress bot face :"]["value"]),1))
                    worksheet.write_comment(row,5,case_dict["Min conc. stress bot face :"]["part"]+" "+case_dict["Min conc. stress bot face :"]["node"]+'\n'+case_dict["Min conc. stress bot face :"]["LC"][-12:])
                    
                    worksheet.write(row,6,round(float(case_dict["Min conc. stress top face :"]["value"]),1))
                    worksheet.write_comment(row,6,case_dict["Min conc. stress top face :"]["part"]+' '+case_dict["Min conc. stress top face :"]["node"]+"\n"+case_dict["Min conc. stress top face :"]["LC"][-12:])
                    
                    
                    worksheet.write(row,7, round(float(case_dict["Max steel stress layer  1 :"]["value"]),0))
                    worksheet.write_comment(row,7,case_dict["Max steel stress layer  1 :"]["part"]+' '+case_dict["Max steel stress layer  1 :"]["node"]+"\n"+case_dict["Max steel stress layer  1 :"]["LC"][-12:])

                    worksheet.write(row,8, round(float(case_dict["Max steel stress layer  2 :"]["value"]),0))
                    worksheet.write_comment(row,8,case_dict["Max steel stress layer  2 :"]["part"]+' '+case_dict["Max steel stress layer  2 :"]["node"]+"\n"+case_dict["Max steel stress layer  2 :"]["LC"][-12:])
                    
                    worksheet.write(row,9, round(float(case_dict["Max steel stress layer  3 :"]["value"]),0))
                    worksheet.write_comment(row,9,case_dict["Max steel stress layer  3 :"]["part"]+' '+case_dict["Max steel stress layer  3 :"]["node"]+"\n"+case_dict["Max steel stress layer  3 :"]["LC"][-12:])
                    
                    worksheet.write(row,10,round(float(case_dict["Max steel stress layer  4 :"]["value"]),0))
                    worksheet.write_comment(row,10,case_dict["Max steel stress layer  4 :"]["part"]+' '+case_dict["Max steel stress layer  4 :"]["node"]+"\n"+case_dict["Max steel stress layer  4 :"]["LC"][-12:])
                    
#                           
                    show_l5 = worksheet_write(worksheet,case_dict,"Max steel stress layer  5 :",row,11,show_l5)
                    show_l6 = worksheet_write(worksheet,case_dict,"Max steel stress layer  6 :",row,12,show_l6)
                    show_l7 = worksheet_write(worksheet,case_dict,"Max steel stress layer  7 :",row,13,show_l7)
                    show_l8 = worksheet_write(worksheet,case_dict,"Max steel stress layer  8 :",row,14,show_l8)
                    show_p1 = worksheet_write(worksheet,case_dict,"Max prestress adstress  1 :",row,15,show_p1)
                    show_p2 = worksheet_write(worksheet,case_dict,"Max prestress adstress  2 :",row,16,show_p2)                                          
                        
                    worksheet.write(row,17, float(case_dict["Min compression depth (mm):"]["value"]))
                    worksheet.write_comment(row,17,case_dict["Min compression depth (mm):"]["part"]+' '+case_dict["Min compression depth (mm):"]["node"]+"\n"+case_dict["Min compression depth (mm):"]["LC"][-12:])
                    
                    worksheet.write(row,18, round(float(case_dict["wk_static"]["value"]),2))
                    #worksheet.write_comment(row,18,thermal_dict[case]['wk_static']['part']+' '+thermal_dict[case]['wk_static']['node']+"\n"+thermal_dict[case]['wk_static']['LC'][-12:])
                    worksheet.write(row,19, float(case_dict["Min Therm comp  depth (mm):_thermal"]["value"]))
                    worksheet.write_comment(row,19,case_dict["Min Therm comp  depth (mm):_thermal"]["part"]+' '+case_dict["Min Therm comp  depth (mm):_thermal"]["node"]+"\n"+case_dict["Min Therm comp  depth (mm):_thermal"]["LC"][-12:])
                    
                    worksheet.write(row,20, round(float(case_dict["wk_thermal"]["value"]),2))
                    #worksheet.write_comment(row,20,thermal_dict[case]['wk_thermal']['part']+' '+thermal_dict[case]['wk_thermal']['node']+"\n"+thermal_dict[case]['wk_thermal']['LC'][-12:])
                    
                                    
                    
                    if Requirements_dict[dg][case]["C"] == 'n/a':
                        worksheet.write(row,23,-1)
                    else:
                        worksheet.write(row,23, float(Requirements_dict[dg][case]["C"]))
                    if Requirements_dict[dg][case]["wk+"]!='n/a' or Requirements_dict[dg][case]["wk-"]!='n/a':
                        worksheet.write(row,24, min(float(Requirements_dict[dg][case]["wk+"]),float(Requirements_dict[dg][case]["wk-"])))
                    else:
                        worksheet.write(row,24, Requirements_dict[dg][case]["wk+"])
                        
                      
                    row = row +1
    if row ==2:
        row = row+1                           
    
    
    set_table_blue_table_sls(row,worksheet,workbook)    
    width_columns_blue_table(worksheet)
    workbook.formats[0].set_font_size(8)
    worksheet.set_zoom(130)

    #    Hide columns with all 'NA'
   
    if show_p1 == False:
        worksheet.set_column(15,15, None, None, {'hidden': 1})
    if show_p2 == False:
        worksheet.set_column(16,16, None, None, {'hidden': 1})
    if show_l5 == False:
        worksheet.set_column(11,11, None, None, {'hidden': 1})
    if show_l6 == False:
        worksheet.set_column(12,12, None, None, {'hidden': 1})
    if show_l7 == False:
        worksheet.set_column(13,13, None, None, {'hidden': 1})
    if show_l8 == False:
        worksheet.set_column(14,14, None, None, {'hidden': 1})
        
    
    red = workbook.add_format({'bg_color': '#F6B7B7'})   
    worksheet.conditional_format('R3:R'+str(row), {'type': 'cell',
                                             'criteria': '<',
                                             'value': 'X3:X'+str(row),
                                             'format': red})
        
    worksheet.conditional_format('T3:T'+str(row), {'type': 'cell',
                                             'criteria': '<',
                                             'value': 'X3:X'+str(row),
                                             'format': red})
    
    worksheet.conditional_format('S3:S'+str(row), {'type': 'cell',
                                             'criteria': '>',
                                             'value': 'Y3:Y'+str(row),
                                             'format': red})
        
    worksheet.conditional_format('U3:U'+str(row), {'type': 'cell',
                                             'criteria': '>',
                                             'value': 'Y3:Y'+str(row),
                                             'format': red})
        
            
    
def envelope_SLS(limit_state):
    Envelope_result_dict[dg]['env'][limit_state]["Min conc. stress bot face :"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Min conc. stress top face :"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  1 :"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  2 :"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  3 :"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  4 :"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  5 :"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  6 :"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  7 :"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  8 :"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  1 :"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  2 :"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  3 :"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  4 :"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  5 :"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  6 :"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  7 :"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  8 :"] = {}

    Envelope_result_dict[dg]['env'][limit_state]["Max prestress adstress  1 :"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Max prestress adstress  2 :"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["wk_static"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["wk_thermal"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Min compression depth (mm):"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Min Therm comp  depth (mm):_thermal"] = {}

    Envelope_result_dict[dg]['env'][limit_state]["Min conc. stress bot face :"]["value"] = str('100')
    Envelope_result_dict[dg]['env'][limit_state]["Min conc. stress top face :"]["value"] = str('100')
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  1 :"]["value"] = str('0')
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  2 :"]["value"] = str('0')
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  3 :"]["value"] = str('0')
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  4 :"]["value"] = str('0')
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  5 :"]["value"] = str('0')
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  6 :"]["value"] = str('0')
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  7 :"]["value"] = str('0')
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  8 :"]["value"] = str('0')
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  1 :"]["value"] = str('0')
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  2 :"]["value"] = str('0')
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  3 :"]["value"] = str('0')
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  4 :"]["value"] = str('0')
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  5 :"]["value"] = str('0')
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  6 :"]["value"] = str('0')
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  7 :"]["value"] = str('0')
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  8 :"]["value"] = str('0')

    Envelope_result_dict[dg]['env'][limit_state]["Max prestress adstress  1 :"]["value"] = str('-1000')
    Envelope_result_dict[dg]['env'][limit_state]["Max prestress adstress  2 :"]["value"] = str('-1000')
    
    Envelope_result_dict[dg]['env'][limit_state]["wk_static"]["value"] = str(-1)
    Envelope_result_dict[dg]['env'][limit_state]["wk_static"]["UR"] = str(-1)
    Envelope_result_dict[dg]['env'][limit_state]["wk_static"]["req"] = str(-1)
    
    Envelope_result_dict[dg]['env'][limit_state]["wk_thermal"]["value"] = str(-1)
    Envelope_result_dict[dg]['env'][limit_state]["wk_thermal"]["UR"] = str(-1)
    Envelope_result_dict[dg]['env'][limit_state]["wk_thermal"]["req"] = str(-1)
    
    Envelope_result_dict[dg]['env'][limit_state]["Min compression depth (mm):"]["value_with_req"] = str(1000)
    Envelope_result_dict[dg]['env'][limit_state]["Min compression depth (mm):"]["value_without_req"] = str(1000)
    Envelope_result_dict[dg]['env'][limit_state]["Min compression depth (mm):"]["req"] = str(-1)
    Envelope_result_dict[dg]['env'][limit_state]["Min compression depth (mm):"]["no_req"] = str(-1)
    
    Envelope_result_dict[dg]['env'][limit_state]["Min Therm comp  depth (mm):_thermal"]["value_with_req"] = str(1000)
    Envelope_result_dict[dg]['env'][limit_state]["Min Therm comp  depth (mm):_thermal"]["value_without_req"] = str(1000)
    Envelope_result_dict[dg]['env'][limit_state]["Min Therm comp  depth (mm):_thermal"]["req"] = str(-1)
    Envelope_result_dict[dg]['env'][limit_state]["Min Therm comp  depth (mm):_thermal"]["no_req"] = str(-1)
    
    
    Envelope_result_dict[dg]['env'][limit_state]["Min conc. stress bot face :"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Min conc. stress top face :"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  1 :"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  2 :"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  3 :"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  4 :"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  5 :"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  6 :"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  7 :"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  8 :"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  1 :"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  2 :"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  3 :"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  4 :"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  5 :"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  6 :"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  7 :"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  8 :"]["changed"] = False

    Envelope_result_dict[dg]['env'][limit_state]["Max prestress adstress  1 :"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Max prestress adstress  2 :"]["changed"] = False
    
def envelope_ULS_ALS(limit_state):
    Envelope_result_dict[dg]['env'][limit_state]["Min conc. stress bot face :"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Min conc. stress top face :"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Min conc. strain bot face :"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Min conc. strain top face :"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  1 :"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  2 :"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  3 :"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  4 :"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  5 :"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  6 :"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  7 :"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  8 :"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  1 :"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  2 :"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  3 :"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  4 :"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  5 :"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  6 :"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  7 :"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  8 :"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Max prestress adstress  1 :"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Max prestress adstress  2 :"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Steel utilisation layer  1:"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Steel utilisation layer  2:"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Steel utilisation layer  3:"] = {}
    Envelope_result_dict[dg]['env'][limit_state]["Steel utilisation layer  4:"] = {} 
    Envelope_result_dict[dg]['env'][limit_state]["Steel utilisation layer  5:"] = {} 
    Envelope_result_dict[dg]['env'][limit_state]["Steel utilisation layer  6:"] = {} 
    Envelope_result_dict[dg]['env'][limit_state]["Steel utilisation layer  7:"] = {} 
    Envelope_result_dict[dg]['env'][limit_state]["Steel utilisation layer  8:"] = {} 
    
    Envelope_result_dict[dg]['env'][limit_state]["Min conc. stress bot face :"]["value"] = str('100')
    Envelope_result_dict[dg]['env'][limit_state]["Min conc. stress top face :"]["value"] = str('100')
    Envelope_result_dict[dg]['env'][limit_state]["Min conc. strain bot face :"]["value"] = str('100')
    Envelope_result_dict[dg]['env'][limit_state]["Min conc. strain top face :"]["value"] = str('100')
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  1 :"]["value"] = str('0')
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  2 :"]["value"] = str('0')
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  3 :"]["value"] = str('0')
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  4 :"]["value"] = str('0')
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  5 :"]["value"] = str('0')
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  6 :"]["value"] = str('0')
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  7 :"]["value"] = str('0')
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  8 :"]["value"] = str('0')
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  1 :"]["value"] = str('0')
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  2 :"]["value"] = str('0')
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  3 :"]["value"] = str('0')
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  4 :"]["value"] = str('0')
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  5 :"]["value"] = str('0')
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  6 :"]["value"] = str('0')
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  7 :"]["value"] = str('0')
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  8 :"]["value"] = str('0')
    Envelope_result_dict[dg]['env'][limit_state]["Max prestress adstress  1 :"]["value"] = str('-1000')
    Envelope_result_dict[dg]['env'][limit_state]["Max prestress adstress  2 :"]["value"] = str('-1000')
    Envelope_result_dict[dg]['env'][limit_state]["Steel utilisation layer  1:"]["value"] = str('-1')
    Envelope_result_dict[dg]['env'][limit_state]["Steel utilisation layer  2:"]["value"] = str('-1')
    Envelope_result_dict[dg]['env'][limit_state]["Steel utilisation layer  3:"]["value"] = str('-1')
    Envelope_result_dict[dg]['env'][limit_state]["Steel utilisation layer  4:"]["value"] = str('-1')      
    Envelope_result_dict[dg]['env'][limit_state]["Steel utilisation layer  5:"]["value"] = str('-1')      
    Envelope_result_dict[dg]['env'][limit_state]["Steel utilisation layer  6:"]["value"] = str('-1')      
    Envelope_result_dict[dg]['env'][limit_state]["Steel utilisation layer  7:"]["value"] = str('-1')      
    Envelope_result_dict[dg]['env'][limit_state]["Steel utilisation layer  8:"]["value"] = str('-1')  
    
    Envelope_result_dict[dg]['env'][limit_state]["Min conc. stress bot face :"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Min conc. stress top face :"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Min conc. strain bot face :"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Min conc. strain top face :"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  1 :"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  2 :"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  3 :"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  4 :"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  5 :"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  6 :"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  7 :"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Max steel stress layer  8 :"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  1 :"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  2 :"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  3 :"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  4 :"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  5 :"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  6 :"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  7 :"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Min steel stress layer  8 :"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Max prestress adstress  1 :"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Max prestress adstress  2 :"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Steel utilisation layer  1:"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Steel utilisation layer  2:"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Steel utilisation layer  3:"]["changed"] = False
    Envelope_result_dict[dg]['env'][limit_state]["Steel utilisation layer  4:"]["changed"] = False   
    Envelope_result_dict[dg]['env'][limit_state]["Steel utilisation layer  5:"]["changed"] = False   
    Envelope_result_dict[dg]['env'][limit_state]["Steel utilisation layer  6:"]["changed"] = False   
    Envelope_result_dict[dg]['env'][limit_state]["Steel utilisation layer  7:"]["changed"] = False   
    Envelope_result_dict[dg]['env'][limit_state]["Steel utilisation layer  8:"]["changed"] = False    

def check_envelope(result,limit_state,bracket):
    if bracket == 'abs' and 'steel stress layer' in result:
        if abs(float(case_dict[result]["value"])) >= abs(float(Envelope_result_dict[dg]['env'][limit_state][result]["value"])):
            Envelope_result_dict[dg]['env'][limit_state][result]["value"] = case_dict[result]["value"]
            Envelope_result_dict[dg]['env'][limit_state][result]["changed"] = True
    if bracket == 'larger':
        if float(case_dict[result]["value"]) > float(Envelope_result_dict[dg]['env'][limit_state][result]["value"]):
            Envelope_result_dict[dg]['env'][limit_state][result]["value"] = case_dict[result]["value"]
            Envelope_result_dict[dg]['env'][limit_state][result]["changed"] = True
    elif bracket == 'smaller':
        if float(case_dict[result]["value"]) < float(Envelope_result_dict[dg]['env'][limit_state][result]["value"]):
            Envelope_result_dict[dg]['env'][limit_state][result]["value"] = case_dict[result]["value"]
            Envelope_result_dict[dg]['env'][limit_state][result]["changed"] = True

def check_envelope_try(result,limit_state,bracket):
    try:
        if bracket == 'larger':
            if float(case_dict[result]["value"]) > float(Envelope_result_dict[dg]['env'][limit_state][result]["value"]):
                Envelope_result_dict[dg]['env'][limit_state][result]["value"] = case_dict[result]["value"]
                Envelope_result_dict[dg]['env'][limit_state][result]["changed"] = True
        elif bracket == 'smaller':
            if float(case_dict[result]["value"]) < float(Envelope_result_dict[dg]['env'][limit_state][result]["value"]):
                Envelope_result_dict[dg]['env'][limit_state][result]["value"] = case_dict[result]["value"]
                Envelope_result_dict[dg]['env'][limit_state][result]["changed"] = True
    except:
        pass    
    
def check_envelope_try_wk(dg,case,result,limit_state,bracket):
    try:
        if bracket == 'larger':
            wk_req = min(float(Requirements_dict[dg][case]["wk+"]),float(Requirements_dict[dg][case]["wk-"])) 
            wk = float(case_dict[result]["value"])
            ur_wk = wk/wk_req
            
            if ur_wk > float(Envelope_result_dict[dg]['env'][limit_state][result]["UR"]):

                Envelope_result_dict[dg]['env'][limit_state][result]["value"] = case_dict[result]["value"]
                
                Envelope_result_dict[dg]['env'][limit_state][result]["UR"] = ur_wk
                Envelope_result_dict[dg]['env'][limit_state][result]["req"] = wk_req
                
                Envelope_result_dict[dg]['env'][limit_state][result]["changed"] = True
                            
    except:
        pass       
    
def check_envelope_try_c(dg,case_dict,case,result,limit_state,bracket):
    try:
        if bracket == 'smaller':
            if Requirements_dict[dg][case]["C"] == 'n/a':
                c_req = 'NA'
            else:
                c_req = float(Requirements_dict[dg][case]["C"])
            
            if c_req != 'NA':
                if float(case_dict[result]["value"]) < float(Envelope_result_dict[dg]['env'][limit_state][result]["value_with_req"]):
                    Envelope_result_dict[dg]['env'][limit_state][result]["value_with_req"] = case_dict[result]["value"]
                    Envelope_result_dict[dg]['env'][limit_state][result]["req"] = c_req
            else:
                if float(case_dict[result]["value"]) < float(Envelope_result_dict[dg]['env'][limit_state][result]["value_without_req"]):
                    Envelope_result_dict[dg]['env'][limit_state][result]["value_without_req"] = case_dict[result]["value"]
    except:
        pass       
    
    
def worksheet_write_envelope(worksheet,dg,result,limit_state,column,row,show_factor):
    if Envelope_result_dict[dg]["env"][limit_state][result]["changed"] == True:
        worksheet.write(row,column,round(float(Envelope_result_dict[dg]["env"][limit_state][result]["value"]),0))
        show_factor = True
    else:
        worksheet.write(row,column,'NA')
    return show_factor
        
def worksheet_write(worksheet,case_dict,result,row,column,show_factor):      
    try: 
        worksheet.write(row,column,round(float(case_dict[result]["value"]),0))
        show_factor = True
    except:
        worksheet.write(row,column,'NA')      
    
    return show_factor
        
                            


#%% --- MAIN SCRIPT STARTS BELOW: ----------------------------------------------#

## Loop through .csmr files to extract summary tables for appendix A
SLS_check = ['wk','partial crack']
LC = '00'
directory = os.getcwd()



# --- Find all .csmr files, .shmr files and properties files 
csmr_files_list, req_link,folder_req = list_csmr_files(directory)
shmr_files,res_files = list_shmr_files(directory)


#%% --- All results dictionary and requirement dictionary
All_result_dict = {}
Requirements_dict = {}
soil_series = '1'
soil_series_emf = '1'
for file in csmr_files_list:

    with open(file) as f:      

        wk_list = []
        part_list = []
        node_list = []
        LC_list = []

        thermal_effects = 'False'
        for line_no, line in enumerate(f):

            if "DG NUMBER" in line:
                
                ##
                #Input to all_result_dict
                dg = line.split()[5]
                case = line.split()[9] +'.'+ line.split()[12]
                lc_thermal = line.split()[len(line.split())-1]
                
                if case[len(case)-1] == '.':
                    case = case[:len(case)-1]
                temp_case = temp_case_function(file)
            
                try:
                    test = All_result_dict[dg]
                    test = Requirements_dict[dg]
                except:
                    All_result_dict[dg] = {}
                    Requirements_dict[dg] = {}
                
                try:
                    test = All_result_dict[dg][lc_thermal]
                    test = Requirements_dict[dg][lc_thermal]
                except:
                    All_result_dict[dg][lc_thermal] = {}
                    Requirements_dict[dg][lc_thermal] = {}
                
            
                All_result_dict[dg][lc_thermal][case] = {}
                Requirements_dict[dg][case] = {}
                
                
                ##
                
                All_result_dict[dg][lc_thermal][case]['limit_state'] = line.split()[6]
                limit_state = line.split()[6]
                All_result_dict[dg][lc_thermal][case]['lc_thermal'] = line.split()[len(line.split())-1]
                
                soil_series = line.split()[12] 
                
                if case[0] == '1' or case[0:2] == '20':
                    soil_series_emf = line.split()[12]                

            if "DATA_BASE VERSION" in line:
                All_result_dict[dg][lc_thermal][case]['Data_base_version'] = line.split()[3]
                Data_base_version= line.split()[3]
                links_and_requirement(All_result_dict,dg,case,file,limit_state,Data_base_version,temp_case,folder_req,soil_series_emf)

            # Read csmr datas
            if line_no > 4:
                excluded_points = 'False'
                excluded_points, thermal_effects = read_csmr_all(line,dg,case, All_result_dict,LC,line_no,thermal_effects)
                if excluded_points == 'True':
                    break

                #Find largest wk
                for req in SLS_check:
                    if req in line:
                         if len(line[34:].split()) > 3:
                             wk_list.append(line[34:].split()[0])
                             part_list.append(line[34:].split()[1])
                             node_list.append(line[34:].split()[2])
                         if len(line[34:].split()) > 4:
                             LC_list.append(line[34:].split()[3] + line[34:].split()[4])
                         elif len(line[34:].split()) > 3:
                             LC_list.append(line[34:].split()[3])
        
        if not wk_list:
            continue
        else:
            (wk_max,i) = max((v,i) for i,v in enumerate(wk_list))
            req = 'wk'
            All_result_dict[dg][lc_thermal][case][req] = {}
            All_result_dict[dg][lc_thermal][case][req]['value'] = wk_max
            All_result_dict[dg][lc_thermal][case][req]['part'] = part_list[i]
            All_result_dict[dg][lc_thermal][case][req]['node'] = node_list[i]
            All_result_dict[dg][lc_thermal][case][req]['LC'] = LC_list[i]

            
            
######################     REad static and thermal wk        

SLS_check = ['wk','partial']


for dg, dg_dict in All_result_dict.items():
    for thermal_case,thermal_dict in dg_dict.items():    
        if thermal_case == 'req_link':
            continue
        for case, case_dict in thermal_dict.items():
            wk_list_thermal = []
            part_list_thermal = []
            node_list_thermal = []
            LC_list_thermal = []
            
            wk_list_static = []
            part_list_static = []
            node_list_static = []
            LC_list_static = []
            
            wk_list = []
            part_list = []
            node_list = []
            LC_list = []
            for req, req_dict in case_dict.items():
                #Find largest wk
                for check in SLS_check:
                    if check in req:
                         try:
                             part_list.append(req_dict['part'])
                             wk_list.append(req_dict['value'])
                             node_list.append(req_dict['node'])
                             LC_list.append(req_dict['LC'])
                         except:
                             continue
                    if check in req and 'thermal' in req:
                         if req == 'wk':
                             continue
                         try:
                             part_list_thermal.append(req_dict['part'])
                             wk_list_thermal.append(req_dict['value'])
                             node_list_thermal.append(req_dict['node'])
                             LC_list_thermal.append(req_dict['LC'])
                         except:
                             continue 
                              
                    if check in req and 'thermal' not in req:
                         if req == 'wk':
                             continue
                         try:
                             part_list_static.append(req_dict['part'])
                             wk_list_static.append(req_dict['value'])
                             node_list_static.append(req_dict['node'])
                             LC_list_static.append(req_dict['LC'])
                         except:
                             continue                      
    
            if not wk_list:
                req = 'wk_2'
                All_result_dict[dg][thermal_case][case][req] = {}
                All_result_dict[dg][thermal_case][case][req]['value'] = '0'
#                continue
            else:
                (wk_max,i) = max((v,i) for i,v in enumerate(wk_list))
                req = 'wk_2'
                All_result_dict[dg][thermal_case][case][req] = {}
                All_result_dict[dg][thermal_case][case][req]['value'] = wk_max
                All_result_dict[dg][thermal_case][case][req]['part'] = part_list[i]
                All_result_dict[dg][thermal_case][case][req]['node'] = node_list[i]
                All_result_dict[dg][thermal_case][case][req]['LC'] = LC_list[i]
                
            if not wk_list_static:
                req = 'wk_static'
                All_result_dict[dg][thermal_case][case][req] = {}
                All_result_dict[dg][thermal_case][case][req]['value'] = '0'
#                continue
            else:
                (wk_max,i) = max((v,i) for i,v in enumerate(wk_list_static))
                req = 'wk_static'
                All_result_dict[dg][thermal_case][case][req] = {}
                All_result_dict[dg][thermal_case][case][req]['value'] = wk_max
                All_result_dict[dg][thermal_case][case][req]['part'] = part_list_static[i]
                All_result_dict[dg][thermal_case][case][req]['node'] = node_list_static[i]
                All_result_dict[dg][thermal_case][case][req]['LC'] = LC_list_static[i]
                
            if not wk_list_thermal:
                req = 'wk_thermal'
                All_result_dict[dg][thermal_case][case][req] = {}
                All_result_dict[dg][thermal_case][case][req]['value'] = '0'
#                continue
            else:
                (wk_max,i) = max((v,i) for i,v in enumerate(wk_list_thermal))
                req = 'wk_thermal'
                All_result_dict[dg][thermal_case][case][req] = {}
                All_result_dict[dg][thermal_case][case][req]['value'] = wk_max
                All_result_dict[dg][thermal_case][case][req]['part'] = part_list_thermal[i]
                All_result_dict[dg][thermal_case][case][req]['node'] = node_list_thermal[i]
                All_result_dict[dg][thermal_case][case][req]['LC'] = LC_list_thermal[i]
            
################## Create Envelope dict
Envelope_result_dict = {}                
for dg, dg_dict in All_result_dict.items():
    Envelope_result_dict[dg] = {}                
    Envelope_result_dict[dg]['env'] = {}                
    Envelope_result_dict[dg]['env']['SLS'] = {}                
    Envelope_result_dict[dg]['env']['ALS'] = {}                
    Envelope_result_dict[dg]['env']['ULS'] = {}                
    
    envelope_ULS_ALS('ULS')    
    envelope_ULS_ALS('ALS')    
    envelope_SLS('SLS')
    
    for thermal_case,thermal_dict in dg_dict.items():    
        if thermal_case == 'req_link':
            continue
        for case, case_dict in thermal_dict.items(): 
            limit_state = case_dict["limit_state"]
            if limit_state == 'ULS' or limit_state == 'ALS':
                check_envelope("Min conc. stress bot face :",limit_state,"smaller")
                check_envelope("Min conc. stress top face :",limit_state,"smaller")
                check_envelope("Min conc. strain bot face :",limit_state,"smaller")
                check_envelope("Min conc. strain top face :",limit_state,"smaller")
                check_envelope("Max steel stress layer  1 :",limit_state,"abs")
                check_envelope("Max steel stress layer  2 :",limit_state,"abs")
                check_envelope("Max steel stress layer  3 :",limit_state,"abs")
                check_envelope("Max steel stress layer  4 :",limit_state,"abs")
                check_envelope("Min steel stress layer  1 :",limit_state,"abs")
                check_envelope("Min steel stress layer  2 :",limit_state,"abs")
                check_envelope("Min steel stress layer  3 :",limit_state,"abs")
                check_envelope("Min steel stress layer  4 :",limit_state,"abs")

                
                check_envelope_try("Max steel stress layer  5 :",limit_state,"abs")
                check_envelope_try("Max steel stress layer  6 :",limit_state,"abs")
                check_envelope_try("Max steel stress layer  7 :",limit_state,"abs")
                check_envelope_try("Max steel stress layer  8 :",limit_state,"abs")
                check_envelope_try("Min steel stress layer  5 :",limit_state,"abs")
                check_envelope_try("Min steel stress layer  6 :",limit_state,"abs")
                check_envelope_try("Min steel stress layer  7 :",limit_state,"abs")
                check_envelope_try("Min steel stress layer  8 :",limit_state,"abs")
                check_envelope_try("Max prestress adstress  1 :",limit_state,"larger")
                check_envelope_try("Max prestress adstress  2 :",limit_state,"larger")
                
                check_envelope("Steel utilisation layer  1:",limit_state,"larger")
                check_envelope("Steel utilisation layer  2:",limit_state,"larger")
                check_envelope("Steel utilisation layer  3:",limit_state,"larger")
                check_envelope("Steel utilisation layer  4:",limit_state,"larger")
                
                check_envelope_try("Steel utilisation layer  5:",limit_state,"larger")
                check_envelope_try("Steel utilisation layer  6:",limit_state,"larger")
                check_envelope_try("Steel utilisation layer  7:",limit_state,"larger")
                check_envelope_try("Steel utilisation layer  8:",limit_state,"larger")            
                
            if limit_state == 'SLS':
                check_envelope("Min conc. stress bot face :",limit_state,"smaller")
                check_envelope("Min conc. stress top face :",limit_state,"smaller")
                check_envelope("Max steel stress layer  1 :",limit_state,"abs")
                check_envelope("Max steel stress layer  2 :",limit_state,"abs")
                check_envelope("Max steel stress layer  3 :",limit_state,"abs")
                check_envelope("Max steel stress layer  4 :",limit_state,"abs")
                check_envelope("Min steel stress layer  1 :",limit_state,"abs")
                check_envelope("Min steel stress layer  2 :",limit_state,"abs")
                check_envelope("Min steel stress layer  3 :",limit_state,"abs")
                check_envelope("Min steel stress layer  4 :",limit_state,"abs")
                              
                check_envelope_try("Max steel stress layer  5 :",limit_state,"abs")
                check_envelope_try("Max steel stress layer  6 :",limit_state,"abs")
                check_envelope_try("Max steel stress layer  7 :",limit_state,"abs")
                check_envelope_try("Max steel stress layer  8 :",limit_state,"abs")
                check_envelope_try("Min steel stress layer  5 :",limit_state,"abs")
                check_envelope_try("Min steel stress layer  6 :",limit_state,"abs")
                check_envelope_try("Min steel stress layer  7 :",limit_state,"abs")
                check_envelope_try("Min steel stress layer  8 :",limit_state,"abs")
                check_envelope_try("Max prestress adstress  1 :",limit_state,"larger")
                check_envelope_try("Max prestress adstress  2 :",limit_state,"larger")
                
                check_envelope_try_wk(dg,case,"wk_static",limit_state,"larger")
                check_envelope_try_wk(dg,case,"wk_thermal",limit_state,"larger")
                
                check_envelope_try_c(dg,case_dict,case,"Min compression depth (mm):",limit_state,"smaller")
                check_envelope_try_c(dg,case_dict,case,"Min Therm comp  depth (mm):_thermal",limit_state,"smaller")
                
                
#####################                

 #--- Summary .xlsx files
file_root='result_tables_m_'
date=str(datetime.date.today())
filename=file_root+date+'_v6.8_41d'+'.xlsx'             
wb = xlsxwriter.Workbook(filename)                
    
summary_blue_table_uls_als_envelope(Envelope_result_dict,wb,'ULS')          
summary_blue_table_uls_als_envelope(Envelope_result_dict,wb,'ALS')          
summary_blue_table_sls_envelope(Envelope_result_dict,wb,'SLS')          

summary_blue_table_uls_als(All_result_dict,wb,'ULS')          
summary_blue_table_uls_als(All_result_dict,wb,'ALS')      
summary_blue_table_sls(All_result_dict,wb)

wb.close()  
#print('☪ ')
print('Your results are ready in '+directory)
