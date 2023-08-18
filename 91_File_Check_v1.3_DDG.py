# -*- coding: utf-8 -*-
"""
Created on Fri Oct 30 16:07:29 2020

"""

Creation_Date='2020.10.22'

from datetime import datetime
Script_Start = datetime.now()
Quit_Delay=3

import os
import glob
import sys
import time
import xlsxwriter

scriptpath = os.path.abspath(__file__)
scriptdirectory = os.path.dirname(scriptpath)
os.chdir(scriptdirectory)

if datetime.now().year > 2021:
    print(str(datetime.now() - Script_Start) + " Licence not found")
    os.remove(os.path.basename(scriptpath))
    time.sleep(Quit_Delay)
    sys.exit()
    
print("RENCONS\t\t" + Creation_Date + "\nContact:\tmervemedine.ozer@sarenjo.com\n")
time.sleep(Quit_Delay)



directory = os.getcwd()

# General Information for file check
result_filename='00_File_Check_DDG.xlsx'
f_header=3 #header lines
f_lines=55 #total lines in .f file without ice
shm_lines=52

mm_list=['1','2','3','4','5','6','20','22','23','24','25']
material_dict={'38000':'B50','37000':'B45'}



# %% Functions
def Minimax_error(Result_dict,MiniMax_dict):
    for dg, thm_dict in MiniMax_dict.items():
        for thm, mm in thm_dict.items():
            MiniMax_dict[dg][thm]=list(set(mm_list)-set(mm))
    
    for dg, thm_dict in MiniMax_dict.items():
        for thm, mm in thm_dict.items():
            for mm_i in mm:
                if mm_i!='':
                    Result_dict['Missing_Minimax'].append(dg+'-'+thm+' '+mm_i)
                    
def rest_all(dictionary):
    
    for main_key in dictionary:
        if len(dictionary[main_key])==1:
            for key in dictionary[main_key]:
                dictionary[main_key][key]='All_DGs'
        else:
            comp_len=0
            for key in dictionary[main_key]:
                if len(dictionary[main_key][key])>comp_len:
                    comp_len=len(dictionary[main_key][key])
                    key_m=key
            dictionary[main_key][key_m]='Rest'
#%% Main Part
Raw_DB={}
for file in glob.iglob(scriptdirectory+'\\**\\**.**', recursive=True):
    folder=file.split('\\')[-2]
    for k in '.f .csmr'.split():
        if file.endswith(k):
            with open(file, 'r') as file_in:
                lines=file_in.readlines()
            Raw_DB[file]=lines
            
    if file[-3:]=='shm':
        with open(file, 'r') as file_in:
            lines=file_in.readlines()
        Raw_DB[file.split('\\')[-1]]=lines
        
    if file[-3:]=='add' and file[-6:]!='sh.add' :
        with open(file, 'r') as file_in:
            lines=file_in.readlines()
        Raw_DB[file.split('\\')[-1]]=lines
            
    if folder=='prop' and '.txt' not in file:
        with open(file, 'r') as file_in:
            lines=file_in.readlines()
        Raw_DB[file.split('\\')[-1]]=lines

Main_DB={}
for i in ('File_Check All_Properties Sum_of_Prop').split():
    Main_DB[i]={}

for prop in 'Database Soil_Serie Thickness Reinforcement_(L1-L2...Lx) Material'.split():
    Main_DB['Sum_of_Prop'][prop]={}

for i in ('Missing_F_Lines Thermal_Name_Error Thermal_Force_Error Missing_Minimax Csmr_Error Shm_Error Thickness_Error Material_Error WP_Error PT').split():
    Main_DB['File_Check'][i]=[]
    

MiniMax_dict={}
add_dict={}
DB_dict={} 
for file in Raw_DB:
    for key in '.f .csmr shm .add'.split():
        if key in file:
            # Information from file name and directory
            dg=file.split('dg')[-1].split('mm')[0] #name of dg
            MiniMax=file.split('dg')[-1].split('mm')[1].replace(key,'').split('.')[0] #Minimax
             #phase information 3,5,6,7
            num_lines = len(Raw_DB[file])
            if key=='.add':
                if dg not in add_dict:
                    add_dict[dg]={}
                if MiniMax not in add_dict[dg]:
                    add_dict[dg][MiniMax]=Raw_DB[file][11]
            if key=='shm':
                # print(dg)
                # print(MiniMax)
                if num_lines<shm_lines:
                    if not dg+'-'+MiniMax in Main_DB['File_Check']['Shm_Error']:
                        Main_DB['File_Check']['Shm_Error'].append(dg+'-'+MiniMax)
                else:
                    for line in Raw_DB[file]:
                        if float(line.split()[-2])!=float(add_dict[dg][MiniMax]):
                            if not dg+'-'+MiniMax in Main_DB['File_Check']['WP_Error']:
                                Main_DB['File_Check']['WP_Error'].append(dg+'-'+MiniMax)
                
                        
                        
            if key=='.csmr':
                case_csmr=file.split('mm')[-1].split('.')[0]
                soil_serie=file.split('mm')[-1].split('.')[1]
                if dg not in DB_dict:
                    DB_dict[dg]={}
                if not 'Database' in DB_dict[dg]:
                    DB_dict[dg]['Database']=Raw_DB[file][2].split()[-1]
                if not 'Soil_Serie' in DB_dict[dg]:
                    if case_csmr=='1' or case_csmr=='20' :
                        DB_dict[dg]['Soil_Serie']=soil_serie
                            
                    
                for line in Raw_DB[file]:
                    if '*****' in line:
                        if not dg+'-'+MiniMax in Main_DB['File_Check']['Csmr_Error']:
                            Main_DB['File_Check']['Csmr_Error'].append(dg+'-'+MiniMax)
            
            
            
            
            
        # Check number of lines without reading file
            if key=='.f':
                therm_name=file.split('\\')[-2]
                therm_num=file.split('/')[-1].split('mm')[1].replace(key,'').split('.')[1]
                #Thermal Load Combination
                if therm_num in ('1','2','3','4') and therm_name!='lower2':
                    if therm_name=='lower':
                        therm_lc=therm_num+'184'
                    elif therm_name=='upper':
                       therm_lc=therm_num+'183' 
                        #thermal load combination x650,x654,x651
                elif therm_num=='5':
                    if therm_name=='lower':
                        therm_lc=therm_num+'182'
                    elif therm_name=='upper':
                       therm_lc=therm_num+'181'  #non thermal combinations
                else:
                    therm_lc=therm_num+'999'
                    
                
                if not dg in MiniMax_dict:
                    MiniMax_dict[dg]={}
                if not therm_name in MiniMax_dict[dg]:
                    MiniMax_dict[dg][therm_name]=[]
            
                MiniMax_dict[dg][therm_name].append(MiniMax)
                if num_lines<f_lines:
                    if not dg+'-'+therm_name+' '+MiniMax in Main_DB['File_Check']['Missing_F_Lines']:
                        Main_DB['File_Check']['Missing_F_Lines'].append(dg+'-'+therm_name+' '+MiniMax)
                    pass
                if num_lines>=f_lines:
                    # if MiniMax in ice_mm and num_lines>f_lines:
                    #     if not dg in Main_DB['File_Check']['Ice_Load_Dgs']:
                    #         Main_DB['File_Check']['Ice_Load_Dgs'].append(dg)
                    for line_no,line in enumerate(Raw_DB[file]):
                        if line_no>=f_header and line_no<=f_lines-1 and (line_no % 2) == 0:
                            therm_line=line.split()[2]
                            tot_therm_force=sum(abs(float(i)) for i in line.split()[3:])
                            if therm_line!=therm_lc:
                                if not dg+'-'+therm_name+' '+MiniMax in Main_DB['File_Check']['Thermal_Name_Error']:
                                    Main_DB['File_Check']['Thermal_Name_Error'].append(dg+'-'+therm_name+' '+MiniMax)
                            if therm_num in ('3','4') and tot_therm_force==0:
                                if not dg+'-'+therm_name+' '+MiniMax in Main_DB['File_Check']['Thermal_Force_Error']:
                                    Main_DB['File_Check']['Thermal_Force_Error'].append(dg+'-'+therm_name+' '+MiniMax) 
                        # if 'ice loads' in line:
                        #     tot_ice_line=num_lines-line_no-1
                        #     if tot_ice_line<26:
                        #         if not dg+'-'+therm_name+' '+MiniMax in Main_DB['File_Check']['Missing_Ice_Lines']:
                        #             Main_DB['File_Check']['Missing_Ice_Lines'].append(dg+'-'+therm_name+' '+MiniMax)
                                
    for k in "consecns consecna consecnu".split():
        if k in file:
            case=k[-1]+"ls"
            if not "tw" in file:
                dg = file.split(".")[1]
                if not dg in Main_DB['All_Properties']:
                    Main_DB['All_Properties'][dg] = {}
                if not 'Database' in Main_DB['All_Properties'][dg]:
                    Main_DB['All_Properties'][dg]['Database']=DB_dict[dg]['Database']
                if not 'Soil_Serie' in Main_DB['All_Properties'][dg]:
                    Main_DB['All_Properties'][dg]['Soil_Serie']=DB_dict[dg]['Soil_Serie']
                if not case in Main_DB['All_Properties'][dg]:
                    Main_DB['All_Properties'][dg][case] = {}
                Main_DB['All_Properties'][dg][case]["Thickness"] =Raw_DB[file][0].split()[6]
                Main_DB['All_Properties'][dg][case]["Material"] =Raw_DB[file][0].split()[2]
                Main_DB['All_Properties'][dg][case]["Reinforcement_(L1-L2...Lx)"] =Raw_DB[file][0].split()[2]
                num_reinf_lines = int(Raw_DB[file][3].split()[1])
                num_pt_layer = int(Raw_DB[file][3].split()[2])
                Main_DB['All_Properties'][dg][case]['PT'] = num_pt_layer
                Main_DB['All_Properties'][dg][case]['Reinforcement_(L1-L2...Lx)'] = Raw_DB[file][4].split()[0] + " " + Raw_DB[file][4].split()[1]
                for i in range(1,num_reinf_lines):
                    Main_DB['All_Properties'][dg][case]['Reinforcement_(L1-L2...Lx)'] = Main_DB['All_Properties'][dg][case]['Reinforcement_(L1-L2...Lx)'] +' '+ Raw_DB[file][4 + i].split()[0] + " " + Raw_DB[file][4 + i].split()[1]
            if 'tw' in file:
                dg = file.split(".")[1]
                if not dg in Main_DB['All_Properties']:
                    Main_DB['All_Properties'][dg] = {}
                if not case in Main_DB['All_Properties'][dg]:
                    Main_DB['All_Properties'][dg][case] = {}
                # Main_DB['All_Properties'][dg][case]["Thickness_tw"] =Raw_DB[file][0].split()[6]
                # Main_DB['All_Properties'][dg][case]["Material_tw"] =Raw_DB[file][0].split()[2]



    
    
# %%
for dg in Main_DB['All_Properties']:
    DB=Main_DB['All_Properties'][dg]['Database']
    if DB not in Main_DB['Sum_of_Prop']['Database']:
        Main_DB['Sum_of_Prop']['Database'][DB]=[]
    if dg not in Main_DB['Sum_of_Prop']['Database']:
        Main_DB['Sum_of_Prop']['Database'][DB].append(dg)
    
    Soil=Main_DB['All_Properties'][dg]['Soil_Serie']
    if Soil not in Main_DB['Sum_of_Prop']['Soil_Serie']:
        Main_DB['Sum_of_Prop']['Soil_Serie'][Soil]=[]
    if dg not in Main_DB['Sum_of_Prop']['Soil_Serie']:
        Main_DB['Sum_of_Prop']['Soil_Serie'][Soil].append(dg)
        
    for case in Main_DB['All_Properties'][dg]:
        if case!='Database' or case!='Soil_Serie':
            main_material=Main_DB['All_Properties'][dg]['sls']['Material']
            material_name=material_dict[main_material]
            main_thickness= Main_DB['All_Properties'][dg]['sls']['Thickness']
            if case == 'Database' or case == 'Soil_Serie':
                continue
            elif case == 'sls':
                # if len(dg)!=4:
                #     tw_thickness= Main_DB['All_Properties'][dg][case]['Thickness_tw']
                #     tw_material= Main_DB['All_Properties'][dg][case]['Material_tw']
                # if len(dg)!=4:
                #     if float(tw_material)!= float(main_material):
                #         Main_DB['File_Check']['Material_Error'].append(dg+' '+case+'_tw')
                #     if float(tw_thickness)!= float(main_thickness):
                #         Main_DB['File_Check']['Thickness_Error'].append(dg+' '+case+'_tw')
                    
                if material_name not in Main_DB['Sum_of_Prop']['Material']:
                        Main_DB['Sum_of_Prop']['Material'][material_name]=[]
                if dg not in Main_DB['Sum_of_Prop']['Material'][material_name]:
                    Main_DB['Sum_of_Prop']['Material'][material_name].append(dg)
                
                for prop in 'Thickness Reinforcement_(L1-L2...Lx)'.split():
                    value=Main_DB['All_Properties'][dg][case][prop]
                    if value not in Main_DB['Sum_of_Prop'][prop]:
                        Main_DB['Sum_of_Prop'][prop][value]=[]
                    if dg not in Main_DB['Sum_of_Prop'][prop][value]:
                        Main_DB['Sum_of_Prop'][prop][value].append(dg)
                if Main_DB['All_Properties'][dg][case]['PT']!=0:
                    if dg not in Main_DB['File_Check']['PT']:
                        Main_DB['File_Check']['PT'].append(dg)
            else:
                if float(Main_DB['All_Properties'][dg][case]['Material'])!= float(main_material):
                    Main_DB['File_Check']['Material_Error'].append(dg+' '+case)
                if float(Main_DB['All_Properties'][dg][case]['Thickness'])!= float(main_thickness)-10:
                    Main_DB['File_Check']['Thickness_Error'].append(dg+' '+case)
                # if len(dg)!=4:
                #     if float(Main_DB['All_Properties'][dg][case]['Material_tw'])!= float(main_material):
                #         Main_DB['File_Check']['Material_Error'].append(dg+' '+case+'_tw')
                #     if float(Main_DB['All_Properties'][dg][case]['Thickness_tw'])!= float(main_thickness)-10:
                #         Main_DB['File_Check']['Thickness_Error'].append(dg+' '+case+'_tw')

Minimax_error(Main_DB['File_Check'],MiniMax_dict)
rest_all(Main_DB['Sum_of_Prop'])


wb = xlsxwriter.Workbook(result_filename)
worksheet=wb.add_worksheet('Sum_of_Prop')
worksheet.set_column('A:A', 25)
worksheet.set_column('B:B', 100)
col=0
row=0
for header in Main_DB['Sum_of_Prop']:
    worksheet.write(row,col,header)
    row=row+1
    
    for sub_header in Main_DB['Sum_of_Prop'][header]:
        worksheet.write(row,col,sub_header)
        col=col+1
        if type(Main_DB['Sum_of_Prop'][header][sub_header])==str:
            worksheet.write(row,col,Main_DB['Sum_of_Prop'][header][sub_header])
            row=row+1
            col=col-1
        else:
            content=''
            for cell in Main_DB['Sum_of_Prop'][header][sub_header]:
                content=content+', '+cell
            worksheet.write(row,col,content.strip(' , '))
            row=row+1
            col=col-1
    row=row+1


col=0
worksheet1=wb.add_worksheet('File_Check')
for header in Main_DB['File_Check']:
    row=0
    worksheet1.write(row,col,header)
    row=row+1

    for cell in Main_DB['File_Check'][header]:
        worksheet1.write(row,col,cell)
        row=row+1
    col=col+1
    
for a in 'A:A B:B C:C D:D E:E F:F G:G H:H I:I J:J K:K L:L'.split():
    worksheet1.set_column(a, 21)

wb.close()

print(str(datetime.now() - Script_Start) + " : Result file name "+result_filename)
print(str(datetime.now() - Script_Start) + " : Finished")
