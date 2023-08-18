# -*- coding: utf-8 -*-
"""
Created on Thu Jan  9 11:57:30 2020

@author: Antl
"""

import os 
import xlsxwriter
import time


req_link=r".\rcz_req_61.txt"
#%%Main Functions
def list_csmr_files(directory,cond,csmr_files,case_list):
# List csmr file in current directory
    for root, dirs, files in os.walk(directory):
        for dg in dg_list:
            for file in files:
                if dg in file and file.endswith('.csmr'):
                    case_fs=file.split('mm')[1].replace('.res.csmr','')
                    case_fs=case_fs.split('.')[0]
                    for case in rcz_req_dict[dg]: 
                        if case_fs==case and rcz_req_dict[dg][case][cond]!='n/a':
                            case_list.append(case_fs)
                            csmr = root +str('/') +file
                            csmr_files.append(csmr)
    csmr_files=list(dict.fromkeys(csmr_files))
    case_list=list(dict.fromkeys(case_list))

    return csmr_files,case_list

def list_res_files(directory,res_files):
# List csmr file in current directory
    for root, dirs, files in os.walk(directory):
        for dg in dg_list:
            for file in files:
                if dg in file and file.endswith('.res'):
                    case_fs=file.split('mm')[-1].replace('.res','')
                    case_fs=case_fs.split('.')[0]
                    for case in rcz_req_dict[dg]: 
                        if case_fs=='2':
                            res = root + str("/") + file
                            res_files.append(res)
    res_files=list(dict.fromkeys(res_files)) 
                           
    return res_files


def read_csmr(csmr_files,cond,dictionary):
    for csmr_file in csmr_files:
        dg=csmr_file.split('mm')[0][-5:]
        case=csmr_file.split('mm')[-1][0:1]
        a=1
        for req_dg, req_dg_dict in rcz_req_dict.items():
                     for case_req, case_req_dict in req_dg_dict.items():
                         if dg==req_dg and case_req==case:
                             req= rcz_req_dict[dg][case][cond]
                             t_dg= rcz_req_dict[dg]['t']
        if a==1:
            try:
                test = dictionary[case]
            except:
                dictionary[case] = {}
        with open(csmr_file) as file:
            for line_no,line in enumerate(file):
                if "DG NUMBER" in line:
                    lc_thermal=line.split()[-1]
                
                if cond=="C":
                    if "Min compression depth (mm):" in line:
                        value = (line[34:].split()[0])
                        value = value[:len(value)]
                        member=line[34:].split()[1]
                        node=line[34:].split()[2]
                        LC=line.split()[-1]
                    elif "Min Therm comp  depth (mm):" in line:
                        value_th = (line[34:].split()[0])
                        value_th = value_th[:len(value_th)]
                        member_th=line[34:].split()[1]
                        node_th=line[34:].split()[2]
                        LC_th=line.split()[-1]
                if cond=='csigp':
                    if "Max princ. membrane force :" in line:
                        value = (line[34:].split()[0])
                        value = value[:len(value)]
                        member=line[34:].split()[1]
                        node=line[34:].split()[2]
                        LC=line.split()[-1]
                    elif "Max pr. Therm. memb force :" in line:
                        value_th = (line[34:].split()[0])
                        value_th = value_th[:len(value_th)]
                        member_th=line[34:].split()[1]
                        node_th=line[34:].split()[2]
                        LC_th=line.split()[-1]
                        
            if  cond=='C':
                if (float(value)<=float(req) or float(value_th)<=float(req)):
                    try:
                        dictionary[case][dg]
                    except:
                        dictionary[case][dg]= {}
                    dictionary[case][dg][lc_thermal]={}    
                    if float(value)<=float(value_th):
                        
                        dictionary[case][dg][lc_thermal]['value']=value
                        dictionary[case][dg][lc_thermal]['value_th']=value_th
                        dictionary[case][dg][lc_thermal]['member']=member
                        dictionary[case][dg][lc_thermal]['node']=node
                        dictionary[case][dg][lc_thermal]['LC']=LC
                        dictionary[case][dg][lc_thermal]['req']=req
                        dictionary[case][dg][lc_thermal]['t']=t_dg
                    else:
                        
                        dictionary[case][dg][lc_thermal]['value']=value
                        dictionary[case][dg][lc_thermal]['value_th']=value_th
                        dictionary[case][dg][lc_thermal]['member']=member_th
                        dictionary[case][dg][lc_thermal]['node']=node_th
                        dictionary[case][dg][lc_thermal]['LC']=LC_th
                        dictionary[case][dg][lc_thermal]['req']=req
                        dictionary[case][dg][lc_thermal]['t']=t_dg
            
            if cond=='csigp':
                if (float(value)/float(t_dg)<=float(req) or float(value_th)/float(t_dg)<=float(req)):
                    try:
                        dictionary[case][dg]
                    except:
                        dictionary[case][dg]= {}
                    dictionary[case][dg][lc_thermal]={}    
                    if float(value)<=float(value_th):
                        
                        dictionary[case][dg][lc_thermal]['value']=float(value)/float(t_dg)
                        dictionary[case][dg][lc_thermal]['value_th']=float(value_th)/float(t_dg)
                        dictionary[case][dg][lc_thermal]['member']=member
                        dictionary[case][dg][lc_thermal]['node']=node
                        dictionary[case][dg][lc_thermal]['LC']=LC
                        dictionary[case][dg][lc_thermal]['req']=req
                        dictionary[case][dg][lc_thermal]['t']=t_dg
                    else:
                        
                        dictionary[case][dg][lc_thermal]['value']=float(value)/float(t_dg)
                        dictionary[case][dg][lc_thermal]['value_th']=float(value_th)/float(t_dg)
                        dictionary[case][dg][lc_thermal]['member']=member_th
                        dictionary[case][dg][lc_thermal]['node']=node_th
                        dictionary[case][dg][lc_thermal]['LC']=LC_th
                        dictionary[case][dg][lc_thermal]['req']=req
                        dictionary[case][dg][lc_thermal]['t']=t_dg
                            
    return dictionary

def envelope(env_dictionary,dictionary,txt_cond,cond):
    for case, case_dict in dictionary.items():
        env_dictionary[case]={}
        for dg, dg_dict in case_dict.items():
            env_dictionary[case][dg]={}
            if cond=="C":
                env_dictionary[case][dg]['value']="1500"
                env_dictionary[case][dg]['value_th']="1500"
            elif cond=="csigp":
                env_dictionary[case][dg]['value']="-1500"
                env_dictionary[case][dg]['value_th']="-1500"
            for term, term_dict in dg_dict.items():
                for value in ['value','value_th']:
                    if cond=='C' :
                        if float(term_dict[value])<float(env_dictionary[case][dg][value]):
                            env_dictionary[case][dg][value]=term_dict[value]
                            env_dictionary[case][dg]['member']=term_dict['member']
                            env_dictionary[case][dg]['node']=term_dict['node']
                            env_dictionary[case][dg]['LC']=term_dict['LC']
                            env_dictionary[case][dg]['req']=term_dict['req']
                            env_dictionary[case][dg]['t']=term_dict['t']
                    elif cond=='csigp':
                        if float(term_dict[value])>float(env_dictionary[case][dg][value]):
                            env_dictionary[case][dg][value]=term_dict[value]
                            env_dictionary[case][dg]['member']=term_dict['member']
                            env_dictionary[case][dg]['node']=term_dict['node']
                            env_dictionary[case][dg]['LC']=term_dict['LC']
                            env_dictionary[case][dg]['req']=term_dict['req']
                            env_dictionary[case][dg]['t']=term_dict['t']
    return env_dictionary
                           
#%% Excel Writing
def set_table(counter_rows,worksheet,workbook,cond):
    cell_format = workbook.add_format({'bold': True})
    worksheet.write('B2', 'Case', cell_format)
    worksheet.write('C2', 'DG', cell_format)
    worksheet.write('D2', 'Member', cell_format)
    worksheet.write('E2', 'Node', cell_format)
    worksheet.write('F2', 'LC', cell_format)
    worksheet.write('J2', 't (mm)', cell_format)
    if cond=="C":
        worksheet.write('G2', 'Cu (mm)', cell_format)
        worksheet.write('H2', 'Cu+T (mm)', cell_format)
        worksheet.write('I2', 'Req RCZ [mm]', cell_format)
    elif cond=="csigp":
        worksheet.write('G2', 'Memb. Stress (MPa)', cell_format)
        worksheet.write('H2', 'Memb. Stress+T (MPa)', cell_format)
        worksheet.write('I2', 'Req. Membrane Stress [MPa]', cell_format)
        
    worksheet.set_column('B:B', 12)
    worksheet.set_column('C:C', 12)
    worksheet.set_column('D:D', 12)
    worksheet.set_column('E:E', 12)
    worksheet.set_column('F:F', 17)
    worksheet.set_column('G:G', 12)
    worksheet.set_column('H:H', 12)
    worksheet.set_column('I:I', 17)
    worksheet.set_column('J:J', 12)
def write_table(env_dict,case_list,cond):
    
    if cond=='C':
        wb=wb_c
    elif cond=='csigp':
        wb=wb_cs
        

    for case_w in case_list:
        row=2
        worksheet=wb.add_worksheet("mm"+case_w)
        for case, case_dict in env_dict.items():
            if case_w==case:
                for dg, dg_dict in case_dict.items():
                    worksheet.write(row,1,float(case))
                    worksheet.write(row,2,float(dg))
                    worksheet.write(row,3,dg_dict["member"])
                    worksheet.write(row,4,dg_dict['node'])
                    worksheet.write(row,5,dg_dict['LC'])
                    worksheet.write(row,6,float(dg_dict['value']))
                    worksheet.write(row,7,float(dg_dict['value_th']))
                    worksheet.write(row,8,float(dg_dict['req']))
                    worksheet.write(row,9,float(dg_dict['t']))
                    row=row+1
        set_table(row,worksheet,wb,cond)

                    
#%%Main Part
directory=os.getcwd()
dg_list=[]
dgm=[]
result_file_name_c='tightness_criteria_c.xlsx'
result_file_name_cs='tightness_criteria_csigp.xlsx'                  
wb_c=xlsxwriter.Workbook(result_file_name_c)
wb_cs=xlsxwriter.Workbook(result_file_name_cs)
for root, dirs, files in os.walk(directory):
    for filename in dirs: #reads all folder names in directory
        if 'dg' in filename:  #reads all dg folder names in directory
            dgf=filename.split('dg')[1]  #reads all dg names in directory
            if dgf!='':
                dgm.append(dgf)
                dg_list=dgm #creates dg list

list(set(dg_list)) #removes duplicated dg names

#res_files=[]
#res_files= list_res_files(directory,res_files)

            

rcz_req_dict={}

for dg in dg_list:
    rcz_req_dict[dg]={}
    with open(req_link) as req_file:
        file=req_file.readlines()
        for line in file:
            if dg in line:
                case=line.split()[1]
                rcz_req_dict[dg][case]={}
                C=line.split()[2]
                csigp=line.split()[-1]
                rcz_req_dict[dg][case]["C"]=C
                rcz_req_dict[dg][case]["csigp"]=csigp
    
 
res_files=[]
res_files= list_res_files(directory,res_files)

for res_f in res_files:
    dgp=res_f.split('dg')[-1][:5]
    for dg, dgd in rcz_req_dict.items():
        if dgp==dg:
            with open(res_f) as file:
                for line_no,line in enumerate(file):
                    if line_no==30:
                        rcz_req_dict[dg]['t']=line.split(' ')[-1]       

        

for dg in dg_list:
    if rcz_req_dict.get(dg)==None:
        dg_list.remove(dg)


csmr_files_c=[]
case_list_c=[]
csmr_files_c, case_list_c = list_csmr_files(directory,"C",csmr_files_c, case_list_c)
csmr_files_cs=[]
case_list_cs=[]
csmr_files_cs, case_list_cs = list_csmr_files(directory,"csigp",csmr_files_cs, case_list_cs)

#t_dict_cs={}
#t_dict_cs=thickness(t_dict_cs,res_files_cs)

C_dict={}
cond="C"
C_dict=read_csmr(csmr_files_c,cond,C_dict)

Cs_dict={}
cond="csigp"
Cs_dict=read_csmr(csmr_files_cs,cond,Cs_dict)

txt_cond=['C','C_th','csigp','csigp_th']

C_dict_env={}
C_dict_env=envelope(C_dict_env,C_dict,txt_cond,"C")

Cs_dict_env={}
Cs_dict_env=envelope(Cs_dict_env,Cs_dict,txt_cond,"csigp")
cs_check=bool(Cs_dict_env)

write_table(C_dict_env,case_list_c,'C')

if cs_check!='False':
    write_table(Cs_dict_env,case_list_cs,'csigp')




wb_c.close()
wb_cs.close()
















