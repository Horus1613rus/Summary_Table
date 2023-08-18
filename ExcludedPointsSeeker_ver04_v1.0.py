#
# Alex Dalle
# 12/02/2019
# Antl
# 21/02/2019
# Function : Read all .csmr files in the directory 'repR' (recursively) and
#            return filename if it contains excluded integration point. 
#
#Note that "tab" can not be used while making the script, only spaces work.

import os
import glob
import numpy
import pandas

repR = os.getcwd()
#repR = "\\saipemnet.saipem.intranet\saf\ALNG2-GBS-DB\11_ENGINEERING\11.02_DISCP\01_CIVIL\02-Civil\95-Users\GBS_Consect_results"
listCSMR = []

######print all excluded points to a table and all failure points to a table
rejected_points_sho_lon_list=[]
rejected_points_up_low_list=[]
dg_rejected_points_list=[]
rejected_points_LC_list=[]
rejected_points_EL_list=[]
rejected_points_NODE_list=[]
rejected_points_PHASE_list=[]

failure_points_sho_lon_list=[]
failure_points_up_low_list=[]
dg_failure_points_list=[]
failure_points_LC_list=[]
failure_points_EL_list=[]
failure_points_NODE_list=[]
failure_points_PHASE_list=[]
################################

excluded_points=False
failures=False
print_failures_only=True
add_failure_to_list=False


impr = '\n\nList of .csmr files containing excluded points :\n'

for f in glob.iglob(repR + '/**/*.csmr', recursive=True) : #loops through every file csmr-file in the folder and subfolders

    listCSMR.append(f)
  
    with open(f, 'r') as fIn :

        content = fIn.read()

        
        
	
    if 'POINTS EXCLUDED FROM SUMMARY' in content or 'FAILURES:' in content:

        impr = impr + f.replace(repR, '') + '\n' #adds dg with rejected points to the resulting text file
        File_Address=f.replace(repR, '')
        fIn=open(f,"r")
        lines=fIn.readlines()

        add_rejected_to_list=True 		
        
        for line in lines:
            if  'POINTS EXCLUDED FROM SUMMARY'  in line :
                excluded_points=True
                header=True
                print_failures_only=False
            if line == "\n":
                pass
            elif excluded_points==True:
                #print(line, end="")
                impr = impr + line          #adds a line to the resulting text file
                
                if  'POINTS EXCLUDED FROM SUMMARY'  in line:
                    pass
                elif 'FAILURES' in line:
                    add_rejected_to_list=False

                elif add_rejected_to_list:
                    
                    rejected_points_sho_lon_list.append(File_Address.split('mm')[1].replace('.res.csmr',''))
                    rejected_points_up_low_list.append(File_Address.split('mm')[0].strip()[len(File_Address.split('mm')[0])-13:len(File_Address.split('mm')[0])-8])
                    dg_rejected_points_list.append(File_Address.split('mm')[0].strip()[len(File_Address.split('mm')[0])-5:len(File_Address.split('mm')[0])])
                    rejected_points_EL_list.append(line[0:17].strip())
                    rejected_points_NODE_list.append(line[18:24].strip())
                    rejected_points_LC_list.append(line[24:42].strip())               
                
                if 'FAILURES' in line:
                    add_failure_to_list=True
                elif add_failure_to_list:
                    
                    failure_points_sho_lon_list.append(File_Address.split('mm')[1].replace('.res.csmr',''))
                    dg_failure_points_list.append(File_Address.split('mm')[0].strip()[len(File_Address.split('mm')[0])-5:len(File_Address.split('mm')[0])])
                    failure_points_up_low_list.append(File_Address.split('mm')[0].strip()[len(File_Address.split('mm')[0])-13:len(File_Address.split('mm')[0])-8])
                    failure_points_EL_list.append(line[0:17].strip())
                    failure_points_NODE_list.append(line[18:24].strip())
                    failure_points_LC_list.append(line[24:42].strip())                   
                
                if header==True:
                    impr = impr + " ELEMENT:     NODE:   LOADCASE:" + "\n" #adding a header for each DG
                    header=False
            elif failures==True:
                impr = impr + line          #adds a line to the resulting text file
                
                failure_points_sho_lon_list.append(File_Address.split('mm')[1].replace('.res.csmr',''))
                dg_failure_points_list.append(File_Address.split('mm')[0].strip()[len(File_Address.split('mm')[0])-5:len(File_Address.split('mm')[0])])
                failure_points_up_low_list.append(File_Address.split('mm')[0].strip()[len(File_Address.split('mm')[0])-13:len(File_Address.split('mm')[0])-8])
                failure_points_EL_list.append(line[0:17].strip())
                failure_points_NODE_list.append(line[18:24].strip())
                failure_points_LC_list.append(line[24:42].strip())
                

            if 'FAILURES' in line and print_failures_only==True:

                failures=True
                impr = impr + line          #adds a line to the resulting text file
                impr = impr + " ELEMENT:     NODE:   LOADCASE:" + "\n" #adding a header for each DG
                
                #dg_failure_points_list.append(dg_temp[3:7])
                #failure_points_EL_list.append(line[0:17].strip())
                #failure_points_NODE_list.append(line[18:24].strip())
                #failure_points_LC_list.append(line[25:42].strip())
                                

        fIn.close()
    excluded_points=False
    failures=False
    print_failures_only=True
    add_failure_to_list=False



impr = impr + '\nNumber of .csmr files analysed : ' + str(len(listCSMR)) + '\n'

#print(impr) #prints in command window
with open('0_ExcludedPoints.txt', 'w') as fOut : #writes the impr to a text file
        fOut.write(impr)

##### create tables
        
table_rejected_points=numpy.array([dg_rejected_points_list,rejected_points_sho_lon_list, rejected_points_up_low_list,rejected_points_EL_list, rejected_points_NODE_list, rejected_points_LC_list ])
numpy.column_stack(table_rejected_points)

df_rejected_points=pandas.DataFrame(numpy.column_stack(table_rejected_points), columns=['DG','PHASE','THER','ELEMENT','NODE','LC'])
table_rejected_points_string=df_rejected_points.to_string()
        
table_failure_points=numpy.array([dg_failure_points_list,failure_points_sho_lon_list,failure_points_up_low_list,failure_points_EL_list, failure_points_NODE_list, failure_points_LC_list ])
numpy.column_stack(table_failure_points)

df_failure_points=pandas.DataFrame(numpy.column_stack(table_failure_points), columns=['DG','PHASE','THER','ELEMENT','NODE','LC'])
table_failure_points_string=df_failure_points.to_string()


#with open('0_FailurePointsTable.txt', 'w') as text_file_table: #writes it without an index if needed
#    numpy.savetxt(text_file_table,df_failure_points.values,fmt='%s')
        
with open('0_FailurePointsTable.txt', 'w') as text_file_table:
    text_file_table.write(table_failure_points_string) 
    
#with open('0_ExcludedPointsTable.txt', 'w') as text_file_table: #writes it without an index if needed
#    numpy.savetxt(text_file_table,df_excluded_points.values,fmt='%s')
        
with open('0_ExcludedPointsTable.txt', 'w') as text_file_table:
    text_file_table.write(table_rejected_points_string) 
