"""
@author: Antl
Version: 1.00
"""
import os
import os.path
import glob
import codecs
import datetime
import xlwt

#Run Script at the folder containing new DG results.

#For Old runs:

#Replace "\" Character with "\\" as seen below!!!
#Example for Network    = '\\\\fs-saren\\07_DESIGN&ENGINEERING\\Calculation Note\\500\\502\\App.C - Transversal_External_Walls\\09_31D\\Run_Final\\dg'
#Example for Local Drive= 'C:\\Users\\Egemen.Sicim\\Desktop\\_GBS\\502\\App.C - Transversal_External_Walls\\09_31D\\Run_Final\\dg'
Old_File_Location = '\\\\fs-saren\\07_DESIGN&ENGINEERING\\Calculation Note\\500\\502\\App.C - Transversal_External_Walls\\09_31D\\Run_Final\\dg'


Settings_Switch = True
if Settings_Switch:
    Output_File = "02_F_File_Compare_Results.txt"
    Output_File_TimeStamp = False
    scriptpath = os.path.abspath(__file__)
    scriptdirectory = os.path.dirname(scriptpath)
    os.chdir(scriptdirectory)
print('Script Start')
with open(Output_File, 'w') as Clear_Data:
    Clear_Data.write('DG\tUp-Low\tPhase\tWall\tNode\tLC_New\tLC_Old\tFxx\tFyy\tFxy\tMxx\tMyy\tMxy\tVxz\tVyz\n')
print('Comparing')
for f_new in glob.iglob(os.getcwd() + '/**/*.f', recursive=True):
    with open(f_new, 'r',encoding='ansi') as f_new_In:
        f_new_lines = f_new_In.readlines()
        DG_new = f_new.split('mm')[0].strip()[-5:]
        Phase_new = f_new.split('mm')[1].replace('.f','')
        UpLow_new = f_new.split('mm')[0].strip()[-13:-8]
        #New_File_Location = f_new.split('dg'+str(DG_new))[0]+'dg'+str(DG_new)+f_new.split('dg'+str(DG_new))[1]
        New_File_Location = f_new.split('dg'+str(DG_new))[0].strip()[0:-1]
        f_new_line_number = 0
        Main_Read_f_new_Switch = False
        Result_Write_Switch = False
        for f_new_line in f_new_lines:            
            if len(f_new_line) < 5:
                Main_Read_f_new_Switch = False
                Result_Write_Switch = False
            if Main_Read_f_new_Switch:
                New_Data_00 = str(f_new_line.strip().split()[0])#wall name
                New_Data_01 = int(f_new_line.strip().split()[1])#Node
                New_Data_02 = int(f_new_line.strip().split()[2])#LC
                New_Data_03 = float(f_new_line.strip().split()[3])
                New_Data_04 = float(f_new_line.strip().split()[4])
                New_Data_05 = float(f_new_line.strip().split()[5])
                New_Data_06 = float(f_new_line.strip().split()[6])
                New_Data_07 = float(f_new_line.strip().split()[7])
                New_Data_08 = float(f_new_line.strip().split()[8])
                New_Data_09 = float(f_new_line.strip().split()[9])
                New_Data_10 = float(f_new_line.strip().split()[10])
                print(f_new.replace(New_File_Location,Old_File_Location))
                with open(f_new.replace(New_File_Location,Old_File_Location), 'r',encoding='ansi') as f_old_In:
                    f_old_lines = f_old_In.readlines()
                    f_old_line_number = 0
                    for f_old_line in f_old_lines:
                        if f_old_line_number == f_new_line_number:
                            Old_Data_00 = str(f_old_line.strip().split()[0])
                            Old_Data_01 = int(f_old_line.strip().split()[1])
                            Old_Data_02 = int(f_old_line.strip().split()[2])
                            Old_Data_03 = float(f_old_line.strip().split()[3])
                            Old_Data_04 = float(f_old_line.strip().split()[4])
                            Old_Data_05 = float(f_old_line.strip().split()[5])
                            Old_Data_06 = float(f_old_line.strip().split()[6])
                            Old_Data_07 = float(f_old_line.strip().split()[7])
                            Old_Data_08 = float(f_old_line.strip().split()[8])
                            Old_Data_09 = float(f_old_line.strip().split()[9])
                            Old_Data_10 = float(f_old_line.strip().split()[10])
                        f_old_line_number = f_old_line_number + 1
                Result_Text_Compare_00 = 'Ok'
                Result_Text_Compare_01 = 'Ok'
                Result_Text_Compare_02 = 'Ok'
                if not New_Data_00 == Old_Data_00:
                    Result_Text_Compare_00 = New_Data_00 + '/' + Old_Data_00
                if not New_Data_01 == Old_Data_01:
                    Result_Text_Compare_01 = str(New_Data_01) + '/' + str(Old_Data_01)
                if not New_Data_02 == Old_Data_02:
                    Result_Text_Compare_02 = str(New_Data_02) + '\t' + str(Old_Data_02)
                if Old_Data_03 == 0:
                    if New_Data_03 == Old_Data_03:
                        Result_Text_Compare_03 = 1
                    if not New_Data_03 == Old_Data_03:
                        Result_Text_Compare_03 = str(New_Data_03)+'/0'
                if not Old_Data_03 == 0:
                    Result_Text_Compare_03 = New_Data_03 / Old_Data_03
                if Old_Data_04 == 0:
                    if New_Data_04 == Old_Data_04:
                        Result_Text_Compare_04 = 1
                    if not New_Data_04 == Old_Data_04:
                        Result_Text_Compare_04 = str(New_Data_04)+'/0'
                if not Old_Data_04 == 0:
                    Result_Text_Compare_04 = New_Data_04 / Old_Data_04
                if Old_Data_05 == 0:
                    if New_Data_05 == Old_Data_05:
                        Result_Text_Compare_05 = 1
                    if not New_Data_05 == Old_Data_05:
                        Result_Text_Compare_05 = str(New_Data_05)+'/0'
                if not Old_Data_05 == 0:
                    Result_Text_Compare_05 = New_Data_05 / Old_Data_05
                if Old_Data_06 == 0:
                    if New_Data_06 == Old_Data_06:
                        Result_Text_Compare_06 = 1
                    if not New_Data_06 == Old_Data_06:
                        Result_Text_Compare_06 = str(New_Data_06)+'/0'
                if not Old_Data_06 == 0:
                    Result_Text_Compare_06 = New_Data_06 / Old_Data_06
                if Old_Data_07 == 0:
                    if New_Data_07 == Old_Data_07:
                        Result_Text_Compare_07 = 1
                    if not New_Data_07 == Old_Data_07:
                        Result_Text_Compare_07 = str(New_Data_07)+'/0'
                if not Old_Data_07 == 0:
                    Result_Text_Compare_07 = New_Data_07 / Old_Data_07
                if Old_Data_08 == 0:
                    if New_Data_08 == Old_Data_08:
                        Result_Text_Compare_08 = 1
                    if not New_Data_08 == Old_Data_08:
                        Result_Text_Compare_08 = str(New_Data_08)+'/0'
                if not Old_Data_08 == 0:
                    Result_Text_Compare_08 = New_Data_08 / Old_Data_08
                if Old_Data_09 == 0:
                    if New_Data_09 == Old_Data_09:
                        Result_Text_Compare_09 = 1
                    if not New_Data_09 == Old_Data_09:
                        Result_Text_Compare_09 = str(New_Data_09)+'/0'
                if not Old_Data_09 == 0:
                    Result_Text_Compare_09 = New_Data_09 / Old_Data_09
                if Old_Data_10 == 0:
                    if New_Data_10 == Old_Data_10:
                        Result_Text_Compare_10 = 1
                    if not New_Data_10 == Old_Data_10:
                        Result_Text_Compare_10 = str(New_Data_10)+'/0'
                if not Old_Data_10 == 0:
                    Result_Text_Compare_10 = New_Data_10 / Old_Data_10
                Result_Write_Switch = True
                if Result_Text_Compare_00 == 'Ok' and Result_Text_Compare_01 == 'Ok' and Result_Text_Compare_02 == 'Ok' and Result_Text_Compare_03 == 1 and Result_Text_Compare_04 == 1 and Result_Text_Compare_05 == 1 and Result_Text_Compare_06 == 1 and Result_Text_Compare_07 == 1 and Result_Text_Compare_08 == 1 and Result_Text_Compare_09 == 1 and Result_Text_Compare_10 == 1:
                    Result_Write_Switch = False
            f_new_line_number = f_new_line_number + 1
            if Result_Write_Switch:
                Result_Text = str(DG_new) + '\t' + str(UpLow_new) + '\t' + str(Phase_new) + '\t' + str(Result_Text_Compare_00) + '\t' + str(Result_Text_Compare_01) + '\t' + str(Result_Text_Compare_02) + '\t' + str(Result_Text_Compare_03) + '\t' + str(Result_Text_Compare_04) + '\t' + str(Result_Text_Compare_05) + '\t' + str(Result_Text_Compare_06) + '\t' + str(Result_Text_Compare_07) + '\t' + str(Result_Text_Compare_08) + '\t' + str(Result_Text_Compare_09) + '\t' + str(Result_Text_Compare_10)
                with open(Output_File, 'a') as Write_Result:
                    Write_Result.write(Result_Text + '\n')
            if '(KN/ml)' in f_new_line:
                Main_Read_f_new_Switch = True
#Convert To Excel
print('Creating XLS')
Output_File_Basic = Output_File.replace('.txt','')
from datetime import datetime
Output_File_Time_Stamp = str(datetime.today().strftime('%Y%m%d'))
Output_File_XLS = Output_File_Basic + '_' + Output_File_Time_Stamp + '.xls'
Workbook = xlwt.Workbook(Output_File)
Worksheet = Workbook.add_sheet('Compare_Results',cell_overwrite_ok=True)
with open (Output_File, 'r',encoding='ansi') as Output_File_In:
    Output_Lines=Output_File_In.readlines()
    Output_Line_Number = 0
    for Output_Line in Output_Lines:
        Output_Line_Cell_Number = len(Output_Line.strip().split('\t'))
        for Column in range(Output_Line_Cell_Number):
            Output_Line_Data  = Output_Line.strip().split('\t')[Column]
            Worksheet.write(Output_Line_Number,Column,Output_Line_Data)
        Output_Line_Number = Output_Line_Number + 1
Workbook.save(Output_File_XLS)
os.remove(Output_File)
