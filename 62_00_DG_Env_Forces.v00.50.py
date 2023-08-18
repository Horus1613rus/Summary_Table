'''
@author: Antl
Date   :
'''

import os
import glob
import sys
import subprocess

Database_Ver = '09.41d'         # '09.33d'
DG = 22651



Settings = True
if Settings:
    Use_DG_Run = False
    scriptpath = os.path.abspath(__file__)
    scriptdirectory = os.path.dirname(scriptpath)
    os.chdir(scriptdirectory)

    SLS_MM_List = ['1.3','2.3','3.5','4.5','5.6','6.7','7.3','8.5']
    ULS_ALS_MM_List = ['20.3','22.5','23.5','24.6','25.7','26.3','32.5','27.3','31.3','33.5','34.6']

try:
    os.system('copy Z:\\Database\\alias\\alias.'+str(Database_Ver)+' alias.dat')
    print("Database set for ver " + str(Database_Ver))
except:
    sys.exit('Problem copying alias.dat. Please check your connection.')

SLS_Database={}
for Force_Line_Number in range(1,27):
    SLS_Database[str(Force_Line_Number)]={}
for SLS_MM in SLS_MM_List:
    with open(os.getcwd() + '\\62_99_Ncredb4_Run', 'w' , encoding='utf-8') as Data_Out:
        Data_Out.write('minimax exminit_soil ' + SLS_MM.replace('.',' ') + ' ' + str(DG) + ' 6999\n' + 'Quit\n')
    Ncredb4_Input = open(os.getcwd() + '\\62_99_Ncredb4_Run')
    Ncredb4_Process = subprocess.Popen('Ncredb4', stdin=Ncredb4_Input)
    Ncredb4_Process.wait()
    Ncredb4_Input.close()
    with open(os.getcwd() + '\\exminit.' + SLS_MM + '.' + str(DG), 'r' , encoding='utf-8') as Data_In:
        Data_In_Lines = Data_In.readlines()
        Force_Line_Number = 0
        for Data_In_Line in Data_In_Lines:
            if not Data_In_Line.split()[2]=='6999':
                Force_Line_Number = Force_Line_Number + 1
                if Force_Line_Number == 1:
                    try:
                        if SLS_Database[str(Force_Line_Number)]['Value'] < float(Data_In_Line.split()[3]):
                            SLS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[3])
                            SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Max Fxx' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        SLS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[3])
                        SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Max Fxx' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                if Force_Line_Number == 2:
                    try:
                        if SLS_Database[str(Force_Line_Number)]['Value'] > float(Data_In_Line.split()[3]):
                            SLS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[3])
                            SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Min Fxx' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        SLS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[3])
                        SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Min Fxx' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                
                if Force_Line_Number == 3:
                    try:
                        if SLS_Database[str(Force_Line_Number)]['Value'] < float(Data_In_Line.split()[4]):
                            SLS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[4])
                            SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Max Fyy' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        SLS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[4])
                        SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Max Fyy' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                if Force_Line_Number == 4:
                    try:
                        if SLS_Database[str(Force_Line_Number)]['Value'] > float(Data_In_Line.split()[4]):
                            SLS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[4])
                            SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Min Fyy' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        SLS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[4])
                        SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Min Fyy' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                
                if Force_Line_Number == 5:
                    try:
                        if SLS_Database[str(Force_Line_Number)]['Value'] < float(Data_In_Line.split()[5]):
                            SLS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[5])
                            SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Max Fxy' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        SLS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[5])
                        SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Max Fxy' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                if Force_Line_Number == 6:
                    try:
                        if SLS_Database[str(Force_Line_Number)]['Value'] > float(Data_In_Line.split()[5]):
                            SLS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[5])
                            SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Min Fxy' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        SLS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[5])
                        SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Min Fxy' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                
                if Force_Line_Number == 7:
                    try:
                        if SLS_Database[str(Force_Line_Number)]['Value'] < float(Data_In_Line.split()[6]):
                            SLS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[6])
                            SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Max Mxx' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        SLS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[6])
                        SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Max Mxx' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                if Force_Line_Number == 8:
                    try:
                        if SLS_Database[str(Force_Line_Number)]['Value'] > float(Data_In_Line.split()[6]):
                            SLS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[6])
                            SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Min Mxx' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        SLS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[6])
                        SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Min Mxx' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                
                if Force_Line_Number == 9:
                    try:
                        if SLS_Database[str(Force_Line_Number)]['Value'] < float(Data_In_Line.split()[7]):
                            SLS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[7])
                            SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Max Myy' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        SLS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[7])
                        SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Max Myy' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                if Force_Line_Number == 10:
                    try:
                        if SLS_Database[str(Force_Line_Number)]['Value'] > float(Data_In_Line.split()[7]):
                            SLS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[7])
                            SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Min Myy' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        SLS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[7])
                        SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Min Myy' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                
                if Force_Line_Number == 11:
                    try:
                        if SLS_Database[str(Force_Line_Number)]['Value'] < float(Data_In_Line.split()[8]):
                            SLS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[8])
                            SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Max Mxy' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        SLS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[8])
                        SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Max Mxy' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                if Force_Line_Number == 12:
                    try:
                        if SLS_Database[str(Force_Line_Number)]['Value'] > float(Data_In_Line.split()[8]):
                            SLS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[8])
                            SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Min Mxy' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        SLS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[8])
                        SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Min Mxy' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                
                if Force_Line_Number == 13:
                    try:
                        if SLS_Database[str(Force_Line_Number)]['Value'] < float(Data_In_Line.split()[9]):
                            SLS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[9])
                            SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Max Vxz' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        SLS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[9])
                        SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Max Vxz' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                if Force_Line_Number == 14:
                    try:
                        if SLS_Database[str(Force_Line_Number)]['Value'] > float(Data_In_Line.split()[9]):
                            SLS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[9])
                            SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Min Vxz' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        SLS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[9])
                        SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Min Vxz' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                
                if Force_Line_Number == 15:
                    try:
                        if SLS_Database[str(Force_Line_Number)]['Value'] < float(Data_In_Line.split()[10]):
                            SLS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[10])
                            SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Max Vyz' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        SLS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[10])
                        SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Max Vyz' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                if Force_Line_Number == 16:
                    try:
                        if SLS_Database[str(Force_Line_Number)]['Value'] > float(Data_In_Line.split()[10]):
                            SLS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[10])
                            SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Min Vyz' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        SLS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[10])
                        SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Min Vyz' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]

                if Force_Line_Number == 17:
                    Faverage = (float(Data_In_Line.split()[3])+float(Data_In_Line.split()[4]))/2
                    R=pow( (((float(Data_In_Line.split()[3])-float(Data_In_Line.split()[4]))**2)/4) + (float(Data_In_Line.split()[5])**2) ,1/2)                    
                    try:
                        if SLS_Database[str(Force_Line_Number)]['Value'] < (Faverage + R):
                            SLS_Database[str(Force_Line_Number)]['Value'] = (Faverage + R)
                            SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Max F1' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        SLS_Database[str(Force_Line_Number)]['Value'] = (Faverage + R)
                        SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Max F1' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                if Force_Line_Number == 18:
                    Faverage = (float(Data_In_Line.split()[3])+float(Data_In_Line.split()[4]))/2
                    R=pow( (((float(Data_In_Line.split()[3])-float(Data_In_Line.split()[4]))**2)/4) + (float(Data_In_Line.split()[5])**2) ,1/2)                    
                    try:
                        if SLS_Database[str(Force_Line_Number)]['Value'] > (Faverage - R):
                            SLS_Database[str(Force_Line_Number)]['Value'] = (Faverage - R)
                            SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Min F1' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        SLS_Database[str(Force_Line_Number)]['Value'] = (Faverage + R)
                        SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Min F1' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                
                if Force_Line_Number == 19:
                    Faverage = (float(Data_In_Line.split()[6])+float(Data_In_Line.split()[7]))/2
                    R=pow( (((float(Data_In_Line.split()[6])-float(Data_In_Line.split()[7]))**2)/4) + (float(Data_In_Line.split()[8])**2) ,1/2)                    
                    try:
                        if SLS_Database[str(Force_Line_Number)]['Value'] < (Faverage + R):
                            SLS_Database[str(Force_Line_Number)]['Value'] = (Faverage + R)
                            SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Max M1' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        SLS_Database[str(Force_Line_Number)]['Value'] = (Faverage + R)
                        SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Max M1' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                if Force_Line_Number == 20:
                    Faverage = (float(Data_In_Line.split()[6])+float(Data_In_Line.split()[7]))/2
                    R=pow( (((float(Data_In_Line.split()[6])-float(Data_In_Line.split()[7]))**2)/4) + (float(Data_In_Line.split()[8])**2) ,1/2)                    
                    try:
                        if SLS_Database[str(Force_Line_Number)]['Value'] > (Faverage - R):
                            SLS_Database[str(Force_Line_Number)]['Value'] = (Faverage - R)
                            SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Min M1' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        SLS_Database[str(Force_Line_Number)]['Value'] = (Faverage + R)
                        SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Min M1' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]

                if Force_Line_Number == 25:
                    R=pow(   (float(Data_In_Line.split()[9])**2) +  (float(Data_In_Line.split()[10])**2) , 1/2)
                    try:
                        if SLS_Database[str(Force_Line_Number)]['Value'] < R:
                            SLS_Database[str(Force_Line_Number)]['Value'] = R
                            SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Max V1' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        SLS_Database[str(Force_Line_Number)]['Value'] = R
                        SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Max V1' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                if Force_Line_Number == 26:
                    R=pow(   (float(Data_In_Line.split()[9])**2) +  (float(Data_In_Line.split()[10])**2) , 1/2)                    
                    try:
                        if SLS_Database[str(Force_Line_Number)]['Value'] > R:
                            SLS_Database[str(Force_Line_Number)]['Value'] = R
                            SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Min V1' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        SLS_Database[str(Force_Line_Number)]['Value'] = R
                        SLS_Database[str(Force_Line_Number)]['FullLine'] = 'Min V1' + '\t' + SLS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]

    os.remove('exminit.' + SLS_MM + '.' + str(DG))

with open(os.getcwd() + '\\62_10_SLS_env_Forces.txt', 'w' , encoding='utf-8') as Data_Out:
    Data_Out.write('Sort\tMinimax\tMemb\tNode\tLC\tFxx\tFyy\tFxy\tMxx\tMyy\tMxy\tVxz\tVyz\n')
    for i in range(1,21):
        Data_Out.write(SLS_Database[str(i)]['FullLine'] + '\n')
    Data_Out.write(SLS_Database['25']['FullLine'] + '\n')
    Data_Out.write(SLS_Database['26']['FullLine'] + '\n')

    
ULS_ALS_Database={}
for Force_Line_Number in range(1,27):
    ULS_ALS_Database[str(Force_Line_Number)]={}
for ULS_ALS_MM in ULS_ALS_MM_List:
    with open(os.getcwd() + '\\62_99_Ncredb4_Run', 'w' , encoding='utf-8') as Data_Out:
        Data_Out.write('minimax exminit_soil ' + ULS_ALS_MM.replace('.',' ') + ' ' + str(DG) + ' 6999\n' + 'Quit\n')
    Ncredb4_Input = open(os.getcwd() + '\\62_99_Ncredb4_Run')
    Ncredb4_Process = subprocess.Popen('Ncredb4', stdin=Ncredb4_Input)
    Ncredb4_Process.wait()
    Ncredb4_Input.close()
    with open(os.getcwd() + '\\exminit.' + ULS_ALS_MM + '.' + str(DG), 'r' , encoding='utf-8') as Data_In:
        Data_In_Lines = Data_In.readlines()
        Force_Line_Number = 0
        for Data_In_Line in Data_In_Lines:
            if not Data_In_Line.split()[2]=='6999':
                Force_Line_Number = Force_Line_Number + 1
                if Force_Line_Number == 1:
                    try:
                        if ULS_ALS_Database[str(Force_Line_Number)]['Value'] < float(Data_In_Line.split()[3]):
                            ULS_ALS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[3])
                            ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Max Fxx' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        ULS_ALS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[3])
                        ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Max Fxx' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                if Force_Line_Number == 2:
                    try:
                        if ULS_ALS_Database[str(Force_Line_Number)]['Value'] > float(Data_In_Line.split()[3]):
                            ULS_ALS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[3])
                            ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Min Fxx' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        ULS_ALS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[3])
                        ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Min Fxx' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                
                if Force_Line_Number == 3:
                    try:
                        if ULS_ALS_Database[str(Force_Line_Number)]['Value'] < float(Data_In_Line.split()[4]):
                            ULS_ALS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[4])
                            ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Max Fyy' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        ULS_ALS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[4])
                        ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Max Fyy' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                if Force_Line_Number == 4:
                    try:
                        if ULS_ALS_Database[str(Force_Line_Number)]['Value'] > float(Data_In_Line.split()[4]):
                            ULS_ALS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[4])
                            ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Min Fyy' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        ULS_ALS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[4])
                        ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Min Fyy' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                
                if Force_Line_Number == 5:
                    try:
                        if ULS_ALS_Database[str(Force_Line_Number)]['Value'] < float(Data_In_Line.split()[5]):
                            ULS_ALS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[5])
                            ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Max Fxy' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        ULS_ALS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[5])
                        ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Max Fxy' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                if Force_Line_Number == 6:
                    try:
                        if ULS_ALS_Database[str(Force_Line_Number)]['Value'] > float(Data_In_Line.split()[5]):
                            ULS_ALS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[5])
                            ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Min Fxy' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        ULS_ALS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[5])
                        ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Min Fxy' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                
                if Force_Line_Number == 7:
                    try:
                        if ULS_ALS_Database[str(Force_Line_Number)]['Value'] < float(Data_In_Line.split()[6]):
                            ULS_ALS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[6])
                            ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Max Mxx' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        ULS_ALS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[6])
                        ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Max Mxx' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                if Force_Line_Number == 8:
                    try:
                        if ULS_ALS_Database[str(Force_Line_Number)]['Value'] > float(Data_In_Line.split()[6]):
                            ULS_ALS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[6])
                            ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Min Mxx' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        ULS_ALS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[6])
                        ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Min Mxx' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                
                if Force_Line_Number == 9:
                    try:
                        if ULS_ALS_Database[str(Force_Line_Number)]['Value'] < float(Data_In_Line.split()[7]):
                            ULS_ALS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[7])
                            ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Max Myy' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        ULS_ALS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[7])
                        ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Max Myy' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                if Force_Line_Number == 10:
                    try:
                        if ULS_ALS_Database[str(Force_Line_Number)]['Value'] > float(Data_In_Line.split()[7]):
                            ULS_ALS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[7])
                            ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Min Myy' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        ULS_ALS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[7])
                        ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Min Myy' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                
                if Force_Line_Number == 11:
                    try:
                        if ULS_ALS_Database[str(Force_Line_Number)]['Value'] < float(Data_In_Line.split()[8]):
                            ULS_ALS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[8])
                            ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Max Mxy' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        ULS_ALS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[8])
                        ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Max Mxy' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                if Force_Line_Number == 12:
                    try:
                        if ULS_ALS_Database[str(Force_Line_Number)]['Value'] > float(Data_In_Line.split()[8]):
                            ULS_ALS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[8])
                            ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Min Mxy' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        ULS_ALS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[8])
                        ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Min Mxy' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                
                if Force_Line_Number == 13:
                    try:
                        if ULS_ALS_Database[str(Force_Line_Number)]['Value'] < float(Data_In_Line.split()[9]):
                            ULS_ALS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[9])
                            ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Max Vxz' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        ULS_ALS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[9])
                        ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Max Vxz' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                if Force_Line_Number == 14:
                    try:
                        if ULS_ALS_Database[str(Force_Line_Number)]['Value'] > float(Data_In_Line.split()[9]):
                            ULS_ALS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[9])
                            ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Min Vxz' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        ULS_ALS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[9])
                        ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Min Vxz' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                
                if Force_Line_Number == 15:
                    try:
                        if ULS_ALS_Database[str(Force_Line_Number)]['Value'] < float(Data_In_Line.split()[10]):
                            ULS_ALS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[10])
                            ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Max Vyz' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        ULS_ALS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[10])
                        ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Max Vyz' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                if Force_Line_Number == 16:
                    try:
                        if ULS_ALS_Database[str(Force_Line_Number)]['Value'] > float(Data_In_Line.split()[10]):
                            ULS_ALS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[10])
                            ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Min Vyz' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        ULS_ALS_Database[str(Force_Line_Number)]['Value'] = float(Data_In_Line.split()[10])
                        ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Min Vyz' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]

                if Force_Line_Number == 17:
                    Faverage = (float(Data_In_Line.split()[3])+float(Data_In_Line.split()[4]))/2
                    R=pow( (((float(Data_In_Line.split()[3])-float(Data_In_Line.split()[4]))**2)/4) + (float(Data_In_Line.split()[5])**2) ,1/2)                    
                    try:
                        if ULS_ALS_Database[str(Force_Line_Number)]['Value'] < (Faverage + R):
                            ULS_ALS_Database[str(Force_Line_Number)]['Value'] = (Faverage + R)
                            ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Max F1' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        ULS_ALS_Database[str(Force_Line_Number)]['Value'] = (Faverage + R)
                        ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Max F1' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                if Force_Line_Number == 18:
                    Faverage = (float(Data_In_Line.split()[3])+float(Data_In_Line.split()[4]))/2
                    R=pow( (((float(Data_In_Line.split()[3])-float(Data_In_Line.split()[4]))**2)/4) + (float(Data_In_Line.split()[5])**2) ,1/2)                    
                    try:
                        if ULS_ALS_Database[str(Force_Line_Number)]['Value'] > (Faverage - R):
                            ULS_ALS_Database[str(Force_Line_Number)]['Value'] = (Faverage - R)
                            ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Min F1' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        ULS_ALS_Database[str(Force_Line_Number)]['Value'] = (Faverage + R)
                        ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Min F1' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                
                if Force_Line_Number == 19:
                    Faverage = (float(Data_In_Line.split()[6])+float(Data_In_Line.split()[7]))/2
                    R=pow( (((float(Data_In_Line.split()[6])-float(Data_In_Line.split()[7]))**2)/4) + (float(Data_In_Line.split()[8])**2) ,1/2)                    
                    try:
                        if ULS_ALS_Database[str(Force_Line_Number)]['Value'] < (Faverage + R):
                            ULS_ALS_Database[str(Force_Line_Number)]['Value'] = (Faverage + R)
                            ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Max M1' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        ULS_ALS_Database[str(Force_Line_Number)]['Value'] = (Faverage + R)
                        ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Max M1' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                if Force_Line_Number == 20:
                    Faverage = (float(Data_In_Line.split()[6])+float(Data_In_Line.split()[7]))/2
                    R=pow( (((float(Data_In_Line.split()[6])-float(Data_In_Line.split()[7]))**2)/4) + (float(Data_In_Line.split()[8])**2) ,1/2)                    
                    try:
                        if ULS_ALS_Database[str(Force_Line_Number)]['Value'] > (Faverage - R):
                            ULS_ALS_Database[str(Force_Line_Number)]['Value'] = (Faverage - R)
                            ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Min M1' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        ULS_ALS_Database[str(Force_Line_Number)]['Value'] = (Faverage + R)
                        ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Min M1' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]

                if Force_Line_Number == 25:
                    R=pow(   (float(Data_In_Line.split()[9])**2) +  (float(Data_In_Line.split()[10])**2) , 1/2)
                    try:
                        if ULS_ALS_Database[str(Force_Line_Number)]['Value'] < R:
                            ULS_ALS_Database[str(Force_Line_Number)]['Value'] = R
                            ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Max V1' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        ULS_ALS_Database[str(Force_Line_Number)]['Value'] = R
                        ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Max V1' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                if Force_Line_Number == 26:
                    R=pow(   (float(Data_In_Line.split()[9])**2) +  (float(Data_In_Line.split()[10])**2) , 1/2)                    
                    try:
                        if ULS_ALS_Database[str(Force_Line_Number)]['Value'] > R:
                            ULS_ALS_Database[str(Force_Line_Number)]['Value'] = R
                            ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Min V1' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]
                    except:
                        ULS_ALS_Database[str(Force_Line_Number)]['Value'] = R
                        ULS_ALS_Database[str(Force_Line_Number)]['FullLine'] = 'Min V1' + '\t' + ULS_ALS_MM + '\t' + Data_In_Line.split()[0] + '\t' + Data_In_Line.split()[1] + '\t' + Data_In_Line.split()[2] + '\t' + Data_In_Line.split()[3] + '\t' + Data_In_Line.split()[4] + '\t' + Data_In_Line.split()[5] + '\t'+ Data_In_Line.split()[6] + '\t' + Data_In_Line.split()[7] + '\t' + Data_In_Line.split()[8] + '\t' + Data_In_Line.split()[9] + '\t' + Data_In_Line.split()[10]

    os.remove('exminit.' + ULS_ALS_MM + '.' + str(DG))

with open(os.getcwd() + '\\62_10_ULS_ALS_env_Forces.txt', 'w' , encoding='utf-8') as Data_Out:
    Data_Out.write('Sort\tMinimax\tMemb\tNode\tLC\tFxx\tFyy\tFxy\tMxx\tMyy\tMxy\tVxz\tVyz\n')
    for i in range(1,21):
        Data_Out.write(ULS_ALS_Database[str(i)]['FullLine'] + '\n')
    Data_Out.write(ULS_ALS_Database['25']['FullLine'] + '\n')
    Data_Out.write(ULS_ALS_Database['26']['FullLine'] + '\n')
        
os.remove('alias.dat')
os.remove('62_99_Ncredb4_Run')
        





        
