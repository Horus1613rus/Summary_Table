'''
@author: Antl
'''

import os
import glob
import sys
import subprocess

Database_Ver = '09.33d'         # '09.33d'

Ref_DG = 33124                  # 33124
Wall_Name = 'LEWESW'            # 'LEWESW'
Node_List = [150,151]           # [150,151] or [150]

State_List = ['ALS', 'ULS']     # ['ALS', 'ULS'] or ['ULS'] or ['ALS']

Combine_Text = True

Settings = True
if Settings:
    Use_DG_Run = False
    scriptpath = os.path.abspath(__file__)
    scriptdirectory = os.path.dirname(scriptpath)
    os.chdir(scriptdirectory)
    ULS_MM = ['20.3','22.5','23.5','24.6','26.3','32.5']
    ALS_MM = ['27.3','31.3','33.5','34.6']


if Use_DG_Run:
    sys.exit('Under Development...')

if not Use_DG_Run:
    os.system('alnovatek ' + str(Database_Ver))
    LC_Database = {}

    for State in State_List:
        if State == 'ULS':
            MM_List = ULS_MM
        if not State == 'ULS':
            if State == 'ALS':
                MM_List = ALS_MM
            if not State == 'ALS':
                sys.exit('State Error')
        print("--------------------\nAre you disco?\n--------------------")
        for MM in MM_List:
            with open(os.getcwd() + '\\61_Ncredb4_Run', 'w' , encoding='utf-8') as Data_Out:
                Data_Out.write('minimax exminit_soil ' + MM.replace('.',' ') + ' ' + str(Ref_DG) + ' 6999\n' + 'Quit\n')
            Ncredb4_Input = open(os.getcwd() + '\\61_Ncredb4_Run')
            Ncredb4_Process = subprocess.Popen('Ncredb4', stdin=Ncredb4_Input)
            Ncredb4_Process.wait()
            Ncredb4_Input.close()

            with open('exminit.' + MM + '.' + str(Ref_DG), 'r' , encoding='utf-8') as Data_In:
                Data_In_Lines = Data_In.readlines()
                for Data_In_Line in Data_In_Lines:
                    if not Data_In_Line.split()[2] == '6999':
                        if not Data_In_Line.split()[2] in LC_Database:
                            LC_Database[Data_In_Line.split()[2]] = {}
                        LC_Database[Data_In_Line.split()[2]]['MM'] = MM
            for exminit_File in glob.iglob(os.getcwd() + '\\*exminit*', recursive=True):
                os.remove(exminit_File)
    os.remove('61_Ncredb4_Run')
    print("--------------------\nAre you cola?\n--------------------")

    for Node in Node_List:
        if not Combine_Text:
            with open(os.getcwd() + '\\61_99_' + Wall_Name + '_' + str(Node) + '_Forces.txt', 'w' , encoding='utf-8') as Data_Out:
                Data_Out.close()
        if Combine_Text:
            with open(os.getcwd() + '\\61_99_' + Wall_Name + '_Combined_Forces.txt', 'w' , encoding='utf-8') as Data_Out:
                Data_Out.close()
    for Node in Node_List:
        with open(os.getcwd() + '\\61_gt4', 'w' , encoding='utf-8') as Data_Out:
            Data_Out.write(Wall_Name + ' ' + str(Node))
        for LC in LC_Database:
            os.system('gtforc 61_gt4 ' + str(LC))
            with open(os.getcwd() + '\\61_gt4.' + str(LC), 'r' , encoding='utf-8') as Data_In:
                Data_In_Lines = Data_In.readlines()
                for Data_In_Line in Data_In_Lines:
                    if not Combine_Text:
                        with open(os.getcwd() + '\\61_99_' + Wall_Name + '_' + str(Node) + '_Forces.txt', 'a' , encoding='utf-8') as Data_Out:
                            Data_Out.write( Data_In_Line.strip() + ' ' + str(LC_Database[LC]['MM']) + '\n')
                    if Combine_Text:
                        with open(os.getcwd() + '\\61_99_' + Wall_Name + '_Combined_Forces.txt', 'a' , encoding='utf-8') as Data_Out:
                            Data_Out.write( Data_In_Line.strip() + ' ' + str(LC_Database[LC]['MM']) + '\n')
            os.remove('61_gt4.' + str(LC))
        os.remove('61_gt4')
        os.remove('_gtforc.run')
        os.remove('_gtforc.res')

    os.remove('alias.dat')

print("--------------------\nSo what are you?!\n--------------------")
    
        

        





        
