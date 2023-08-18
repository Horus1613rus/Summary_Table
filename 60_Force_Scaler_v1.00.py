"""
@author: Antl
Version:
Date   :
"""
import os
import os.path
import glob

Fxx_Coef = 1
Fyy_Coef = 1
Fxy_Coef = 1
Mxx_Coef = 1
Myy_Coef = 1
Mxy_Coef = 1
Vxz_Coef = 1
Vyz_Coef = 1



Settings_Switch = True
if Settings_Switch:
    scriptpath = os.path.abspath(__file__)
    scriptdirectory = os.path.dirname(scriptpath)
    os.chdir(scriptdirectory)
    Convert_Extentions = ['.3','.5','.6','.f']
for Convert_Extention in Convert_Extentions:
    for Source_File in glob.iglob(os.getcwd() + '/**/*'+Convert_Extention, recursive=True):
        with open(Source_File, 'r',encoding='utf-8') as Source_File_Data :
            Source_File_Lines = Source_File_Data.readlines()
            Header_Enabled = False
            i = 3
            with open(Source_File + '_Scaled', 'w') as Scaled_Data:
                Scaled_Data.close()

            for Source_File_Line in Source_File_Lines:
                if 'Memb' in Source_File_Line:
                    Header_Enabled = True
                    i = 1
                
                if Header_Enabled:
                    i = i+1
                    with open(Source_File + '_Scaled', 'a') as Scaled_Data:
                        Scaled_Data.write(Source_File_Line.strip() + '\n')
                if not Header_Enabled:
                    with open(Source_File + '_Scaled', 'a') as Scaled_Data:
                        Scaled_Data.write(str(Source_File_Line.strip().split()[0]) + ' ')
                        Scaled_Data.write(str(Source_File_Line.strip().split()[1]) + ' ')
                        Scaled_Data.write(str(Source_File_Line.strip().split()[2]) + ' ')                            
                        Scaled_Data.write(str(round((float(Source_File_Line.strip().split()[3])*Fxx_Coef),2)) + ' ')
                        Scaled_Data.write(str(round((float(Source_File_Line.strip().split()[4])*Fyy_Coef),2)) + ' ')
                        Scaled_Data.write(str(round((float(Source_File_Line.strip().split()[5])*Fxy_Coef),2)) + ' ')
                        Scaled_Data.write(str(round((float(Source_File_Line.strip().split()[6])*Mxx_Coef),2)) + ' ')
                        Scaled_Data.write(str(round((float(Source_File_Line.strip().split()[7])*Myy_Coef),2)) + ' ')
                        Scaled_Data.write(str(round((float(Source_File_Line.strip().split()[8])*Mxy_Coef),2)) + ' ')
                        Scaled_Data.write(str(round((float(Source_File_Line.strip().split()[9])*Vxz_Coef),2)) + ' ')
                        Scaled_Data.write(str(round((float(Source_File_Line.strip().split()[10])*Vyz_Coef),2)) + '\n')
                if i == 3:
                    Header_Enabled = False
print('\n--------------------------------------------------\nMay the Force be with you!\n--------------------------------------------------\n')
