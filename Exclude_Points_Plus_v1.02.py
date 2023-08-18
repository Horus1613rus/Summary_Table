"""
@author: Antl
"""
import os
import os.path
import glob
import codecs
import datetime


Settings_Switch = True
if Settings_Switch:    
    Output_File = "00_Egemen_Exclude_Points_Plus.txt"
    Inspection_File = ".res"
    Output_File_TimeStamp = False
    scriptpath = os.path.abspath(__file__)
    scriptdirectory = os.path.dirname(scriptpath)
    os.chdir(scriptdirectory)
    Output_Text = ''

with open(Output_File, 'w') as Clear_Data:
    if Output_File_TimeStamp:
        Clear_Data.write('Script start = ' + str(datetime.datetime.now()) + '\n\n')
    Clear_Data.write('DG\tPhase\tUpLow\tWall\tNode\tLC\tC_Stress_Top_1\tC_Stress_Top_2\tC_Stress_Bot_1\tC_Stress_Bot_2\tFxx_Given\tFxx_Found\tFyy_Given\tFyy_Found\tFxy_Given\tFxy_Found\tMxx_Given\tMxx_Found\tMyy_Given\tMyy_Found\tMxy_Given\tMxy_Found\tSteel_UR\n')

for csmr in glob.iglob(os.getcwd() + '/**/*.csmr', recursive=True):
    with open(csmr, 'r',encoding='ansi') as csmrIn:
        csmrlines = csmrIn.readlines()
        DG = csmr.split('mm')[0].strip()[-5:]
        Phase = csmr.split('mm')[1].replace('.res.csmr','')
        UpLow = csmr.split('mm')[0].strip()[-13:-8]
        if UpLow == 'ower2':
            UpLow == 'lower2'
        Exclude_Evaluation_Switch = False
        for csmrline in csmrlines:
            if 'FAILURES:' in csmrline:
                Exclude_Evaluation_Switch = False
            if Exclude_Evaluation_Switch:
                if len(csmrline.strip()) > 5:
                    Wall = csmrline.strip().split()[0]
                    Node = csmrline.strip().split()[1]
                    LC = csmrline.strip().split()[2]
                    LC_Match = False
                    Node_Match = False
                    Wall_Match = False
                    Finalize_Investigation_Switch = False
                    
                    Steel_UR_Text = ''
                    C_Top_1 = ''
                    C_Top_2 = ''
                    C_Bot_1 = ''
                    C_Bot_2 = ''
                    Fxx_Given = ''
                    Fxx_Found = ''
                    Fyy_Given = ''
                    Fyy_Found = ''
                    Fxy_Given = ''
                    Fxy_Found = ''
                    Mxx_Given = ''
                    Mxx_Found = ''
                    Myy_Given = ''
                    Myy_Found = ''
                    Mxy_Given = ''
                    Mxy_Found = ''

                    with open(csmr.strip()[:-5], 'r',encoding='ansi') as resIn:
                        reslines = resIn.readlines()
                        for resline in reslines:
                            if not Finalize_Investigation_Switch:
                                if 'CONSECT' in resline:
                                    LC_Match = False
                                    Node_Match = False
                                    Wall_Match = False
                                if 'LC: '+LC in resline:
                                    LC_Match = True
                                if 'LC:  '+LC in resline:
                                    LC_Match = True
                                if ' ' + Node + ' ' in resline:
                                    Node_Match = True                        
                                if ' ' + Wall + ' ' in resline:
                                    Wall_Match = True
                                if LC_Match:
                                    if Node_Match:
                                        if Wall_Match:
                                            if 'z+  21' in resline:
                                                C_Top_1 = resline.strip().split()[4]
                                                C_Top_2 = resline.strip().split()[5]
                                            if 'z-   1' in resline:
                                                C_Bot_1 = resline.strip().split()[4]
                                                C_Bot_2 = resline.strip().split()[5]
                                            if 'FXX' in resline:
                                                Fxx_Given = resline.strip().split()[1]
                                                Fxx_Found = resline.strip().split()[2]
                                            if 'FYY' in resline:
                                                Fyy_Given = resline.strip().split()[1]
                                                Fyy_Found = resline.strip().split()[2]
                                            if 'FXY' in resline:
                                                Fxy_Given = resline.strip().split()[1]
                                                Fxy_Found = resline.strip().split()[2]
                                            if 'MXX' in resline:
                                                Mxx_Given = resline.strip().split()[1]
                                                Mxx_Found = resline.strip().split()[2]
                                            if 'MYY' in resline:
                                                Myy_Given = resline.strip().split()[1]
                                                Myy_Found = resline.strip().split()[2]
                                            if 'MXY' in resline:
                                                Mxy_Given = resline.strip().split()[1]
                                                Mxy_Found = resline.strip().split()[2]
                                            if 'utilisation :' in resline:
                                                Steel_Layer = resline.strip().split()[2]
                                                Steel_Utilization = resline.strip().split()[5]
                                                Steel_UR_Text = Steel_UR_Text + 'L' + str(Steel_Layer) + '=' + str(Steel_Utilization) + ' '
                                            if ' Thermal strains:' in resline:
                                                Finalize_Investigation_Switch = True
                                            

                    Exclude_Text = str(DG) + '\t' + str(Phase) + '\t' + str(UpLow) + '\t' + str(Wall) + '\t' + str(Node) + '\t' + str(LC) + '\t' + str(C_Top_1) + '\t' + str(C_Top_2) + '\t' + str(C_Bot_1) + '\t' + str(C_Bot_2) + '\t' + str(Fxx_Given) + '\t' + str(Fxx_Found) + '\t' + str(Fyy_Given) + '\t' + str(Fyy_Found) + '\t' + str(Fxy_Given) + '\t' + str(Fxy_Found) + '\t' + str(Mxx_Given) + '\t' + str(Mxx_Found) + '\t' + str(Myy_Given) + '\t' + str(Myy_Found) + '\t' + str(Mxy_Given) + '\t' + str(Mxy_Found) + '\t' + Steel_UR_Text
                    with open(Output_File, 'a') as Write_Result:
                        Write_Result.write(Exclude_Text + '\n')

                    resIn.close()

            if 'POINTS EXCLUDED FROM SUMMARY:' in csmrline:
                Exclude_Evaluation_Switch = True

    csmrIn.close()





if Output_File_TimeStamp:
    with open(Output_File, 'a') as Compare_Result:
        Compare_Result.write('\n\nScript end = ' + str(datetime.datetime.now()))
