Author= "Antl"
Creation_Date =  "2020.07.13"
Code_Ver = "0.05"

from datetime import datetime
Script_Start = datetime.now()
import os
import glob

Script_Path = os.path.abspath(__file__)
Script_Directory = os.path.dirname(Script_Path)
os.chdir(Script_Directory)

Code_File_Ver = os.path.basename(__file__).split("_v")[1].split("_Egemen")[0]
Code_No = "05_00"

Quit_Delay = 3

SLS_MMs = [1,2,3,4,5,6,7,8]
Ignored_MMs = []
Coef = 0.90

Ignored_MMs = Ignored_MMs + SLS_MMs

DB = {}
for Res_File in glob.iglob(Script_Directory + "\\**\\**.res", recursive=True):
    if not "sh." in os.path.basename(Res_File) and not "soil" in os.path.basename(Res_File):
        DG = os.path.basename(Res_File).split("mm")[0].replace("dg","")
        MM = int(os.path.basename(Res_File).split("mm")[1].split(".")[0])
        Thermal = Res_File.split("\\dg" + str(DG) + "\\")[1].replace("\\" + os.path.basename(Res_File),"")
        
        if not MM in Ignored_MMs:
            Mat_Info_Switch = False
            Concrete_Switch = False

            Line_Iteration = 0
            LC = "0"
            Node = "0"

            if not DG in DB:
                DB[DG] = {}
            if not Thermal in DB[DG]:
                DB[DG][Thermal] = {}
            if not MM in DB[DG][Thermal]:
                DB[DG][Thermal][MM] = {}
                DB[DG][Thermal][MM]["Props"] = {}
                DB[DG][Thermal][MM]["Results"] = {}
            with open(Res_File, "r" , encoding="utf-8") as File_In:
                File_Lines = File_In.readlines()
                for File_Line in File_Lines:
                    if Mat_Info_Switch:
                        fcn = float(File_Line.split()[0])
                        gam = float(File_Line.split()[2])
                        ec = float(File_Line.split()[3])
                        eclim = float(File_Line.split()[4])
                        DB[DG][Thermal][MM]["Props"]["fcd"] = 0.85 * fcn / gam
                        DB[DG][Thermal][MM]["Props"]["ec"] = ec
                        DB[DG][Thermal][MM]["Props"]["eclim"] = eclim
                        Mat_Info_Switch = False
                    if "eclim" in File_Line:
                        Mat_Info_Switch = True

                    if "LC" in File_Line:
                        LC = File_Line.split("LC:")[1].split()[0]
                    if "Node" in File_Line:
                        Node = File_Line.split("Node :")[1].split()[0]
                    
                    if not "T" in LC and not "Y" in LC and not "X" in LC:
                        if "z+" in File_Line:
                            Concrete_Switch = True
                        if Concrete_Switch:
                            e1 = float(File_Line.replace("z+","").replace("z-","").split()[1])
                            e2 = float(File_Line.replace("z+","").replace("z-","").split()[2])
                            s1 = float(File_Line.replace("z+","").replace("z-","").split()[3])
                            s2 = float(File_Line.replace("z+","").replace("z-","").split()[4])
                            if e1 >= 0 or e2 >= 0:
                                if e1 >= 0:
                                    fc2d = round(DB[DG][Thermal][MM]["Props"]["fcd"] / (0.8 + (0.1 * e1)) * -1, 2)
                                    if abs(s2) > abs(Coef * fc2d):
                                        DB[DG][Thermal][MM]["Results"][LC + "\t" + Node] = "S2=\t" + str(s2) + "\tFc2d=\t" + str(fc2d)
                                else:
                                    fc2d = round(DB[DG][Thermal][MM]["Props"]["fcd"] / (0.8 + (0.1 * e2)) * -1, 2)
                                    if abs(s1) > abs(Coef * fc2d):
                                        DB[DG][Thermal][MM]["Results"][LC + "\t" + Node] = "S1=\t" + str(s1) + "\tFc2d=\t" + str(fc2d)
                        if "z-" in File_Line:
                            Concrete_Switch = False

                            


with open(Code_No + "_Fc2d_Output.txt", "w" , encoding="utf-8") as File_Out:
    File_Out.write("DG\tThermal\tMM\tLC\tInformation\t" + str(Coef * 100) + "%\n")
    i = 0
    for DG in DB:
        for Thermal in DB[DG]:
            for MM in DB[DG][Thermal]:
                for LC in DB[DG][Thermal][MM]["Results"]:
                    File_Out.write(DG + "\t" + Thermal + "\t" + str(MM) + "\t" + LC + "\t" + DB[DG][Thermal][MM]["Results"][LC] + "\n")
                    i = i + 1
    if i == 0:
         File_Out.write("There is no Fc2d problem\n")






