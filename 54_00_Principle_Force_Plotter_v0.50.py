from datetime import datetime
Script_Start = datetime.now()
Author= "Antl"
Creation_Date =  "2020.10.10"
Code_Ver = "0.50"
##--------------------------------------------------


Database_Ver = "09.42d"

Wall_List = ["LEW18SW", "LEW19SWC", "LEW19SEC", "LEW18SE"]
LC_List = [3063114200000]
Max_Stress = 0
Min_Stress = -0.6


Thermal = [3658]


##--------------------------------------------------
import os
import sys
import glob
import time
import math
import subprocess
import matplotlib.pyplot as PLT
import matplotlib.patches as mpatches

scriptpath = os.path.abspath(__file__)
Script_Directory = os.path.dirname(scriptpath)
os.chdir(Script_Directory)

Debug = False
Export_Txt = True
Invert = False
Local_DB = False

Quit_Delay = 3

FNULL = open(os.devnull, "w")

Working_Directory = "Principle_Membrane_Stress_Files"

dpi = 1000




Code_File_Ver = os.path.basename(__file__).split("_v")[1].split("_Egemen")[0]
if not Code_Ver == Code_File_Ver:
    print(str(datetime.now() - Script_Start) + " ! Error-001 : Version does not match.")
    time.sleep(Quit_Delay)
    sys.exit()

if not os.path.isdir(Working_Directory):
    os.makedirs(Working_Directory)

with open("54_10_Prop_SLS.txt", 'r' , encoding='utf-8') as File_In:
    Lines = File_In.readlines()
    Rep_Prop_SLS = ""
    Wall_Thickness = int(Lines[0].split()[6])
    Con_Prop = Lines[0].split()[2] + " " + Lines[0].split()[4]
    for Line in Lines:
        Rep_Prop_SLS = Rep_Prop_SLS + Line.strip() + "\n"

os.chdir(Working_Directory)
print("egemen.sicim@rencons.com\n")
try:
    if Local_DB:
        os.system("copy D:\\_Local_Tools\Alias\\alias.local." + str(Database_Ver)+" alias.dat")
        print("Database set to ver " + str(Database_Ver))
    else:
        os.system("copy Z:\\Database\\alias\\alias." + str(Database_Ver)+" alias.dat")
        print("Database set to ver " + str(Database_Ver))
except:
    print(str(datetime.now() - Script_Start) + " : Problem copying alias.dat. Please check your connection.")
    time.sleep(Quit_Delay)
    sys.exit()

Thermal_Text = ""
for TH in Thermal:
    Thermal_Text = Thermal_Text + " " + str(TH)

Coordinates = {}
Position = {}

print(str(datetime.now() - Script_Start) + " : Gathering Coordinates")
for Wall in Wall_List:
    if not Wall in Coordinates:
        Coordinates[Wall] = {}
        
    for Coor_File in glob.iglob(Script_Directory + "\\**\\"+ str(Wall) + ".dg.txt", recursive=True):
        with open(Coor_File, 'r' , encoding='utf-8') as Coor_File_In:
            Coor_File_Lines = Coor_File_In.readlines()
            for Coor_File_Line in Coor_File_Lines:
                Node = Coor_File_Line.split()[1]
                Coordinates[Wall][Node] = {}
                Coordinates[Wall][Node]["X"] = float(Coor_File_Line.split()[2])
                Coordinates[Wall][Node]["Y"] = float(Coor_File_Line.split()[3])
                Coordinates[Wall][Node]["Z"] = float(Coor_File_Line.split()[4])

    Temp = {}
    for Variable in "X Y Z".split():
        Temp[Variable + "_Max"] = -1000000000000
        Temp[Variable + "_Min"] = 1000000000000

    for Node in Coordinates[Wall]:
        for Variable in "X Y Z".split():
            if Coordinates[Wall][Node][Variable] < Temp[Variable + "_Min"]:
                Temp[Variable + "_Min"] = Coordinates[Wall][Node][Variable]
            if Coordinates[Wall][Node][Variable] > Temp[Variable + "_Max"]:
                Temp[Variable + "_Max"] = Coordinates[Wall][Node][Variable]

    for Variable in "X Y Z".split():
        Temp[Variable + "_Diff"] = abs(Temp[Variable + "_Max"] - Temp[Variable + "_Min"])
    if Temp["X_Diff"] > Temp["Z_Diff"] and Temp["Y_Diff"] > Temp["Z_Diff"]:
        Position[Wall] = "X_Y"
        PlotPlane = "X_Y"
    elif Temp["X_Diff"] > Temp["Y_Diff"]:
        Position[Wall] = "X_Z"
        PlotPlane = "X_Z"
    else:
        Position[Wall] = "Y_Z"
        PlotPlane = "Y_Z"

for Wall in Position:
    if not Position[Wall] == PlotPlane:
        print(str(datetime.now() - Script_Start) + " : Walls are not in the same plane")
        time.sleep(Quit_Delay)
        sys.exit()


print(str(datetime.now() - Script_Start) + " : Preparing Node List")
Text = ""
for Wall in Coordinates:
    for Node in Coordinates[Wall]:
        Text = Text + str(Wall) + " " + str(Node) + "\n"
with open("Node_List", "w" , encoding = "utf-8") as Data_Out:
    Data_Out.write(Text)

LC_Database = {}
for LC in LC_List + Thermal:
    if not LC in LC_Database:
        LC_Database[LC] = {}
    Skipping = False
    for root, dirs, files in os.walk("."):
        Search_File = "LC_" + str(LC)
        if Search_File in files:
            print(str(datetime.now() - Script_Start) + " : Skipping " + str(LC))
            Skipping = True
    if not Skipping:
        print(str(datetime.now() - Script_Start) + " : Preparing forces for " + str(LC))
        if Local_DB:
            with open("Ncredb4_Run", "w" , encoding="utf-8") as Data_Out:
                Data_Out.write("fbas gtforc " + "Node_List " + str(LC) + " LC_" + str(LC) +" 1")
            Ncredb4_Input = open("Ncredb4_Run")
            Ncredb4_Process = subprocess.Popen('ncredb4_local', stdout = FNULL, stdin = Ncredb4_Input)
            Ncredb4_Process.wait()
            Ncredb4_Input.close()
        else:
            with open("Ncredb4_Run", "w" , encoding="utf-8") as Data_Out:
                Data_Out.write("fbas gtforc " + "Node_List " + str(LC) + " LC_" + str(LC) +" 1")
            Ncredb4_Input = open("Ncredb4_Run")
            Ncredb4_Process = subprocess.Popen('ncredb4', stdout = FNULL, stdin = Ncredb4_Input)
            Ncredb4_Process.wait()
            Ncredb4_Input.close()

    if not LC in Thermal:
        with open("LC_" + str(LC), "r" , encoding="utf-8") as Data_In:
            Lines = Data_In.readlines()
            for Line in Lines:
                Wall = Line.split()[0]
                Node = Line.split()[1]
                if not Wall in LC_Database[LC]:
                    LC_Database[LC][Wall] = {}
                if not Node in LC_Database[LC][Wall]:
                    LC_Database[LC][Wall][Node] = {}
                LC_Database[LC][Wall][Node]["Fxx"] = float(Line.split()[3])
                LC_Database[LC][Wall][Node]["Fyy"] = float(Line.split()[4])
                LC_Database[LC][Wall][Node]["Fxy"] = float(Line.split()[5])

with open("Prop_SLS", "w" , encoding="utf-8") as File_Out:
    File_Out.write(Rep_Prop_SLS)

LC_Th_Database = {}
for LC in LC_Database:

    with open("LC_" + str(LC), "r" , encoding="utf-8") as Data_In:
        LC_Lines = Data_In.readlines()

    if not LC in LC_Th_Database:
        LC_Th_Database[LC] = {}

    for LC_Th in Thermal:
        print(str(datetime.now() - Script_Start) + " : Converting forces of " + str(LC) + " + " + str(LC_Th))

        if not LC_Th in LC_Th_Database[LC]:
            LC_Th_Database[LC][LC_Th] = {}

        with open("LC_" + str(LC_Th), "r" , encoding="utf-8") as Data_In:
            LC_Th_Lines = Data_In.readlines()

        Text = ""
        for i in range(0,len(LC_Th_Lines)):
            Text = Text + LC_Lines[i].strip() + "\n" + LC_Th_Lines[i].strip() + "\n"

        with open("F", "w" , encoding="utf-8") as Data_Out:
            Data_Out.write(Text)

        with open("Consec_Run", "w" , encoding="utf-8") as Data_Out:
            Data_Out.write("Automated\nProp_SLS\nF\nResult.txt\n1\nsls long\ndnv\n" + Con_Prop + "\nn\n1 1 1 1 1 1 1 1\n0 0 0 0 0 0 0 0\n0.0")

        ConSec_Input = open("Consec_Run")
        ConSec_Process = subprocess.Popen("nconsectdnv", stdout=FNULL, stdin=ConSec_Input)
        ConSec_Process.wait()
        ConSec_Input.close()

        with open("Result.txt", "r" , encoding="utf-8") as Data_In:
            Lines = Data_In.readlines()
            res_read_switch = False
            for Line in Lines:
                if res_read_switch:
                    if "Element :" in Line:
                        Wall = Line.split("Element :")[1].split()[0]
                        Node = Line.split("Node :")[1].split()[0]
                        if not Wall in LC_Th_Database[LC][LC_Th]:
                            LC_Th_Database[LC][LC_Th][Wall] = {}
                        if not Node in LC_Th_Database[LC][LC_Th][Wall]:
                            LC_Th_Database[LC][LC_Th][Wall][Node] = {}
                    if "FXX" in Line:
                        LC_Th_Database[LC][LC_Th][Wall][Node]["Fxx"] = float(Line.strip()[25:34].strip())
                    if "FYY" in Line:
                        LC_Th_Database[LC][LC_Th][Wall][Node]["Fyy"] = float(Line.strip()[25:34].strip())
                    if "FXY" in Line:
                        LC_Th_Database[LC][LC_Th][Wall][Node]["Fxy"] = float(Line.strip()[25:34].strip())
                    if "MXX" in Line:
                        res_read_switch = False
                if "LC: T" in Line:
                    res_read_switch = True



def Principle_func(Fxx, Fyy, Fxy):
    R = round(math.sqrt(math.pow(0.5 * (Fxx - Fyy), 2) + math.pow(Fxy, 2)), 3)
    Favg = round(0.5 * (Fxx + Fyy), 3)
    Fp1 = round(Favg + R, 3)
    Fp2 = round(Favg - R, 3)
    #print(Fxx, Fyy, Fxy)
    #print(R, Favg)
    #print(Fp1, Fp2)
    return Fp1, Fp2


Principle_DB = {}
for LC in LC_Database:
    print(str(datetime.now() - Script_Start) + " : Calculating principles forces for " + str(LC))
    for Wall in LC_Database[LC]:
        if not Wall in Principle_DB:
            Principle_DB[Wall] = {}
        for Node in LC_Database[LC][Wall]:
            Temp_Value = {}
            if not Node in Principle_DB[Wall]:
                Principle_DB[Wall][Node] = {}
                Principle_DB[Wall][Node]["No_Th"] = {}
                Principle_DB[Wall][Node]["No_Th"]["Fp1"] = -10000000
                Principle_DB[Wall][Node]["No_Th"]["Fp2"] = 10000000
                Principle_DB[Wall][Node]["Env"] = {}
                Principle_DB[Wall][Node]["Env"]["Fp1"] = -10000000
                Principle_DB[Wall][Node]["Env"]["Fp2"] = 10000000
                for TH in Thermal:
                    Principle_DB[Wall][Node][str(TH)] = {}
                    Principle_DB[Wall][Node][str(TH)]["Fp1"] = -10000000
                    Principle_DB[Wall][Node][str(TH)]["Fp2"] = 10000000
            
            Temp_Value["Fxx"] = round(LC_Database[LC][Wall][Node]["Fxx"], 3)
            Temp_Value["Fyy"] = round(LC_Database[LC][Wall][Node]["Fyy"], 3)
            Temp_Value["Fxy"] = round(LC_Database[LC][Wall][Node]["Fxy"], 3)
            
            Temp_Value["Fp1"], Temp_Value["Fp2"] = Principle_func(Temp_Value["Fxx"], Temp_Value["Fyy"], Temp_Value["Fxy"])

            if Temp_Value["Fp1"] > Principle_DB[Wall][Node]["No_Th"]["Fp1"]:
                Principle_DB[Wall][Node]["No_Th"]["Fp1"] = Temp_Value["Fp1"]
            if Temp_Value["Fp2"] < Principle_DB[Wall][Node]["No_Th"]["Fp2"]:
                Principle_DB[Wall][Node]["No_Th"]["Fp2"] = Temp_Value["Fp2"]
            
            Temp_Value = {}
            for TH_temp in Thermal:
                Temp_Value["Fxx"] = round(LC_Database[LC][Wall][Node]["Fxx"] + LC_Th_Database[LC][TH_temp][Wall][Node]["Fxx"], 3)
                Temp_Value["Fyy"] = round(LC_Database[LC][Wall][Node]["Fyy"] + LC_Th_Database[LC][TH_temp][Wall][Node]["Fyy"], 3)
                Temp_Value["Fxy"] = round(LC_Database[LC][Wall][Node]["Fxy"] + LC_Th_Database[LC][TH_temp][Wall][Node]["Fxy"], 3)
                Temp_Value["Fp1"], Temp_Value["Fp2"] = Principle_func(Temp_Value["Fxx"], Temp_Value["Fyy"], Temp_Value["Fxy"])

                if Temp_Value["Fp1"] > Principle_DB[Wall][Node][str(TH_temp)]["Fp1"]:
                    Principle_DB[Wall][Node][str(TH_temp)]["Fp1"] = Temp_Value["Fp1"]
                if Temp_Value["Fp2"] < Principle_DB[Wall][Node][str(TH_temp)]["Fp2"]:
                    Principle_DB[Wall][Node][str(TH_temp)]["Fp2"] = Temp_Value["Fp2"]

            for i in ("No_Th" + Thermal_Text).split():
                if Principle_DB[Wall][Node][i]["Fp1"] > Principle_DB[Wall][Node]["Env"]["Fp1"]:
                    Principle_DB[Wall][Node]["Env"]["Fp1"] = Principle_DB[Wall][Node][i]["Fp1"]
                    Principle_DB[Wall][Node]["Env"]["Fp1_Res"] = i
                    Principle_DB[Wall][Node]["Env"]["Fp1_LC"] = LC
                if Principle_DB[Wall][Node][i]["Fp2"] < Principle_DB[Wall][Node]["Env"]["Fp2"]:
                    Principle_DB[Wall][Node]["Env"]["Fp2"] = Principle_DB[Wall][Node][i]["Fp2"]
                    Principle_DB[Wall][Node]["Env"]["Fp2_Res"] = i
                    Principle_DB[Wall][Node]["Env"]["Fp2_LC"] = LC
            if Debug:
                if Wall == "LEW8SW":
                    if Node == "803":
                        print(LC)
                        print(LC_Database[LC][Wall][Node]["Fxx"], LC_Database[LC][Wall][Node]["Fyy"],LC_Database[LC][Wall][Node]["Fxy"])
                        for TH_temp in Thermal:
                            print(LC_Th_Database[LC][TH_temp][Wall][Node]["Fxx"], LC_Th_Database[LC][TH_temp][Wall][Node]["Fyy"],LC_Th_Database[LC][TH_temp][Wall][Node]["Fxy"])
                        input(Principle_DB[Wall][Node])


Text = {}

for Wall in Principle_DB:
    print(str(datetime.now() - Script_Start) + " : Combining principles forces for " + str(Wall))
    for Node in Principle_DB[Wall]:
        for i in Principle_DB[Wall][Node]:
            if not i in Text:
                Text[i] = ""
            if not i == "Env":
                Text[i] = Text[i] + str(Wall) + " " + str(Node) + " " + str(round(Principle_DB[Wall][Node][i]["Fp1"],3)) + " " + str(round(Principle_DB[Wall][Node][i]["Fp2"],3)) + "\n"
            else:
                Text[i] = Text[i] + str(Wall) + " " + str(Node) + " " + str(round(Principle_DB[Wall][Node][i]["Fp1"],3)) + " " + str(round(Principle_DB[Wall][Node][i]["Fp2"],3)) + " " + str(Principle_DB[Wall][Node][i]["Fp1_Res"]) + " " + str(Principle_DB[Wall][Node][i]["Fp2_Res"]) + " " + str(Principle_DB[Wall][Node][i]["Fp1_LC"]) + " " + str(Principle_DB[Wall][Node][i]["Fp2_LC"]) + "\n"

os.chdir(Script_Directory)

if Export_Txt:
    for out in Text:
        print(str(datetime.now() - Script_Start) + " : Preparing txt output for " + str(out))
        with open("Out_" + out + ".txt", "w" , encoding="utf-8") as Data_Out:
            Data_Out.write(Text[out])


                


X_Min = 10000000
X_Max = -10000000
Y_Min = 10000000
Y_Max = -10000000

for Wall in Coordinates:
    for Node in Coordinates[Wall]:
        if Coordinates[Wall][Node][PlotPlane.split("_")[0]] > X_Max:
            X_Max = Coordinates[Wall][Node][PlotPlane.split("_")[0]]
        if Coordinates[Wall][Node][PlotPlane.split("_")[1]] > Y_Max:
            Y_Max = Coordinates[Wall][Node][PlotPlane.split("_")[1]]
        if Coordinates[Wall][Node][PlotPlane.split("_")[0]] < X_Min:
            X_Min = Coordinates[Wall][Node][PlotPlane.split("_")[0]]
        if Coordinates[Wall][Node][PlotPlane.split("_")[1]] < Y_Min:
            Y_Min = Coordinates[Wall][Node][PlotPlane.split("_")[1]]

Plot_X_Min = (((X_Min / 1000) - ((X_Min % 1000) / 1000)) * 1000) - 500
Plot_X_Max = (((X_Max / 1000) - ((X_Max % 1000) / 1000)) * 1000) + 500

Ticks_Y = []
Ticks_X = []
i = Plot_X_Min + 500
while i < Plot_X_Max + 501:
    if Invert:
        Ticks_Y.append(i)
    else:
        Ticks_X.append(i)
    i = i + 1000

Plot_Y_Min = (((Y_Min / 1000) - ((Y_Min % 1000) / 1000)) * 1000) - 500
Plot_Y_Max = (((Y_Max / 1000) - ((Y_Max % 1000) / 1000)) * 1000) + 500


i = Plot_Y_Min + 500
while i < Plot_Y_Max + 501:
    if Invert:
        Ticks_X.append(i)
    else:
        Ticks_Y.append(i)
    i = i + 1000


Man_Legend = []
figure_number = 0



for Value in ("No_Th" + Thermal_Text + " Env").split():
    for i in "Fp1 Fp2".split():
        print(str(datetime.now() - Script_Start) + " : Preparing Plot for " + Value + " " + i)
        figure_number = figure_number + 1
        PLT.figure(figure_number)
        for Wall in Principle_DB:
            for Node in Principle_DB[Wall]:
                if Principle_DB[Wall][Node][Value][i] / Wall_Thickness > Max_Stress:
                    dot_color = "#ff0000"
                    dot_label  = "Stress > " + str(Max_Stress)
                elif Principle_DB[Wall][Node][Value][i] / Wall_Thickness > Max_Stress - (1 * (Max_Stress - Min_Stress) / 6):
                    dot_color = "#ff7600"
                    dot_label  = str(Max_Stress) + " >= Stress > " + str(round(Max_Stress - (1 * (Max_Stress - Min_Stress) / 6), 2))
                elif Principle_DB[Wall][Node][Value][i] / Wall_Thickness > Max_Stress - (2 * (Max_Stress - Min_Stress) / 6):
                    dot_color = "#ffbc00" 
                    dot_label  = str(round(Max_Stress - (1 * (Max_Stress - Min_Stress) / 6), 2)) + " >= Stress > " + str(round(Max_Stress - (2 * (Max_Stress - Min_Stress) / 6), 2))
                elif Principle_DB[Wall][Node][Value][i] / Wall_Thickness > Max_Stress - (3 * (Max_Stress - Min_Stress) / 6):
                    dot_color = "#ffe500" 
                    dot_label  = str(round(Max_Stress - (2 * (Max_Stress - Min_Stress) / 6), 2)) + " >= Stress > " + str(round(Max_Stress - (3 * (Max_Stress - Min_Stress) / 6), 2))
                elif Principle_DB[Wall][Node][Value][i] / Wall_Thickness > Max_Stress - (4 * (Max_Stress - Min_Stress) / 6):
                    dot_color = "#80c300" 
                    dot_label  = str(round(Max_Stress - (3 * (Max_Stress - Min_Stress) / 6), 2)) + " >= Stress > " + str(round(Max_Stress - (4 * (Max_Stress - Min_Stress) / 6), 2))
                elif Principle_DB[Wall][Node][Value][i] / Wall_Thickness > Max_Stress - (5 * (Max_Stress - Min_Stress) / 6):
                    dot_color = "#00a000" 
                    dot_label  = str(round(Max_Stress - (4 * (Max_Stress - Min_Stress) / 6), 2)) + " >= Stress > " + str(round(Max_Stress - (5 * (Max_Stress - Min_Stress) / 6), 2))
                elif Principle_DB[Wall][Node][Value][i] / Wall_Thickness > Max_Stress - (6 * (Max_Stress - Min_Stress) / 6):
                    dot_color = "#005080" 
                    dot_label  = str(round(Max_Stress - (5 * (Max_Stress - Min_Stress) / 6), 2)) + " >= Stress > " + str(round(Max_Stress - (6 * (Max_Stress - Min_Stress) / 6), 2))
                else:
                    dot_color = "#0000FF" 
                    dot_label  = str(round(Max_Stress - (6 * (Max_Stress - Min_Stress) / 6), 2)) + " >= Stress"


                if Invert:
                    PLT.scatter(Coordinates[Wall][Node][PlotPlane.split("_")[1]], Coordinates[Wall][Node][PlotPlane.split("_")[0]], marker = ",", label = dot_label, color = dot_color, s = 15*72./dpi, lw = 0)
                else:
                    PLT.scatter(Coordinates[Wall][Node][PlotPlane.split("_")[0]], Coordinates[Wall][Node][PlotPlane.split("_")[1]], marker = ",", label = dot_label, color = dot_color, s = 15*72./dpi, lw = 0)

        R01 = mpatches.Patch(color = "#ff0000", label = "Stress > " + str(Max_Stress))
        R02 = mpatches.Patch(color = "#ff7600", label = str(Max_Stress) + " >= Stress > " + str(round(Max_Stress - (1 * (Max_Stress - Min_Stress) / 6), 2)))
        R03 = mpatches.Patch(color = "#ffbc00", label = str(round(Max_Stress - (1 * (Max_Stress - Min_Stress) / 6), 2)) + " >= Stress > " + str(round(Max_Stress - (2 * (Max_Stress - Min_Stress) / 6), 2)))
        R04 = mpatches.Patch(color = "#ffe500", label = str(round(Max_Stress - (2 * (Max_Stress - Min_Stress) / 6), 2)) + " >= Stress > " + str(round(Max_Stress - (3 * (Max_Stress - Min_Stress) / 6), 2)))
        R05 = mpatches.Patch(color = "#80c300", label = str(round(Max_Stress - (3 * (Max_Stress - Min_Stress) / 6), 2)) + " >= Stress > " + str(round(Max_Stress - (4 * (Max_Stress - Min_Stress) / 6), 2)))
        R06 = mpatches.Patch(color = "#00a000", label = str(round(Max_Stress - (4 * (Max_Stress - Min_Stress) / 6), 2)) + " >= Stress > " + str(round(Max_Stress - (5 * (Max_Stress - Min_Stress) / 6), 2)))
        R07 = mpatches.Patch(color = "#005080", label = str(round(Max_Stress - (5 * (Max_Stress - Min_Stress) / 6), 2)) + " >= Stress > " + str(round(Max_Stress - (6 * (Max_Stress - Min_Stress) / 6), 2)))
        R08 = mpatches.Patch(color = "#0000FF", label = str(round(Max_Stress - (6 * (Max_Stress - Min_Stress) / 6), 2)) + " >= Stress")


        PLT.legend(handles = [R01, R02, R03, R04, R05, R06, R07, R08], loc="upper left", bbox_to_anchor=(1, 1))

        PLT.grid(True, linewidth = 0.1 * 300 / dpi, linestyle = "--")
        

        PLT.xlim(Plot_X_Min, Plot_X_Max)
        PLT.xticks(Ticks_X, rotation = 90, fontsize = 0.7 * 3000 / dpi)
        PLT.yticks(Ticks_Y, fontsize = 0.7 * 3000 / dpi)
        PLT.ylim(Plot_Y_Min, Plot_Y_Max)
        PLT.axis("equal")
        PLT.title(Value + " " + i)
        PLT.savefig(Value + "_" + i + ".png", dpi = dpi, bbox_inches='tight')     
            
print(str(datetime.now() - Script_Start) + " : Finished")
