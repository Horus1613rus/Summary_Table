"""
@author: Antl
"""
import os
import glob
import sys
import matplotlib.pyplot as PLT
import matplotlib.patches as mpatches


Crack_Limit = 0.2
 
Only_Consider_These = []
#Only_Consider_These = ["LEWE1SE","LEWE2SE"]


PNG_Plot = True
dpi = 900
Dot_Coef = 4
Grid_Coef = 0.7
Invert = False
Excel_Plot = True
Txt_Plot = True
Rima = False


Additional_Settings_Switch = True
if Additional_Settings_Switch:
    Script_Path = os.path.abspath(__file__)
    Script_Directory = os.path.dirname(Script_Path)
    os.chdir(Script_Directory)
    Script_Ver = os.path.basename(__file__).split("_v")[1].split("_Egemen")[0]
    from datetime import datetime
    Script_Start = datetime.now()
    print(str(datetime.now()-Script_Start) + " : Script Start...")

    Coordinates_Folder = "00_Wall_Coordinates_with_DG"
    Plot_Folder = "52_00_PNG_Plots"

    Database = {}
    LC_List = []
    Wall_List = []
    Coordinates = {}

print(str(datetime.now()-Script_Start) + " : Gathering Results")
for Res_File in glob.iglob(os.getcwd() + '\\**\\*.res', recursive=True):
    with open(Res_File, 'r' , encoding='utf-8') as Res_File_In:
        Res_File_Lines = Res_File_In.readlines()
        Read_Switch_01 = False
        Read_Switch_02 = False
        for Res_File_Line in Res_File_Lines:
            if "LC:" in Res_File_Line:
                if not Res_File_Line.split("LC:")[1].split()[0].replace("T","").replace("W","").replace("Y","").replace("X","") in LC_List:
                    LC_List.append(Res_File_Line.split("LC:")[1].split()[0].replace("T","").replace("W","").replace("Y","").replace("X",""))
            if "Element :" in Res_File_Line:
                Wall = Res_File_Line.split("Element :")[1].split()[0]
                if not Wall in Database:
                    Database[Wall]={}
            if "Node :" in Res_File_Line:
                Node = Res_File_Line.split("Node :")[1].split()[0]
                if not Node in Database[Wall]:
                    Database[Wall][Node]={}
                    Database[Wall][Node]["wk+"]=0                    
                    Database[Wall][Node]["wk-"]=0
                    Database[Wall][Node]["wke"]=0
                    wk_top = 0
                    wk_bot = 0
            if Read_Switch_02:
                wk_bot = float(Res_File_Line.strip()[29:37])
                if wk_bot > Database[Wall][Node]["wk-"]:
                    Database[Wall][Node]["wk-"] = wk_bot
                if (wk_top + wk_bot) == 0:
                    wk_e = 0
                if not (wk_top + wk_bot) == 0:
                    wk_e = pow(abs(2*wk_top*wk_top*wk_bot*wk_bot / (wk_top + wk_bot)),1/3)
                    if wk_e > Database[Wall][Node]["wke"]:
                        Database[Wall][Node]["wke"] = wk_e
                Read_Switch_02 = False
            if Read_Switch_01:
                wk_top = float(Res_File_Line.strip()[29:37])
                if wk_top > Database[Wall][Node]["wk+"]:
                    Database[Wall][Node]["wk+"] = wk_top
                Read_Switch_01 = False
                Read_Switch_02 = True
            if "wok" in Res_File_Line:
                Read_Switch_01 = True

if not len(Only_Consider_These) == 0:
    Wall_List_Temp = []
    for Wall in Database:
        if not Wall in Only_Consider_These:
            Wall_List_Temp.append(Wall)
    for Wall in Wall_List_Temp:
        del Database[Wall]

if Txt_Plot:
    with open(os.getcwd() + '\\52_01_ConPro_Helper_SLS_Crack_v'+Script_Ver+".txt", 'w' , encoding='utf-8') as Data_Out:
        Data_Out.write( "Wall\tNode\twk+\twk-\twke\n" )
    for Wall in Database:
        for Node in Database[Wall]:
            with open(os.getcwd() + '\\52_01_ConPro_Helper_SLS_Crack_v'+Script_Ver+".txt", 'a' , encoding='utf-8') as Data_Out:
                Data_Out.write( Wall + "\t" + Node + "\t" + str(Database[Wall][Node]["wk+"]) + "\t" + str(Database[Wall][Node]["wk-"]) + "\t" + str(round(Database[Wall][Node]["wke"],3)) + "\n")

X_Value_List = []
Y_Value_List = []
Z_Value_List = []


if PNG_Plot:
    print(str(datetime.now()-Script_Start) + " : Gathering Coordinates")
    if not os.path.isdir(Plot_Folder):
        os.makedirs(Plot_Folder)
    for Wall in Database:
        X_Value_List = []
        Y_Value_List = []
        Z_Value_List = []
        Coordinates[Wall]={}
        with open(os.getcwd() + "\\" + Coordinates_Folder + "\\" +Wall +".dg.txt", 'r' , encoding='utf-8') as Coor_In:
            Coor_Lines = Coor_In.readlines()
            for Coor_Line in Coor_Lines:
                Node = Coor_Line.split()[1]
                Coordinates[Wall][Node]={}
                Coordinates[Wall][Node]["X"] = float(Coor_Line.split()[2])
                Coordinates[Wall][Node]["Y"] = float(Coor_Line.split()[3])
                Coordinates[Wall][Node]["Z"] = float(Coor_Line.split()[4])
                if not Coordinates[Wall][Node]["X"] in X_Value_List:
                    X_Value_List.append(Coordinates[Wall][Node]["X"])
                if not Coordinates[Wall][Node]["Y"] in X_Value_List:
                    Y_Value_List.append(Coordinates[Wall][Node]["Y"])
                if not Coordinates[Wall][Node]["Y"] in X_Value_List:
                    Z_Value_List.append(Coordinates[Wall][Node]["Z"])
    X_Value_List.sort()
    Y_Value_List.sort()
    Z_Value_List.sort()
    
    dif_X = X_Value_List[len(X_Value_List)-1]-X_Value_List[0]
    dif_Y = Y_Value_List[len(Y_Value_List)-1]-Y_Value_List[0]
    dif_Z = Z_Value_List[len(Z_Value_List)-1]-Z_Value_List[0]

    if dif_Z > dif_X or dif_Z > dif_Y:
        if dif_X > dif_Y:
            Plane = "XZ"
        if dif_Y > dif_X:
            Plane = "YZ"
    if dif_Z < dif_X and dif_Z < dif_Y:
        Plane = "XY"

    print(str(datetime.now()-Script_Start) + " : Calculated plane for element is " + Plane + ".")


    #fig, ax = PLT.subplots()
    #Color_Map = PLT.cm.get_cmap('nipy_spectral')
    X_Min = 10000000
    X_Max = -10000000
    Y_Min = 10000000
    Y_Max = -10000000
    for Wall in Database:
        for Node in Database[Wall]:
            if Plane == "XZ":
                PNG_X = Coordinates[Wall][Node]["X"]
                PNG_Y = Coordinates[Wall][Node]["Z"]
            if Plane == "YZ":
                PNG_X = Coordinates[Wall][Node]["Y"]
                PNG_Y = Coordinates[Wall][Node]["Z"]
            if Plane == "XY":
                PNG_X = Coordinates[Wall][Node]["X"]
                PNG_Y = Coordinates[Wall][Node]["Y"]
            if PNG_X > X_Max:
                X_Max = PNG_X
            if PNG_Y > Y_Max:
                Y_Max = PNG_Y
            if PNG_X < X_Min:
                X_Min = PNG_X
            if PNG_Y < Y_Min:
                Y_Min = PNG_Y
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
    



    for Crack in "wk+ wk- wke".split():
        Red_Values = 0
        Orange_Values = 0
        Yellow_Values = 0
        Green_Values = 0
        Gray_Values = 0
        Legend_Off = False
        
        if Crack == "wke":
            Crack_Limit = Crack_Limit/2

        R01 = mpatches.Patch(color = "red", label = str(round(Crack_Limit,2)) + " ≤ " + str(Crack))
        R02 = mpatches.Patch(color = "orange", label = str(round(3*Crack_Limit/4,2)) + " ≤ " + str(Crack) + " < "+str(round(Crack_Limit,2)))
        R03 = mpatches.Patch(color = "yellow", label = str(round(Crack_Limit/2,2)) + " ≤ " + str(Crack) + " < "+str(round(3*Crack_Limit/4,2)))
        R04 = mpatches.Patch(color = "green", label = str(round(Crack_Limit/4,2)) + " ≤ " + str(Crack) + " < "+str(round(Crack_Limit/2,2)))
        R05 = mpatches.Patch(color = "gray", label = str(Crack) + " < "+str(round(Crack_Limit/4,2)))
        
        R06 = mpatches.Patch(color = "#ff0000", label = "0.2 ≤ " + str(Crack))
        R07 = mpatches.Patch(color = "#ff9900", label = "0.1 ≤ " + str(Crack) + " < 0.2")
        R08 = mpatches.Patch(color = "#cccccc", label = "0.0 ≤ " + str(Crack) + " < 0.1")

        print(str(datetime.now()-Script_Start) + " : Preparing Plots for " + Crack)
        
        for Wall in Database:
            

            for Node in Database[Wall]:
                if Plane == "XZ":
                    PNG_X = Coordinates[Wall][Node]["X"]
                    PNG_Y = Coordinates[Wall][Node]["Z"]
                if Plane == "YZ":
                    PNG_X = Coordinates[Wall][Node]["Y"]
                    PNG_Y = Coordinates[Wall][Node]["Z"]
                if Plane == "XY":
                    PNG_X = Coordinates[Wall][Node]["X"]
                    PNG_Y = Coordinates[Wall][Node]["Y"]

                
                if Rima:
                    PLT.legend(handles = [R06, R07, R08],loc="upper left", bbox_to_anchor=(1, 1))
                    if Database[Wall][Node][Crack] >= 0.2:
                        Yellow_Values = PLT.scatter(PNG_X,PNG_Y ,marker = "," , color='#ff0000',s=Dot_Coef*72./dpi, lw = 0)
                    else:
                        if Database[Wall][Node][Crack] >= 0.1:
                            Green_Values = PLT.scatter(PNG_X,PNG_Y ,marker = "," , color='#ff9900',s=Dot_Coef*72./dpi, lw = 0)
                        else:
                            Gray_Values = PLT.scatter(PNG_X,PNG_Y ,marker = "," , color='#cccccc',s=Dot_Coef*72./dpi, lw = 0)
                else:
                    PLT.legend(handles = [R01, R02, R03, R04, R05],loc="upper left", bbox_to_anchor=(1, 1))
                    if Database[Wall][Node][Crack] >= Crack_Limit:
                        Red_Values = PLT.scatter(PNG_X,PNG_Y ,marker = "," , color='red',s=Dot_Coef*72./dpi, lw = 0)

                    if Database[Wall][Node][Crack] < Crack_Limit:
                        if Database[Wall][Node][Crack] >= 3*Crack_Limit/4:
                            Orange_Values = PLT.scatter(PNG_X,PNG_Y ,marker = "," , color='orange',s=Dot_Coef*72./dpi, lw = 0)
                        else:
                            if Database[Wall][Node][Crack] >= Crack_Limit/2:
                                Yellow_Values = PLT.scatter(PNG_X,PNG_Y ,marker = "," , color='yellow',s=Dot_Coef*72./dpi, lw = 0)
                            else:
                                if Database[Wall][Node][Crack] >= Crack_Limit/4:
                                    Green_Values = PLT.scatter(PNG_X,PNG_Y ,marker = "," , color='green',s=Dot_Coef*72./dpi, lw = 0)
                                else:
                                    Gray_Values = PLT.scatter(PNG_X,PNG_Y ,marker = "," , color='gray',s=Dot_Coef*72./dpi, lw = 0)


        PLT.title(str(Crack) + " Envelope")
        PLT.grid(True, linewidth = 0.1 * 300 / dpi, linestyle = "--")
        PLT.xlim(Plot_X_Min, Plot_X_Max)
        PLT.xticks(Ticks_X, rotation = 90, fontsize = Grid_Coef * 3000 / dpi)
        PLT.yticks(Ticks_Y, fontsize = Grid_Coef * 3000 / dpi)
        PLT.ylim(Plot_Y_Min, Plot_Y_Max)
        PLT.axis("equal")
        PLT.savefig(os.getcwd() + "\\" +Plot_Folder + "\\" +str(Crack) + "_Env.png", bbox_inches='tight', dpi = dpi)

print(str(datetime.now()-Script_Start) + " : Done!")
