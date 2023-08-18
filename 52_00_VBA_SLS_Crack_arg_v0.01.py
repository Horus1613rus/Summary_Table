"""
@author: Antl
"""
import os
import glob
import sys
import matplotlib.pyplot as PLT

if not len(sys.argv) == 3:
    print("52_00_VBA_... 0.2 input.txt")
    sys.exit()

Crack_Limit = float(sys.argv[1])
File = sys.argv[2]


PNG_Plot = True
Excel_Plot = True
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
with open(File, 'r' , encoding='utf-8') as Res_File_In:
    Res_File_Lines = Res_File_In.readlines()
    for Res_File_Line in Res_File_Lines:
        if not "Wall" in Res_File_Line:
            Wall = Res_File_Line.split()[0]
            Node = Res_File_Line.split()[1]
            if not Wall in Database:
                Database[Wall]= {}
                Coordinates[Wall]= {}
            if not Node in Database[Wall]:
                Database[Wall][Node]={}
                Coordinates[Wall][Node]= {}
            Database[Wall][Node]["wk+"] = float(Res_File_Line.split()[2])
            Database[Wall][Node]["wk-"] = float(Res_File_Line.split()[3])
            Database[Wall][Node]["wke"] = float(Res_File_Line.split()[4])

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
    
    for Crack in "wk+ wk- wke".split():
        Red_Values = 0
        Orange_Values = 0
        Yellow_Values = 0
        Green_Values = 0
        Gray_Values = 0
        Legend_Off = False
        print(str(datetime.now()-Script_Start) + " : Preparing Plots for " + Crack)
        if Crack == "wke":
            Crack_Limit = Crack_Limit/2
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
                    if Database[Wall][Node][Crack] >= 0.3:
                        Orange_Values = PLT.scatter(PNG_X,PNG_Y ,marker = "s" , color='#800000',s=5,label= "0.3 ≤ " + str(Crack) )

                    if Database[Wall][Node][Crack] < 0.3:
                        if Database[Wall][Node][Crack] >= 0.2:
                            Yellow_Values = PLT.scatter(PNG_X,PNG_Y ,marker = "s" , color='#ff0000',s=5,label="0.2 ≤ " + str(Crack) + " < 0.3")
                        else:
                            if Database[Wall][Node][Crack] >= 0.1:
                                Green_Values = PLT.scatter(PNG_X,PNG_Y ,marker = "s" , color='#ff9900',s=5,label="0.1 ≤ " + str(Crack) + " < 0.2")
                            else:
                                Gray_Values = PLT.scatter(PNG_X,PNG_Y ,marker = "s" , color='#cccccc',s=5,label="0.0 ≤ " + str(Crack) + " < 0.1")
                else:
                    if Database[Wall][Node][Crack] >= Crack_Limit:
                        Red_Values = PLT.scatter(PNG_X,PNG_Y ,marker = "s" , color='red',s=5,label= str(round(Crack_Limit,2)) + " ≤ " + str(Crack) )

                    if Database[Wall][Node][Crack] < Crack_Limit:
                        if Database[Wall][Node][Crack] >= 3*Crack_Limit/4:
                            Orange_Values = PLT.scatter(PNG_X,PNG_Y ,marker = "s" , color='orange',s=5,label=str(round(3*Crack_Limit/4,2)) + " ≤ " + str(Crack) + " < "+str(round(Crack_Limit,2)))
                        else:
                            if Database[Wall][Node][Crack] >= Crack_Limit/2:
                                Yellow_Values = PLT.scatter(PNG_X,PNG_Y ,marker = "s" , color='yellow',s=5,label=str(round(Crack_Limit/2,2)) + " ≤ " + str(Crack) + " < "+str(round(3*Crack_Limit/4,2)))
                            else:
                                if Database[Wall][Node][Crack] >= Crack_Limit/4:
                                    Green_Values = PLT.scatter(PNG_X,PNG_Y ,marker = "s" , color='green',s=5,label=str(round(Crack_Limit/4,2)) + " ≤ " + str(Crack) + " < "+str(round(Crack_Limit/2,2)))
                                else:
                                    Gray_Values = PLT.scatter(PNG_X,PNG_Y ,marker = "s" , color='gray',s=5,label= str(Crack) + " < "+str(round(Crack_Limit/4,2)))


        PLT.title(str(Crack) + " Envelope")
        try:
            Color_Info = PLT.legend(loc='upper left', bbox_to_anchor=(1, 0.5),handles=[Red_Values,Orange_Values,Yellow_Values,Green_Values,Gray_Values])
        except:
            try:
                Color_Info = PLT.legend(loc='upper left', bbox_to_anchor=(1, 0.5),handles=[Orange_Values,Yellow_Values,Green_Values,Gray_Values])
            except:
                try:
                    Color_Info = PLT.legend(loc='upper left', bbox_to_anchor=(1, 0.5),handles=[Yellow_Values,Green_Values,Gray_Values])
                except:
                    try:
                        Color_Info = PLT.legend(loc='upper left', bbox_to_anchor=(1, 0.5),handles=[Green_Values,Gray_Values])
                    except:
                        try:
                            Color_Info = PLT.legend(loc='upper left', bbox_to_anchor=(1, 0.5),handles=[Gray_Values])
                        except:
                            Legend_Off = True
        if not Legend_Off:
            PLT.savefig(os.getcwd() + "\\" +Plot_Folder + "\\" +str(Crack) + "_Env.png",bbox_extra_artists=(Color_Info,), bbox_inches='tight', dpi = 300)
        else:
            PLT.savefig(os.getcwd() + "\\" +Plot_Folder + "\\" +str(Crack) + "_Env.png", dpi = 300)

print(str(datetime.now()-Script_Start) + " : Done!")
