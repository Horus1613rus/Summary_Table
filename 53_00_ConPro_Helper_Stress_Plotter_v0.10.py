"""
@author: Antl
"""
import os
import glob
import sys
import matplotlib.pyplot as PLT


Max_Stress = 100
Min_Stress = -100

# Max_Stress    <= Red
# Max_Stress/2  <= Yellow   <  Max_Stress
# Min_Stress/2  <  Green    <  Max_Stress/2
# Min_Stress    <  Blue     <= Min_Stress/2
#                  Purple   <= Min_Stress

PNG_Plot = True
Txt_Plot = True

Additional_Settings_Switch = True
if Additional_Settings_Switch:
    Script_Path = os.path.abspath(__file__)
    Script_Directory = os.path.dirname(Script_Path)
    os.chdir(Script_Directory)
    Script_Ver = os.path.basename(__file__).split("_v")[1].split("_Egemen")[0]
    from datetime import datetime
    Script_Start = datetime.now()
    print(str(datetime.now()-Script_Start) + " : Script Start...")

    Plot_Folder = "53_00_PNG_Plots"
    Coordinates_Folder = "00_Wall_Coordinates_with_DG"

    Database = {}
    Coordinates = {}
    LC_List = []

print(str(datetime.now()-Script_Start) + " : Gathering Results")
for Res_File in glob.iglob(os.getcwd() + '\\**\\*.res', recursive=True):
    with open(Res_File, 'r' , encoding='utf-8') as Res_File_In:
        Res_File_Lines = Res_File_In.readlines()
        Read_Switch = False
        for Res_File_Line in Res_File_Lines:
            #if "steel      :" in Res_File_Line:
                #Steel_Total_Layers = int(Res_File_Line.split()[2])
            if "LC:" in Res_File_Line:
                if not Res_File_Line.split("LC:")[1].split()[0].replace("T","").replace("W","").replace("Y","").replace("X","") in LC_List:
                    LC_List.append(Res_File_Line.split("LC:")[1].split()[0].replace("T","").replace("W","").replace("Y","").replace("X",""))
            if "Element :" in Res_File_Line:
                Wall = Res_File_Line.split("Element :")[1].split()[0]
                if not Wall in Database:
                    Database[Wall]= {}
                    Coordinates[Wall]= {}
            if "Node :" in Res_File_Line:
                Node = Res_File_Line.split("Node :")[1].split()[0]
                if not Node in Database[Wall]:
                    Database[Wall][Node]={}
                    Coordinates[Wall][Node]= {}
                    for i in range(1,5):
                        Database[Wall][Node]["L"+str(i)]=0

            if Read_Switch:
                try:
                    Layer = int(Res_File_Line.split()[0])
                    Stress = float(Res_File_Line.split()[2])
                    if abs(Stress) > abs(Database[Wall][Node]["L"+str(Layer)]):
                        Database[Wall][Node]["L" + str(Layer)] = Stress
                except:
                    Read_Switch = False
            if "layer      strain     stress    angle       zb         as" in Res_File_Line:
                Read_Switch = True


if Txt_Plot:
    with open(os.getcwd() + '\\53_01_ConPro_Helper_Stress_Plotter_v'+Script_Ver+".txt", 'w' , encoding='utf-8') as Data_Out:
        Data_Out.write( "Wall\tNode\tL1\tL2\tL3\tL4\n" )
    for Wall in Database:
        for Node in Database[Wall]:
            with open(os.getcwd() + '\\53_01_ConPro_Helper_Stress_Plotter_v'+Script_Ver+".txt", 'a' , encoding='utf-8') as Data_Out:
                Data_Out.write( Wall + "\t" + Node + "\t")
            for Layer in Database[Wall][Node]:
                with open(os.getcwd() + '\\53_01_ConPro_Helper_Stress_Plotter_v'+Script_Ver+".txt", 'a' , encoding='utf-8') as Data_Out:
                    Data_Out.write( str(Database[Wall][Node][Layer]) + "\t" )
            with open(os.getcwd() + '\\53_01_ConPro_Helper_Stress_Plotter_v'+Script_Ver+".txt", 'a' , encoding='utf-8') as Data_Out:
                Data_Out.write( "\n" )

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
    for Layer in "L1 L2 L3 L4".split():
        print(str(datetime.now()-Script_Start) + " : Preparing Stress Plots for " + Layer)
        Red_Values = 0
        Blue_Values = 0
        Yellow_Values = 0
        Green_Values = 0
        Purple_Values = 0
        Legend_Off = False
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

                if Database[Wall][Node][Layer] >= Max_Stress:
                    Red_Values = PLT.scatter(PNG_X,PNG_Y ,marker = "s" , color='red',s=5,label= str(Max_Stress) + " ≤ Stress")
                if Database[Wall][Node][Layer] < Max_Stress:
                    if Database[Wall][Node][Layer] >= Max_Stress/2:
                        Yellow_Values = PLT.scatter(PNG_X,PNG_Y ,marker = "s" , color='yellow',s=5,label=str(Max_Stress/2) + " ≤ Stress < "+str(Max_Stress))

                if Database[Wall][Node][Layer] <= Min_Stress:
                    Purple_Values = PLT.scatter(PNG_X,PNG_Y ,marker = "s" , color='purple',s=5,label=  "  Stress ≤ " + str(Min_Stress))
                if Database[Wall][Node][Layer] > Min_Stress:
                    if Database[Wall][Node][Layer] <= Min_Stress/2:
                        Blue_Values = PLT.scatter(PNG_X,PNG_Y ,marker = "s" , color='blue',s=5,label=str(Min_Stress) + " < Stress ≤ "+str(Min_Stress/2))

                if Database[Wall][Node][Layer] > Min_Stress/2:
                    if Database[Wall][Node][Layer] < Max_Stress/2:
                        Green_Values = PLT.scatter(PNG_X,PNG_Y ,marker = "s" , color='green',s=5,label=str(Min_Stress/2) + " < Stress < "+str(Max_Stress/2))


        PLT.title(str(Layer) + " Envelope Stress")
        try:
            Color_Info = PLT.legend(loc='upper left', bbox_to_anchor=(1, 0.5),handles=[Red_Values, Yellow_Values,Green_Values,Blue_Values,Purple_Values])
        except:
            try:
                Color_Info = PLT.legend(loc='upper left', bbox_to_anchor=(1, 0.5),handles=[Yellow_Values,Green_Values,Blue_Values,Purple_Values])
            except:
                try:
                    Color_Info = PLT.legend(loc='upper left', bbox_to_anchor=(1, 0.5),handles=[Red_Values, Yellow_Values,Green_Values,Blue_Values])
                except:
                    try:
                        Color_Info = PLT.legend(loc='upper left', bbox_to_anchor=(1, 0.5),handles=[ Yellow_Values,Green_Values,Blue_Values])
                    except:
                        try:
                            Color_Info = PLT.legend(loc='upper left', bbox_to_anchor=(1, 0.5),handles=[Green_Values,Blue_Values])
                        except:
                            try:
                                Color_Info = PLT.legend(loc='upper left', bbox_to_anchor=(1, 0.5),handles=[ Yellow_Values,Green_Values])
                            except:
                                try:
                                    Color_Info = PLT.legend(loc='upper left', bbox_to_anchor=(1, 0.5),handles=[ Green_Values])
                                except:
                                    Legend_Off = True
            

        if Legend_Off:
            PLT.savefig(os.getcwd() + "\\" +Plot_Folder + "\\" +str(Layer) + "_Env_Stress.png", dpi = 300)
        else:
            PLT.savefig(os.getcwd() + "\\" +Plot_Folder + "\\" +str(Layer) + "_Env_Stress.png",bbox_extra_artists=(Color_Info,), bbox_inches='tight', dpi = 300)


print(str(datetime.now()-Script_Start) + " : Done!")
