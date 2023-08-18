"""
@author: Antl
Date   :
"""
import os
import glob
import sys
import subprocess
import datetime
import xlwt
import matplotlib.pyplot as PLT

Override_Database_Ver = True                   # False Uses DGs Alias.dat
Database_Ver = '09.41d_1'                         # True Uses this database

Properties_DG = 33421                           # Representative Dg (for thickness Mat PT ref)
Override_Reinforcement = True ; Con_Cover = 50 # False representative DG rebar conf
Reinforcement_L1 = '1 25'                       # True Uses Rebar that set here
Reinforcement_L2 = '1 25'
Reinforcement_L3 = '1 25'
Reinforcement_L4 = '1 25'

DG_Wall = 'LEWESE'                             # Wall Name
Wall_Plane = 'XZ'                               # Wall Positioning in global axis "XZ" "YZ" "XY"
Wall_Plot_Partial_File = False                  # True for limited nodes defined at "51_01_Wall.txt"
                                                # False plots whole wall

Coordinate_Repair = False                       # Some wall's nodes not aligned
Coordinate_Repair_Range = 120                   # Set true for better alignement

Inspect_ULS_ALS = False                          # ULS/ALS/Shear Switch
Inspect_Shear = False                            # Shear Switch
State = 'ALS'
ULS_ALS_LC = '3604241800000'
ULS_ALS_Thermal = 3650

Inspect_Cracks = True                           # Crack Switch
SLS_LC = '5091116913310'
SLS_Thermal = 5999

Additional_Info = 'n'                           # Additional values
Load_Scale_Factor = '1 1 1 1 1 1'
Water_Pressure = 0.00                           # Water Pressure

Converge_Limit = 5.0                            # Above this percentage considered not converged

Settings_Switch = True
if Settings_Switch:
    #Additional Options
    Remove_Files = True                         # True gtforc res file junktion
    Plot_Nodes = True
    Crack_Per_Face = True
    Inspect_CPZ = False

    Wall_Node_File = '51_11_Wall.txt'

    ULS_Scaler = 1.5        #Plot Variable
    ULS_Power_Multiplier = 2#Plot Variable
    Shear_Scaler = 0.005    #Plot Variable
    Crack_Scaler = 5        #Plot Variable

    Ignore_Converge_Errors = [0]

    Plot_PNG = True
    Plot_Folder = "51_Plots"
    Shear_Limit_01 = 1414   #12f200c400         # Green Dot limit
    Shear_Limit_02 = 2827   #12f200c200         # Red Dot limit
    Wk_Limit_01 = 0.4       #Standard           # Green Dot limit
    Wk_Limit_02 = 0.6       #Standard           # Red Dot limit
    
    #Development
    Plot_Reinforcement_UR_Sheet_Env = True
    Debug_01 = False
    Debug_02 = False

    #if Inspect_Shear:
    #    if not Water_Pressure == 0:
    #        sys.exit("This version can not calculate water pressure for shear. Please set it to 0")

    #End
    Script_Path = os.path.abspath(__file__)
    Script_Directory = os.path.dirname(Script_Path)
    os.chdir(Script_Directory)
    Script_Ver = os.path.basename(__file__).split("_v")[1].split("_Egemen")[0]
    FNULL = open(os.devnull, 'w')

    from datetime import datetime
    Script_Start = datetime.now()

    Output_File = "51_19_DG"
    DG_Database_Location = "\\\\fs-saren\\07_DESIGN&ENGINEERING\\Project Documents to Review\\Calculation Note Documents\\All_Scripts&Tables\\_Data\\"
    Wall_Coordinates_with_Dg_Folder = "000_01_02_Wall_Coordinates_with_DG"
    Wall_Coordinates_with_Dg_Extention = ".dg.txt"

    DG_Database_DG_Sorted_Extention = ".dgsort"

    Additional_Extention = 'DG_Configuration'

    if Override_Reinforcement:
        Additional_Extention = Reinforcement_L1.split()[0] + 'f' + Reinforcement_L1.split()[1] + '_' + Reinforcement_L2.split()[0] + 'f' + Reinforcement_L2.split()[1] + '_' + Reinforcement_L3.split()[0] + 'f' + Reinforcement_L3.split()[1] + '_' + Reinforcement_L4.split()[0] + 'f' + Reinforcement_L4.split()[1]

    if State == 'ALS':
        Conf_File = 'prop\\consecna_1.'
        Conf_File_Shr = 'prop\\conshrna_1.'
    if State == 'ULS':
        Conf_File = 'prop\\consecnu_1.'
        Conf_File_Shr = 'prop\\conshrnu_1.'
    Conf_File_SLS = 'prop\\consecns_1.'
            
    Internal_Force_List = [0,1,2,3,4,5,6]
    for Ignore_Converge_Error in Ignore_Converge_Errors:
        Internal_Force_List.remove(Ignore_Converge_Error)
    
    Bold_Style = xlwt.easyxf('font: bold True')
    Red_Style = xlwt.easyxf('font: colour red')
    Bold_Red_Style = xlwt.easyxf('font: bold True, colour red')
    Coordinates_Style = xlwt.easyxf('font: bold True')

    Node_Database = {}
    ULS_ALS_Database = {}
    SLS_Database = {}
    Shear_Database = {}
    Other_Database = {}

#Ref. DG Check
Control_rinput_Count = 0
for root, dirs, files in os.walk('.'):
    Search_File = 'rinput_1.' + str(Properties_DG) + '.txt'
    if Search_File in files:
        Control_rinput_Count = Control_rinput_Count + 1
if not Control_rinput_Count == 1:
    sys.exit('Please Control DG ' + str(Properties_DG) + '...')
#if Inspect_ULS_ALS:
    #if not  int(ULS_ALS_LC) > 1000000000000:
        #sys.exit('Please Control LC...')
    #if not  int(ULS_ALS_LC) < 10000000000000:
        #sys.exit('Please Control LC...')
#if Inspect_Cracks:
    #if not  int(SLS_LC) > 1000000000000:
        #sys.exit('Please Control LC...')
    #if not  int(SLS_LC) < 10000000000000:
        #sys.exit('Please Control LC...')

#Alnovatek
if Override_Database_Ver:
    try:
        os.system('copy Z:\\Database\\alias\\alias.'+str(Database_Ver)+' alias.dat')
        print("Database set for ver " + str(Database_Ver))
    except:
        sys.exit('Problem copying alias.dat. Please check your connection.')

if not Override_Database_Ver:
    for Alias_File in glob.iglob(os.getcwd() + '\\**\\*' + str(Properties_DG)+'\\lower\\alias.dat', recursive=True):
        os.system('Copy "' + Alias_File + '" "' + Script_Directory + '"')

#Configure Section Properties
for DG_Conf_File in glob.iglob(os.getcwd() + '\\**\\*'+ str(Conf_File) + str(Properties_DG), recursive=True):
    with open(DG_Conf_File, 'r' , encoding='utf-8') as DG_Conf_File_In:
        DG_Conf_File_Lines = DG_Conf_File_In.readlines()
        with open(os.getcwd() + '\\51_99_DGConf', 'w',encoding='utf-8') as Data_Out:
            Data_Out.close()
        DG_Conf_File_Line_Number = 1
        Write_Switch = True
        for DG_Conf_File_Line in DG_Conf_File_Lines:
            if DG_Conf_File_Line_Number == 1:
                Ec = DG_Conf_File_Line.split()[2]
                Po = DG_Conf_File_Line.split()[4]
                hcon = float(DG_Conf_File_Line.split()[6])
            if Write_Switch:
                with open(os.getcwd() + '\\51_99_DGConf', 'a',encoding='utf-8') as Data_Out:
                    Data_Out.write(DG_Conf_File_Line.strip() + '\n')
            if not Write_Switch:
                Con_Cover_Cal = Con_Cover + 35
                if DG_Conf_File_Line_Number == 5:
                    Loc_L1 =  round((hcon / 2) - Con_Cover_Cal - (float(Reinforcement_L1.split()[1]) * 1.2 / 2 ),1)
                    Line_Text_Temp = Reinforcement_L1 + ' 200 ' + str(Loc_L1) + ' 0'
                if DG_Conf_File_Line_Number == 6:
                    Loc_L2 = round((hcon / 2) - Con_Cover_Cal - (float(Reinforcement_L1.split()[1])* 1.2) - (float(Reinforcement_L2.split()[1])* 1.2 / 2 ),1)
                    Line_Text_Temp = Reinforcement_L2 + ' 200 ' + str(Loc_L2) + ' 90'
                if DG_Conf_File_Line_Number == 7:
                    Loc_L3 = round(-1*((hcon / 2) - Con_Cover_Cal - (float(Reinforcement_L4.split()[1])* 1.2) - (float(Reinforcement_L3.split()[1])* 1.2 / 2 )),1)
                    Line_Text_Temp = Reinforcement_L3 + ' 200 ' + str(Loc_L3) + ' 90'
                if DG_Conf_File_Line_Number == 8:
                    Loc_L4 =  round(-1*((hcon / 2) - Con_Cover_Cal - (float(Reinforcement_L4.split()[1])* 1.2 / 2 )),1)
                    Line_Text_Temp = Reinforcement_L4 + ' 200 ' + str(Loc_L4) + ' 0'
                    Write_Switch = True
                with open(os.getcwd() + '\\51_99_DGConf', 'a',encoding='utf-8') as Data_Out:
                    Data_Out.write(str(Line_Text_Temp) + '\n')
            if DG_Conf_File_Line_Number == 4:
                P_T_Quantity = DG_Conf_File_Line.split()[2]
                Reinf_Quantity = DG_Conf_File_Line.split()[1]
                if int(Reinf_Quantity) > 4:
                    sys.exit('Only capable for 4 Layers of Reinforcement...')
            if Override_Reinforcement:
                if DG_Conf_File_Line_Number == 4:
                    Write_Switch  = False
            DG_Conf_File_Line_Number = DG_Conf_File_Line_Number + 1

for DG_Conf_File in glob.iglob(os.getcwd() + '\\**\\*'+ str(Conf_File_Shr) + str(Properties_DG), recursive=True):
    with open(DG_Conf_File, 'r' , encoding='utf-8') as DG_Conf_File_In:
        DG_Conf_File_Lines = DG_Conf_File_In.readlines()
        with open(os.getcwd() + '\\51_99_DGConf_Shr', 'w',encoding='utf-8') as Data_Out:
            Data_Out.close()
        Write_Switch = True
        DG_Conf_File_Line_Number = 1
        for DG_Conf_File_Line in DG_Conf_File_Lines:
            if Write_Switch:
                with open(os.getcwd() + '\\51_99_DGConf_Shr', 'a',encoding='utf-8') as Data_Out:
                    Data_Out.write(DG_Conf_File_Line.strip() + '\n')
            if not Write_Switch:
                if DG_Conf_File_Line_Number == 3:
                    Line_Text_Temp = Reinforcement_L1 + ' 200 ' + str(Loc_L1) + ' 0'
                if DG_Conf_File_Line_Number == 4:
                    Line_Text_Temp = Reinforcement_L2 + ' 200 ' + str(Loc_L2) + ' 90'
                if DG_Conf_File_Line_Number == 5:
                    Line_Text_Temp = Reinforcement_L3 + ' 200 ' + str(Loc_L3) + ' 90'
                if DG_Conf_File_Line_Number == 6:
                    Line_Text_Temp = Reinforcement_L4 + ' 200 ' + str(Loc_L4) + ' 0'
                    Write_Switch = True
                with open(os.getcwd() + '\\51_99_DGConf_Shr', 'a',encoding='utf-8') as Data_Out:
                    Data_Out.write(str(Line_Text_Temp) + '\n')
            if Override_Reinforcement:
                if DG_Conf_File_Line_Number == 2:
                    Write_Switch  = False
            DG_Conf_File_Line_Number = DG_Conf_File_Line_Number + 1

for DG_Conf_File in glob.iglob(os.getcwd() + '\\**\\*'+ str(Conf_File_SLS) + str(Properties_DG), recursive=True):
    with open(DG_Conf_File, 'r' , encoding='utf-8') as DG_Conf_File_In:
        DG_Conf_File_Lines = DG_Conf_File_In.readlines()
        with open(os.getcwd() + '\\51_99_DGConf_SLS', 'w',encoding='utf-8') as Data_Out:
            Data_Out.close()
        DG_Conf_File_Line_Number = 1
        Write_Switch = True
        for DG_Conf_File_Line in DG_Conf_File_Lines:
            if Write_Switch:
                with open(os.getcwd() + '\\51_99_DGConf_SLS', 'a',encoding='utf-8') as Data_Out:
                    Data_Out.write(DG_Conf_File_Line.strip() + '\n')
            if not Write_Switch:
                if DG_Conf_File_Line_Number == 5:
                    Line_Text_Temp = Reinforcement_L1 + ' 200 ' + str(Loc_L1+25) + ' 0'
                if DG_Conf_File_Line_Number == 6:
                    Line_Text_Temp = Reinforcement_L2 + ' 200 ' + str(Loc_L2+25) + ' 90'
                if DG_Conf_File_Line_Number == 7:
                    Line_Text_Temp = Reinforcement_L3 + ' 200 ' + str(Loc_L3-25) + ' 90'
                if DG_Conf_File_Line_Number == 8:
                    Line_Text_Temp = Reinforcement_L4 + ' 200 ' + str(Loc_L4-25) + ' 0'
                    Write_Switch = True
                with open(os.getcwd() + '\\51_99_DGConf_SLS', 'a',encoding='utf-8') as Data_Out:
                    Data_Out.write(str(Line_Text_Temp) + '\n')
            if Override_Reinforcement:
                if DG_Conf_File_Line_Number == 4:
                    Write_Switch  = False
            DG_Conf_File_Line_Number = DG_Conf_File_Line_Number + 1


#Nodes
X_Vary_List = []
Y_Vary_List = []
Z_Vary_List = []
with open(os.getcwd() + '\\51_99_Wall', 'w' , encoding='utf-8') as Wall_Node_File_Out:
    Wall_Node_File_Out.close()

if not Wall_Plot_Partial_File:
    with open(os.getcwd() + '\\'+ Wall_Node_File, 'w' , encoding='utf-8') as Wall_Node_File_Out:
        Wall_Node_File_Out.close()
    with open(DG_Database_Location + Wall_Coordinates_with_Dg_Folder + '\\'+ DG_Wall + Wall_Coordinates_with_Dg_Extention, 'r' , encoding='utf-8') as Wall_Node_File_In:
        Wall_Node_File_Lines = Wall_Node_File_In.readlines()
        for Wall_Node_File_Line in Wall_Node_File_Lines:
            with open(os.getcwd() + '\\'+ Wall_Node_File, 'a' , encoding='utf-8') as Wall_Node_File_Out:
                Wall_Node_File_Out.write(Wall_Node_File_Line)

with open(Wall_Node_File, 'r' , encoding='utf-8') as Wall_Node_File_In:
    Wall_Node_File_Lines = Wall_Node_File_In.readlines()
    if Coordinate_Repair:
        X_Repair_List = []
        Y_Repair_List = []
        Z_Repair_List = []
        for Wall_Node_File_Line in Wall_Node_File_Lines:
            DG_Wall = Wall_Node_File_Line.split()[0]
            Node = Wall_Node_File_Line.split()[1]
            Node_List = []
            if not Node in Node_List:
                Node_List.append(Node)
                for WallRaw_File in glob.iglob(DG_Database_Location+ '\\' + Wall_Coordinates_with_Dg_Folder +'\\*'+str(DG_Wall) + '*dg.txt', recursive=True):
                    with open(WallRaw_File, 'r' , encoding='utf-8') as WallRaw_File_In:
                        WallRaw_File_Lines = WallRaw_File_In.readlines()
                        for WallRaw_File_Line in WallRaw_File_Lines:
                            if str(DG_Wall) == WallRaw_File_Line.split()[0]:
                                if str(Node) == WallRaw_File_Line.split()[1]:
                                    X = int(float(WallRaw_File_Line.split()[2]))
                                    Y = int(float(WallRaw_File_Line.split()[3]))
                                    Z = int(float(WallRaw_File_Line.split()[4]))
                                    X_Repair_List.append(X)
                                    Y_Repair_List.append(Y)
                                    Z_Repair_List.append(Z)
                                    X_Repair_List.sort()
                                    Y_Repair_List.sort()
                                    Z_Repair_List.sort()
        for X in X_Repair_List:
            X_Sum = 0
            Repair_List = []
            for X_Check in X_Repair_List:
                if abs(X_Check - X) < Coordinate_Repair_Range:
                    X_Sum = X_Sum + X_Check
                    Repair_List.append(X_Check)
            X_Average = X_Sum / len(Repair_List)
            for Repair in Repair_List:
                X_Repair_List.remove(Repair)
            X_Repair_List.append(X_Average)
            X_Repair_List.sort()
        for Y in Y_Repair_List:
            Y_Sum = 0
            Repair_List = []
            for Y_Check in Y_Repair_List:
                if abs(Y_Check - Y) < Coordinate_Repair_Range:
                    Y_Sum = Y_Sum + Y_Check
                    Repair_List.append(Y_Check)
            Y_Average = Y_Sum / len(Repair_List)
            for Repair in Repair_List:
                Y_Repair_List.remove(Repair)
            Y_Repair_List.append(Y_Average)
            Y_Repair_List.sort()
        for Z in Z_Repair_List:
            Z_Sum = 0
            Repair_List = []
            for Z_Check in Z_Repair_List:
                if abs(Z_Check - Z) < Coordinate_Repair_Range:
                    Z_Sum = Z_Sum + Z_Check
                    Repair_List.append(Z_Check)
            Z_Average = Z_Sum / len(Repair_List)
            for Repair in Repair_List:
                Z_Repair_List.remove(Repair)
            Z_Repair_List.append(Z_Average)
            Z_Repair_List.sort()
        for Wall_Node_File_Line in Wall_Node_File_Lines:
            DG_Wall = Wall_Node_File_Line.split()[0]
            Node = Wall_Node_File_Line.split()[1]
            Node_List = []
            if not Node in Node_List:
                Node_List.append(Node)
                for WallRaw_File in glob.iglob(DG_Database_Location+ '\\' + Wall_Coordinates_with_Dg_Folder +'\\*'+str(DG_Wall) + '*dg.txt', recursive=True):
                    with open(WallRaw_File, 'r' , encoding='utf-8') as WallRaw_File_In:
                        WallRaw_File_Lines = WallRaw_File_In.readlines()
                        for WallRaw_File_Line in WallRaw_File_Lines:
                            if str(DG_Wall) == WallRaw_File_Line.split()[0]:
                                if str(Node) == WallRaw_File_Line.split()[1]:
                                    X = int(float(WallRaw_File_Line.split()[2]))
                                    for Control in X_Repair_List:
                                        if abs(Control - X) <Coordinate_Repair_Range:
                                            X = Control
                                    Y = int(float(WallRaw_File_Line.split()[3]))
                                    for Control in Y_Repair_List:
                                        if abs(Control - Y) <Coordinate_Repair_Range:
                                            Y = Control
                                    Z = int(float(WallRaw_File_Line.split()[4]))
                                    for Control in Z_Repair_List:
                                        if abs(Control - Z) <Coordinate_Repair_Range:
                                            Z = Control
                                    with open(os.getcwd() + '\\51_99_Wall', 'a' , encoding='utf-8') as Wall_Node_File_Out:
                                        Wall_Node_File_Out.write(str(DG_Wall) + ' ' + str(Node) + ' ' + str(X) + ' ' + str(Y) + ' ' + str(Z) + '\n' )
    if not Coordinate_Repair:
        for Wall_Node_File_Line in Wall_Node_File_Lines:
            DG_Wall = Wall_Node_File_Line.split()[0]
            Node = Wall_Node_File_Line.split()[1]
            Node_List = []
            if not Node in Node_List:
                Node_List.append(Node)
                for WallRaw_File in glob.iglob(DG_Database_Location+ '\\' + Wall_Coordinates_with_Dg_Folder +'\\*'+str(DG_Wall) + '*dg.txt', recursive=True):
                    with open(WallRaw_File, 'r' , encoding='utf-8') as WallRaw_File_In:
                        WallRaw_File_Lines = WallRaw_File_In.readlines()
                        for WallRaw_File_Line in WallRaw_File_Lines:
                            if str(DG_Wall) == WallRaw_File_Line.split()[0]:
                                if str(Node) == WallRaw_File_Line.split()[1]:
                                    X = int(float(WallRaw_File_Line.split()[2]))
                                    Y = int(float(WallRaw_File_Line.split()[3]))
                                    Z = int(float(WallRaw_File_Line.split()[4]))
                                    with open(os.getcwd() + '\\51_99_Wall', 'a' , encoding='utf-8') as Wall_Node_File_Out:
                                        Wall_Node_File_Out.write(str(DG_Wall) + ' ' + str(Node) + ' ' + str(X) + ' ' + str(Y) + ' ' + str(Z) + '\n' )

with open(os.getcwd() + '\\51_99_Wall', 'r' , encoding='utf-8') as DGSort_File_In:
    DGSort_File_Lines = DGSort_File_In.readlines()
    Node_Quatitiy = len(DGSort_File_Lines)
    for DGSort_File_Line in DGSort_File_Lines:
        Node_Database[DGSort_File_Line.strip().split()[1]]={}
        Node_Database[DGSort_File_Line.strip().split()[1]]['X'] = DGSort_File_Line.strip().split()[2]
        Node_Database[DGSort_File_Line.strip().split()[1]]['Y'] = DGSort_File_Line.strip().split()[3]
        Node_Database[DGSort_File_Line.strip().split()[1]]['Z'] = DGSort_File_Line.strip().split()[4]
        if not float(DGSort_File_Line.strip().split()[2]) in X_Vary_List:
            X_Vary_List.append(float(DGSort_File_Line.strip().split()[2]))
            X_Vary_List.sort()
        if not float(DGSort_File_Line.strip().split()[3]) in Y_Vary_List:
            Y_Vary_List.append(float(DGSort_File_Line.strip().split()[3]))
            Y_Vary_List.sort()
        if not float(DGSort_File_Line.strip().split()[4]) in Z_Vary_List:
            Z_Vary_List.append(float(DGSort_File_Line.strip().split()[4]))
            Z_Vary_List.sort()
        with open(os.getcwd() + '\\51_99_gtForce', 'a',encoding='utf-8') as Data_Out:
            Data_Out.write(DGSort_File_Line.split()[0]+' '+DGSort_File_Line.split()[1] + '\n')

if Wall_Plane == 'XY':
    Excel_X = int(len(X_Vary_List)) #WallX
    Excel_X_Values = X_Vary_List
    Excel_Y = int(len(Y_Vary_List)) #WallY
    Excel_Y_Values = Y_Vary_List
if Wall_Plane == 'XZ':
    Excel_X = int(len(X_Vary_List)) #WallX
    Excel_X_Values = X_Vary_List
    Excel_Y = int(len(Z_Vary_List)) #WallZ
    Excel_Y_Values = Z_Vary_List
if Wall_Plane == 'YZ':
    Excel_X = int(len(Y_Vary_List)) #WallY
    Excel_X_Values = Y_Vary_List
    Excel_Y = int(len(Z_Vary_List)) #WallZ
    Excel_Y_Values = Z_Vary_List

if Excel_X>255:
    for Value_Temp in '51_99_DGConf alias.dat 51_99_Wall 51_99_DGConf_Shr 51_99_DGConf_SLS 51_99_gtForce'.split():
        os.remove(Value_Temp)
    sys.exit( DG_Wall + ' is too wide. Please split the wall.')

#Internal Loads and ConSec
if Inspect_ULS_ALS:
    print(str(datetime.now()-Script_Start) + " : Preparing Loads...")
    with open(os.getcwd() + '\\51_Ncredb4_Run', 'w' , encoding='utf-8') as Data_Out:
        Data_Out.write('fbas gtforc 51_99_gtForce ' + str(ULS_ALS_LC) + ' 51_99_gtForce.' + str(ULS_ALS_LC) +' 1\nfbas gtforc 51_99_gtForce ' + str(ULS_ALS_Thermal) + ' 51_99_gtForce.' + str(ULS_ALS_Thermal) +' 1')
    Ncredb4_Input = open(os.getcwd() + '\\51_Ncredb4_Run')
    Ncredb4_Process = subprocess.Popen('Ncredb4', stdout=FNULL, stdin=Ncredb4_Input)
    Ncredb4_Process.wait()
    Ncredb4_Input.close()
    
    with open(os.getcwd() + '\\51_99_Loads.LC', 'w' , encoding='utf-8') as Data_Out:
        Data_Out.close()
    for i in range(Node_Quatitiy):
        with open(os.getcwd() + '\\51_99_gtForce.'+ str(ULS_ALS_LC), 'r' , encoding='utf-8') as Force_File_In:
            Force_File_Lines = Force_File_In.readlines()
            Line_LC = Force_File_Lines[i].strip()
        with open(os.getcwd() + '\\51_99_gtForce.'+ str(ULS_ALS_Thermal), 'r' , encoding='utf-8') as Force_File_In:
            Force_File_Lines = Force_File_In.readlines()
            Line_T = Force_File_Lines[i].strip()                
        with open(os.getcwd() + '\\51_99_Loads.LC', 'a' , encoding='utf-8') as Data_Out:
            Data_Out.write(Line_LC + '\n' + Line_T + '\n')
    #Consec
    print(str(datetime.now()-Script_Start) + " : Consec Run ULS/ALS...")
    with open(os.getcwd() + '\\51_99_Main', 'w' , encoding='utf-8') as Data_Out:
        Data_Out.write('Automated\n51_99_DGConf\n51_99_Loads.LC\n51_99_Results.LC\n1\nchk\ndnv FC2D\n' + str(Ec) + ' ' + str(Po) + '\n'+Additional_Info+'\n'+Load_Scale_Factor+'\n0 0 0 0 0 0 0 0\n'+str(Water_Pressure))
    ConSectInput = open(os.getcwd() + '\\51_99_Main')
    ConSect_Process = subprocess.Popen('nconsectdnv', stdout=FNULL, stdin=ConSectInput)
    ConSect_Process.wait()
    ConSectInput.close()

    with open(os.getcwd() + '\\51_99_Results.LC', 'r' , encoding='utf-8') as Data_In:
        Data_Lines = Data_In.readlines()
        Read_Switch = False
        Read_Switch_Converge = False
        Node_Appear_List = []
        Force_Line = 1
        Steel_Layer = 1
        for Data_Line in Data_Lines:                    
            if 'Node :' in Data_Line:
                Node = Data_Line.split('Node :')[1].strip().split()[0]
                if not Node in Node_Appear_List:
                    ULS_ALS_Database[Node]={}
                    Shear_Database[Node]={}
                    Other_Database[Node]={}
                    Node_Appear_List.append(Node)
        for Data_Line in Data_Lines:
            if 'LC:' in Data_Line:
                Load_Prefix = Data_Line.split('LC:')[1].split()[0].strip().replace(str(ULS_ALS_LC),'_LC0')
            if 'Node :' in Data_Line:
                Node = Data_Line.split('Node :')[1].strip().split()[0]
                ULS_ALS_Database[Node][str(Load_Prefix)]={}
                Shear_Database[Node][str(Load_Prefix)]={}
                Other_Database[Node]['ULS_ALS_ENV'] = 0
                Other_Database[Node]['ULS_ALS_Converge_Ratio'] = 0
                for Steel_Layer in range(1,5):
                    Other_Database[Node]['ENV_L' + str(Steel_Layer)] = 0

        for Data_Line in Data_Lines:
            if 'LC:' in Data_Line:
                Load_Prefix = Data_Line.split('LC:')[1].split()[0].strip().replace(str(ULS_ALS_LC),'_LC0')
            if 'Node :' in Data_Line:
                Node = Data_Line.split('Node :')[1].strip().split()[0]
            if not 'T' == Load_Prefix.strip()[:1]:
                if not 'Y' == Load_Prefix.strip()[:1]:
                    if Read_Switch_Converge:
                        if Force_Line in Internal_Force_List:
                            if abs(float(Data_Line.split()[4])) > abs(Other_Database[Node]['ULS_ALS_Converge_Ratio']):
                                Other_Database[Node]['ULS_ALS_Converge_Ratio'] = float(Data_Line.split()[4])
                        if Force_Line == 6:
                            Read_Switch_Converge = False
                        Force_Line = Force_Line + 1
                    
                    if 'given      found     diff    diff' in Data_Line:
                        Read_Switch_Converge = True
                        Force_Line = 1
            
            if Read_Switch:
                ULS_ALS_Database[Node][str(Load_Prefix)]['S_L' + str(Steel_Layer)] = float( Data_Line.strip().split()[1] )
                if abs(Other_Database[Node]['ULS_ALS_ENV']) < abs( ULS_ALS_Database[Node][str(Load_Prefix)]['S_L' + str(Steel_Layer)] ):
                    Other_Database[Node]['ULS_ALS_ENV'] = float( ULS_ALS_Database[Node][str(Load_Prefix)]['S_L' + str(Steel_Layer)] )
                if abs(Other_Database[Node]['ENV_L' + str(Steel_Layer)]) < abs( ULS_ALS_Database[Node][str(Load_Prefix)]['S_L' + str(Steel_Layer)] ):
                    Other_Database[Node]['ENV_L' + str(Steel_Layer)] = float( ULS_ALS_Database[Node][str(Load_Prefix)]['S_L' + str(Steel_Layer)] )
                if Steel_Layer == 4:
                    Read_Switch = False
                Steel_Layer = Steel_Layer + 1
                
            if 'layer      strain     stress    angle       zb         as' in Data_Line:
                Read_Switch = True
                Steel_Layer = 1

    if Inspect_Shear:
        with open(os.getcwd() + '\\51_99_LoadsShr.LC', 'w' , encoding='utf-8') as Data_Out:
            Data_Out.close()
        for i in range(Node_Quatitiy):
            with open(os.getcwd() + '\\51_99_gtForce.'+ str(ULS_ALS_LC), 'r' , encoding='utf-8') as Force_File_In:
                Force_File_Lines = Force_File_In.readlines()
                Node = Force_File_Lines[i].strip().split()[1]
                LC_Text = '_LC0'
                Fxx = round(float(Force_File_Lines[i].strip().split()[3]),2)
                Fyy = round(float(Force_File_Lines[i].strip().split()[4]),2)
                Fxy = round(float(Force_File_Lines[i].strip().split()[5]),2)
                Mxx = round(float(Force_File_Lines[i].strip().split()[6]),2)
                Myy = round(float(Force_File_Lines[i].strip().split()[7]),2)
                Mxy = round(float(Force_File_Lines[i].strip().split()[8]),2)
                Vxz = round(float(Force_File_Lines[i].strip().split()[9]),2)
                Vyz = round(float(Force_File_Lines[i].strip().split()[10]),2)
                UR_Env = round(float(Other_Database[Node]['ULS_ALS_ENV']),2)
                if Other_Database[Node]['ULS_ALS_ENV'] > 2.5:
                    UR_Env = 2.5
            with open(os.getcwd() + '\\51_99_gtForce.'+ str(ULS_ALS_Thermal), 'r' , encoding='utf-8') as Force_File_In:
                Force_File_Lines = Force_File_In.readlines()
                LC_Text_Th = 'T_LC0'
                Fxx_Th = round(float(Force_File_Lines[i].strip().split()[3]),2)*0.5 + Fxx
                Fyy_Th = round(float(Force_File_Lines[i].strip().split()[4]),2)*0.5 + Fyy
                Fxy_Th = round(float(Force_File_Lines[i].strip().split()[5]),2)*0.5 + Fxy
                Mxx_Th = round(float(Force_File_Lines[i].strip().split()[6]),2)*0.5 + Mxx
                Myy_Th = round(float(Force_File_Lines[i].strip().split()[7]),2)*0.5 + Myy
                Mxy_Th = round(float(Force_File_Lines[i].strip().split()[8]),2)*0.5 + Mxy
                Vxz_Th = round(float(Force_File_Lines[i].strip().split()[9]),2)*0.5 + Vxz
                Vyz_Th = round(float(Force_File_Lines[i].strip().split()[10]),2)*0.5 + Vyz                    
                UR_Th_Env = round(float(Other_Database[Node]['ULS_ALS_ENV']),2)
                if Other_Database[Node]['ULS_ALS_ENV'] > 2.5:
                    UR_Th_Env = 2.5
            with open(os.getcwd() + '\\51_99_LoadsShr.LC', 'a' , encoding='utf-8') as Data_Out:
                Data_Out.write(
                    DG_Wall + ' ' + Node + ' ' + LC_Text + ' ' +
                    str(round(Fxx,2)) + ' ' + str(round(Fyy,2)) + ' ' + str(round(Fxy,2)) + ' ' + str(round(Mxx,2)) + ' ' + str(round(Myy,2)) + ' ' + str(round(Mxy,2)) + ' ' + str(round(Vxz,2)) + ' ' + str(round(Vyz,2)) + ' ' + str(Water_Pressure) + ' ' + str(UR_Env) +
                    '\n' + 
                    DG_Wall + ' ' + Node + ' ' + LC_Text_Th + ' ' +
                    str(round(Fxx_Th,2)) + ' ' + str(round(Fyy_Th,2)) + ' ' + str(round(Fxy_Th,2)) + ' ' + str(round(Mxx_Th,2)) + ' ' + str(round(Myy_Th,2)) + ' ' + str(round(Mxy_Th,2)) + ' ' + str(round(Vxz_Th,2)) + ' ' + str(round(Vyz_Th,2)) + ' ' + str(Water_Pressure) + ' ' + str(UR_Th_Env) +
                    '\n')

        subprocess.run('nconshr_DNV 51_99_DGConf_Shr 51_99_LoadsShr.LC 51_99_ResultsShr.LC', input='n', encoding='utf-8')

        with open(os.getcwd() + '\\51_99_ResultsShr.LC', 'r' , encoding='utf-8') as Data_In:
            Data_Lines = Data_In.readlines()
            for Data_Line in Data_Lines:
                if 'LC:' in Data_Line:
                    Load_Prefix = Data_Line.split('LC:')[1].split()[0].strip() + '0'
                    if Data_Line.split('LC:')[1].split()[0].strip() =="W":
                        Load_Prefix = "W_LC0"
                        if Data_Line.split('LC:')[1].split()[1].strip() == "T_LC":
                            Load_Prefix = "Y_LC0"
                            
                if 'Node :' in Data_Line:
                    Node = Data_Line.split('Node :')[1].strip().split()[0]
                try:
                    Shear_Database[Node][str(Load_Prefix)]['Asv'] = 0
                except:
                    pass
                Other_Database[Node]['Shr_Style'] = 0
                Other_Database[Node]['Shear_ENV'] = 0
            LC_Read_Switch = True

            for Data_Line in Data_Lines:
                if 'Table of overall results' in Data_Line:
                    LC_Read_Switch = False
                if LC_Read_Switch:
                    if 'LC:' in Data_Line:
                        Load_Prefix = Data_Line.split('LC:')[1].split()[0].strip() + '0'
                        if Data_Line.split('LC:')[1].split()[0].strip() =="W":
                            Load_Prefix = "W_LC0"
                            if Data_Line.split('LC:')[1].split()[1].strip() == "T_LC":
                                Load_Prefix = "Y_LC0"
                    if 'Node :' in Data_Line:
                        Node = Data_Line.split('Node :')[1].strip().split()[0]
                    if '** Maximum required stirrup area is' in Data_Line:
                        try:
                            if float(Data_Line.strip()[35:45]) > float(Shear_Database[Node][str(Load_Prefix)]['Asv']):
                                Shear_Database[Node][str(Load_Prefix)]['Asv'] = int(Data_Line.strip()[35:45].replace('.',''))
                        except:
                            pass
                    if Other_Database[Node]['ULS_ALS_ENV'] >= 2.5 :
                        if 'WARNING ' in Data_Line:
                            Other_Database[Node]['Shear_ENV'] = 1
                            try:
                                if float(Data_Line.strip().split('asvm =')[1]) > float(Shear_Database[Node][str(Load_Prefix)]['Asv']):
                                    Shear_Database[Node][str(Load_Prefix)]['Asv'] = int(Data_Line.strip().split('asvm =')[1].replace('.',''))
                            except:
                                pass
                    try:
                        if float(Shear_Database[Node][str(Load_Prefix)]['Asv']) > float(Other_Database[Node]['Shear_ENV']):
                            Other_Database[Node]['Shear_ENV'] = int(Shear_Database[Node][str(Load_Prefix)]['Asv'])
                    except:
                        pass


if Inspect_Cracks:
    print(str(datetime.now()-Script_Start) + " : Preparing Loads...")
    with open(os.getcwd() + '\\51_Ncredb4_Run', 'w' , encoding='utf-8') as Data_Out:
        Data_Out.write('fbas gtforc 51_99_gtForce ' + str(SLS_LC) + ' 51_99_gtForce.' + str(SLS_LC) +' 1\nfbas gtforc 51_99_gtForce ' + str(SLS_Thermal) + ' 51_99_gtForce.' + str(SLS_Thermal) +' 1')
    Ncredb4_Input = open(os.getcwd() + '\\51_Ncredb4_Run')
    Ncredb4_Process = subprocess.Popen('Ncredb4', stdout=FNULL, stdin=Ncredb4_Input)
    Ncredb4_Process.wait()
    Ncredb4_Input.close()

    with open(os.getcwd() + '\\51_99_Loads_SLS.LC', 'w' , encoding='utf-8') as Data_Out:
        Data_Out.close()
    for i in range(Node_Quatitiy):
        with open(os.getcwd() + '\\51_99_gtForce.'+ str(SLS_LC), 'r' , encoding='utf-8') as Force_File_In:
            Force_File_Lines = Force_File_In.readlines()
            Line_LC = Force_File_Lines[i].strip()
        with open(os.getcwd() + '\\51_99_gtForce.'+ str(SLS_Thermal), 'r' , encoding='utf-8') as Force_File_In:
            Force_File_Lines = Force_File_In.readlines()
            Line_T = Force_File_Lines[i].strip()                
        with open(os.getcwd() + '\\51_99_Loads_SLS.LC', 'a' , encoding='utf-8') as Data_Out:
            Data_Out.write(Line_LC + '\n' + Line_T + '\n')
#SLS Consec
    print(str(datetime.now()-Script_Start) + " : Consec Run SLS...")
    with open(os.getcwd() + '\\51_99_Main', 'w' , encoding='utf-8') as Data_Out:
        Data_Out.write('Automated\n51_99_DGConf_SLS\n51_99_Loads_SLS.LC\n51_99_Results_SLS.LC\n1\nsls long\ndnv\n' + str(Ec) + ' ' + str(Po) + '\n'+Additional_Info+'\n'+Load_Scale_Factor+'\n0 0 0 0 0 0 0 0\n'+str(Water_Pressure))
    ConSectInput = open(os.getcwd() + '\\51_99_Main')
    ConSect_Process = subprocess.Popen('nconsectdnv', stdout=FNULL, stdin=ConSectInput)
    ConSect_Process.wait()
    ConSectInput.close()
#SLS Results
    with open(os.getcwd() + '\\51_99_Results_SLS.LC', 'r' , encoding='utf-8') as Data_In:
        Data_Lines = Data_In.readlines()
        Read_Switch_Top = False
        Read_Switch_Bot = False
        Node_Appear_List = []
        wk_top=0
        wk_bot=0
        for Data_Line in Data_Lines:                    
            if 'Node :' in Data_Line:
                Node = Data_Line.split('Node :')[1].strip().split()[0]
                if not Node in Node_Appear_List:
                    SLS_Database[Node]={}                    
                    SLS_Database[Node]['CPZ'] = 10000
                    for Value_Temp in 'Wk+ Wk- Wk_ENV Wk2_ENV Wke'.split():
                        SLS_Database[Node][Value_Temp] = 0
                    Node_Appear_List.append(Node)
        for Data_Line in Data_Lines:
            if 'LC:' in Data_Line:
                Load_Prefix = Data_Line.split('LC:')[1].split()[0].strip().replace(str(ULS_ALS_LC),'_LC0')
            if 'Node :' in Data_Line:
                Node = Data_Line.split('Node :')[1].strip().split()[0]
            if 'compression depth =' in Data_Line:
                if  float(Data_Line.split('compression depth =')[1].split('.mm')[0]) < SLS_Database[Node]['CPZ']:
                    SLS_Database[Node]['CPZ'] = float(Data_Line.split('compression depth =')[1].split('.mm')[0])
            if Read_Switch_Bot:
                wk_bot = float(Data_Line.strip()[29:37])
                if  float(Data_Line.strip()[29:37]) > SLS_Database[Node]['Wk-']:
                    SLS_Database[Node]['Wk-'] = float(Data_Line.strip()[29:37])
                if SLS_Database[Node]['Wk-'] > SLS_Database[Node]['Wk_ENV']:
                    SLS_Database[Node]['Wk_ENV'] = SLS_Database[Node]['Wk-']
                Read_Switch_Bot = False
                if (wk_top + wk_bot) == 0:
                    wke=0
                if not (wk_top + wk_bot) == 0:
                    wke= pow(2*wk_top*wk_top*wk_bot*wk_bot / (wk_top + wk_bot),1/3)
                if wke > SLS_Database[Node]['Wke']:
                    SLS_Database[Node]['Wke'] = wke
            if Read_Switch_Top:
                wk_top = float(Data_Line.strip()[29:37])
                if  float(Data_Line.strip()[29:37]) > SLS_Database[Node]['Wk+']:
                    SLS_Database[Node]['Wk+'] = float(Data_Line.strip()[29:37])
                if SLS_Database[Node]['Wk+'] > SLS_Database[Node]['Wk_ENV']:
                    SLS_Database[Node]['Wk_ENV'] = SLS_Database[Node]['Wk+']
                Read_Switch_Top = False
                Read_Switch_Bot = True
            if 'wok' in Data_Line:
                Read_Switch_Top = True
            
for Value_Temp in '51_Ncredb4_Run alias.dat 51_99_Wall 51_99_gtForce'.split():
    os.remove(Value_Temp)
if Remove_Files:
    for Value_Temp in '51_99_Main 51_99_DGConf 51_99_DGConf_Shr 51_99_DGConf_SLS'.split():
        os.remove(Value_Temp)
    if Inspect_ULS_ALS or Inspect_Shear:
        os.remove('51_99_gtForce.'+ str(ULS_ALS_LC))
        for Value_Temp in '51_99_Loads.LC 51_99_Results.LC 51_99_Results.LC.s51'.split():
            os.remove(Value_Temp)
        os.remove('51_99_gtForce.'+ str(ULS_ALS_Thermal))
        if Inspect_Shear:
            os.remove('51_99_LoadsShr.LC')
            os.remove('51_99_ResultsShr.LC')
    if Inspect_Cracks:
        os.remove('51_99_gtForce.'+ str(SLS_LC))
        for Value_Temp in '51_99_Loads_SLS.LC 51_99_Results_SLS.LC 51_99_Results_SLS.LC.s51'.split():
            os.remove(Value_Temp)
        if not ULS_ALS_Thermal == SLS_Thermal:
            os.remove('51_99_gtForce.'+ str(SLS_Thermal))

print(str(datetime.now()-Script_Start) + " : Preparing Excels...")
Worksheet = {}
Worksheet_SLS = {}
if Inspect_ULS_ALS:    
    if not Inspect_Shear:
        Workbook = xlwt.Workbook(Output_File + str(Properties_DG) + '_' +str(DG_Wall) + "_" +  str(ULS_ALS_LC) +"_ULS_ALS_"+Additional_Extention+".xls")
    if Inspect_Shear:
        Workbook = xlwt.Workbook(Output_File + str(Properties_DG) + '_' +str(DG_Wall) + "_" +  str(ULS_ALS_LC) +"_ULS_ALS_Shear_"+Additional_Extention+".xls")
    if Plot_Nodes:
        Worksheet['Nodes'] = Workbook.add_sheet('Nodes',cell_overwrite_ok=True)
        Worksheet['Nodes'].write(0,0,'Plane '+ Wall_Plane)
        for X in Excel_X_Values:
            Worksheet['Nodes'].write(0,Excel_X_Values.index(X)+1,X,Coordinates_Style)
        for Y in Excel_Y_Values:
            Worksheet['Nodes'].write(len(Excel_Y_Values)-Excel_Y_Values.index(Y),0,Y,Coordinates_Style)
    if Inspect_ULS_ALS:
        for L in range(1,5):
            Worksheet['Layer '+ str(L)] = Workbook.add_sheet('Layer '+ str(L),cell_overwrite_ok=True)
            Worksheet['Layer '+ str(L)].write(0,0,'Plane '+ Wall_Plane)
            for X in Excel_X_Values:
                Worksheet['Layer '+ str(L)].write(0,Excel_X_Values.index(X)+1,X,Coordinates_Style)
            for Y in Excel_Y_Values:
                Worksheet['Layer '+ str(L)].write(len(Excel_Y_Values)-Excel_Y_Values.index(Y),0,Y,Coordinates_Style)
    if Inspect_Shear:
        Worksheet['Shear'] = Workbook.add_sheet('Shear',cell_overwrite_ok=True)
        Worksheet['Shear'].write(0,0,'Plane '+ Wall_Plane)
        for X in Excel_X_Values:
            Worksheet['Shear'].write(0,Excel_X_Values.index(X)+1,X,Coordinates_Style)
        for Y in Excel_Y_Values:
            Worksheet['Shear'].write(len(Excel_Y_Values)-Excel_Y_Values.index(Y),0,Y,Coordinates_Style)

    for Node in Node_Database:
        if Wall_Plane == 'XY':
            Excel_X_Index = float(Node_Database[Node]['X'])
            Excel_Y_Index = float(Node_Database[Node]['Y'])
        if Wall_Plane == 'XZ':
            Excel_X_Index = float(Node_Database[Node]['X'])
            Excel_Y_Index = float(Node_Database[Node]['Z'])
        if Wall_Plane == 'YZ':
            Excel_X_Index = float(Node_Database[Node]['Y'])
            Excel_Y_Index = float(Node_Database[Node]['Z'])
        if Plot_Nodes:
            Worksheet['Nodes'].write(len(Excel_Y_Values)-Excel_Y_Values.index(Excel_Y_Index),Excel_X_Values.index(Excel_X_Index)+1,str(Node))
        if Inspect_ULS_ALS:
            for Steel_Layer in range (1,5):
                if abs(float(Other_Database[Node]['ULS_ALS_Converge_Ratio'])) >= float(Converge_Limit):
                    Worksheet['Layer '+ str(Steel_Layer)].write(len(Excel_Y_Values)-Excel_Y_Values.index(Excel_Y_Index),Excel_X_Values.index(Excel_X_Index)+1,round(float(Other_Database[Node]['ENV_L' + str(Steel_Layer)]),2),style = Bold_Style)
                if abs(float(Other_Database[Node]['ULS_ALS_Converge_Ratio'])) < float(Converge_Limit):
                    Worksheet['Layer '+ str(Steel_Layer)].write(len(Excel_Y_Values)-Excel_Y_Values.index(Excel_Y_Index),Excel_X_Values.index(Excel_X_Index)+1,round(float(Other_Database[Node]['ENV_L' + str(Steel_Layer)]),2))
        if Inspect_Shear:
            if abs(float(Other_Database[Node]['ULS_ALS_Converge_Ratio'])) >= float(Converge_Limit):
                if abs(Other_Database[Node]['ULS_ALS_ENV']) >= 2.5:
                    Worksheet['Shear'].write(len(Excel_Y_Values)-Excel_Y_Values.index(Excel_Y_Index),Excel_X_Values.index(Excel_X_Index)+1,int(Other_Database[Node]['Shear_ENV']),style = Bold_Red_Style)
                if abs(Other_Database[Node]['ULS_ALS_ENV']) < 2.5:
                    Worksheet['Shear'].write(len(Excel_Y_Values)-Excel_Y_Values.index(Excel_Y_Index),Excel_X_Values.index(Excel_X_Index)+1,int(Other_Database[Node]['Shear_ENV']),style = Red_Style)
            if abs(float(Other_Database[Node]['ULS_ALS_Converge_Ratio'])) < float(Converge_Limit):
                if abs(Other_Database[Node]['ULS_ALS_ENV']) >= 2.5:
                    Worksheet['Shear'].write(len(Excel_Y_Values)-Excel_Y_Values.index(Excel_Y_Index),Excel_X_Values.index(Excel_X_Index)+1,int(Other_Database[Node]['Shear_ENV']),style = Bold_Style)
                if abs(Other_Database[Node]['ULS_ALS_ENV']) < 2.5:
                    Worksheet['Shear'].write(len(Excel_Y_Values)-Excel_Y_Values.index(Excel_Y_Index),Excel_X_Values.index(Excel_X_Index)+1,int(Other_Database[Node]['Shear_ENV']))
    if Inspect_Shear:
        Workbook.save(Output_File + str(Properties_DG) + '_' +str(DG_Wall) + "_" +  str(ULS_ALS_LC) +"_ULS_ALS_Shear_"+Additional_Extention+".xls")
    if not Inspect_Shear:
        Workbook.save(Output_File + str(Properties_DG) + '_' +str(DG_Wall) + "_" +  str(ULS_ALS_LC) +"_ULS_ALS_"+Additional_Extention+".xls")
    


if Inspect_Cracks:
    Workbook_SLS = xlwt.Workbook(Output_File + str(Properties_DG) + '_' +str(DG_Wall) + "_" +  str(SLS_LC) +"_SLS_"+Additional_Extention+".xls")
    if Plot_Nodes:
        Worksheet_SLS['Nodes'] = Workbook_SLS.add_sheet('Nodes',cell_overwrite_ok=True)
        Worksheet_SLS['Nodes'].write(0,0,'Plane '+ Wall_Plane)
        for X in Excel_X_Values:
            Worksheet_SLS['Nodes'].write(0,Excel_X_Values.index(X)+1,X,Coordinates_Style)
        for Y in Excel_Y_Values:
            Worksheet_SLS['Nodes'].write(len(Excel_Y_Values)-Excel_Y_Values.index(Y),0,Y,Coordinates_Style)
    if Crack_Per_Face:
        Worksheet_SLS['Wk+'] = Workbook_SLS.add_sheet('Wk+',cell_overwrite_ok=True)
        Worksheet_SLS['Wk+'].write(0,0,'Plane '+ Wall_Plane)
        Worksheet_SLS['Wk-'] = Workbook_SLS.add_sheet('Wk-',cell_overwrite_ok=True)
        Worksheet_SLS['Wk-'].write(0,0,'Plane '+ Wall_Plane)
        Worksheet_SLS['WKe'] = Workbook_SLS.add_sheet('WKe',cell_overwrite_ok=True)
        Worksheet_SLS['WKe'].write(0,0,'Plane '+ Wall_Plane)
    if not Crack_Per_Face:
        Worksheet_SLS['Wk_ENV'] = Workbook_SLS.add_sheet('Wk_ENV',cell_overwrite_ok=True)
        Worksheet_SLS['Wk_ENV'].write(0,0,'Plane '+ Wall_Plane)
    if Inspect_CPZ:
        Worksheet_SLS['CPZ'] = Workbook_SLS.add_sheet('CPZ',cell_overwrite_ok=True)
        Worksheet_SLS['CPZ'].write(0,0,'Plane '+ Wall_Plane)

    for X in Excel_X_Values:
        if Crack_Per_Face:
            Worksheet_SLS['Wk+'].write(0,Excel_X_Values.index(X)+1,X,Coordinates_Style)
            Worksheet_SLS['Wk-'].write(0,Excel_X_Values.index(X)+1,X,Coordinates_Style)
            Worksheet_SLS['WKe'].write(0,Excel_X_Values.index(X)+1,X,Coordinates_Style)
        if not Crack_Per_Face:
            Worksheet_SLS['Wk_ENV'].write(0,Excel_X_Values.index(X)+1,X,Coordinates_Style)
        if Inspect_CPZ:
            Worksheet_SLS['CPZ'].write(0,Excel_X_Values.index(X)+1,X,Coordinates_Style)
    for Y in Excel_Y_Values:
        if Crack_Per_Face:
            Worksheet_SLS['Wk+'].write(len(Excel_Y_Values)-Excel_Y_Values.index(Y),0,Y,Coordinates_Style)
            Worksheet_SLS['Wk-'].write(len(Excel_Y_Values)-Excel_Y_Values.index(Y),0,Y,Coordinates_Style)
            Worksheet_SLS['WKe'].write(len(Excel_Y_Values)-Excel_Y_Values.index(Y),0,Y,Coordinates_Style)
        if not Crack_Per_Face:
            Worksheet_SLS['Wk_ENV'].write(len(Excel_Y_Values)-Excel_Y_Values.index(Y),0,Y,Coordinates_Style)
        if Inspect_CPZ:
            Worksheet_SLS['CPZ'].write(len(Excel_Y_Values)-Excel_Y_Values.index(Y),0,Y,Coordinates_Style)


    for Node in Node_Database:
        if Wall_Plane == 'XY':
            Excel_X_Index = float(Node_Database[Node]['X'])
            Excel_Y_Index = float(Node_Database[Node]['Y'])
        if Wall_Plane == 'XZ':
            Excel_X_Index = float(Node_Database[Node]['X'])
            Excel_Y_Index = float(Node_Database[Node]['Z'])
        if Wall_Plane == 'YZ':
            Excel_X_Index = float(Node_Database[Node]['Y'])
            Excel_Y_Index = float(Node_Database[Node]['Z'])
        if Plot_Nodes:
            Worksheet_SLS['Nodes'].write(len(Excel_Y_Values)-Excel_Y_Values.index(Excel_Y_Index),Excel_X_Values.index(Excel_X_Index)+1,str(Node))
        if Crack_Per_Face:
            Worksheet_SLS['Wk+'].write(len(Excel_Y_Values)-Excel_Y_Values.index(Excel_Y_Index),Excel_X_Values.index(Excel_X_Index)+1,float(SLS_Database[Node]['Wk+']))
            Worksheet_SLS['Wk-'].write(len(Excel_Y_Values)-Excel_Y_Values.index(Excel_Y_Index),Excel_X_Values.index(Excel_X_Index)+1,float(SLS_Database[Node]['Wk-']))
            Worksheet_SLS['WKe'].write(len(Excel_Y_Values)-Excel_Y_Values.index(Excel_Y_Index),Excel_X_Values.index(Excel_X_Index)+1,float(SLS_Database[Node]['Wke']))
        if not Crack_Per_Face:
            Worksheet_SLS['Wk_ENV'].write(len(Excel_Y_Values)-Excel_Y_Values.index(Excel_Y_Index),Excel_X_Values.index(Excel_X_Index)+1,float(SLS_Database[Node]['Wk_ENV']))
        if Inspect_CPZ:
            Worksheet_SLS['CPZ'].write(len(Excel_Y_Values)-Excel_Y_Values.index(Excel_Y_Index),Excel_X_Values.index(Excel_X_Index)+1,int(SLS_Database[Node]['CPZ']))

    Workbook_SLS.save(Output_File + str(Properties_DG) + '_' +str(DG_Wall) + "_" +  str(SLS_LC) +"_SLS_"+Additional_Extention+".xls")

print(str(datetime.now()-Script_Start) + " : Preparing PNGs...")
if Plot_PNG:
    if not os.path.isdir(Plot_Folder):
        os.makedirs(Plot_Folder)    
    if Inspect_ULS_ALS:
        PLT.figure(1)
        for Node in Node_Database:
            if Wall_Plane == 'XY':
                Excel_X_Index = float(Node_Database[Node]['X'])
                Excel_Y_Index = float(Node_Database[Node]['Y'])
            if Wall_Plane == 'XZ':
                Excel_X_Index = float(Node_Database[Node]['X'])
                Excel_Y_Index = float(Node_Database[Node]['Z'])
            if Wall_Plane == 'YZ':
                Excel_X_Index = float(Node_Database[Node]['Y'])
                Excel_Y_Index = float(Node_Database[Node]['Z'])

            if abs(Other_Database[Node]['ULS_ALS_ENV']) >= 2.5:
                if abs(float(Other_Database[Node]['ULS_ALS_Converge_Ratio'])) >= float(Converge_Limit):
                    PLT.scatter(Excel_X_Index,Excel_Y_Index ,marker = "s" , color='red',label='Error',s=ULS_Scaler**(ULS_Power_Multiplier*abs(Other_Database[Node]['ULS_ALS_ENV'])))
                if not abs(float(Other_Database[Node]['ULS_ALS_Converge_Ratio'])) >= float(Converge_Limit):
                    PLT.scatter(Excel_X_Index,Excel_Y_Index ,marker = "s" , color='green',label='Non-Linear',s=ULS_Scaler**(ULS_Power_Multiplier*abs(Other_Database[Node]['ULS_ALS_ENV'])))
            if not abs(Other_Database[Node]['ULS_ALS_ENV']) >= 2.5:
                PLT.scatter(Excel_X_Index,Excel_Y_Index ,marker = "s" , color='gray',label='Linear',s=ULS_Scaler**(ULS_Power_Multiplier*abs(Other_Database[Node]['ULS_ALS_ENV'])))
        PLT.title(str(Properties_DG) + ' ' +str(DG_Wall) + " " +  str(ULS_ALS_LC) + " " +str(ULS_ALS_Thermal) + " ULS/ALS Convergence")
        PLT.savefig(os.getcwd() + "\\" + Plot_Folder + "\\" +Output_File + str(Properties_DG) + '_' +str(DG_Wall) + "_" +  str(ULS_ALS_LC) +"_ULS_ALS_Convergence"+Additional_Extention+".png",dpi = 300)

    if Inspect_Shear:
        PLT.figure(2)
        for Node in Node_Database:
            if Wall_Plane == 'XY':
                Excel_X_Index = float(Node_Database[Node]['X'])
                Excel_Y_Index = float(Node_Database[Node]['Y'])
            if Wall_Plane == 'XZ':
                Excel_X_Index = float(Node_Database[Node]['X'])
                Excel_Y_Index = float(Node_Database[Node]['Z'])
            if Wall_Plane == 'YZ':
                Excel_X_Index = float(Node_Database[Node]['Y'])
                Excel_Y_Index = float(Node_Database[Node]['Z'])

            if int(Other_Database[Node]['Shear_ENV']) < Shear_Limit_01:
                if abs(float(Other_Database[Node]['ULS_ALS_Converge_Ratio'])) >= float(Converge_Limit):
                    PLT.scatter(Excel_X_Index,Excel_Y_Index ,marker = "s" , color='gray',label='Zero',s=Shear_Scaler*(Other_Database[Node]['Shear_ENV']))
                if not abs(float(Other_Database[Node]['ULS_ALS_Converge_Ratio'])) >= float(Converge_Limit):
                    PLT.scatter(Excel_X_Index,Excel_Y_Index ,marker = "s" , color='gray',label='Zero',s=Shear_Scaler*(Other_Database[Node]['Shear_ENV']))
            if not int(Other_Database[Node]['Shear_ENV']) < Shear_Limit_01:
                if int(Other_Database[Node]['Shear_ENV']) < Shear_Limit_02:
                    if abs(float(Other_Database[Node]['ULS_ALS_Converge_Ratio'])) >= float(Converge_Limit):
                        PLT.scatter(Excel_X_Index,Excel_Y_Index ,marker = "s" , color='green',label='12f200/400',s=Shear_Scaler*(Other_Database[Node]['Shear_ENV']))
                    if not abs(float(Other_Database[Node]['ULS_ALS_Converge_Ratio'])) >= float(Converge_Limit):
                        PLT.scatter(Excel_X_Index,Excel_Y_Index ,marker = "s" , color='green',label='12f200/400',s=Shear_Scaler*(Other_Database[Node]['Shear_ENV']))
                if not int(Other_Database[Node]['Shear_ENV']) < Shear_Limit_02:
                    if abs(float(Other_Database[Node]['ULS_ALS_Converge_Ratio'])) >= float(Converge_Limit):
                        PLT.scatter(Excel_X_Index,Excel_Y_Index ,marker = "s" , color='red',label='12f200/200',s=Shear_Scaler*(Other_Database[Node]['Shear_ENV']))
                    if not abs(float(Other_Database[Node]['ULS_ALS_Converge_Ratio'])) >= float(Converge_Limit):
                        PLT.scatter(Excel_X_Index,Excel_Y_Index ,marker = "s" , color='red',label='12f200/200',s=Shear_Scaler*(Other_Database[Node]['Shear_ENV']))
        PLT.title(str(Properties_DG) + ' ' +str(DG_Wall) + " " +  str(ULS_ALS_LC) + " " +str(ULS_ALS_Thermal) + " Shear")
        PLT.savefig(os.getcwd() + "\\" + Plot_Folder + "\\" +Output_File + str(Properties_DG) + '_' +str(DG_Wall) + "_" +  str(ULS_ALS_LC) +"_Shear_"+Additional_Extention+".png",dpi = 300)

    if Inspect_Cracks:
        Crck_Plt = PLT.figure(3)
        if not Crack_Per_Face:
            for Node in Node_Database:
                if Wall_Plane == 'XY':
                    Excel_X_Index = float(Node_Database[Node]['X'])
                    Excel_Y_Index = float(Node_Database[Node]['Y'])
                if Wall_Plane == 'XZ':
                    Excel_X_Index = float(Node_Database[Node]['X'])
                    Excel_Y_Index = float(Node_Database[Node]['Z'])
                if Wall_Plane == 'YZ':
                    Excel_X_Index = float(Node_Database[Node]['Y'])
                    Excel_Y_Index = float(Node_Database[Node]['Z'])
                    PLT.figure(3)
                    if float(SLS_Database[Node]['Wk_ENV']) < Wk_Limit_01:
                        PLT.scatter(Excel_X_Index,Excel_Y_Index ,marker = "s" , color='gray',label='Zero',s=Crack_Scaler*float(SLS_Database[Node]['Wk_ENV']))
                    if not float(SLS_Database[Node]['Wk_ENV']) < Wk_Limit_01:
                        if float(SLS_Database[Node]['Wk_ENV']) < Wk_Limit_02:
                            PLT.scatter(Excel_X_Index,Excel_Y_Index ,marker = "s" , color='green',label='0.4',s=Crack_Scaler*float(SLS_Database[Node]['Wk_ENV']))
                        if not float(SLS_Database[Node]['Wk_ENV']) < Wk_Limit_02:
                            PLT.scatter(Excel_X_Index,Excel_Y_Index ,marker = "s" , color='red',label='0.6',s=Crack_Scaler*float(SLS_Database[Node]['Wk_ENV']))
            PLT.title(str(Properties_DG) + ' ' +str(DG_Wall) + " " +  str(SLS_LC) + " " +str(SLS_Thermal) + " Wk_ENV")
            PLT.savefig(os.getcwd() + "\\" + Plot_Folder + "\\" +Output_File + str(Properties_DG) + '_' +str(DG_Wall) + "_" +  str(SLS_LC) +"Wk_ENV_"+Additional_Extention+".png",dpi = 300)

        if Crack_Per_Face:
            for Node in Node_Database:
                if Wall_Plane == 'XY':
                    Excel_X_Index = float(Node_Database[Node]['X'])
                    Excel_Y_Index = float(Node_Database[Node]['Y'])
                if Wall_Plane == 'XZ':
                    Excel_X_Index = float(Node_Database[Node]['X'])
                    Excel_Y_Index = float(Node_Database[Node]['Z'])
                if Wall_Plane == 'YZ':
                    Excel_X_Index = float(Node_Database[Node]['Y'])
                    Excel_Y_Index = float(Node_Database[Node]['Z'])
                PLT.figure(3)
                if float(SLS_Database[Node]['Wk+']) < Wk_Limit_01:
                    PLT.scatter(Excel_X_Index,Excel_Y_Index ,marker = "s" , color='gray',label='Zero',s=Crack_Scaler*float(SLS_Database[Node]['Wk+']))
                if not float(SLS_Database[Node]['Wk+']) < Wk_Limit_01:
                    if float(SLS_Database[Node]['Wk+']) < Wk_Limit_02:
                        PLT.scatter(Excel_X_Index,Excel_Y_Index ,marker = "s" , color='green',label='0.4',s=Crack_Scaler*float(SLS_Database[Node]['Wk+']))
                    if not float(SLS_Database[Node]['Wk+']) < Wk_Limit_02:
                        PLT.scatter(Excel_X_Index,Excel_Y_Index ,marker = "s" , color='red',label='0.6',s=Crack_Scaler*float(SLS_Database[Node]['Wk+']))
            PLT.title(str(Properties_DG) + ' ' +str(DG_Wall) + " " +  str(SLS_LC) + " " +str(SLS_Thermal) + " WK+")
            PLT.savefig(os.getcwd() + "\\" + Plot_Folder + "\\" +Output_File + str(Properties_DG) + '_' +str(DG_Wall) + "_" +  str(SLS_LC) +"_WK+_"+Additional_Extention+".png",dpi = 300)

        if Crack_Per_Face:
            for Node in Node_Database:
                if Wall_Plane == 'XY':
                    Excel_X_Index = float(Node_Database[Node]['X'])
                    Excel_Y_Index = float(Node_Database[Node]['Y'])
                if Wall_Plane == 'XZ':
                    Excel_X_Index = float(Node_Database[Node]['X'])
                    Excel_Y_Index = float(Node_Database[Node]['Z'])
                if Wall_Plane == 'YZ':
                    Excel_X_Index = float(Node_Database[Node]['Y'])
                    Excel_Y_Index = float(Node_Database[Node]['Z'])
                PLT.figure(4)
                if float(SLS_Database[Node]['Wk-']) < Wk_Limit_01:
                    PLT.scatter(Excel_X_Index,Excel_Y_Index ,marker = "s" , color='gray',label='Zero',s=Crack_Scaler*float(SLS_Database[Node]['Wk-']))
                if not float(SLS_Database[Node]['Wk-']) < Wk_Limit_01:
                    if float(SLS_Database[Node]['Wk-']) < Wk_Limit_02:
                        PLT.scatter(Excel_X_Index,Excel_Y_Index ,marker = "s" , color='green',label='0.4',s=Crack_Scaler*float(SLS_Database[Node]['Wk-']))
                    if not float(SLS_Database[Node]['Wk-']) < Wk_Limit_02:
                        PLT.scatter(Excel_X_Index,Excel_Y_Index ,marker = "s" , color='red',label='0.6',s=Crack_Scaler*float(SLS_Database[Node]['Wk-']))
            PLT.title(str(Properties_DG) + ' ' +str(DG_Wall) + " " +  str(SLS_LC) + " " +str(SLS_Thermal) + " WK-")
            PLT.savefig(os.getcwd() + "\\" + Plot_Folder + "\\" +Output_File + str(Properties_DG) + '_' +str(DG_Wall) + "_" +  str(SLS_LC) +"_WK-_"+Additional_Extention+".png",dpi = 300)
            
            for Node in Node_Database:
                if Wall_Plane == 'XY':
                    Excel_X_Index = float(Node_Database[Node]['X'])
                    Excel_Y_Index = float(Node_Database[Node]['Y'])
                if Wall_Plane == 'XZ':
                    Excel_X_Index = float(Node_Database[Node]['X'])
                    Excel_Y_Index = float(Node_Database[Node]['Z'])
                if Wall_Plane == 'YZ':
                    Excel_X_Index = float(Node_Database[Node]['Y'])
                    Excel_Y_Index = float(Node_Database[Node]['Z'])
                PLT.figure(5)
                if float(SLS_Database[Node]['Wke']) < (Wk_Limit_01/4):
                    PLT.scatter(Excel_X_Index,Excel_Y_Index ,marker = "s" , color='gray',label='Zero',s=Crack_Scaler*float(SLS_Database[Node]['Wke']))
                if not float(SLS_Database[Node]['Wke']) < (Wk_Limit_01/4):
                    if float(SLS_Database[Node]['Wke']) < (Wk_Limit_01/2):
                        PLT.scatter(Excel_X_Index,Excel_Y_Index ,marker = "s" , color='green',label='0.2',s=Crack_Scaler*float(SLS_Database[Node]['Wke']))
                    if not float(SLS_Database[Node]['Wke']) < (Wk_Limit_01/2):
                        PLT.scatter(Excel_X_Index,Excel_Y_Index ,marker = "s" , color='red',label='0.3',s=Crack_Scaler*float(SLS_Database[Node]['Wke']))
            PLT.title(str(Properties_DG) + ' ' +str(DG_Wall) + " " +  str(SLS_LC) + " " +str(SLS_Thermal) + " Wke")
            PLT.savefig(os.getcwd() + "\\" + Plot_Folder + "\\" +Output_File + str(Properties_DG) + '_' +str(DG_Wall) + "_" +  str(SLS_LC) +"_Wke_"+Additional_Extention+".png",dpi = 300)
                
print(str(datetime.now()-Script_Start) + " : Done...")


