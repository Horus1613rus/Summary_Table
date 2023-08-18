"""
@author: Antl
Date   :
"""

Wall_Name = "CASETS2SW"

Excel_Output = True
List_Output = True
# Excel Limits switch to true to activate.
X_Limit = True
X_Limit_Min = 0
X_Limit_Max = 5000
Y_Limit = False
Y_Limit_Min = 0
Y_Limit_Max = 0



from datetime import datetime
Script_Start = datetime.now()
import os
import sys
import glob
import time
import xlwt
Script_Path = os.path.abspath(__file__)
Script_Directory = os.path.dirname(Script_Path)
os.chdir(Script_Directory)
Script_Ver = os.path.basename(__file__).split("_v")[1].split("_Egemen")[0]

time.sleep(1)



Temp_Values = {}

Workbook = {} ; Worksheet = {}

# Raw Coordinates
print(str(datetime.now()-Script_Start) + " : Collecting Coordinates" )
Raw_Coor = {}
Raw_Coor_Values = {}
for Axis in "X Y Z".split():
    Raw_Coor_Values[Axis] = []
for Coor_File in glob.iglob(os.getcwd() + "\\**\\"+ str(Wall_Name) + ".dg.txt", recursive=True):
    with open(Coor_File, 'r' , encoding='utf-8') as Coor_File_In:
        Coor_File_Lines = Coor_File_In.readlines()
        for Coor_File_Line in Coor_File_Lines:
            Temp_Node = Coor_File_Line.split()[1]
            if not Temp_Node in Raw_Coor:
                Raw_Coor[Temp_Node] = {}
            Raw_Coor[Temp_Node]["X"] = float(Coor_File_Line.split()[2])
            Raw_Coor[Temp_Node]["Y"] = float(Coor_File_Line.split()[3])
            Raw_Coor[Temp_Node]["Z"] = float(Coor_File_Line.split()[4])
            Raw_Coor[Temp_Node]["DG"] = Coor_File_Line.split()[5]
            for Axis in "X Y Z".split():
                if not Raw_Coor[Temp_Node][Axis] in Raw_Coor_Values[Axis]:
                    Raw_Coor_Values[Axis].append(Raw_Coor[Temp_Node][Axis])

# Wall Positioning
print(str(datetime.now()-Script_Start) + " : Setting Work Plane" )
for Axis in "X Y Z".split():
    Raw_Coor_Values[Axis].sort
    Temp_Values["Diff_" + Axis] = abs(Raw_Coor_Values[Axis][len(Raw_Coor_Values[Axis]) - 1] - Raw_Coor_Values[Axis][0])
if Temp_Values["Diff_X"] > Temp_Values["Diff_Z"] and Temp_Values["Diff_Y"] > Temp_Values["Diff_Z"]:
    Plot_Position = "X Y"
else:
    if Temp_Values["Diff_X"] > Temp_Values["Diff_Y"]:
        Plot_Position = "X Z"
    else:
        Plot_Position = "Y Z"
print(str(datetime.now()-Script_Start) + " : " + Plot_Position.split()[0] + Plot_Position.split()[1])

# Plot Coordinates Raw
Raw_Plot_Values = {}
for Axis in "Column Row".split():
    Raw_Plot_Values[Axis] = []
Raw_Plot_Values["Column"] = Raw_Coor_Values[Plot_Position.split()[0]]
Raw_Plot_Values["Row"] = Raw_Coor_Values[Plot_Position.split()[1]]
for Axis in "Column Row".split():
    Raw_Plot_Values[Axis].sort

# Node Quantities
Node_Quantity = {}
for Axis in "Column Row".split():
    Temp_Values[Axis + "_Max_Qty"] = 0
    if not Axis in Node_Quantity:
        Node_Quantity[Axis] = {}
    if Axis == "Column":
        i = 0
    else:
        i = 1
    for Value in Raw_Plot_Values[Axis]:
        Temp_Values["Qty"] = 0
        if not Value in Node_Quantity[Axis]:
            Node_Quantity[Axis][Value] = []
        for Node in Raw_Coor:
            if Raw_Coor[Node][Plot_Position.split()[i]] == Value:
                Node_Quantity[Axis][Value].append(Node)
                Temp_Values["Qty"] = Temp_Values["Qty"] + 1
        if Temp_Values["Qty"] > Temp_Values[Axis + "_Max_Qty"]:
            Temp_Values[Axis + "_Max_Qty"] = Temp_Values["Qty"]

# Plot Coordinates Alligning
print(str(datetime.now()-Script_Start) + " : Alligning Nodes" )
Plot_Coor = {}
Plot_Values = {}
for Axis in "Column Row".split():
    Plot_Values[Axis] = []
    if Axis == "Column":
        i = 0
    else:
        i = 1
    Align_Coordinates = True
    j = 0
    Temp_Values["List of Nodes"] = []
    while j < len(Raw_Plot_Values[Axis]):
        for Node in Node_Quantity[Axis][Raw_Plot_Values[Axis][j]]:
            Temp_Values["List of Nodes"].append(Node)
        if len(Temp_Values["List of Nodes"]) < Temp_Values[Axis + "_Max_Qty"]:
            j = j + 1
        else:
            j = j + 1
            Temp_Values["Average_Value"] = 0
            for Node in Temp_Values["List of Nodes"]:
                if not Node in Plot_Coor:
                    Plot_Coor[Node] = {}
                Temp_Values["Average_Value"] = Temp_Values["Average_Value"] + Raw_Coor[Node][Plot_Position.split()[i]]
            for Node in Temp_Values["List of Nodes"]:
                Plot_Coor[Node][Axis] = round(Temp_Values["Average_Value"] / len(Temp_Values["List of Nodes"]),1)
            Temp_Values["List of Nodes"] = []
    for Node in Plot_Coor:
        if not Plot_Coor[Node][Axis] in Plot_Values[Axis]:
            Plot_Values[Axis].append(Plot_Coor[Node][Axis])
            Plot_Values[Axis].sort()

Remove_Nodes = []
for Node in Plot_Coor:
    if X_Limit:
        if Plot_Coor[Node]["Column"] > float(X_Limit_Max) or Plot_Coor[Node]["Column"] < float(X_Limit_Min):
            if not Node in Remove_Nodes:
                Remove_Nodes.append(Node)
for Node in Plot_Coor:
    if Y_Limit:
        if Plot_Coor[Node]["Row"] > float(Y_Limit_Max) or Plot_Coor[Node]["Row"] < float(Y_Limit_Min):
            if not Node in Remove_Nodes:
                Remove_Nodes.append(Node)
for Node in Remove_Nodes:
    del Plot_Coor[Node]

Name = Wall_Name + "_Nodes"
if List_Output:
    print(str(datetime.now()-Script_Start) + " : Preparing txt" )
    with open(Name + ".txt", 'w' , encoding='utf-8') as List_File_Out:
        for Node in Plot_Coor:
            List_File_Out.write(str(Wall_Name) + "\t" + Node + "\n")

if Excel_Output:
    print(str(datetime.now()-Script_Start) + " : Preparing Excel" )
    Workbook[Name] = xlwt.Workbook(Name + ".xls")
    for Value in "Nodes DG".split():
        Worksheet[Value] = Workbook[Name].add_sheet(Value,cell_overwrite_ok=True)
        for Column_Value in Plot_Values["Column"]:
            Worksheet[Value].write(0,Plot_Values["Column"].index(Column_Value)+1,Column_Value, xlwt.easyxf("font: bold True; align: horiz center"))
        for Row_Value in Plot_Values["Row"]:
            Worksheet[Value].write(len(Plot_Values["Row"])-Plot_Values["Row"].index(Row_Value),0,Row_Value, xlwt.easyxf("font: bold True; align: horiz center"))
    for Node in Plot_Coor:
        Worksheet["Nodes"].write(len(Plot_Values["Row"])-Plot_Values["Row"].index(Plot_Coor[Node]["Row"]),Plot_Values["Column"].index(Plot_Coor[Node]["Column"])+1,int(Node))
        try:
            Worksheet["DG"].write(len(Plot_Values["Row"])-Plot_Values["Row"].index(Plot_Coor[Node]["Row"]),Plot_Values["Column"].index(Plot_Coor[Node]["Column"])+1,int(Raw_Coor[Node]["DG"]))
        except:
            Dummy = True
    Workbook[Name].save(Name + ".xls")
