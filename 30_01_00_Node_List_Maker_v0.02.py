"""
@author: Antl
Date   :
"""
Wall_Name_List = ["IW2NW", "IW2NE"]

import os

Script_Path = os.path.abspath(__file__)
Script_Directory = os.path.dirname(Script_Path)
os.chdir(Script_Directory)
Script_Ver = os.path.basename(__file__).split("_v")[1].split("_Egemen")[0]
Script_Prefix = os.path.basename(__file__).split("_")[0] + "_" + os.path.basename(__file__).split("_")[1] + "_"


Node_List = []
try:
    with open(Script_Prefix + "01_Raw_List.txt", 'r' , encoding='utf-8') as File_In:
        File_Lines = File_In.readlines()
        for File_Line in File_Lines:
            try:
                Qty = len(File_Line.split())
                for i in range(0, Qty):
                    Node_List.append(int(File_Line.split()[i]))
            except:
                Dummy = True
    Node_List.sort()
    with open(Script_Prefix + "02_Node_List.txt", 'w' , encoding='utf-8') as File_Out:
        for Node in Node_List:
            for Wall_Name in Wall_Name_List:
                File_Out.write(Wall_Name + "\t" + str(Node) + "\n")
except:
    with open(Script_Prefix + "01_Raw_List.txt", 'w' , encoding='utf-8') as File_Out:
        File_Out.close()



