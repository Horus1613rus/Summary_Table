Author= "Antl"
Creation_Date =  "2020.10.23"
Code_Ver = "06.85"

import os
import glob
import sys
import datetime
import xlwt
import sqlite3

Summary_for_Slab = False
Digit = 5
#Do not change values below this line.

Settings_Switch = True
if Settings_Switch:
#Dev
    Enable_Concrete_Check = True
#__________________________________________________
    from datetime import datetime
    Script_Start = datetime.now()
    print("\n" + str(datetime.now()-Script_Start) + " : Script Start...")
    
    if len(sys.argv) == 3:
        os.chdir(sys.argv[2])
        with open (sys.argv[1],"r",encoding="utf-8") as Input:
            Lines = Input.readlines()
            Desc_Ver = Lines[0].split("=")[1].strip()
            SLS_Op_MMs = []
            for i in range(0, len(Lines[3].split("=")[1].split(";")[0].split())):
                SLS_Op_MMs.append(int(Lines[3].split("=")[1].split(";")[0].split()[i]))
            SLS_To_MMs = []
            for i in range(0, len(Lines[3].split("=")[1].split(";")[1].split())):
                SLS_To_MMs.append(int(Lines[3].split("=")[1].split(";")[1].split()[i]))
            SLS_In_MMs = []
            for i in range(0, len(Lines[3].split("=")[1].split(";")[2].split())):
                SLS_In_MMs.append(int(Lines[3].split("=")[1].split(";")[2].split()[i]))
            ULS_Op_MMs = []
            for i in range(0, len(Lines[4].split("=")[1].split(";")[0].split())):
                ULS_Op_MMs.append(int(Lines[4].split("=")[1].split(";")[0].split()[i]))
            ULS_To_MMs = []
            for i in range(0, len(Lines[4].split("=")[1].split(";")[1].split())):
                ULS_To_MMs.append(int(Lines[4].split("=")[1].split(";")[1].split()[i]))
            ULS_In_MMs = []
            for i in range(0, len(Lines[4].split("=")[1].split(";")[2].split())):
                ULS_In_MMs.append(int(Lines[4].split("=")[1].split(";")[2].split()[i]))
            ALS_Co_MMs = []
            for i in range(0, len(Lines[5].split("=")[1].split())):
                ALS_Co_MMs.append(int(Lines[5].split("=")[1].split()[i]))
            Field_DGs = []
            for i in range(0, len(Lines[7].split("=")[1].split())):
                Field_DGs.append(int(Lines[7].split("=")[1].split()[i]))
            
            Excluded_DGs = []
            for i in range(0, len(Lines[8].split("=")[1].split())):
                Excluded_DGs.append(int(Lines[8].split("=")[1].split()[i]))
            if len(Field_DGs) == 0:
                sys.exit(str(datetime.now()-Script_Start) + " : There must be at least 1 field DG in " + os.path.basename(sys.argv[2]) + "\n")
    else:
        Desc_Ver = "09.42D"
        Field_DGs_File = "01_2_Field_DGs.txt"
        Exclude_DG_File = "01_3_Excluded_DGs.txt"
        

        SLS_Op_MMs = [1,2,7]    ; SLS_To_MMs = [3,4,8]      ; SLS_In_MMs = [5]
        ULS_Op_MMs = [20,29,26] ; ULS_To_MMs = [22,23,32]   ; ULS_In_MMs = [24]
        ALS_Co_MMs = [27,31,33,34]

        scriptpath = os.path.abspath(__file__)
        scriptdirectory = os.path.dirname(scriptpath)
        os.chdir(scriptdirectory)


        if not len(str(datetime.now()-Script_Start)) < 8:
            print(str(datetime.now()-Script_Start) + " : Processing Field Dgs...")
        else:
            print(str(datetime.now()-Script_Start) + ".000000 : Processing Field Dgs...")
        #Field_DGs
        Field_DGs = []
        with open (Field_DGs_File,"r",encoding="utf-8") as Field_DGs_In:
            Field_DGs_Lines=Field_DGs_In.readlines()
            for Field_DGs_Line in Field_DGs_Lines:
                if len(Field_DGs_Line.strip())>0:
                    if not int(Field_DGs_Line.strip()) in Field_DGs:
                        Field_DGs.append(int(Field_DGs_Line.strip()))


        for Field_DG in Field_DGs:
            if not len(str(Field_DG)) == Digit:
                sys.exit(str(datetime.now()-Script_Start) + " : Field DG must have " + str(Digit) + " digits in " + os.path.basename(Field_DGs_File) + "\n")
        if len(Field_DGs) == 0:
            sys.exit(str(datetime.now()-Script_Start) + " : There must be at least 1 field DG in " + os.path.basename(Field_DGs_File) + "\n")

        #Exclude_Dgs
        Excluded_DGs = []
        print(str(datetime.now()-Script_Start) + " : Processing Exclude Dg...")
        with open (Exclude_DG_File,"r",encoding="utf-8") as Exclude_DG_In:
            Exclude_DG_Lines=Exclude_DG_In.readlines()
            for Exclude_DG_Line in Exclude_DG_Lines:
                if len(Exclude_DG_Line)>2:
                    if not int(Exclude_DG_Line.strip()) in Excluded_DGs:
                        Excluded_DGs.append(int(Exclude_DG_Line.strip()))

    
    Result_File_Name = "01_9_Summary_Table"
    Script_Ver = os.path.basename(__file__).split("_v")[1].split("_Egemen")[0]
    if not os.path.isfile("\\\\fs-saren\\07_DESIGN&ENGINEERING\\Project Documents to Review\\Calculation Note Documents\All_Scripts&Tables\\01_Summary_Table_Experimental\\LC_Description_" + Desc_Ver + ".db"):
        sys.exit(str(datetime.now()-Script_Start) + " : Description not connected to R:\n")
    Connect_SQLite_Desc_DB = sqlite3.connect("\\\\fs-saren\\07_DESIGN&ENGINEERING\\Project Documents to Review\\Calculation Note Documents\All_Scripts&Tables\\01_Summary_Table_Experimental\\LC_Description_" + Desc_Ver + ".db")
    SQLite_Desc_DB = Connect_SQLite_Desc_DB.cursor()

    


    Database={}
    for State in "ULS SLS ALS".split():
        Database[State]={}
        for Phase in "Op To In Co".split():
            Database[State][Phase]={}
            for Side in "Field Edge".split():
                Database[State][Phase][Side]={}
                for Face in "Top Bot Env".split():
                    Database[State][Phase][Side][Face]={}
                    Database[State][Phase][Side][Face]["Concrete"]={}
                    Database[State][Phase][Side][Face]["Steel"]={}
                    Database[State][Phase][Side][Face]["Steel"]["Dir_X"]={}
                    Database[State][Phase][Side][Face]["Steel"]["Dir_X"]["Strain"]=0
                    Database[State][Phase][Side][Face]["Steel"]["Dir_Y"]={}
                    Database[State][Phase][Side][Face]["Steel"]["Dir_Y"]["Strain"]=0
    
    Fcd=0
    FcdALS=0
    




#DG Properties
DG_Prop = {}
Temp_Values = {}
print(str(datetime.now()-Script_Start) + " : Processing DG Info...")
for conshrnu_1 in glob.iglob(os.getcwd() + "\\**\\prop\\conshrnu_1.**", recursive=True):
    Temp_Values["DG"] = os.path.basename(conshrnu_1).replace("conshrnu_1.","")
    DG_Prop[Temp_Values["DG"]] = {}
    DG_Prop[Temp_Values["DG"]]["Reinf"] = {}
    with open (conshrnu_1,"r",encoding="utf-8") as conshrnu_1_In:
        conshrnu_1_Lines=conshrnu_1_In.readlines()
        DG_Prop[Temp_Values["DG"]]["Fcd"] = float(conshrnu_1_Lines[0].strip().split()[0])
        Layer = 1
        for i in range(2, len(conshrnu_1_Lines) - 1):
            DG_Prop[Temp_Values["DG"]]["Reinf"][Layer] = conshrnu_1_Lines[i].split()[0] + "Ø" + conshrnu_1_Lines[i].split()[1] + "c" + conshrnu_1_Lines[i].split()[2]
            Layer = Layer + 1


for conshrna_1 in glob.iglob(os.getcwd() + "\\**\\prop\\conshrna_1.**", recursive=True):
    with open (conshrna_1,"r",encoding="utf-8") as conshrna_1_In:
        Temp_Values["DG"] = os.path.basename(conshrna_1).replace("conshrna_1.","")
        conshrna_1_Lines=conshrna_1_In.readlines()
        DG_Prop[Temp_Values["DG"]]["FcdALS"] = float(conshrna_1_Lines[0].strip().split()[0])


#First Value
for Phase in "Op To In".split():
    for Side in "Field Edge".split():
        for Direction in "Dir_X Dir_Y".split():
            for Face in "Top Bot".split():
                Database["ULS"][Phase][Side][Face]["Steel"][Direction]["Strain"] = 0
                Database["ALS"]["Co"][Side][Face]["Steel"][Direction]["Strain"] = 0
                Database["SLS"][Phase][Side][Face]["Concrete"]["Wk"] = 0
                Database["SLS"][Phase][Side]["Env"]["Concrete"]["Wk"] = 0
                Database["SLS"][Phase][Side][Face]["Concrete"]["LC_Desc"] = "No Crack Occurred"
                Database["SLS"][Phase][Side]["Env"]["Concrete"]["LC_Desc"] = "No Crack Occurred"
                for Value in "DG LC LC_Desc Reinf".split():
                    Database["SLS"][Phase][Side][Face]["Concrete"][Value] = ""
                    Database["SLS"][Phase][Side]["Env"]["Concrete"][Value] = ""

#SLS
print(str(datetime.now()-Script_Start) + " : Processing SLS Info...")
for csm in glob.iglob(os.getcwd() + "/**/*.csm", recursive=True):
    DG = int(csm.split("mm")[0].strip()[len(csm.split("mm")[0]) - Digit: len(csm.split("mm")[0])])
    MM = int(csm.split("mm")[1].replace(".res.csmr","").split(".")[0])
    if not DG in Excluded_DGs:        
        if MM in SLS_Op_MMs + SLS_To_MMs + SLS_In_MMs:
            State = "SLS"
            Side = "Edge"
            if DG in Field_DGs:
                Side = "Field"
            if MM in SLS_Op_MMs:
                Phase = "Op"
            if MM in SLS_To_MMs:
                Phase = "To"
            if MM in SLS_In_MMs:
                Phase = "In"
            with open(csm, "r",encoding="utf-8") as csmIn:
                csmlines=csmIn.readlines()
                Red_Switch = False
                for csmline in csmlines:
                    if "Summary results for" in csmline:
                        LC_Temp = csmline.strip().replace("T","").replace("W","").replace("X","").replace("Y","").split("loadcase :")[1].strip()
                    if Red_Switch:
                        if not "Wk" in Database[State][Phase][Side][Face]["Concrete"]:
                            Database[State][Phase][Side][Face]["Concrete"]["Wk"] = 0
                        if not "Wk" in Database[State][Phase][Side]["Env"]["Concrete"]:
                            Database[State][Phase][Side]["Env"]["Concrete"]["Wk"] = 0
                        if float(csmline.strip().split()[4]) > Database[State][Phase][Side][Face]["Concrete"]["Wk"]:
                            Database[State][Phase][Side][Face]["Concrete"]["Wk"] = float(csmline.strip().split()[4])
                            Database[State][Phase][Side][Face]["Concrete"]["LC"] = int(LC_Temp)
                            Database[State][Phase][Side][Face]["Concrete"]["DG"] = int(DG)
                            if Summary_for_Slab:
                                Database[State][Phase][Side][Face]["Concrete"]["Reinf"] = "Xt:" + DG_Prop[str(DG)]["Reinf"][1] + " Yt:" + DG_Prop[str(DG)]["Reinf"][2] + " Xb:" + DG_Prop[str(DG)]["Reinf"][len(DG_Prop[str(DG)]["Reinf"])] + " Yb:" + DG_Prop[str(DG)]["Reinf"][len(DG_Prop[str(DG)]["Reinf"])-1]
                            if not Summary_for_Slab:
                                Database[State][Phase][Side][Face]["Concrete"]["Reinf"] = "X1:" + DG_Prop[str(DG)]["Reinf"][1] + " Y1:" + DG_Prop[str(DG)]["Reinf"][2] + " X2:" + DG_Prop[str(DG)]["Reinf"][len(DG_Prop[str(DG)]["Reinf"])] + " Y2:" + DG_Prop[str(DG)]["Reinf"][len(DG_Prop[str(DG)]["Reinf"])-1]
                            
                            
                        if float(csmline.strip().split()[4]) > Database[State][Phase][Side]["Env"]["Concrete"]["Wk"]:
                            Database[State][Phase][Side]["Env"]["Concrete"]["Wk"] = float(csmline.strip().split()[4])
                            Database[State][Phase][Side]["Env"]["Concrete"]["LC"] = int(LC_Temp)
                            Database[State][Phase][Side]["Env"]["Concrete"]["DG"] = int(DG)
                            if Summary_for_Slab:
                                Database[State][Phase][Side]["Env"]["Concrete"]["Reinf"] = "Xt:" + DG_Prop[str(DG)]["Reinf"][1] + " Yt:" + DG_Prop[str(DG)]["Reinf"][2] + " Xb:" + DG_Prop[str(DG)]["Reinf"][len(DG_Prop[str(DG)]["Reinf"])] + " Yb:" + DG_Prop[str(DG)]["Reinf"][len(DG_Prop[str(DG)]["Reinf"])-1]
                            if not Summary_for_Slab:
                                Database[State][Phase][Side]["Env"]["Concrete"]["Reinf"] = "X1:" + DG_Prop[str(DG)]["Reinf"][1] + " Y1:" + DG_Prop[str(DG)]["Reinf"][2] + " X2:" + DG_Prop[str(DG)]["Reinf"][len(DG_Prop[str(DG)]["Reinf"])] + " Y2:" + DG_Prop[str(DG)]["Reinf"][len(DG_Prop[str(DG)]["Reinf"])-1]
                            

                    if Face == "Bot":
                        Red_Switch = False
                    if Face == "Top":
                        Face = "Bot"
                    if "wok" in csmline:
                        Red_Switch = True
                        Face = "Top"

#ULS/ALS
print(str(datetime.now()-Script_Start) + " : Processing ULS/ALS Info...")
for csmr in glob.iglob(os.getcwd() + "/**/*.csmr", recursive=True):
    DG = int(csmr.split("mm")[0].strip()[len(csmr.split("mm")[0]) - Digit: len(csmr.split("mm")[0])])
    MM = int(csmr.split("mm")[1].replace(".res.csmr","").split(".")[0])
    if not DG in Excluded_DGs:
        if MM in ULS_Op_MMs + ULS_To_MMs + ULS_In_MMs + ALS_Co_MMs:
            State = "ULS"
            if MM in  ALS_Co_MMs:
                State = "ALS"
                Phase = "Co"
            Side = "Edge"
            if DG in Field_DGs:
                Side = "Field"
            if MM in ULS_Op_MMs:
                Phase = "Op"
            if MM in ULS_To_MMs:
                Phase = "To"
            if MM in ULS_In_MMs:
                Phase = "In"
            with open(csmr, "r",encoding="utf-8") as csmrIn:
                csmrlines=csmrIn.readlines()
                Read_Switch_Con = False
                Steel_Temp_Database={}
                for i in range(1,9):
                    Steel_Temp_Database[i]={}
                    Steel_Temp_Database[i]["Strain"] = 0
                for csmrline in csmrlines:
                    if "Min conc. strain" in csmrline:
                        if "top face" in csmrline:
                            Face = "Top"
                        if "bot face" in csmrline:
                            Face = "Bot"
                        if not "Strain" in Database[State][Phase][Side][Face]["Concrete"]:
                            Database[State][Phase][Side][Face]["Concrete"]["Strain"] = 0
                        if abs(float(csmrline.strip()[27:].split()[0])) > abs(Database[State][Phase][Side][Face]["Concrete"]["Strain"]):
                            Database[State][Phase][Side][Face]["Concrete"]["Strain"] = float(csmrline.strip()[27:].split()[0])
                            Database[State][Phase][Side][Face]["Concrete"]["LC"] = int(csmrline.strip()[-13:].strip().replace("T","").replace("W","").replace("X","").replace("Y",""))
                            Database[State][Phase][Side][Face]["Concrete"]["DG"] = int(DG)
                            if Summary_for_Slab:
                                Database[State][Phase][Side][Face]["Concrete"]["Reinf"] = "Xt:" + DG_Prop[str(DG)]["Reinf"][1] + " Yt:" + DG_Prop[str(DG)]["Reinf"][2] + " Xb:" + DG_Prop[str(DG)]["Reinf"][len(DG_Prop[str(DG)]["Reinf"])] + " Yb:" + DG_Prop[str(DG)]["Reinf"][len(DG_Prop[str(DG)]["Reinf"])-1]
                            if not Summary_for_Slab:
                                Database[State][Phase][Side][Face]["Concrete"]["Reinf"] = "X1:" + DG_Prop[str(DG)]["Reinf"][1] + " Y1:" + DG_Prop[str(DG)]["Reinf"][2] + " X2:" + DG_Prop[str(DG)]["Reinf"][len(DG_Prop[str(DG)]["Reinf"])] + " Y2:" + DG_Prop[str(DG)]["Reinf"][len(DG_Prop[str(DG)]["Reinf"])-1]

                            Read_Switch_Con = True
                    if Read_Switch_Con:
                        if "Min conc. stress" in csmrline:
                            if "top face" in csmrline:
                                Face = "Top"
                            if "bot face" in csmrline:
                                Face = "Bot"
                            Database[State][Phase][Side][Face]["Concrete"]["Stress"] = float(csmrline.strip()[27:].split()[0])
                            if MM in  ALS_Co_MMs:
                                Database[State][Phase][Side][Face]["Concrete"]["UR"] = float(csmrline.strip()[27:].split()[0])/DG_Prop[str(DG)]["FcdALS"]
                            if not MM in  ALS_Co_MMs:
                                Database[State][Phase][Side][Face]["Concrete"]["UR"] = float(csmrline.strip()[27:].split()[0])/DG_Prop[str(DG)]["Fcd"]
                            Read_Switch_Con = False
                    if "steel strain layer" in csmrline:
                        Layer = int(csmrline.strip().split()[4])
                        if abs(float(csmrline.strip()[27:].split()[0])) > abs(Steel_Temp_Database[Layer]["Strain"]):
                            Steel_Temp_Database[Layer]["Strain"] = float(csmrline.strip()[27:].split()[0])
                            Steel_Temp_Database[Layer]["LC"] = int(csmrline.strip()[-13:].strip().replace("T","").replace("W","").replace("X","").replace("Y",""))
                            Steel_Temp_Database[Layer]["DG"] = int(DG)
                            Steel_Temp_Database[Layer]["UR"] = float(csmrline.strip()[27:].split()[0])/ 2.5
                            Steel_Temp_Database[Layer]["Reinf"] = DG_Prop[str(DG)]["Reinf"][Layer]

                            
                            
            if abs(Steel_Temp_Database[1]["Strain"]) > abs(Database[State][Phase][Side]["Top"]["Steel"]["Dir_X"]["Strain"]):   
                for Value in "Reinf Strain UR DG LC".split():                 
                    Database[State][Phase][Side]["Top"]["Steel"]["Dir_X"][Value] = Steel_Temp_Database[1][Value]
            if abs(Steel_Temp_Database[2]["Strain"]) > abs(Database[State][Phase][Side]["Top"]["Steel"]["Dir_Y"]["Strain"]):
                for Value in "Reinf Strain UR DG LC".split():
                    Database[State][Phase][Side]["Top"]["Steel"]["Dir_Y"][Value] = Steel_Temp_Database[2][Value]
            if abs(Steel_Temp_Database[Layer-1]["Strain"]) > abs(Database[State][Phase][Side]["Bot"]["Steel"]["Dir_Y"]["Strain"]):
                for Value in "Reinf Strain UR DG LC".split():
                    Database[State][Phase][Side]["Bot"]["Steel"]["Dir_Y"][Value] = Steel_Temp_Database[Layer-1][Value]
            if abs(Steel_Temp_Database[Layer]["Strain"]) > abs(Database[State][Phase][Side]["Bot"]["Steel"]["Dir_X"]["Strain"]):
                for Value in "Reinf Strain UR DG LC".split():
                    Database[State][Phase][Side]["Bot"]["Steel"]["Dir_X"][Value] = Steel_Temp_Database[Layer][Value]

#Excel_Preparing
print(str(datetime.now()-Script_Start) + " : Preparing Excel...")
Time_Stamp = str(datetime.today().strftime('%Y%m%d'))
Worksheet = {}
Workbook = xlwt.Workbook(Result_File_Name + "_" + str(Script_Ver) + "_" + Time_Stamp +".xls")
Worksheet["Info"] = Workbook.add_sheet("Info",cell_overwrite_ok=True)
Worksheet["Info"].col(1).width = 256 * 20
Worksheet["Info"].col(2).width = 256 * 10
Worksheet["Info"].write( 0 , 1 , "Information" , xlwt.easyxf('font: bold True') )
Worksheet["Info"].write( 2 , 1 , "XLS Creation Date" , xlwt.easyxf('font: bold True') )
Worksheet["Info"].write( 2 , 2 , str(datetime.today().strftime('%d-%m-%Y')) )
Worksheet["Info"].write( 3 , 1 , "Script Version" , xlwt.easyxf('font: bold True') )
Worksheet["Info"].write( 3 , 2 , "v"+str(Script_Ver) )
Worksheet["Info"].write( 4 , 1 , "LC_Desc Version" , xlwt.easyxf('font: bold True') )
Worksheet["Info"].write( 4 , 2 , Desc_Ver )

Worksheet["Info"].col(4).width = 256 * 20
Worksheet["Info"].write( 2 , 4 , "Excluded Dgs" , xlwt.easyxf('font: bold True')  )
for i in range(3,len(Excluded_DGs)+3):
    Worksheet["Info"].write( i , 4 , Excluded_DGs[i-3])
Worksheet["Info"].col(6).width = 256 * 20
Worksheet["Info"].write( 2 , 6 , "Field Dgs" , xlwt.easyxf('font: bold True')  )
for i in range(3,len(Field_DGs)+3):
    Worksheet["Info"].write( i , 6 , Field_DGs[i-3])

Op_MMs = SLS_Op_MMs + ULS_Op_MMs
Worksheet["Info"].write( 2 , 8 , "Op MMs" , xlwt.easyxf('font: bold True') )
for i in range(3,len(Op_MMs)+3):
    Worksheet["Info"].write( i , 8 , Op_MMs[i-3])
To_MMs = SLS_To_MMs + ULS_To_MMs
Worksheet["Info"].write( 2 , 9 , "To MMs" , xlwt.easyxf('font: bold True') )
for i in range(3,len(To_MMs)+3):
    Worksheet["Info"].write( i , 9 , To_MMs[i-3])
In_MMs = SLS_In_MMs + ULS_In_MMs
Worksheet["Info"].write( 2 , 10 , "In MMs" , xlwt.easyxf('font: bold True') )
for i in range(3,len(In_MMs)+3):
    Worksheet["Info"].write( i , 10 , In_MMs[i-3])

Worksheet["Info"].write( 2 , 11 , "Ac MMs" , xlwt.easyxf('font: bold True') )
for i in range(3,len(ALS_Co_MMs)+3):
    Worksheet["Info"].write( i , 11 , ALS_Co_MMs[i-3])




Cell_Style = {}
Cell_Style[5] = xlwt.XFStyle()
Cell_Style[5].num_format_str = '0.00'
Cell_Style[6] = xlwt.XFStyle()
Cell_Style[6].num_format_str = '0%'
Cell_Style[7] = xlwt.XFStyle()
Cell_Style[7].num_format_str = '0'
Cell_Style[8] = xlwt.XFStyle()
Cell_Style[8].num_format_str = '0'
Cell_Style[9] = xlwt.XFStyle()
Cell_Style[9].num_format_str = '0'

for State in Database:
    for Phase in Database[State]:
        for Side in Database[State][Phase]:
            for Face in Database[State][Phase][Side]:
                for i in Database[State][Phase][Side][Face]:
                    if i == "Concrete":
                        if not State == "SLS":
                            if Face == "Env":
                                Database[State][Phase][Side][Face][i]["LC_Desc"] = "Desc Error - 01"
                                break
                        else:
                            if not Face == "Env":
                                Database[State][Phase][Side][Face][i]["LC_Desc"] = "Desc Error - 02"
                                break
                        if not State == "ALS":
                            if Phase == "Co":
                                Database[State][Phase][Side][Face][i]["LC_Desc"] = "Desc Error - 03"
                                break
                        else:
                            if not Phase == "Co":
                                Database[State][Phase][Side][Face][i]["LC_Desc"] = "Desc Error - 04"
                                break
                        ##print(str(i) + " " + str(State) + " " + str(Phase) + " " + str(Side) + " " + str(Face))
                        try:

                            LC_Temp = str(Database[State][Phase][Side][Face][i]["LC"])
                            if not int(LC_Temp.strip()[-5:])== 0:
                                SQLite_Desc_DB.execute("SELECT * FROM main WHERE LC = '" + str(int(LC_Temp.strip()[:-5])*100000) +"'")
                                Temp_LC_Desc = SQLite_Desc_DB.fetchone()
                                Database[State][Phase][Side][Face][i]["LC_Desc"] = Temp_LC_Desc[1] + " " + LC_Temp.strip()[-4:-3] + " Phase: " + LC_Temp.strip()[-3:]
                            if int(LC_Temp.strip()[-5:])== 0:
                                SQLite_Desc_DB.execute("SELECT * FROM main WHERE LC = '" + LC_Temp +"'")
                                Temp_LC_Desc = SQLite_Desc_DB.fetchone()
                                Database[State][Phase][Side][Face][i]["LC_Desc"] = Temp_LC_Desc[1]
                        except:
                            Database[State][Phase][Side][Face][i]["LC_Desc"] = "Desc Error - 05"
                    if i == "Steel":
                        if not State == "SLS":
                            for Direction in Database[State][Phase][Side][Face][i]:
                                ##print(str(i) + " " + str(State) + " " + str(Phase) + " " + str(Side) + " " + str(Face))
                                try:
                                    LC_Temp = str(Database[State][Phase][Side][Face][i][Direction]["LC"])
                                    if not int(LC_Temp.strip()[-5:])== 0:
                                        SQLite_Desc_DB.execute("SELECT * FROM main WHERE LC = '" + str(int(LC_Temp.strip()[:-5])*100000) +"'")
                                        Temp_LC_Desc = SQLite_Desc_DB.fetchone()
                                        Database[State][Phase][Side][Face][i][Direction]["LC_Desc"] = Temp_LC_Desc[1] + " " + LC_Temp.strip()[-4:-3] + " Phase: " + LC_Temp.strip()[-3:]
                                    if int(LC_Temp.strip()[-5:])== 0:
                                        SQLite_Desc_DB.execute("SELECT * FROM main WHERE LC = '" + LC_Temp +"'")
                                        Temp_LC_Desc = SQLite_Desc_DB.fetchone()
                                        Database[State][Phase][Side][Face][i][Direction]["LC_Desc"] = Temp_LC_Desc[1]
                                except:
                                    Database[State][Phase][Side][Face][i]["LC_Desc"] = "Desc Error - 06"




for State in Database:
    if State == "ALS":
        Phase = "Co"
        Worksheet[State+"_"+Phase] = Workbook.add_sheet(State+"_"+Phase,cell_overwrite_ok=True)
        Worksheet[State+"_"+Phase].write( 0 , 1 , "Accidental" , xlwt.easyxf('font: bold True') )

        Worksheet[State+"_"+Phase].col(1).width = 256 * 10  #Limit_State
        Worksheet[State+"_"+Phase].col(2).width = 256 * 8  #Location
        Worksheet[State+"_"+Phase].col(3).width = 256 * 8  #Face
        Worksheet[State+"_"+Phase].col(4).width = 256 * 25  #Reinf
        Worksheet[State+"_"+Phase].col(5).width = 256 * 8  #Strain
        Worksheet[State+"_"+Phase].col(6).width = 256 * 8  #UR
        Worksheet[State+"_"+Phase].col(7).width = 256 * 8  #DG
        Worksheet[State+"_"+Phase].col(8).width = 256 * 15  #LC
        Worksheet[State+"_"+Phase].col(9).width = 256 * 50  #Description

        column_num = 1
        for i in "Limit_State Location Face Reinforcement Strain UR DG LC Description".split():
            Worksheet[State+"_"+Phase].write( 2 , column_num , i , xlwt.easyxf('font: bold True') )
            column_num = column_num + 1
        column_num = 1
        for i in "Limit_State Location Face Reinforcement Stress UR DG LC Description".split():
            Worksheet[State+"_"+Phase].write( 11 , column_num , i , xlwt.easyxf('font: bold True') )
            column_num = column_num + 1
        
        row_num = 3
        for Side in Database[State][Phase]:
            for Face in "Top Bot".split():
                if Face=="Top":
                    Face_Name = "Outer"
                    Face_Number = "1"
                    if Summary_for_Slab:
                        Face_Name = "Top"
                        Face_Number = "t"
                if Face=="Bot":
                    Face_Name = "Inner"
                    Face_Number = "2"
                    if Summary_for_Slab:
                        Face_Name = "Bottom"
                        Face_Number = "b"
                for Direction in "Dir_X Dir_Y".split():
                    Worksheet[State+"_"+Phase].write( row_num , 1 , "ALS_Reinf." , xlwt.easyxf('font: bold True') )
                    Worksheet[State+"_"+Phase].write( row_num , 2 , Side , xlwt.easyxf('font: bold True') )
                    Worksheet[State+"_"+Phase].write( row_num , 3 , Face_Name , xlwt.easyxf('font: bold True') )
                    column_num = 5
                    Worksheet[State+"_"+Phase].write( row_num , 4 , Direction.replace("Dir_","") + Face_Number + ": " + Database[State][Phase][Side][Face]["Steel"][Direction]["Reinf"])
                    for Value in "Strain UR DG LC LC_Desc".split():
                        Worksheet[State+"_"+Phase].write( row_num , column_num , Database[State][Phase][Side][Face]["Steel"][Direction][Value] , Cell_Style[column_num])
                        column_num = column_num + 1
                    row_num = row_num + 1
        row_num = row_num + 1
        for Side in Database[State][Phase]:
            for Face in "Top Bot".split():
                if Face=="Top":
                    Face_Name = "Outer"
                    Face_Number = "1"
                    if Summary_for_Slab:
                        Face_Name = "Top"
                        Face_Number = "t"
                if Face=="Bot":
                    Face_Name = "Inner"
                    Face_Number = "2"
                    if Summary_for_Slab:
                        Face_Name = "Bottom"
                        Face_Number = "b"
                Worksheet[State+"_"+Phase].write( row_num , 1 , "ALS_C_Comp." , xlwt.easyxf('font: bold True') )
                Worksheet[State+"_"+Phase].write( row_num , 2 , Side , xlwt.easyxf('font: bold True') )
                Worksheet[State+"_"+Phase].write( row_num , 3 , Face_Name , xlwt.easyxf('font: bold True') )
                column_num = 5
                Worksheet[State+"_"+Phase].write( row_num , 4 , Database[State][Phase][Side][Face]["Concrete"]["Reinf"] , xlwt.easyxf('font: height 160 ; alignment: wrap True'))
                for Value in "Stress UR DG LC LC_Desc".split():
                    Worksheet[State+"_"+Phase].write( row_num , column_num , Database[State][Phase][Side][Face]["Concrete"][Value] , Cell_Style[column_num])
                    column_num = column_num + 1
                # Concrete Material
                if DG_Prop[str(Database[State][Phase][Side][Face]["Concrete"]["DG"])]["Fcd"] == 29.656:
                    Worksheet[State+"_"+Phase].write( row_num , column_num , "MND2")
                elif DG_Prop[str(Database[State][Phase][Side][Face]["Concrete"]["DG"])]["Fcd"] == 31.481:
                    Worksheet[State+"_"+Phase].write( row_num , column_num , "MND")
                elif DG_Prop[str(Database[State][Phase][Side][Face]["Concrete"]["DG"])]["Fcd"] == 25.72:
                    Worksheet[State+"_"+Phase].write( row_num , column_num , "LWA")
                else:
                    Worksheet[State+"_"+Phase].write( row_num , column_num , "Unidentified")

                row_num = row_num + 1
        
    if not State == "ALS":
        for Phase in Database[State]:
            if not Phase == "Co":
                if Phase == "Op":
                    Exc_Name = "Operation"
                if Phase == "To":
                    Exc_Name = "Towing"
                if Phase == "In":
                    Exc_Name = "Installation"
                try:
                    Worksheet[Phase] = Workbook.add_sheet(Phase,cell_overwrite_ok=True)
                    Worksheet[Phase].write( 0 , 1 , Exc_Name , xlwt.easyxf('font: bold True') )
                except:
                    pass
                Worksheet[Phase].col(1).width = 256 * 10  #Limit_State
                Worksheet[Phase].col(2).width = 256 * 8  #Location
                Worksheet[Phase].col(3).width = 256 * 8  #Face
                Worksheet[Phase].col(4).width = 256 * 25  #Reinf
                Worksheet[Phase].col(5).width = 256 * 8  #Strain
                Worksheet[Phase].col(6).width = 256 * 8  #UR
                Worksheet[Phase].col(7).width = 256 * 8  #DG
                Worksheet[Phase].col(8).width = 256 * 15  #LC
                Worksheet[Phase].col(9).width = 256 * 50  #Description

                column_num = 1
                for i in "Limit_State Location Face Reinforcement Value MaxCr(mm) DG LC Description".split():
                    Worksheet[Phase].write( 2 , column_num , i , xlwt.easyxf('font: bold True') )
                    column_num = column_num + 1
                column_num = 1
                for i in "Limit_State Location Face Reinforcement Strain UR DG LC Description".split():
                    Worksheet[Phase].write( 5 , column_num , i , xlwt.easyxf('font: bold True') )
                    column_num = column_num + 1
                column_num = 1
                for i in "Limit_State Location Face Reinforcement Stress UR DG LC Description".split():
                    Worksheet[Phase].write( 14 , column_num , i , xlwt.easyxf('font: bold True') )
                    column_num = column_num + 1
                if State == "SLS":
                    row_num = 3
                    for Side in Database[State][Phase]:
                        Face = "Env"
                        Worksheet[Phase].write( row_num , 1 , "SLS" , xlwt.easyxf('font: bold True') )
                        Worksheet[Phase].write( row_num , 2 , Side , xlwt.easyxf('font: bold True') )
                        Worksheet[Phase].write( row_num , 3 , "" , xlwt.easyxf('font: bold True') )
                        column_num = 5
                        Worksheet[Phase].write( row_num , 4 , Database[State][Phase][Side][Face]["Concrete"]["Reinf"] , xlwt.easyxf('font: height 160; alignment: wrap True'))
                        for Value in "Wk MaxCr DG LC LC_Desc".split():
                            if not Value == "MaxCr":
                                Worksheet[Phase].write( row_num , column_num , Database[State][Phase][Side][Face]["Concrete"][Value] , Cell_Style[column_num])
                            if Value == "MaxCr":
                                Worksheet[Phase].write( row_num , column_num , "Please Fill" , Cell_Style[column_num])
                            column_num = column_num + 1
                        row_num = row_num + 1

                if State == "ULS":
                    row_num = 6
                    for Side in Database[State][Phase]:
                        for Face in "Top Bot".split():
                            if Face=="Top":
                                Face_Name = "Outer"
                                Face_Number = "1"
                                if Summary_for_Slab:
                                    Face_Name = "Top"
                                    Face_Number = "t"
                            if Face=="Bot":
                                Face_Name = "Inner"
                                Face_Number = "2"
                                if Summary_for_Slab:
                                    Face_Name = "Bottom"
                                    Face_Number = "b"
                            for Direction in "Dir_X Dir_Y".split():
                                Worksheet[Phase].write( row_num , 1 , "ULS_Reinf." , xlwt.easyxf('font: bold True') )
                                Worksheet[Phase].write( row_num , 2 , Side , xlwt.easyxf('font: bold True') )
                                Worksheet[Phase].write( row_num , 3 , Face_Name , xlwt.easyxf('font: bold True') )
                                column_num = 5
                                Worksheet[Phase].write( row_num , 4 , Direction.replace("Dir_","") + Face_Number + ": " + Database[State][Phase][Side][Face]["Steel"][Direction]["Reinf"])
                                for Value in "Strain UR DG LC LC_Desc".split():
                                    Worksheet[Phase].write( row_num , column_num , Database[State][Phase][Side][Face]["Steel"][Direction][Value] , Cell_Style[column_num])
                                    column_num = column_num + 1
                                row_num = row_num + 1
                    row_num = row_num + 1
                    for Side in Database[State][Phase]:
                        for Face in "Top Bot".split():
                            if Face=="Top":
                                Face_Name = "Outer"
                                Face_Number = "1"
                                if Summary_for_Slab:
                                    Face_Name = "Top"
                                    Face_Number = "t"
                            if Face=="Bot":
                                Face_Name = "Inner"
                                Face_Number = "1"
                                if Summary_for_Slab:
                                    Face_Name = "Bottom"
                                    Face_Number = "b"
                            Worksheet[Phase].write( row_num , 1 , "ULS_C_Comp." , xlwt.easyxf('font: bold True') )
                            Worksheet[Phase].write( row_num , 2 , Side , xlwt.easyxf('font: bold True') )
                            Worksheet[Phase].write( row_num , 3 , Face_Name , xlwt.easyxf('font: bold True') )
                            column_num = 5
                            Worksheet[Phase].write( row_num , 4 , Database[State][Phase][Side][Face]["Concrete"]["Reinf"] , xlwt.easyxf('font: height 160 ; alignment: wrap True'))
                            for Value in "Stress UR DG LC LC_Desc".split():
                                Worksheet[Phase].write( row_num , column_num , Database[State][Phase][Side][Face]["Concrete"][Value] , Cell_Style[column_num])
                                column_num = column_num + 1
                            # Concrete Material
                            if DG_Prop[str(Database[State][Phase][Side][Face]["Concrete"]["DG"])]["Fcd"] == 29.656:
                                Worksheet[Phase].write( row_num , column_num , "MND2")
                            elif DG_Prop[str(Database[State][Phase][Side][Face]["Concrete"]["DG"])]["Fcd"] == 31.481:
                                Worksheet[Phase].write( row_num , column_num , "MND")
                            elif DG_Prop[str(Database[State][Phase][Side][Face]["Concrete"]["DG"])]["Fcd"] == 25.72:
                                Worksheet[Phase].write( row_num , column_num , "LWA")
                            else:
                                Worksheet[Phase].write( row_num , column_num , "Unidentified")
                            row_num = row_num + 1

Workbook.save(Result_File_Name + "_" + str(Script_Ver) + "_" + Time_Stamp +".xls")
