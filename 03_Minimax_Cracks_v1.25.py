"""
@author: Antl
Version: 1.00
"""
import os
import os.path
import glob
import codecs
import datetime
import xlwt


Operation_Crack_Req = 0.4
Towing_Crack_Req = 0.6
Installation_Crack_Req = 0.6


Settings_Switch = True
if Settings_Switch:
    Output_File = "03_Minimax_Cracks"
    Inspection_File = ".res"
    Minimaxes=[1,2,3,4,5,7,8]
    Op_Minimaxes=[1,2,7]
    In_Minimaxes=[5]
    To_Minimaxes=[3,4,8]
    Minimaxes.sort()
    print("\n")
    Output_File_TimeStamp = False
    scriptpath = os.path.abspath(__file__)
    scriptdirectory = os.path.dirname(scriptpath)
    os.chdir(scriptdirectory)
    DG_List = []
    Database = {}
    from datetime import datetime
    Delta_Time_Start = datetime.now()


for consecna_1_file in glob.iglob(os.getcwd() + '/**/*consecna_1**', recursive=True):
    DG_N = len(os.path.basename(consecna_1_file).split(".")[1])
    break

    
for csm in glob.iglob(os.getcwd() + '/**/*.csm', recursive=True):
    DG = int(csm.split('mm')[0].strip()[len(csm.split('mm')[0]) - DG_N:len(csm.split('mm')[0])])
    Minimax = int(csm.split('mm')[1].replace('.res.csmr','').split('.')[0])
    if not DG in Database:
        Database[DG] = {}
        DG_List.append(DG)
        DG_List.sort()
    if Minimax in Minimaxes:
        if not Minimax in Database[DG]:
            Database[DG][Minimax] = {}
            Database[DG][Minimax]["Crack_Value"] = 0
            Database[DG][Minimax]["Crack_LC"] = 'No Crack'
            Database[DG][Minimax]["Crack_Element"] = ''
            Database[DG][Minimax]["Crack_Node"] = ''
            Database[DG][Minimax]["Thermal"] = ''
    with open(csm, 'r',encoding='ansi') as csm_In:
        csm_lines=csm_In.readlines()
        if '\\upper\\' in csm:
            Thermal = "Up"
        if '\\lower2\\' in csm:
            Thermal = "L2"
        if '\\lower\\' in csm:
            Thermal = "Lo"
        DG = int(csm.split('mm')[0].strip()[len(csm.split('mm')[0]) - DG_N:len(csm.split('mm')[0])])        
        Minimax = int(csm.split('mm')[1].replace('.res.csmr','').split('.')[0])
        if Minimax in Minimaxes:
            Outer_Read_Switch = False
            Inner_Read_Switch = False
            for csm_line in csm_lines:
                if 'Summary results for' in csm_line:
                    Crack_LC_Temp =  csm_line.strip().replace('T','').replace('W','').replace('Y','').replace('X','').split('loadcase :')[1].strip()
                    Element_Temp = csm_line.split("Summary results for el :")[1].strip().split()[0]
                    Node_Temp = csm_line.split("point :")[1].strip().split()[0]
                if Inner_Read_Switch:
                    Crack_Value_Temp = float(csm_line.strip().split()[4])
                    Inner_Read_Switch = False
                    if abs(Crack_Value_Temp) > abs(float(Database[DG][Minimax]["Crack_Value"])):
                        Database[DG][Minimax]["Crack_Value"] = Crack_Value_Temp
                        Database[DG][Minimax]["Crack_LC"] = Crack_LC_Temp
                        Database[DG][Minimax]["Crack_Element"] = Element_Temp
                        Database[DG][Minimax]["Crack_Node"] = Node_Temp
                        Database[DG][Minimax]["Thermal"] = " " + Thermal
                if Outer_Read_Switch:
                    Crack_Value_Temp = float(csm_line.strip().split()[4])
                    Outer_Read_Switch = False
                    Inner_Read_Switch = True
                    if abs(Crack_Value_Temp) > abs(float(Database[DG][Minimax]["Crack_Value"])):
                        Database[DG][Minimax]["Crack_Value"] = Crack_Value_Temp
                        Database[DG][Minimax]["Crack_LC"] = Crack_LC_Temp  
                        Database[DG][Minimax]["Crack_Element"] = Element_Temp
                        Database[DG][Minimax]["Crack_Node"] = Node_Temp
                        Database[DG][Minimax]["Thermal"] = " " + Thermal
                if "wok" in csm_line:
                    Outer_Read_Switch = True


from datetime import datetime
Output_File_Time_Stamp = str(datetime.today().strftime('%Y%m%d'))
Output_File = Output_File + '_' + Output_File_Time_Stamp + '.xls'
Workbook = xlwt.Workbook(Output_File)
Worksheet_Op = Workbook.add_sheet('Operation Cracks',cell_overwrite_ok=True)
Worksheet_To = Workbook.add_sheet('Towing Cracks',cell_overwrite_ok=True)
Worksheet_In = Workbook.add_sheet('Installation Cracks',cell_overwrite_ok=True)
Worksheet_Op.write(0,0,"DG"),Worksheet_To.write(0,0,"DG"),Worksheet_In.write(0,0,"DG")
Worksheet_Op.write(0,10,"Max Cr"),Worksheet_To.write(0,10,"Max Cr"),Worksheet_In.write(0,10,"Max Cr")
i=0
for i in range(len(Op_Minimaxes)):
    Worksheet_Op.write(0,(i*3)+1,"Crack")
    Worksheet_Op.col((i*3)+1).width = 2560
    Worksheet_Op.write(0,(i*3)+2,"Minimax"+str(Op_Minimaxes[i]))
    Worksheet_Op.col((i*3)+2).width = 3840
    Worksheet_Op.write(0,(i*3)+3,"LC")
    Worksheet_Op.col((i*3)+3).width = 5120
    
    
i=0
for i in range(len(To_Minimaxes)):
    Worksheet_To.write(0,(i*3)+1,"Crack")
    Worksheet_To.col((i*3)+1).width = 2560
    Worksheet_To.write(0,(i*3)+2,"Minimax"+str(To_Minimaxes[i]))
    Worksheet_To.col((i*3)+2).width = 3840
    Worksheet_To.write(0,(i*3)+3,"LC")
    Worksheet_To.col((i*3)+3).width = 5120
i=0
for i in range(len(In_Minimaxes)):
    Worksheet_In.write(0,(i*3)+1,"Crack")
    Worksheet_In.col((i*3)+1).width = 2560
    Worksheet_In.write(0,(i*3)+2,"Minimax"+str(In_Minimaxes[i]))
    Worksheet_In.col((i*3)+2).width = 3840
    Worksheet_In.write(0,(i*3)+3,"LC")
    Worksheet_In.col((i*3)+3).width = 5120

for DG in Database:
    for Minimax in Op_Minimaxes:
            Row = DG_List.index(DG) + 1 
            Column = (3 * Op_Minimaxes.index(Minimax)) + 1
            Worksheet_Op.write(Row,0,DG)
            Worksheet_Op.write(Row,Column,Database[DG][Minimax]["Crack_Value"])
            Worksheet_Op.write(Row,Column+1,str(Database[DG][Minimax]["Crack_Element"]) + " " + str(Database[DG][Minimax]["Crack_Node"]))
            Worksheet_Op.write(Row,Column+2,str(Database[DG][Minimax]["Crack_LC"]) + str(Database[DG][Minimax]["Thermal"]))
            Worksheet_Op.write(Row,10,xlwt.Formula("MAX(B" + str(Row+1) + ",E" + str(Row+1) + ",H" + str(Row+1) + ")"))
for DG in Database:
    for Minimax in To_Minimaxes:
            Row = DG_List.index(DG) + 1 
            Column = (3 * To_Minimaxes.index(Minimax)) + 1
            Worksheet_To.write(Row,0,DG)
            Worksheet_To.write(Row,Column,Database[DG][Minimax]["Crack_Value"])
            Worksheet_To.write(Row,Column+1,str(Database[DG][Minimax]["Crack_Element"]) + " " + str(Database[DG][Minimax]["Crack_Node"]))
            Worksheet_To.write(Row,Column+2,str(Database[DG][Minimax]["Crack_LC"]) + str(Database[DG][Minimax]["Thermal"]))
            Worksheet_To.write(Row,10,xlwt.Formula("MAX(B" + str(Row+1) + ",E" + str(Row+1) + ",H" + str(Row+1) + ")"))
for DG in Database:
    for Minimax in In_Minimaxes:
            Row = DG_List.index(DG) + 1 
            Column = (3 * In_Minimaxes.index(Minimax)) + 1
            Worksheet_In.write(Row,0,DG)
            Worksheet_In.write(Row,Column,Database[DG][Minimax]["Crack_Value"])
            Worksheet_In.write(Row,Column+1,str(Database[DG][Minimax]["Crack_Element"]) + " " + str(Database[DG][Minimax]["Crack_Node"]))
            Worksheet_In.write(Row,Column+2,str(Database[DG][Minimax]["Crack_LC"]) + str(Database[DG][Minimax]["Thermal"]))
            Worksheet_In.write(Row,10,xlwt.Formula("MAX(B" + str(Row+1) + ",E" + str(Row+1) + ",H" + str(Row+1) + ")"))


Workbook.save(Output_File)

print("Done!!! " + str(datetime.now()-Delta_Time_Start) + "...")
