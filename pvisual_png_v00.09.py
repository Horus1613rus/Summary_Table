# -*- coding: utf-8 -*-
"""
Created on Sat Mar 21 17:08:13 2020

@author: Antl
"""
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os

#%%
#Definition of Member
member_name="CENTS1NE"                #Plotting member 
Plane="XY"                          #Plane of member 
LC="3651"                           #LoadCombination
database_version='09.41d'           #Version of Database

#Definition of limit condition 
run_gtforc=True                     #True while first run, False while limit run 
force_limit=False                   #Condition of limiting force (True=limit, False=not limit) 
limit_force="MXX"                   #Force to be limited (forces: FXX, FYY, FXY, MXX, MYY, MXY, VXZ, VYZ)
f_max=100                           #max limit of force
f_min=-100                          #min limit of force

#Label Sizes
node_label_size=1                   #Label of nodes size for node plot (Recommended between 1-2)
f_label_size=0.3                    #Label of forces size for force plot (Recommended between 0.3-0.5)



#%%
Plot_Folder= member_name + "_" + LC + "_Plots"
scriptpath=os.path.abspath(__file__)
scriptdirectory=os.path.dirname(scriptpath)
os.chdir(scriptdirectory)
directory=os.getcwd()
path=os.path.join(directory,Plot_Folder)
Dev_Switch = False

try:  
    os.mkdir(path)  
except OSError as error: 
    pass



for root, dirs, files in os.walk(directory):
    for file in files:
        if file.endswith('.dg.txt') and member_name in file:
            member_file_name=root + "/" +file
            
#%%repairing function 
def coordinates_repair(df_member,df_repair,limit):
    for ind in df_member.index:
        ver_sum=0
        hor_sum=0
        ver_repair_list=[]
        hor_repair_list=[]
        ver=df_member[ver_axis][ind]
        hor=df_member[hor_axis][ind]
        for indd in df_member.index:
            ver_check=df_member[ver_axis][indd]
            hor_check=df_member[hor_axis][indd]
            if abs(ver_check-ver)<limit:
                ver_sum=ver_sum+ver_check
                ver_repair_list.append(ver_check)
            if abs( hor_check- hor)<limit:
                 hor_sum= hor_sum+ hor_check
                 hor_repair_list.append( hor_check)
        ver_average=ver_sum/len(ver_repair_list)
        hor_average=hor_sum/len(hor_repair_list)
        df_repair["Wall"][ind]=df_member["Wall"][ind]
        df_repair["Node"][ind]=df_member["Node"][ind]
        df_repair[out_axis][ind]=df_member[out_axis][ind]
        df_repair[hor_axis][ind]=int(hor_average)
        df_repair[ver_axis][ind]=int(ver_average)
        df_repair["DG"][ind]=df_member["DG"][ind]
    table_repair=df_repair.pivot(index=ver_axis, columns=hor_axis, values="Node" )
    return df_repair, table_repair




#%% Read member coordinates text
#member_file_name=member_name+".dg.txt"
member_fig=LC+".png"
member_fig2="00_Nodes.png"


gt_file=member_name+".txt"
out_gt_file=gt_file+"."+LC
run_gt="gtforc "+gt_file+" "+LC

df_member= pd.read_csv(member_file_name, sep=" ", header=None, names=["Wall","Node","X","Y","Z","DG"])
index=list(df_member.index.values)
ver_axis=Plane[1]
hor_axis=Plane[0]

for axis in ["X","Y","Z"]:
    if axis!=ver_axis and axis!=hor_axis:
        out_axis=axis

value_ver_axis=df_member[ver_axis].tolist()
value_ver_axis=list( dict.fromkeys(value_ver_axis))

value_hor_axis=df_member[hor_axis].tolist()
value_hor_axis=list( dict.fromkeys(value_hor_axis))

df_member=df_member.drop_duplicates([hor_axis,ver_axis,out_axis])
table_node=df_member.pivot(index=ver_axis, columns=hor_axis, values="Node" )
repair_coordinates=table_node.isnull().values.any()

if repair_coordinates==True:
    limit=100
    df_repair=pd.DataFrame(index=index,columns=["Wall","Node","X","Y","Z","DG"])
    df_repair,table_repair=coordinates_repair(df_member,df_repair,limit)
    




#%%plot
if run_gtforc==True:

    if Dev_Switch:
        fig = plt.figure()
        plt.scatter(df_member[Plane[0]], df_member[Plane[1]], s=5)
        plt.title(member_name+" Location of Nodes")
        fig.savefig(path+"\\"+member_fig, format='png', dpi=600)
       
    groups = df_member.groupby('DG')
    fig2, ax =plt.subplots()
    for name, group in groups:
        ax.plot(group[hor_axis], group[ver_axis], marker='o', linestyle='', ms=2, label=name)
    legend=ax.legend(loc='center left', bbox_to_anchor=(1, 0.5))
    plt.title(member_name+" Location and Labels of Nodes")
    for label, x, y in zip(df_member["Node"], df_member[Plane[0]], df_member[Plane[1]]):
        ax.annotate(label, xy = (x + 0, y - 0),fontsize =node_label_size)
    plt.savefig(path+"\\"+member_fig2, format='png', bbox_extra_artists=(legend,), bbox_inches='tight', dpi=600) 

    os.system('alnovatek ' + database_version)
    outf=open(gt_file, "w+")
    with open(member_file_name) as member_file:
        lines=member_file.readlines()
        for line in lines:
            gtforc=line.split()[0]+" "+line.split()[1]
            outf.write(gtforc)
            outf.write("\n")
    outf.close()
    os.system(run_gt)
#%%
member_list=[]
node_list=[]
LC_list=[]
Fxx_list=[]
Fyy_list=[]
Fxy_list=[]
Mxx_list=[]
Myy_list=[]
Mxy_list=[]
Vxz_list=[]
Vyz_list=[]

with open(out_gt_file) as outfile:
    lines=outfile.readlines()
    for line in lines:
        member_list.append(line.split()[0])
        node_list.append(line.split()[1])
        LC_list.append(line.split()[2])
        Fxx_list.append(float(line.split()[3]))
        Fyy_list.append(float(line.split()[4]))
        Fxy_list.append(float(line.split()[5]))
        Mxx_list.append(float(line.split()[6]))
        Myy_list.append(float(line.split()[7]))
        Mxy_list.append(float(line.split()[8]))
        Vxz_list.append(float(line.split()[9]))
        Vyz_list.append(float(line.split()[10]))
        
df_force=pd.DataFrame({"Node":node_list,
                      hor_axis:df_member[hor_axis],
                      ver_axis:df_member[ver_axis],
                      "FXX":Fxx_list,
                      "FYY":Fyy_list,
                      "FXY":Fxy_list,
                      "MXX":Mxx_list,
                      "MYY":Myy_list,
                      "MXY":Mxy_list,
                      "VXZ":Vxz_list,
                      "VYZ":Vyz_list})
forces=["FXX","FYY", "FXY", "MXX", "MYY", "MXY", "VXZ","VYZ"]
if run_gtforc==True:        
    for Force in forces:
        member_fig3="0"+str(forces.index(Force)+1)+"_"+Force+".png"
        
        fig3, ax =plt.subplots()
        cm = plt.cm.get_cmap('nipy_spectral')
        x=df_force[hor_axis]
        y=df_force[ver_axis]
        c=c=df_force[Force]
        sc = plt.scatter(x,y,c=c, cmap=cm, s=10)
        plt.title(member_name+" "+Force+" "+LC)
        plt.colorbar(sc)
        for label, x, y in zip(df_force[Force], df_force[Plane[0]], df_force[Plane[1]]):
            ax.annotate(label, xy = (x + 0, y - 0),fontsize =f_label_size)
        fig3.savefig(path+"\\"+member_fig3, format='png',dpi=1000)
        
    os.remove("_gtforc.res")
    os.remove("_gtforc.run")
    os.remove(member_name + ".txt")
    os.remove("alias.dat")
        
    
#%%    
member_fig5="0"+str(forces.index(limit_force)+1)+"_"+limit_force+ " Scale_Limit "+str(f_min)+" ; "+str(f_max)+".png"
if force_limit==True:
    fig5=plt.figure()
    cm = plt.cm.get_cmap('nipy_spectral')
    x=df_force[hor_axis]
    y=df_force[ver_axis]
    c=c=df_force[limit_force]
    sc = plt.scatter(x,y,c=c, cmap=cm, s=10, vmin=f_min, vmax=f_max)
    plt.title(member_name+" "+limit_force+" "+LC)
    plt.colorbar(sc)
    fig5.savefig(path+"\\"+member_fig5, format='png',dpi=1000) 



