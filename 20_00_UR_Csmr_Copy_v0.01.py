"""
@author: Antl
Date   :
"""
import os
import glob

Ignored_MMs = []

# Run_Final icerisinde calistir.
# Dgler dg klasorunde olsun

Target_Directory = "\\UR_Plot"
scriptpath = os.path.abspath(__file__)
scriptdirectory = os.path.dirname(scriptpath)
os.chdir(scriptdirectory)
if not os.path.isdir(str(scriptdirectory) + Target_Directory):
    os.makedirs(str(scriptdirectory) + Target_Directory)
for Csmr_File in glob.iglob(os.getcwd() +"\\dg\\**\\lower\\**.csmr", recursive=True):
    MM = int(os.path.basename(Csmr_File).split("mm")[1].split(".")[0])
    if not os.path.isdir(str(scriptdirectory) + Target_Directory + "\\" + str(Csmr_File.replace(scriptdirectory,"").replace(os.path.basename(Csmr_File),""))):
        os.makedirs(str(scriptdirectory) + Target_Directory + "\\" + str(Csmr_File.replace(scriptdirectory,"").replace(os.path.basename(Csmr_File),"")))
    if not MM in Ignored_MMs:
        os.system('copy "' + str(Csmr_File)+'" "' + str(scriptdirectory) + str(Target_Directory) + str(Csmr_File.replace(scriptdirectory,"")) + '"')
