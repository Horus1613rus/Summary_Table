import os
import glob
import codecs

scriptpath = os.path.abspath(__file__)
scriptdirectory = os.path.dirname(scriptpath)
os.chdir(scriptdirectory)


encodings = ['ansi','utf-8','ascii','utf-16']


with open('0_Check_FXY.txt', 'w') as F_text_file_table:
    Header='DG\ttherm\tMinimax\tLC\n'
    F_text_file_table.write(Header)


for f in glob.iglob(os.getcwd() + '/**/*.csm', recursive=True):
    for e in encodings:
        try:
            Text_File = codecs.open(f, 'r', encoding=e)
            Text_File.readlines()
            Text_File.seek(0)
            with open(f, 'r',encoding=e) as fIn :
                lines=fIn.readlines()
                Line_Number = 0                    
                DG = f.split('mm')[0].strip()[-5:]
                therm = f.split('mm')[0].strip()[-13:-8]
                Minimax = f.split('mm')[1].split('.')[0]
                for line in lines:
                    Line_Number = Line_Number + 1
                    if 'loadcase :' in line:
                        LC = line.strip().strip()[-13:]
                    if 'Fxy' in line:
                        Fxyfound = line.strip().split()[2]
                        Fxygiven = line.strip().split()[1]
                        if float(Fxyfound)==0:
                            if not float(Fxygiven)==0:
                                text_Fxy = str(DG) +'\t' + str(therm) +'\t' + str(Minimax) +'\t' + str(LC) +'\n'
                                with open('0_Check_FXY.txt', 'a') as F_text_file_table:
                                    F_text_file_table.write(text_Fxy)
        except UnicodeDecodeError:
            print('F File:Got unicode error with %s , trying different encoding' % e)
        else:   
            print('F File:Opening ',DG,' with encoding:  %s ' % e)
            break
    