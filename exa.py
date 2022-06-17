from xlrd import open_workbook
import numpy as np
import sys
import os
from exaC import ExaM
import shutil as shT
import subprocess
import shlex


# Name of excel document

datEx = 'exa.xls'

#Four question of multiple choice

nQ = 5


ff = ExaM(datEx, nQ) #class


with open('list1.data') as f1:
    
    red = f1.readlines()
    c = len(red)
    cc = []
    

    for i in range(c):
        
        nI = i+1
        c1 = 'exa%s.tex'%nI
        cc.append(c1)        
        nuM = i+1
        a = ff.dist(c1,red[i], nuM)

    sys.stdout = open('runP', 'w')
    
        
    for i in range(c):
        
        nI2 = i+1
        print ("xelatex exa%s.tex"%nI2)
        st = os.stat("runP")
        os.chmod("runP", st.st_mode | 0o111)

    sys.stdout = open('runP1', 'w')
    
   
    for i in range(c):

        nI2 = i+1
        print ("xelatex %s.tex"%nI2)
        st = os.stat("runP1")
        os.chmod("runP1", st.st_mode | 0o111)
        sys.stdout = open('Ex.py', 'w')
        print( """
        import os
        import subprocess
        import shlex""")
       
