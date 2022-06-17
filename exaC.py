
from xlrd import open_workbook
import numpy as np
import sys


class ExaM:

    def __init__(self, datEx, nQ):
        
        self.datEx = datEx
        
        self.nQ = nQ

    def dist(self, dat,na, nuE):

        wb = open_workbook(self.datEx)
        
## Select sheet
        
        wsh = wb.sheet_by_name("Sheet1")

## Number of columns rows and colummns
        
        aa1 = np.linspace(2,3,2) #number of columns for questions
        bb1 = len(aa1)
        aa2= np.linspace(2,6,5)
        bb2=len(aa2)
        aa3=np.linspace(2,4,3)
        bb3=len(aa3)
        aa4=np.linspace(7,13,7)
        bb4=len(aa4)
        cc = np.linspace(0,self.nQ,self.nQ,endpoint=False) #number of rows
        b = len(cc)
        c = np.zeros(b)

        sys.stdout = open(dat, 'w')

    
        print ("                                  EXAM")
        print ('')
        print ('')
        print ('')
        print (r"Name: %s"%na )
        print ('')
        print ('')
        print(r"Num: %s"%nuE)
        print("")
        print("")
        

        
        for i in range(b):
            
            c[i] =  int(np.random.randint(0, 27))
            c1 = int(c[i])
            c2 = c[i]
            print("")
            print(r"question {}".format(i+1))

            if c1 <= 15:

                if c1!=c2:
                    vv = wsh.cell(c1,1)
                    print (vv.value)
                    print ('')

                    for j in range(bb1):
                        ab = int(aa1[j])
                        v = wsh.cell(c1,ab)
                        print ( v.value)
                        print ('')
                     
                elif c1==c2:
                    c1=int(c1+1)
                    vv = wsh.cell(c1,1)
                    print (vv.value)
                    print ('')
                

                    for j in range(bb1):
                        
                        ab = int(aa1[j])
                        v = wsh.cell(c1,ab)
                        print ('  ', v.value)
                        print ('')

            elif 16 <=c1 <=20 :

                if c1!=c2:
                    vv = wsh.cell(c1,1)
                    print (vv.value)
                    print ('')
                  

                    for j in range(bb1):
                        ab = int(aa1[j])
                        v = wsh.cell(c1,ab)
                        print ( v.value)
                        print ('')
                     
                elif c1==c2:
                    c1=int(c1+1)
                    vv = wsh.cell(c1,1)
                    print (vv.value)
                    print ('')
                   

                    for j in range(bb1):
                        
                        ab = int(aa1[j])
                        v = wsh.cell(c1,ab)
                        print ('  ', v.value)
                        print ('')
                        

            elif 21<=c1<=27:

                if c1!=c2:
                    vv = wsh.cell(c1,1)
                    print (vv.value)
                    print ('')

                    for j in range(bb3):
                        ab = int(aa3[j])
                        v = wsh.cell(c1,ab)
                        print ( v.value)
                        print("")
                        
                elif c1==c2:
                    c1=int(c1+1)
                    vv = wsh.cell(c1,1)
                    print (vv.value)
                    print ('')

                    for j in range(bb3):
                        
                        ab = int(aa3[j])
                        v = wsh.cell(c1,ab)
                        print ('  ', v.value)
                        print ('')

                elif c1==22:

                    vv=wsh.cell(c1,1)
                    print(vv.value)
                    print("")

                    for j in range(bb2):

                        ab= int(aa2[j])
                        v=wsh.cell(c1,ab)
                        print("".v.value)

                elif c1==25:

                    vv=wsh.cell(c1,1)
                    print(vv.value)
                    print("")

                    for j in range (7,14):

                        ab=int(aa4[j])
                        v=wsh.cell(24,ab)
                        print("",v.value)
                        
                                       

                    
                        

                        
                    

                  
  


  




