#!/usr/bin/env python
# coding: utf-8

# In[1]:

import pandas as pd
from openpyxl import Workbook, load_workbook
wb = load_workbook("CarData.xlsx")
ws = wb.active
f = pd.read_excel("CarData.xlsx")
f


# In[24]:


class state():
    def __init__(self, name):
        self.name = name
        self.num = 0
    def getAllValue(self):
        dict={}
        for i in range(23):       #find better way 
            dict[f.iloc[i,0]]= i
        
        print(f.iloc[dict[self.name]])
    

    def getValue(self, ind):
        a = "\n" + " "*23
        
        dict={}
        for i in range(23):       
            dict[f.iloc[i,0]]= i
        try:
            print(f.iloc[dict[self.name], [ind]])
        except IndexError:
            print("Error: Enter a number. 1 : Age 0-20" + a + "2 : Age 21-3" + a + "3 : Age 35-54" + a + \
                 "4 : 55+" + a + "5 : All ages" + a + "6 : Male" + a + "7 : Female")
            
    def setValue(self, col, newValue):
        dict={}
        for i in range(23):       
            dict[f.iloc[i,0]]= i
        dict2 = {1:"A", 2:"B", 3: "C", 4:"D", 5 : "E", 6 : "F", 7 : "G"}
        ws[str(dict2[col])+str(dict[self.name] + 2)] = newValue
        wb.save("CarData.xlsx")
        print(ws["B3"].value)
        print(dict2[4])
minn = state("Delaware")
minn.getAllValue()
minn.getValue("8")
minn.setValue(4, 3.0)


                


                






