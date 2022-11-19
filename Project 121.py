#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
excel1 = pd.read_excel("CarData.xlsx")
excel1


# In[24]:


class state():
    def __init__(self, name):
        self.name = name
        
    def getAllValue(self):
        dict={}
        for i in range(23):       #find better way 
            dict[f.iloc[i,0]]= i
        
        print(f.iloc[dict[self.name]])
        
    def add(self, name):
        

    
minn = state("Minnesota")
minn.getAllValue()


                


# In[ ]:




