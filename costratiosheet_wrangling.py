# -*- coding: utf-8 -*-
"""
Created on Mon Oct  8 09:26:28 2018

@author: mishr
"""

import pandas as pd 
import numpy as np
#%%
path = 'G:/doubler project/Worksheet to create new cost ratios-DS  6-18 Aco Mod3.xlsx'
sheetname = 'Ratio Creation'
df = pd.read_excel(path, sheetname = sheetname, header = 22)
df = df[['Construction #','Fiber','Twist','Put-up Description']]
df = df.drop(0)
df['Put-up Description'] =df['Put-up Description'].astype(str).slice()
#%%
for i , rows in df.iterrows():
    if 'oz' in df.iloc