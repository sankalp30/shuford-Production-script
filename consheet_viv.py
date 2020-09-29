# -*- coding: utf-8 -*-
"""
Created on Wed Oct  3 12:47:42 2018

@author: mishr

do not run the entire script at once. It contains a test case.
"""

import pandas as pd
import numpy as np
import re
from sklearn.preprocessing import LabelEncoder # use to convert wax variable
#%%
path = 'H:/doubler project/tables/cons_sheet_viv.xlsx'
xlsx = pd.ExcelFile(path)
cons_sheet = []
for sheet in xlsx.sheet_names:
    cons_sheet.append(xlsx.parse(sheet))
print(cons_sheet)
#%%
cols = ['const','tr','yno','ply','mech_tpi','tw_spindle_rpm', 'blend','windspeed','pkg_od','pkg_size','pkg_weight','length','wax',' tube_type', 'cone_type']
list = []
for x in xlsx.sheet_names:
    dfa = pd.read_excel(path, sheetname = x, header = None) # dont use sheet_name in attributes u idiot -_-
    const = dfa.iloc[6][1] #row, column
    tr = dfa.iloc[9][1]
    windspeed = dfa.iloc[16][0]
    pkg_od = dfa.iloc[16][3]
    pkg_size = dfa.iloc[31][1]
    pkg_weight = dfa.iloc[32][1]
    length = dfa.iloc[16][4]
    mech_tpi = dfa.iloc[21][1]
    
    
    if len(dfa.columns)<11:
        yno = dfa.iloc[6][4] #
        ply = dfa.iloc[6][6]#
        blend = dfa.iloc[6][8]#
        wax = dfa.iloc[30][9]#
        tw_spindle_rpm = dfa.iloc[21][9]#
        cone_type = dfa.iloc[27][5]
        tube_type = dfa.iloc[21][5]
    else: 
        yno = dfa.iloc[6][5] #
        ply = dfa.iloc[6][7]#
        blend = dfa.iloc[6][9]#
        wax = dfa.iloc[30][10]#
        tw_spindle_rpm = dfa.iloc[21][10]#
        cone_type = dfa.iloc[27][6]
        tube_type = dfa.iloc[21][6]
    dfb = pd.DataFrame([[const, tr, yno, ply, mech_tpi, tw_spindle_rpm, blend, windspeed, pkg_od, pkg_size, pkg_weight, length, wax, tube_type, cone_type]], columns = cols)
    print(dfb)
    list.append(dfb)
df = pd.concat(list)
#%% work on this block to get columns in correct datatype.
df['pkg_weight'] = df['pkg_weight'].fillna(0)
df['pkg_weight']= df['pkg_weight'].astype(str).str.lower() #df['BLEND'].astype(str).str.lower()
df['pkg_weight_lb'] = df['pkg_weight'].str.split().str.get(0)
df = df.dropna(subset = ['const'])
df['blend'] = df['blend'].astype(str).str.lower()
#%%difference between ix and iloc but ix is deprecated even though it stll works -_-
df = df.reset_index(drop = True)
non_decimal =  re.compile(r'[^\d.]+')
for i, rows in df.iterrows():
    x = df.iloc[i]['pkg_weight']
    y = df.iloc[i]['pkg_weight_lb']
    if str(y) != 'nan' and 'oz' in str(x) and '-' not in str(y) and 'x' not in str(y) and 'lb' not in str(y):
        df.ix[i,'pkg_weight_lb'] = float(y)*0.0625
    if 'lb' in str(y):
        df.ix[i,'pkg_weight_lb'] = str(df.iloc[i]['pkg_weight'])[0]
    if '-' in str(y):
        df.ix[i,'pkg_weight_lb'] = str(df.iloc[i]['pkg_weight'])[0:3]
    if 'x' in str(y):
        df.iloc[i]['pkg_size'] = str(df.iloc[i]['pkg_weight'])
        df.iloc[i]['pkg_weight'] = np.NaN
    if '"' in str(df.iloc[i]['pkg_od']):
        df.iloc[i]['pkg_od'] = non_decimal.sub('', str(df.iloc[i]['pkg_od']))
    if 'sup' in str(df.iloc[i]['blend']):
        df.ix[i, 'blend'] = 'sup'
    if 'cat' in str(df.iloc[i]['blend']):
        df.ix[i, 'blend'] = 'cat'
    #if 'lb' or 'oz' in str(df.iloc[i]['pkg_size']).str.lower():
     #   df.ix['pkg_weight'] = 
#%%
#if wax needs to be a binary variable
#%%


#%% TEST CASE-  Do not run! moron:zz
df = pd.read_excel(path, sheetname = '4-2 SKY 078506', header = None)
const = df.iloc[6][1]
tr = df.iloc[9][1]
yno = df.iloc[6][5]
ply = df.iloc[6][7]
blend = df.iloc[6][9]
windspeed = df.iloc[16][0]
pkg_od = df.iloc[16][3]
pkg_size = df.iloc[31][1]
pkg_weight = df.iloc[32][1]
length = df.iloc[16][4]
wax = df.iloc[30][9]
cone = df.iloc[27][5]

print(const, tr, yno, ply, blend, windspeed, pkg_od, pkg_size, pkg_weight, length, wax)
#%%
out_path = 'G:/doubler project/cons_viv_wrangled_twist.xlsx'
writer = pd.ExcelWriter(out_path , engine='xlsxwriter')
df.to_excel(writer, sheet_name = 'Sheet1', index = False)
#%% in-or corrections without regenerating the entire file. No Need to run
path = 'G:/doubler project/cons_viv_wrangled.xlsx'
df = pd.read_excel(path, header = 0)
#%%
df['pkg_weight_lb'] = df['pkg_weight_lb'].astype(float)
#%%
out_path = 'G:/doubler project/cons_viv_wrangled_twist.xlsx'
writer = pd.ExcelWriter(out_path , engine='xlsxwriter')
df.to_excel(writer, sheet_name = 'Sheet1', index = False)

















