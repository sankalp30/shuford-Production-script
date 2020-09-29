# -*- coding: utf-8 -*-
"""
Created on Tue May 21 14:11:51 2019

@author: SankalpMishra

operator assignments are not calculated for conditioning.
"""

import numpy as np
import pandas as pd
#%%
class ReadData:
    '''
    file names should be passed as strings (in quotations)
    '''
    def __init__(self, path, cno_file, conditiontime_file, pkgincrate_file):
        self.df_cno = pd.read_excel(path + cno_file + '.xlsx')
        
        self.df_pkgincrate = pd.read_excel(path + pkgincrate_file + '.xlsx')
        self.df_conditiontime = pd.read_excel(path + conditiontime_file + '.xlsx')
        
#%%        
class cno:
    '''
    unpacking variables from cno_tbl for the passed value of construction number
    '''
    
    def __init__(self, cno, df_cno):
        self.cno = cno
        self.cno_present = df_cno[df_cno['cno'] == self.cno]
        
#%%    
class condition:
    def __init__(self, cno_present = [], df_conditiontime = [], df_pkgincrate = [],default_time = 180):
        self.cno_present = cno_present
        self.default_time = default_time
        self.df_conditiontime = df_conditiontime
        self.df_pkgincrate  = df_pkgincrate
        self.pkg_wt = self.cno_present['pkg_wt'].values[0]
        self.input_time = self.df_conditiontime['input_time'].values[0]
        self.output_time =self.df_conditiontime['output_time'].values[0]
        
    def rt(self):
        return self.default_time
    
    def ct(self):
        return self.default_time + self.input_time + self.output_time
    
    def pkgincrate(self):
        z = self.df_pkgincrate[(self.df_pkgincrate['ul']>self.pkg_wt) & \
                          (self.df_pkgincrate['ll']<= self.pkg_wt)]
        return z['pkgcrate'].values[0]
    
    def maxlb(self):
        return (2*self.pkg_wt*self.pkgincrate()*480/(self.default_time))
    
    def explb(self):
        return(2*self.pkg_wt*self.pkgincrate()*480/self.ct())

    def mecheff(self):
        return self.default_time*100/self.ct()

    def stdmin(self):
        return (self.input_time + self.output_time)*480/(2*self.pkg_wt*self.pkgincrate()*430)
        
#%%
path_cond = 'G:/Standards/Main/tables_excel/'

cond = ReadData(path_cond, 'dscond_cno_tbl', 'dscond_times_tbl', 'dscond_pkgincrate_tbl')

#%%
summary_list = []
cols = ['cno', 'yno', 'blend', 'pkg_type', \
        'pkg_weight', 'condition time', '100 % lb', 'exp lb', 'mach eff(%)', 'stdmin/lb']

for i, rows in cond.df_cno.iterrows():
    cno_x = cond.df_cno.iloc[i]['cno']
    cno_present = cno(cno_x, cond.df_cno)
    cno_present = cno_present.cno_present
    
    cond_sum = condition(cno_present, cond.df_conditiontime, cond.df_pkgincrate)
    
    print(cond_sum.ct(), cond_sum.maxlb(), cond_sum.explb(), cond_sum.mecheff(), cond_sum.stdmin())
    
     
    x= pd.DataFrame([[cno_present['cno'].values[0], cno_present['yno'].values[0], cno_present['blend'].values[0], \
                      cno_present['pkg_type'].values[0], cno_present['pkg_wt'].values[0], \
                      cond_sum.rt(), cond_sum.maxlb(), cond_sum.explb(), cond_sum.mecheff(), cond_sum.stdmin()]], columns = cols)
    
    summary_list.append(x)
    
df = pd.concat(summary_list)

#%%
for i, row in df.iterrows():
    print(i, row)
#%%

out_path = 'I:\costing_automation\ds_condition_summary.xlsx'
writer = pd.ExcelWriter(out_path , engine='xlsxwriter')
df.to_excel(writer, sheet_name = 'Sheet1', index = False)

    
    
    
    

    