# -*- coding: utf-8 -*-
"""
Created on Thu May 23 10:35:17 2019

@author: SankalpMishra
"""

import numpy as np
import pandas as pd

#%%
class ReadData:
    def __init__(self, path, allowance_file, cno_file, times_file):
        self.df_allowance = pd.read_excel(path + allowance_file +'.xlsx')
        self.df_cno = pd.read_excel(path + cno_file + '.xlsx')
        self.df_times = pd.read_excel(path + times_file + '.xlsx')
        
#%%        
class cno:
    def __init__(self, cno, df_cno):
        self.cno = cno
        self.cno_present = df_cno[df_cno['cno'] == self.cno]

#%%
class xeno:
    def __init__(self, cno_present = [], df_allowance = [], df_times = []):
        
        self.opa = df_allowance['opa'].values[0]
        self.mta = df_allowance['mta'].values[0]
        self.ba = df_allowance['ba'].values[0]
        
        self.cno = cno_present['cno'].values[0]
        self.blend_detail = cno_present['blend_detail'].values[0]
        self.yno = cno_present['yno'].values[0]
        self.ply = cno_present['ply'].values[0]
        self.package_type = cno_present['pkg_type'].values[0]
        self.pkg_wt = cno_present['pkg_wt'].values[0]
        self.windspeed = cno_present['wind_speed'].values[0]
        self.creel_weight = cno_present['creel_weight'].values[0]
        
        self.place_crates = df_times['place_crates'].values[0]
        self.tie_off = df_times['tie_off'].values[0]
        self.change_crates = df_times['auto_doff'].values[0]
        self.creel = df_times['creel'].values[0]
        self.get_creels = df_times['get_creels'].values[0]
        self.tbst = df_times['tbst'].values[0]
        self.mark_tubes = df_times['mark_tubes'].values[0]
        self.delivery = df_times['delivery'].values[0]
        self.intf = df_times['intf'].values[0]
        self.auto_doff = df_times['auto_doff'].values[0]
        
        
    def rt(self):
        return 840*self.yno*self.pkg_wt/(self.ply*self.windspeed)
    
    def ct(self):
        return (self.rt()+ self.place_crates/30 + self.tie_off + self.change_crates/30+\
               self.creel*self.pkg_wt/self.creel_weight + self.get_creels + \
               self.tbst + self.mark_tubes + self.auto_doff + \
               self.delivery)*self.mta*self.opa*self.ba
                
    def maxlb(self):
        return 480*self.pkg_wt/self.rt()
    
    def explb(self):
        return 480*self.pkg_wt/self.ct()
    
    def mecheff(self):
        return self.rt()*100/self.ct()
    
    def opt(self):
        return (self.place_crates + self.tie_off + self.change_crates + \
               self.creel*self.pkg_wt/self.creel_weight + self.get_creels + \
               self.tbst + self.mark_tubes + self.delivery)*self.ba
                
    def stdminperlb(self):
        return self.opt()*480/(430*self.pkg_wt)
    
    def spinass(self):
        return 430/(480*self.opt()/self.ct())
    
    
#%%

path_xeno = 'G:/Standards/Main/tables_excel/'
xeno_data = ReadData(path_xeno, 'dsxeno_allowance_tbl', 'dsxeno_cno_tbl', 'dsxeno_times_tbl')

cols = ['cno', 'yno', 'ply', 'blend', 'pkg_type', \
        'pkg_weight', 'cycle time','100 % lb per spindle', 'exp lb per spindle', 'mach eff(%)', 'stdmin/lb', 'spindle assignment']

summarylist = []

for i, row in xeno_data.df_cno.iterrows():
    cno_x  = xeno_data.df_cno.iloc[i]['cno']
    cno_present = cno(cno_x, xeno_data.df_cno)
    cno_present = cno_present.cno_present
    
    xeno_sum = xeno(cno_present, xeno_data.df_allowance, xeno_data.df_times)
    
    print(xeno_sum.ct(), xeno_sum.maxlb(), xeno_sum.explb(), xeno_sum.mecheff(),\
          xeno_sum.stdminperlb(), xeno_sum.spinass())
    
    
    x = pd.DataFrame([[cno_present['cno'].values[0], cno_present['yno'].values[0], \
                       cno_present['ply'].values[0], cno_present['blend_detail'].values[0], \
                       cno_present['pkg_type'].values[0], cno_present['pkg_wt'].values[0],\
                       xeno_sum.ct(), xeno_sum.maxlb(), xeno_sum.explb(), \
                       xeno_sum.mecheff(), xeno_sum.stdminperlb(), xeno_sum.spinass()]], columns = cols)
    
    summarylist.append(x)
    

df = pd.concat(summarylist)        
#%%
for i, row in df.iterrows():
    print(row)
#%%
out_path = 'I:\costing_automation\ds_xenowinder_summary.xlsx'
writer = pd.ExcelWriter(out_path , engine='xlsxwriter')
df.to_excel(writer, sheet_name = 'Sheet1', index = False)
    
                
        
        
        
        
        
    
    
    