# -*- coding: utf-8 -*-
"""
Created on Wed May 22 13:25:38 2019

@author: SankalpMishra
"""
import numpy as np
import pandas as pd
#%%
class ReadData:
    def __init__(self, path, allowance_file, cno_file, times_file):
        self.df_allowance = pd.read_excel(path + allowance_file + '.xlsx')
        self.df_cno = pd.read_excel(path + cno_file + '.xlsx')
        self.df_times = pd.read_excel(path + times_file + '.xlsx')

#%%
class cno:
    def __init__(self, cno, df_cno):
        self.cno = cno
        self.cno_present = df_cno[df_cno['cno'] == self.cno]

#%%

class duro:
    
    def __init__(self, cno_present = [], df_allowance = [], df_time = []):
        self.opa = df_allowance['opa'].values[0]
        self.mta = df_allowance['mta'].values[0]
        self.ba = df_allowance['ba'].values[0]
        
        self.cno = cno_present['cno'].values[0]
        self.blend = cno_present['blend'].values[0]
        self.yno = cno_present['yno'].values[0]
        self.ply = cno_present['ply'].values[0]
        self.package_type = cno_present['package_type'].values[0]
        self.pkg_wt = cno_present['pkg_wt'].values[0]
        self.windspeed = cno_present['windspeed'].values[0]
        self.blend_det = cno_present['blend_detail'].values[0]
        
        self.collect_package = df_time['collect_package'].values[0]
        self.place_newcreel = df_time['place_newcreel'].values[0]
        self.place_newcone= df_time['place_newcone'].values[0]
        self.get_traypack = df_time['get_traypack'].values[0]
        self.get_divider = df_time['get_divider'].values[0]
        self.get_creels = df_time['get_creels'].values[0]
        self.deliver = df_time['deliver'].values[0]
        self.cleaning = df_time['cleaning'].values[0]
        self.numspin = df_time['num_spin'].values[0]
        self.break_ratio = df_time['break_ratio'].values[0]
        self.intf = df_time['intf'].values[0]
                            
        
        
    def rt(self):
        return self.pkg_wt*self.yno*840/(self.windspeed*self.ply)
    
    def ct(self):
        return (self.rt() + self.collect_package + self.place_newcreel +\
               self.place_newcone + self.cleaning*self.rt()/480 + \
               self.intf*self.rt())*self.mta*((self.ba-1)*self.ply+1)*self.opa
                
    def maxlb(self):
        return 480*self.pkg_wt/(self.rt())
    
    def explb(self):
        return 480*self.pkg_wt/self.ct()
    
    def mecheff(self):
        return self.rt()*100/self.ct()
    
    def opt(self):
        return self.collect_package + self.place_newcreel+ self.place_newcone + \
               self.get_traypack/100 + self.get_divider/25 + self.get_creels + \
               self.deliver/100 + self.cleaning*self.rt()/480 + \
               self.break_ratio*self.ply*self.rt()
               
    def stdminperlb(self):
        return self.opt()*480/(430*self.pkg_wt)
    
    def spinass(self):
        return (430/(480*self.opt()/self.ct()))
    
    
#%%
path_duro = 'G:/Standards/Main/tables_excel/'
duro_data = ReadData(path_duro, 'dsduro_allowance_tbl', 'dsduro_cno_tbl', 'dsduro_times_tbl')

cols = ['cno', 'yno', 'ply', 'blend', 'pkg_type', \
        'pkg_weight', 'cycle time','100 % lb', 'exp lb', 'mach eff(%)', 'stdmin/lb', 'spindle assignment']

summarylist = []

for i, row in duro_data.df_cno.iterrows():
    cno_x  = duro_data.df_cno.iloc[i]['cno']
    cno_present = cno(cno_x, duro_data.df_cno)
    cno_present = cno_present.cno_present
    
    duro_sum = duro(cno_present, duro_data.df_allowance, duro_data.df_times)
    
    print(duro_sum.ct(), duro_sum.maxlb(), duro_sum.explb(), duro_sum.mecheff(),\
          duro_sum.stdminperlb(), duro_sum.spinass())
    
    
    x = pd.DataFrame([[cno_present['cno'].values[0], cno_present['yno'].values[0], \
                       cno_present['ply'].values[0], cno_present['blend_detail'].values[0], \
                       cno_present['package_type'].values[0], cno_present['pkg_wt'].values[0],\
                       duro_sum.ct(), duro_sum.maxlb(), duro_sum.explb(), \
                       duro_sum.mecheff(), duro_sum.stdminperlb(), duro_sum.spinass()]], columns = cols)
    
    summarylist.append(x)
    

df = pd.concat(summarylist)
#%%
for i , row in df.iterrows():
    print(row)

#%%
out_path = 'H:\costing_automation\ds_durowinder_summary.xlsx'
writer = pd.ExcelWriter(out_path , engine='xlsxwriter')
df.to_excel(writer, sheet_name = 'Sheet1', index = False)
    
            
        
        
        
        