# -*- coding: utf-8 -*- 
"""
Created on Fri Jun 21 14:11:02 2019

@author: SankalpMishra
"""

import numpy as np
import pandas as pd

#%%
path_waylon = 'H:/DS drawing/tables'
#%%
class ReadData:
    def __init__(self, path, blend_file, allowance_file, drawtime_file, spintype_file):
        self.df_blend = pd.read_excel(path + blend_file + '.xlsx')
        self.df_allowance = pd.read_excel(path + allowance_file + '.xlsx')
        self.df_drawtime = pd.read_excel(path + drawtime_file + '.xlsx')
        self.df_spintype = pd.read_excel(path + spintype_file + '.xlsx')
                
#%%%
class blendc:
    def __init__(self, blend, spin_type, df_blend = [], df_spin = []):
        self.blend = blend
        self.spin_type = spin_type
        self.df_blend = df_blend[(df_blend['Full Description'] == blend) & (df_blend['Process'] == spin_type)]        
        self.df_spin = df_spin[df_spin['abbrev'] == self.spin_type]
        
#%%
'''
unpacking variables
standards function        
'''

        
class drawing:

    def __init__(self, current_blend = [], df_allowance = [], df_drawtime = [], df_spin = []):
        self.current_blend = current_blend
        self.fiber = current_blend['Fiber/Blend'].values[0]
        self.fulld = current_blend['Full Description'].values[0]
        self.b_s = self.break_speed = current_blend['Breaker Drawing M/min'].values[0]
        self.f_s = current_blend['Finisher Drawing M/min'].values[0]
        self.break_speed = current_blend['Breaker Drawing M/min'].values[0]*1.0931
        self.finish_speed = current_blend['Finisher Drawing M/min'].values[0]*1.0931
        self.spin = current_blend['Process'].values[0]
        self.draw_sliver = current_blend['Sliver Size'].values[0]
        self.draw_can_wt = current_blend['draw_can_weight'].values[0]
        self.card_can_wt = current_blend['card_can_weight'].values[0]
        
        self.creeling = df_drawtime['creeling'].values[0]
        self.measurewt = df_drawtime['measure_sliver_weight'].values[0]
        self.stockcan = df_drawtime['stock_empty_cans'].values[0]
        self.break_repair = df_drawtime['break_repair'].values[0]
        self.lightfix = df_drawtime['stop_light_fix'].values[0]
        self.cleaning = df_drawtime['cleaning'].values[0]
   #    self.recordwt = df_drawtime['entries_on_sheet'].values[0]
        self.doff = df_drawtime['auto_doff'].values[0]
        self.avg_breaks = df_drawtime['avg_breaks'].values[0]
        self.deliver = df_drawtime['deliver'].values[0]
        
        self.mta = df_allowance['mta'].values[0]
        self.opa = df_allowance['opa'].values[0]
        self.cha = df_allowance['cha'].values[0]
        
        self.spin = df_spin['spin_type'].values[0]
        self.spin_type = df_spin['abbrev'].values[0]
        self.passes = current_blend['.Passes'].values[0] #passes depend on fiber. ex: OE spun yarn can have 1 or 2 passes
        
        self.ncc = 6
        self.ndc = 8
        self.delay = 5.28
        
  
    def hankperyard(self):
        return 7000/(840*self.draw_sliver)
        
    def break_runtime(self):
        return (7000*self.draw_can_wt/(self.draw_sliver*self.break_speed))
        
    def break_cycletime(self):
        return (self.break_runtime() + self.creeling*self.draw_can_wt/(self.ncc*self.card_can_wt) + \
               (self.delay + self.break_repair)*self.avg_breaks + self.doff)*self.opa*self.cha*self.mta
        
    def maxlbs(self):
        return (480*self.draw_can_wt/(self.line_runtime()))
            
    def finish_runtime(self):
        return 7000*self.draw_can_wt/(self.draw_sliver*self.finish_speed)
        
    def finish_cycletime(self):
        return (self.finish_runtime() + self.creeling/(self.ndc) + \
               (self.delay +self.break_repair)*self.avg_breaks + \
                self.doff)*self.mta*self.opa*self.cha
                
    def finish_factor(self):
        if self.passes == 1:
            return 0
        else:
            return 1
            
    def breaker_opt(self):
        if self.passes ==1:
            return 0
        else:
            return (self.passes-1)
        
    def line_runtime(self):
        return (max(self.break_runtime(), self.finish_runtime()*(self.finish_factor())))
    
    def line_cycletime(self):
        return max(self.break_cycletime(), self.finish_cycletime()*self.finish_factor())
    
    def explbs(self):
        return (480*self.draw_can_wt/(self.line_cycletime()))      
        
    def line_eff(self):
        return (self.line_runtime()*100/self.line_cycletime())
 
    def opt_time(self):
        return (self.creeling*self.draw_can_wt/(self.ncc*self.card_can_wt) + self.creeling*self.breaker_opt()/self.ndc + \
                self.measurewt*self.passes*self.break_runtime()/480 +\
                self.passes*self.stockcan + self.deliver*self.passes +\
                self.cleaning*self.break_runtime() + self.break_repair*self.avg_breaks*self.passes)*self.cha 
                    
    def stdminperlb(self):
        return (self.opt_time()*480/(430*self.draw_can_wt))
    
    def op_ass(self):
        return (430*self.passes/self.opt_time())/(480/self.line_cycletime())      


            
        
#%% Main 
path_draw = 'H:/tables_excel/' # check everytime
blend_file = 'dsdr_blenddata_tbl'
allowance_file = 'dsdr_allowance_tbl'
drawtime_file = 'dsdr_drawtimes_tbl'
spintype_file = 'dsdr_spintype_tbl'

cols = ['blend', 'draw_speed', 'spin_type', 'passes', 'sliver grain', 'cycletime', 'maxlbs'\
        , 'explbs', 'mech eff', 'stdmin', 'frame assignment']


#%%
draw_one = ReadData(path_draw, blend_file, allowance_file, drawtime_file, spintype_file)
df_blend = draw_one.df_blend
df_spin = draw_one.df_spintype


#%% this part only required to be run for operator standards at one pass for each product

for i , row in df_blend.iterrows():
    df_blend.iloc[i]['.Passes'] = 1   # iloc works fine as a replacement to ix
    


#%%    
summary_list = []
check_spin=[]

for i, row in draw_one.df_blend.iterrows():
    blend = draw_one.df_blend.iloc[i]['Full Description']
    spin_type = draw_one.df_blend.iloc[i]['Process']
    
    
    blend_current = blendc(blend, spin_type, df_blend = df_blend, df_spin = df_spin)
    check_spin.append(blend_current.df_blend)## test list to check if all spintypes are looped on for same blend types
    
    
    drawing_std = drawing(current_blend = blend_current.df_blend, \
                          df_allowance = draw_one.df_allowance, df_drawtime = draw_one.df_drawtime,\
                          df_spin = blend_current.df_spin)
    
    
    print(drawing_std.hankperyard(), drawing_std.break_runtime(), drawing_std.break_cycletime(),\
          drawing_std.maxlbs(), drawing_std.finish_runtime(), drawing_std.finish_cycletime(),\
          drawing_std.breaker_opt(), drawing_std.breaker_opt(), drawing_std.line_runtime(),\
          drawing_std.line_cycletime(), drawing_std.explbs(), drawing_std.line_eff(), drawing_std.stdminperlb(),\
          drawing_std.op_ass())
    
    X = pd.DataFrame([[drawing_std.fulld, drawing_std.b_s, spin_type, \
                       drawing_std.passes, drawing_std.draw_sliver,drawing_std.line_cycletime(),\
                       drawing_std.maxlbs(), drawing_std.explbs(), drawing_std.line_eff(), \
                       drawing_std.stdminperlb(), drawing_std.op_ass()]], columns = cols)
    
    
    summary_list.append(X)
    
df = pd.concat(summary_list)

#%% Test case
try_one = blendc('100% Polyester Recycled', 'AJ', df_blend = df_blend, df_spin = df_spin)
try_draw = drawing(current_blend = try_one.df_blend, df_allowance = draw_one.df_allowance, df_drawtime =  draw_one.df_drawtime, df_spin = draw_one.df_spintype)

print(try_draw.current_blend, try_draw.break_runtime(), try_draw.break_cycletime(), try_draw.maxlbs(), try_draw.explbs(), try_draw.opt_time(), try_draw.op_ass())
#%%
out_path = 'H:\costing_automation\ds_drawing_summary.xlsx'
writer = pd.ExcelWriter(out_path, engine='xlsxwriter')
df.to_excel(writer, sheet_name = 'Sheet1', index = False)        
