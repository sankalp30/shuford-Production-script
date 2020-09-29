# -*- coding: utf-8 -*-
"""
Created on Fri Aug 16 15:11:24 2019

@author: SankalpMishra
"""

import numpy as np
import pandas as pd
#%%
path_justin = ""

#%%
class ReadData:
    def __init__(self, path, blend_file, allowance_file, drawtime_file):
        self.df_blend = pd.read_excel(path + blend_file + '.xlsx')
        self.df_allowance = pd.read_excel(path + allowance_file + '.xlsx')
        self.df_drawtime = pd.read_excel(path + drawtime_file + '.xlsx')
        
#%%
class blend:
    def  __init__(self, abbrev, df_blend = []):
        self.abbrev = abbrev
        self.df_blend = df_blend[df_blend['abbrev'] == abbrev]
        
#%%
class HsDrawing:
    def __init__(self, current_blend = [], df_allowance = [], df_drawtime= []):
        self.current_blend = current_blend
        self.df_allowance = df_allowance
        self.df_drawtime = df_drawtime
        
        self.blend = current_blend['blend'].values[0]
        self.sliver_del = current_blend['sliver_del'].values[0]
        self.sliver_rec = current_blend['sliver_rec'].values[0]
        self.blenddet = current_blend['det'].values[0]
        self.drawcanwt = current_blend['draw can weight'].values[0]
        self.cardcanwt = current_blend['card can weight'].values[0]
        self.oldmachinespeed = current_blend['old_machine_speed'].values[0]
        self.newmachinespeed = current_blend['new machine speed'].values[0]
        
        self.opa = df_allowance['opa'].values[0]
        self.mta = df_allowance['mta'].values[0]
        self.cha = df_allowance['cha'].values[0]
        
        
        
    