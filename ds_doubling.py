# -*- coding: utf-8 +*-
"""
Created on Tue May 12 08:33:43 2020

@author: SankalpMishra
"""

import numpy as np
import pandas as pd

#%%

class cno_doub:
    def __init__(self, cno_a, df_cno = []):
        self.cno_a = cno_a
        self.cno = df_cno
        self.cno_prst = df_cno[df_cno['const'] == cno_a]
        
        
class doubler:
    op_allowance = 1.096692
    mt_allowance = 1.070664
    
    