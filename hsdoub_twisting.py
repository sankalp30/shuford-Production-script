# -*- coding: utf-8 -*-
"""
Created on Tue Apr  9 13:17:42 2019

@author: SankalpMishra
"""

import numpy as np
import pandas as pd
#%%
path = 'H:/costing_automation/tables_hsdoubling_twisting' 
df_hsdoub_cno = pd.read_excel(path+ '/doubconstviv_tbl.xlsx')
df_hsdoub_machine = pd.read_excel(path + '/doubmachine_speed_tbl.xlsx')
df_hsdoub_weighnumber = pd.read_excel(path + '/doub_weighnumber_tbl.xlsx')
df_hsdoub_breaks = pd.read_excel(path + '/doubbreaks_tbl.xlsx')

df_hstwist_cno = pd.read_excel(path + '/hstw_const_tbl.xlsx') # use only hstw construction number table for both doubler and twister
df_hstwist_machine = pd.read_excel(path+ '/hstw_machine_tbl.xlsx')

#%%
class cno:
    def __init__(self, cno_a, df_cno = []):
        self.cno_a = str(cno_a)
        self.df_cno = df_cno
        self.cno_present = df_cno[df_cno['const'] == cno_a]
        
#%%
class  machine_details:
    def __init__(self, cno_present =[], doubmachine_tbl = [], twistmachine_tbl =[]):
        self.cno_present = cno_present
        self.doubmachine_tbl = doubmachine_tbl
        self.doubmachine = doubmachine_tbl[(doubmachine_tbl['package size upperlimit'] >= cno_present['pkg_od'].values[0]) & \
                                           (doubmachine_tbl['plylimit'] >= cno_present['ply'].values[0]) & \
                                           (doubmachine_tbl['tube_length'] == cno_present['tube_length'].values[0])]
        # add twisting machine details
        self.twistmachine_tbl = twistmachine_tbl
        self.twistmachine = twistmachine_tbl[(twistmachine_tbl['ll']<cno_present['stroke'].values[0])&\
                                             (twistmachine_tbl['ul']>cno_present['stroke'].values[0])]
#%%
class doubler_breaks:
    def __init__(self, doubmachine = [], doubbreaks_tbl = []):
        self.doubmachine = doubmachine
        self.doubbreaks_tbl = doubbreaks_tbl
        self.doub_breaks = doubbreaks_tbl[doubbreaks_tbl['machine'] == doubmachine['machine'].values[0]]
        
        
#%%
class doubweighnumber:
    def __init__(self, cno_present = [], doubweighnumber_tbl = []):
        self.cno_present = cno_present
        self.doubweighnumber_tbl = doubweighnumber_tbl
        self.packagesincrate = doubweighnumber_tbl[(doubweighnumber_tbl['ul'] >= cno_present['pkg_od'].values[0]) & \
                                                   (doubweighnumber_tbl['ll']< cno_present['pkg_od'].values[0])]

        

#%%
#class twister breaks
        # twister breaks are unpacked from prodata class using twister machine table data.
#%%
class prod_data():
    """
    twist machine variable input to this class objects are generated from same machine_details class that is used for generation of doubler machine input variabes.
    """
    
    doub_change_time = 0.002
    doub_delay = 1.23
    maintenancetime_doub = 0.066
    doub_weigh_time = 2.5
    doub_creelweight = 7
    
    doffdelayfac_twist = 0.4
    orderchangetime_twist = 40
    interference_twist = 1.167
    
    def __init__(self, cno_present = [], doubmachine = [],doub_breaks = [], packagesincrate = [], \
                 twist_machine = []):
         self.cno_1 = cno_present['const'].values[0]
         self.yno = cno_present['yno'].values[0]
         self.ply = cno_present['ply'].values[0]
         self.tpi = cno_present['mech_tpi'].values[0]
         self.pkg_weight_twist = cno_present['pkg_weight_lb'].values[0]
         self.length_doub = cno_present['length'].values[0]
         self.blend = cno_present['blend'].values[0]
         self.windspeed =cno_present['windspeed'].values[0]
         self.pkg_od = cno_present['windspeed'].values[0]
         self.pkg = cno_present['package'].values[0]
         self.pkg_size = cno_present['pkg_size'].values[0]
         self.stroke_twist = cno_present['stroke'].values[0]
         self.tubelength_doub = cno_present['tube_length'].values[0]
         self.pkg_weightdoub = cno_present['pkg_weight'].values[0]
         self.wax = cno_present['wax'].values[0]
         self.srpm = cno_present['srpm'].values[0]
         self.tie_off = cno_present['tie_off'].values[0]
         self.weave_knot = cno_present['weave_knot'].values[0]
         self.illmansplice = cno_present['illmansplice'].values[0]
         self.newcone = cno_present['newcone'].values[0]
         self.oil = cno_present['oil'].values[0]
         self.label = cno_present['label'].values[0]
         self.label_fact = cno_present['label_fact'].values[0]
         
         self.tpi_twist = cno_present['mech_tpi'].values[0]
         self.brt_twist = twist_machine['brt'].values[0]
         self.twisterbreaks = twist_machine['br_per_hr'].values[0]
         self.dofftime_twist = twist_machine['dt'].values[0]
         self.twistmachine_ul = twist_machine['ul'].values[0]
         self.twistmachine_ll = twist_machine['ll'].values[0]
         self.maintenancetime_twist = twist_machine['mt'].values[0]
         self.spindle_twist = twist_machine['num_spin'].values[0]
         self.weightime_twist = twist_machine['wt'].values[0]
         self.buggychange_twist = twist_machine['buggychange'].values[0]
         self.tbst_twist = twist_machine['tbst'].values[0]
         self.label_twist = twist_machine['label'].values[0]
         

         
         self.doub_machine = doubmachine['machine'].values[-1]
         self.doubspeed = doubmachine['speed'].values[-1]
         self.doubspindles = doubmachine['spindles'].values[-1]
         self.doub_pkgsizell = doubmachine['package size lowerlimit'].values[-1]
         self.doubplylimit = doubmachine['plylimit'].values[-1]
         self.doubdofftime = doubmachine['doff_time'].values[-1]
         self.doubbrt = doubmachine['brt'].values[-1]
         self.doubmachine_tubelength = doubmachine['tube_length'].values[-1]
         self.slipfactor = doubmachine['slip_factor'].values[-1]
         
         self.packages_crate = packagesincrate['n'].values[0] # n is the column name in doub weighnumber table
         
         self.doubbreaks = doub_breaks['average'].values[0]
         self.doubspindles = doub_breaks['spindles'].values[0]
         

         
         
         
         
#%%
class doubling(prod_data):
    
    
    def __init__(self, *args, **kwargs):
        super(doubling, self).__init__(*args, **kwargs)
        
    def doub_length(self):
        return (840*self.pkg_weight_twist*self.yno)
    
    def runtime_doub(self): # you can use self to call super class variables in current class
        return self.yno*840*self.pkg_weight_twist/(self.ply*self.doubspeed)
    
    def doub_slip(self):
        if self.doub_machine == 5 or self.doub_machine == 6:
            return 0
        else: 
            return 1
        
    def cycletime_doub(self): # break time is adjusted as allowance here
        return (self.runtime_doub()+ prod_data.doub_change_time + \
                self.doubdofftime/self.doubspindles + prod_data.doub_delay +\
                ((0.0685*self.doub_length()+ 365.62)*self.doub_slip()/self.doubspeed))*(1/(1-(58 - 0.5*self.runtime_doub())/480))
    
    def opt_doub(self):
        return (prod_data.maintenancetime_doub*self.runtime_doub() + \
               self.doubdofftime + self.doubbrt*self.doubbreaks*self.runtime_doub()*self.doubspindles*self.ply+\
               prod_data.doub_weigh_time*self.doubspindles/self.packages_crate +\
               0.5*self.doubspindles*self.pkg_weight_twist/prod_data.doub_creelweight)*1.20/self.doubspindles
    
    def maxlb_doub(self):
        return 480*self.pkg_weight_twist/self.runtime_doub()
    
    def explb_doub(self):
        return 480*self.pkg_weight_twist/self.cycletime_doub()
    
    def macheff_doub(self): # convert to percentage form in excel
        return self.runtime_doub()/self.cycletime_doub()
    
    def doub_stdmin_spindle(self):
        return self.opt_doub()*480/420
    
    def doub_stdmin_lb(self):
        return self.doub_stdmin_spindle()/self.pkg_weight_twist
    
    def doub_spinass(self):
        return (420/self.opt_doub())/(480/self.cycletime_doub())
    
    
    
#%%

class twisting(prod_data):
    def __init__(self, *args, ** kwargs):
        super(twisting, self).__init__(*args, **kwargs)
        
    def illman_splicefunc(self):
        if str(self.illmansplice).lower() == 'yes':
            return 4
        else: 
            return 1
        
    def newconefunc(self):
        if str(self.newcone).lower() == 'yes':
            return 1
        else: return 0
    
    def oilfunc(self):
        if str(self.oil).lower() == 'yes':
            return 1.15
        else: return 1
    
    def tieofffunc(self):
        if str(self.tie_off).lower() =='yes':
            return 1
        else: return 0
        
    def op_allowance(self):
        return (480/(480 - max(0, (58-0.45*self.runtime_twist()))))
               
        
    def runtime_twist(self):
        return 840*self.yno*self.tpi_twist*self.pkg_weight_twist*18/(self.ply*self.srpm)
    
    def cycletime_twist(self):
        return (self.runtime_twist() + self.runtime_twist()*self.twisterbreaks*self.newconefunc()/2 + \
                self.dofftime_twist*self.oilfunc()/self.spindle_twist + self.dofftime_twist*prod_data.interference_twist*prod_data.doffdelayfac_twist + \
                self.tbst_twist/self.spindle_twist + self.label_twist*self.label_fact/self.spindle_twist + \
                0.1*self.tieofffunc())*(1/(1-self.maintenancetime_twist))*(1/(1-prod_data.orderchangetime_twist/(13*3*8*60)))*(1/(1-self.brt_twist*self.twisterbreaks*self.illman_splicefunc()/60))*\
                (self.op_allowance())
                

    def maxlb_twist(self):
        return 480*self.pkg_weight_twist/self.runtime_twist()
    
    def explb_twist(self):
        return 480*self.pkg_weight_twist/self.cycletime_twist()
    
    def macheff_twist(self):
        return self.explb_twist()/self.maxlb_twist()
    
    def opt_twist(self):
        return (self.dofftime_twist + self.weightime_twist + self.buggychange_twist/4 + \
                self.brt_twist*self.twisterbreaks*self.spindle_twist/60 + self.tbst_twist+ \
                self.label_fact+ 0.1*self.tieofffunc()*self.spindle_twist)*1.05/self.spindle_twist
                
    def stdmin_spin_twist(self):
         return (self.opt_twist()*480/420)
     
    def stdmin_lb_twist(self):
        return self.opt_twist()*480/(420*self.pkg_weight_twist)
    
    def spinass_twist(self): # assuming 58 minutes operator break~
        return 422/(480*self.opt_twist()/self.cycletime_twist())
        
#%% Test case 
        
cno_1 = cno(183102, df_hstwist_cno)
mach_det = machine_details(cno_1.cno_present, df_hsdoub_machine, df_hstwist_machine)
doubbreakx = doubler_breaks(mach_det.doubmachine, df_hsdoub_breaks)
pkg_crate = doubweighnumber(cno_1.cno_present, df_hsdoub_weighnumber)
proddata_try = prod_data(cno_1.cno_present, mach_det.doubmachine, doubbreakx.doub_breaks, pkg_crate.packagesincrate, \
                         mach_det.twistmachine)

doubtry = doubling(cno_1.cno_present, mach_det.doubmachine, doubbreakx.doub_breaks, pkg_crate.packagesincrate, mach_det.twistmachine)  
twist_try = twisting(cno_1.cno_present, mach_det.doubmachine, \
                     doubbreakx.doub_breaks, pkg_crate.packagesincrate, mach_det.twistmachine)

#%%main class

summary_listdoub = []
summary_listtwist = []
columns_dfa = ['cno', 'blend','yno', 'ply', 'pkg_weight', 'pkg_type', 'srpm', 'wax', 'oil', 'machine', 'cycletime', '100 % lbs', 'exp lbs', 'macheff', 'stdminperlb', 'spindle assignment' ]

for i, row in df_hstwist_cno.iterrows():
    cno_a = df_hstwist_cno.iloc[i]['const']
    cno_1 = cno(cno_a, df_hstwist_cno)
    mach_det = machine_details(cno_1.cno_present, df_hsdoub_machine, df_hstwist_machine)
    doubbreakx = doubler_breaks(mach_det.doubmachine, df_hsdoub_breaks)
    pkg_crate = doubweighnumber(cno_1.cno_present, df_hsdoub_weighnumber)
    cno1_doub = doubling(cno_1.cno_present, mach_det.doubmachine, doubbreakx.doub_breaks, pkg_crate.packagesincrate, mach_det.twistmachine)
    cno1_twist = twisting(cno_1.cno_present, mach_det.doubmachine, doubbreakx.doub_breaks, pkg_crate.packagesincrate, mach_det.twistmachine)
    
    print(cno_1.cno_present['const'])
    
    doubx= pd.DataFrame([[cno_1.cno_present['const'].values[0], cno_1.cno_present['blend'].values[0], cno_1.cno_present['yno'].values[0],\
                          cno_1.cno_present['ply'].values[0], cno_1.cno_present['pkg_weight_lb'].values[0], cno_1.cno_present['package'].values[0],\
                          cno_1.cno_present['srpm'].values[0], cno_1.cno_present['wax'].values[0], cno_1.cno_present['oil'].values[0], mach_det.doubmachine['machine'].values[-1], \
                          cno1_doub.cycletime_doub(), cno1_doub.maxlb_doub(), cno1_doub.explb_doub(), cno1_doub.macheff_doub(), cno1_doub.doub_stdmin_lb(), \
                          cno1_doub.doub_spinass()]], columns = columns_dfa)

    twistx = pd.DataFrame([[cno_1.cno_present['const'].values[0], cno_1.cno_present['blend'].values[0], cno_1.cno_present['yno'].values[0],\
                          cno_1.cno_present['ply'].values[0], cno_1.cno_present['pkg_weight_lb'].values[0], cno_1.cno_present['package'].values[0],\
                          cno_1.cno_present['srpm'].values[0], cno_1.cno_present['wax'].values[0], cno_1.cno_present['oil'].values[0], mach_det.twistmachine['machine'].values[0], cno1_twist.cycletime_twist(), cno1_twist.maxlb_twist(), cno1_twist.explb_twist(),\
                          cno1_twist.macheff_twist(), cno1_twist.stdmin_lb_twist(), cno1_twist.spinass_twist()]], columns = columns_dfa)
            
    print(doubx)
    print(twistx)
    
    summary_listdoub.append(doubx)
    summary_listtwist.append(twistx)
    
    
dfdoub = pd.concat(summary_listdoub)
dftwist = pd.concat(summary_listtwist)

#%%
out_path = 'H:\costing_automation\operator_assigment_hsdoublingtwisting.xlsx'
writer = pd.ExcelWriter(out_path , engine='xlsxwriter')
dfdoub.to_excel(writer, sheet_name = 'doubler_summary', index = False)
dftwist.to_excel(writer, sheet_name = 'twister_summary', index = False)
writer.save()
    
            



























