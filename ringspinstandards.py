# -*- coding: utf-8 -*-
"""
Created on Tue May 21 09:53:32 2019

@author: SankalpMishra
"""

# -*- coding: utf-8 -*-
"""
Created on Thu Apr  4 10:19:36 2019

@author: SankalpMishra
siro functions and variables are not included in zinser tables.
So, a column for siro with all zeroes isadded to cno_tbl.
siro has been removed from output table.

"""
#import libraries


import numpy as np
import pandas as pd

#%%
'''
use this class to get standards for other spinning frames.
file names should be passed as strings (in quotations)

'''

class ReadData:
    def __init__(self, path, cno_file, allowance_file, spintime_file, windtime_file, blenddetail_file, hankrov_file):
        self.df_rs_cno = pd.read_excel(path + cno_file + '.xlsx')
        self.df_rs_allowance = pd.read_excel(path + allowance_file + '.xlsx')
        self.df_rs_spinnertimes = pd.read_excel(path + spintime_file +'.xlsx')
        self.df_rs_windtimes = pd.read_excel(path + windtime_file +'.xlsx')
        self.df_rs_blenddetail = pd.read_excel(path + blenddetail_file + '.xlsx')
        self.df_rs_hankrov = pd.read_excel(path + hankrov_file +'.xlsx')
        
        # next part is for siro adjustment on zinser construction table
        self.df_rs_cno['siro'] = 0
        
#%%
class rs_cno:
    """
    extracting row of seletcted construction number from master data table
    """
    
    def __init__(self, cno_1, cno_tbl = []):
        self.cno_1=cno_1
        self.cno_tbl = cno_tbl
        self.cno_present = cno_tbl[cno_tbl['cno_1'] == cno_1]
        
#%%
class rs_blend:
    """
    extacting row of blend data from blend data table based on blend of selected construction number
    """
    
    def __init__(self, blend_type1,blend_tbl = []):
        self.blend_type1 = blend_type1
        self.blend_tbl = blend_tbl
        self.blend_detail = blend_tbl[blend_tbl['blend_type']== blend_type1]

#%%
class rs_spintime:
    """
    unpacking spinning activity time details
    """
    
    def __init__(self, spintimes=[]):
        self.spintimes = spintimes


#%%        
class rs_rovingweight:
    """
    extracting roving weight as a scalar from roving weight tablebased on the hank roving size of the selected construction number.
    """
    
    def __init__(self, hr, hr_tbl = []):
        #self.cno_present = cno_present
        self.hr_tbl = hr_tbl
        self.hr = hr
        self.hra = hr_tbl[hr_tbl['hank_roving'] == hr]
        self.rov_weight = self.hra['roving_weight'].values[0]
    #if cno_present['siro'].values[0] == 1:
        #rov_weight = 5.52
    #else:
        #rov_weight = hr['roving_weight'].values[0]
#%%
class rs_data:
    """
    Data class used tounpackall variables required for further calculations.
    Takes inputs of selected rows from different tables based on the selected construction number.
    """
    
    
    spin_delay = 18.033 #static variables that may be changed based on user experience. not fixed according to construction number
    wind_delay = 15.06
    
    def __init__(self, roving_weight, cno_present=[], blend_detail=[], spintimes = [], windtimes = [], allowance = []):
        self.roving_weight = roving_weight
        self.cno_present = cno_present
        self.blend_detail = blend_detail
        self.spintimes = spintimes
        self.windtimes = windtimes
        self.allowance = allowance  
        #self.runtime  = runtime_spin() -- cannot do this!
        
        # UNPACKING FROM MASTER DATA (CNO _TBL)
        self.cno = cno_present['cno_1'].values[0]
        self.blend = cno_present['blend'].values[0]
        self.yno = cno_present['yno'].values[0]
        self.hank_roving = cno_present['hank_roving'].values[0]
        self.srpm = cno_present['srpm'].values[0]
        self.tpi = cno_present['tpi'].values[0]
        self.pkg_weight = cno_present['pkg_weight'].values[0]
        self.pkg_type = cno_present['pkg_type'].values[0]
        self.wind_speed = cno_present['wind_speed'].values[0]
        self.container = cno_present['container'].values[0]
        self.siro = cno_present['siro'].values[0]
        self.wax = cno_present['wax'].values[0]
        self.bag = cno_present['bag_x'].values[0]
        self.tieoff = cno_present['tieoff'].values[0]
        self.label = cno_present['label'].values[0]
        
        #(UNPACKING ROVING WEIGHT SCALAR)
        self.rov_weight = roving_weight.rov_weight
        
        #(UNPACKING VARABLE FROM BLEND DATA TABLE)
        self.spinner_breaks = blend_detail['spinner_breaks'].values[0]
        self.winder_breaks = blend_detail['winder_breaks'].values[0]
        self.winder_lights = blend_detail['winder_lights'].values[0]
        self.bobbin_weight = blend_detail['bobbin_weight'].values[0]
        
        #(SPINING ACTIVITY TIME VARIABLES)
        self.breakrepairtime_spin = spintimes['break_repair_spin'].values[0]
        self.recreeltime_spin = spintimes['recreel'].values[0]
        self.automaticdofftime_spin = spintimes['automatic_dofftime_spin'].values[0]
        self.avgrepair_delay = spintimes['avg_repair_delay'].values[0]
        self.setuptime_spin = spintimes['setup_time'].values[0]
        self.numspindles_spin = spintimes['num_spin_s'].values[0]
        
        
        #(WINDING ACTIVITY TIME VARIABLES)
        self.w_collectpackages = windtimes['collect_packages'].values[0]
        self.w_tietime = windtimes['tie_time'].values[0]
        self.w_labeltime = windtimes['label_time'].values[0]
        self.w_lightfix = windtimes['light_fix'].values[0]
        self.w_bagtime = windtimes['bag_time'].values[0]
        self.w_repairbreak  = windtimes['repair_break'].values[0]
        self.w_dofftime = windtimes['doff'].values[0]
        self.w_cleanbobbins = windtimes['clean_bobbins'].values[0]
        self.w_placedivider = windtimes['place divider'].values[0]
        self.w_deliver = windtimes['deliver'].values[0]
        self.w_tbst = windtimes['tbst'].values[0]
        self.w_numspindles = windtimes['num_spin_w'].values[0]
        self.w_numpackages = windtimes['num_package_crate'].values[0]
        
        #(VARIABLES FROM ALLOWANCE TABLE)
        self.mta_spin = allowance['mta'].values[0]
        self.opa = allowance['opa'].values[0]
        self.cha = allowance['cha'].values[0]
        self.bqa = allowance['bqa'].values[0]
        self.mta_wind = allowance['mta_w'].values[0]
        
#%%     
class rs_spinning_winding(rs_data):   
    """
    metrics calculating class from spinning and winding
    """
    
    def __init__(self,*args, **kwargs):
        super(rs_spinning_winding, self).__init__(*args, **kwargs)
    

    #spinning functions    
    def getfrs(self): # spinning machine front roll speed
        return self.cno_present['srpm'].values[0]/(self.cno_present['tpi'].values[0]*36)
    
    def runtime_spin(self):
        return 840*self.yno*self.bobbin_weight/(self.srpm/(self.tpi*36))
    
    def rovingweight(self):
        if self.siro == 1:
            return 5.52
        else:
            return self.rov_weight
    
    #runtime = self.runtime_spin()  ## how can we generate instance calculated variables in a class
    
    def cycletime_spin(self):
        return ((self.runtime_spin() + \
                2*self.breakrepairtime_spin*self.spinner_breaks*self.runtime_spin()/self.numspindles_spin)+\
                self.recreeltime_spin*(1+self.siro)*self.bobbin_weight/self.rovingweight() + self.automaticdofftime_spin + \
                2*rs_data.spin_delay*self.spinner_breaks*self.runtime_spin()/self.numspindles_spin)*\
                self.mta_spin*self.opa*self.cha*self.bqa
                
    def maxlb_spin(self):
        return 480*self.bobbin_weight*self.numspindles_spin/self.runtime_spin()
                
    def explb_spin(self):
        return 480*self.bobbin_weight*self.numspindles_spin/self.cycletime_spin()
    
    
     
    def efficiency_spin(self):
        return self.runtime_spin()*100/self.cycletime_spin()
    
    def opt_spin(self):
        return (self.automaticdofftime_spin + \
                2*self.breakrepairtime_spin*self.spinner_breaks*self.runtime_spin() + \
                self.recreeltime_spin*self.bobbin_weight*self.numspindles_spin/self.rovingweight() +\
                (self.mta_spin-1)*self.runtime_spin()/self.mta_spin+\
                (self.cha-1)*self.runtime_spin()/self.cha)
    
    def stdminperlb_spin(self):
        return self.opt_spin()*480/(430*self.bobbin_weight)
    
    def op_assign_spin(self):
        return (430/(self.opt_spin()*0.8))/(480/self.cycletime_spin())
    
    ### winder functions
    
    def runtime_wind(self):
        return 840*self.pkg_weight*self.yno/self.wind_speed
    
    def runtime_windbatch(self):
        return (840*self.pkg_weight*self.yno/self.wind_speed)*self.bobbin_weight*self.numspindles_spin/(self.w_numspindles*self.pkg_weight)

    
    def cycletime_wind(self): #package delivery time is not added!!- not an operator activity.
        return (self.runtime_wind() + (self.winder_breaks*self.runtime_wind()*self.w_repairbreak/self.w_numspindles) +\
               self.w_dofftime + self.winder_lights*self.w_lightfix + self.bag*self.w_bagtime + self.w_cleanbobbins + \
               self.label*self.w_labeltime + self.w_collectpackages)*self.mta_wind
                
    def cycletime_windbatch(self):
        return (self.runtime_wind() + (self.winder_breaks*self.runtime_wind()*self.w_repairbreak/self.w_numspindles) +\
               self.w_dofftime + self.winder_lights*self.w_lightfix + self.bag*self.w_bagtime + self.w_cleanbobbins + \
               self.label*self.w_labeltime + self.w_collectpackages)*self.mta_wind*self.bobbin_weight*self.numspindles_spin/(self.w_numspindles*self.pkg_weight)
                    
    def maxlb_wind(self):
        return (480*self.pkg_weight*self.w_numspindles/self.runtime_wind())
    
    def explb_wind(self):
        return 480*self.pkg_weight*self.w_numspindles/self.cycletime_wind()
    
    def efficiency_wind(self):
        return self.runtime_wind*100/self.cycletime_wind()
    
    def opt_wind(self): #deliverytime not included here as well!
        return self.w_tbst + self.winder_lights*0.0833*self.runtime_wind() + self.bag*self.w_bagtime + self.w_cleanbobbins + self.w_collectpackages+\
               self.label*self.w_labeltime + self.tieoff*self.w_tietime
               
    def opt_windbatch(self):
        return (self.w_tbst + self.winder_lights*0.0833*self.runtime_wind() + self.bag*self.w_bagtime + self.w_cleanbobbins + self.w_collectpackages+\
               self.label*self.w_labeltime + self.tieoff*self.w_tietime)*self.bobbin_weight*self.numspindles_spin/(self.pkg_weight*self.w_numspindles)

    def stdminperlb_wind(self):
        return (self.opt_wind()*480/(430*self.pkg_weight))

    def op_assign_wind(self):
        return (430/(self.opt_wind()*self.w_numspindles*0.8))/(480/(self.cycletime_wind()))       
        
    
    def runtime_line(self):
        return max(self.runtime_spin(), self.runtime_windbatch())    
    
    
#%% seperate class for winding? would require use of multiple inheritance later - not muchfamiliar with it. ask shubha!!
        
    
#%% 
class rs_line(rs_spinning_winding):
    """
    model class for metrics calculation for combined spinning and winding systems.
    """
    
    def __init__(self, *args, **kwargs):
        super(rs_line, self).__init__(*args, **kwargs)
    
    def runtime_line(self):
        return max(rs_spinning_winding.runtime_spin(self), rs_spinning_winding.runtime_windbatch(self))
    
    def bottleneckshift_delay(self):
        a = rs_spinning_winding.cycletime_spin(self)
        b = rs_spinning_winding.cycletime_windbatch(self)
        if b>a:    
            return rs_spinning_winding.cycletime_wind(self)*self.bobbin_weight*592/(self.pkg_weight*self.w_numspindles)
        else:
            return 0
    
    def cycletime_line(self):
        return max(rs_spinning_winding.cycletime_spin(self), rs_spinning_winding.cycletime_windbatch(self) + self.bottleneckshift_delay())
    
    def opt_line(self):
        return (rs_spinning_winding.opt_spin(self) + rs_spinning_winding.opt_windbatch(self))
    
    def maxlb_line(self):
        return min(rs_spinning_winding.maxlb_spin(self), rs_spinning_winding.maxlb_wind(self))
    
    def explb_line(self):
        return min(rs_spinning_winding.explb_spin(self), rs_spinning_winding.explb_wind(self))
    
    def efficiency_line(self):
        return self.runtime_line()/self.cycletime_line()
    
    def stdminperlb_line(self):
        return (rs_spinning_winding.opt_spin(self) + rs_spinning_winding.opt_windbatch(self)*self.w_numspindles)*480/(430*self.bobbin_weight*self.numspindles_spin)
    
    def optassign_line(self):
        return (430/self.opt_line())/(480/(self.cycletime_line()))
    
     
#%%
        
#reading zinser data'
path_zin = 'G:/Standards/Main/tables_excel/'
# path, cno_file, allowance_file, spintime_file, windtime_file, blenddetail_file, hankrov_file
zinser = ReadData(path_zin, 'dszs_cno_tbl', 'dszs_allowance_tbl', 'dszs_spinnertime_tbl', 'dszs_windertime_tbl', \
                  'dszs_blendbobbinwt_tbl', 'dszs_rovingbobbin_tbl')

#%% main class- how to put this loop in a class?
"""
lopping classes in this module togenerate summary for every construciton number in the master table(cno_tbl).
"""
#removing siro for zinser
summary_list = []
cols = ['cno', 'yno', 'blend', 'tpi', 'hank_roving', 'srpm', 'wind_speed', 'pkg_type', \
        'pkg_weight', 'bag', 'siro', 'combined cycle time', '100 % lb', 'exp lb', 'combined efficiency', \
        'stdmin/lb', 'spin assignment', 'wind assignment', 'combined assignment']

for i, row in zinser.df_rs_cno.iterrows():
    cno_x = zinser.df_rs_cno.iloc[i]['cno_1']
    currentcno = rs_cno(cno_x, zinser.df_rs_cno)
    blend_detail = rs_blend(currentcno.cno_present['blend_type'].values[0], zinser.df_rs_blenddetail)
    spintime = rs_spintime(zinser.df_rs_spinnertimes)
    rov_weight = rs_rovingweight(currentcno.cno_present['hank_roving'].values[0], zinser.df_rs_hankrov)
    
    print(rov_weight, currentcno.cno_present['cno_1'].values[0], blend_detail.blend_detail)
    
    spin_data = rs_line(rov_weight, currentcno.cno_present, blend_detail.blend_detail, spintime.spintimes, zinser.df_rs_windtimes, zinser.df_rs_allowance)
    
    x= pd.DataFrame([[currentcno.cno_present['cno_1'].values[0], currentcno.cno_present['yno'].values[0], currentcno.cno_present['blend'].values[0], \
                      currentcno.cno_present['tpi'].values[0], currentcno.cno_present['hank_roving'].values[0],\
                      currentcno.cno_present['srpm'].values[0], currentcno.cno_present['wind_speed'].values[0], \
                      currentcno.cno_present['pkg_type'].values[0], currentcno.cno_present['pkg_weight'].values[0], \
                      currentcno.cno_present['bag'].values[0], currentcno.cno_present['siro'].values[0], spin_data.cycletime_line(), \
                      spin_data.maxlb_line(), spin_data.explb_line(), spin_data.efficiency_line(), spin_data.stdminperlb_line(), \
                      spin_data.op_assign_spin(), spin_data.op_assign_wind(), spin_data.optassign_line()]], columns = cols)
    
    print(x)
    
    summary_list.append(x)
    
df = pd.concat(summary_list)

#%%
out_path = 'I:\costing_automation\operator_assigment_zinser_up.xlsx'
writer = pd.ExcelWriter(out_path , engine='xlsxwriter')
df.to_excel(writer, sheet_name = 'Sheet1', index = False)

    


