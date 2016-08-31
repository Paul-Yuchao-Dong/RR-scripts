
# coding: utf-8

# In[1]:

import pandas as pd
import numpy as np
import re
from calendar import monthrange 
from pandas.tseries.offsets import Second

import xlwings as xw
from xlwings import Workbook, Range, Sheet

from RR_scripts.RR_converter import RR_convert


# In[4]:

def t_month_end(target_month, target_year):
    days_in_month = monthrange(target_year, target_month)[1]
    return pd.to_datetime('{day}/{month}/{year} 12:00:00'.format(day = days_in_month, month = target_month, year = target_year), dayfirst = True)


# In[5]:

def t_month_start(target_month, target_year):
    return pd.to_datetime('{day}/{month}/{year} 00:00:00'.format(day = 1, month = target_month, year = target_year), dayfirst = True) - Second(1)


# In[12]:

def human_quarter(q_string):
    s = re.split('Q|q|\W+', q_string)
    s = list(filter(None, s))
    s = list(map(int, map(str.strip, s)))
    year = max(s)
    quarter = min(s)
    return year, quarter


# In[13]:

def quarter_2_month(quarter):
    return range(quarter * 3 - 2, quarter * 3 + 1)


# In[27]:

def write_to_excel(things, sh, row_flag = 2):
    Range(sh, (row_flag, 2)).value = things
    try:
        offset_rows = things.shape[0]
    except:
        offset_rows = 0
    return row_flag + offset_rows + 2
    


# In[14]:

def ready_excel(df):
    return (df.copy()
            .sort_values("Real_END").reset_index()\
            .assign(L_START = lambda x: x.L_START.dt.strftime('%d-%b-%Y'),
                    L_END = lambda x: x.L_END.dt.strftime('%d-%b-%Y'),
                    ET_DATE = lambda x: x.ET_DATE.dt.strftime('%d-%b-%Y'),
                    Real_END = lambda x: x.Real_END.dt.strftime('%d-%b-%Y')      
                   ))


def ready_func(x):
    try:
        return x.dt.strftime('%d-%b-%Y')
    except:
        return x


# In[216]:

class RR_A():
    def __init__(self, data = RR_convert()):
    #def __init__(self, data = df):
        self.data = data
        self.active_mask = [True] * data.shape[0]
    
    def active_on_the_day(self, t_date):
        # creating masks to filter the rentroll
        '''
        t_date could be anything that can be converted to a pd datetime by pd.to_datetime
        '''
        try:
            t_date = pd.to_datetime(t_date)
        except:
            print("Date Format Issue")
            return None
        t_active_mask = (self.data.L_START<= t_date) & (self.data.L_END>= t_date) # all leases active in the month's end
        et_mask = ~((pd.notnull(self.data.ET_DATE)) & (self.data.ET_DATE<t_date)) # exclude leases that has been early terminated before month end
        
        df = self.data[t_active_mask & et_mask].copy()
        
        df = df.assign(main_ls_num = df['Ls_no'].str.split('-').str[0],
                       sub_ls_num = df['Ls_no'].str.split('-').str[1].astype(int)) # separate the lease no
    
        df = df.reset_index()
        
        init_groups = df.groupby(['main_ls_num', 'FL']) # this only groups the main lease no
        
        ls_maxes = init_groups.sub_ls_num.transform(max) # find the largest lease no in each category
        
        df = df[df.sub_ls_num == ls_maxes] # filter out the older leases
        
        return RR_A(data = df)
    
    def lease_calc(self, active_mask = None):
        #Calculation of performance
        '''
        returns the GFA and weighted avg rent
        active_mask should be a mask on the dataframe saved under self.data, default to None
        '''
        active_mask = active_mask or self.active_mask 
        calc_GFA = self.data.loc[active_mask, 'GFA'].sum()
        try: 
            calc_rent = (self.data.loc[active_mask, 'E_RENT'] * self.data.loc[active_mask, 'GFA']).sum() / calc_GFA
        except ZeroDivisionError:
            calc_rent = 0
        return calc_GFA, calc_rent
    
    def new_analysis(self, s_date, t_date):
         #new leases
        '''
        returns new leases started in the period
        s_date marks the starting date of the period in question
        t_date marks the end date of the period in question
        s_date, t_date could be anything that can be converted to a pd datetime by pd.to_datetime
        '''
        df = self.data
        s_date = pd.to_datetime(s_date)
        t_date = pd.to_datetime(t_date)
        self.active_mask = (df.L_START>= s_date) & (df.L_START<= t_date)
        return RR_A(df[self.active_mask])
    
    def old_analysis(self, s_date, t_date):
        #Expired leases
        '''
        returns leases terminated in the period
        s_date marks the starting date of the period in question
        t_date marks the end date of the period in question
        s_date, t_date could be anything that can be converted to a pd datetime by pd.to_datetime
        '''
        df = self.data
        s_date = pd.to_datetime(s_date)
        t_date = pd.to_datetime(t_date)
        expired_during_mask = (df.L_END<=t_date) & (df.L_END>=s_date)
        early_terminated_during_mask = (pd.notnull(df.ET_DATE)) & ( (df.ET_DATE>= s_date ) & (df.ET_DATE<=t_date ))
        et_2_mask = ~((pd.notnull(df.ET_DATE)) & (df.ET_DATE<s_date)) # exclude leases that has been early terminated before the start
        expire_lease_mask = (expired_during_mask & et_2_mask) | early_terminated_during_mask
        self.active_mask = expire_lease_mask
        return RR_A(df[self.active_mask]) 
    
    def period_calc(self, q_months, test_year):
        # can be used for quarterly analysis
        """
        returns the Leased_GFA	Passing_rent	Occupancy	Started_GFA	Spot_rent	Ended_GFA	Expiring_rent of q_months in test_year
        q_months should be a range of int
        test_year also an int        
        """
        q_df = {(test_month, test_year):self.active_on_the_day(t_month_end(test_month, test_year))                                        .lease_calc()                for test_month in q_months}
        
        q_df = pd.DataFrame(q_df, index= ['Leased_GFA', 'Passing_rent']).T

        q_df = q_df.assign(Occupancy = lambda x: x.Leased_GFA / 120245)
        
        q_new_df = {(test_month, test_year): self.new_analysis(t_month_start(test_month, test_year), t_month_end(test_month, test_year))                                                .lease_calc()                    for test_month in q_months}
        q_new_df = pd.DataFrame(q_new_df, index= ['Started_GFA', 'Spot_rent']).T
        q_df = q_df.join(q_new_df)
 
        q_end_df = {(test_month, test_year): self.old_analysis(t_month_start(test_month, test_year), t_month_end(test_month, test_year))                                                .lease_calc()                    for test_month in q_months}

        q_end_df = pd.DataFrame(q_end_df, index= ['Ended_GFA', 'Expiring_rent']).T
        q_df = q_df.join(q_end_df)

        q_df.index.names = ['Month', 'Year']
        q_df = q_df.swaplevel("Month","Year")
        
        return q_df
    def period_stat(self, q_months, test_year):
        """
        returns further summary stat based on period_calc
        q_months should be a range of int of the months in question
        test_year also an int of the year in question       
        """
        q_df = self.period_calc(q_months, test_year)
        return pd.DataFrame({'Avg_leased_GFA': q_df.Leased_GFA.mean(),
                             'Avg_Passing_rent': q_df.Passing_rent.mean(),
                             'Avg_Occupancy': q_df.Occupancy.mean(),
                             'W_avg_Spot_rent': np.average(q_df.Spot_rent, weights= q_df.Started_GFA),
                             'W_avg_Expiring_rent': np.average(q_df.Expiring_rent, weights= q_df.Ended_GFA)
                            }, index = ['Summary Stat'])
    
    def renewal_a(self, q_months, year):
        # Renewal analysis
        index_col_list = ['BLDG','FL','UNITS']
        s_date = t_month_start(q_months[0], year)
        t_date = t_month_end(q_months[-1], year)
        
        r_new_leases_list = (self.new_analysis(s_date, t_date).data.set_index(index_col_list))

        r_expired_leases_list = (self.old_analysis(s_date, t_date).data.set_index(index_col_list))

        reversion = r_expired_leases_list.join(r_new_leases_list, how = 'inner', lsuffix = '_e', rsuffix = '_n')
        reversion = (reversion.assign(reversion_r = lambda x: x.E_RENT_n / x.E_RENT_e.astype(float) - 1))

        r_new_leases_list = r_new_leases_list.loc[reversion.index]
        r_expired_leases_list = r_expired_leases_list.loc[reversion.index]

        r_new_leases_list['Reversion_Rate'] = reversion.reversion_r
        r_expired_leases_list['Reversion_Rate'] = reversion.reversion_r
        try:
            period_rate = np.average(reversion.reversion_r, weights = reversion.GFA_e)
        except ZeroDivisionError:
            period_rate = 0
        
        return r_expired_leases_list, r_new_leases_list, period_rate
    
    def quarterly_routine(self, q_string):
        """
        
        """
        year, quarter = human_quarter(q_string)
        q_months = quarter_2_month(quarter)
        
        self.months_stat(q_months, year)
    
    def months_stat(self, q_months, year):
        q_df = self.period_calc(q_months, year)
        sum_q = self.period_stat(q_months, year)

        wb = Workbook()

        sh = Sheet.add("Summary", wkb = wb)

        row_flag = write_to_excel(q_df, sh = sh)
        row_flag = write_to_excel(sum_q, sh = sh, row_flag = row_flag)

        sh = Sheet.add("Master", wkb = wb)
        row_flag = write_to_excel(self.active_on_the_day(t_month_end(q_months[-1], year))                                  .data.pipe(ready_excel), 
                                sh = sh)
        
        sh1 = Sheet.add("Aggregate", wkb = wb)
        row_flag = write_to_excel('New Leases During the Period', sh = sh1)
        new_leases_list = self.new_analysis(t_month_start(q_months[0], year), t_month_end(q_months[-1], year))                           .data.pipe(ready_excel)
        row_flag = write_to_excel(new_leases_list, sh = sh1, row_flag = row_flag)

        row_flag = write_to_excel('Expired During the Period', sh = sh1, row_flag = row_flag)
        
        expired_leases_list = self.old_analysis(t_month_start(q_months[0], year), t_month_end(q_months[-1], year))                                   .data.pipe(ready_excel)
        row_flag = write_to_excel(expired_leases_list, sh = sh1, row_flag = row_flag)     
        
        r_expired_leases_list, r_new_leases_list, period_rate = self.renewal_a(q_months, year)
        
        sh1 = Sheet.add("Renewal", wkb = wb)
        row_flag = write_to_excel('Renewed Leases During the Period', sh = sh1)
        row_flag = write_to_excel('Original Leases', sh = sh1, row_flag = row_flag)

        row_flag = write_to_excel(r_expired_leases_list.pipe(ready_excel), sh = sh1, row_flag = row_flag)

        row_flag = write_to_excel('Renewed Leases', sh = sh1, row_flag = row_flag)    
        row_flag = write_to_excel(r_new_leases_list.pipe(ready_excel), sh = sh1, row_flag = row_flag)

        row_flag = write_to_excel('Weighted Average Reversion Rate', sh = sh1, row_flag = row_flag)
        row_flag = write_to_excel(period_rate, sh = sh1, row_flag = row_flag)
        
        quarter = q_months[-1]//3

        for tower in range(1,3):    
            sh_new = Sheet.add("Tower {tower} {year} Q{quarter}".format(tower = tower, year = year, quarter = quarter), wkb = wb)
            row_flag = write_to_excel('Tower {tower} New Leases During the Period'.format(tower = tower), sh = sh_new)   
            new_leases_list_T = new_leases_list.loc[new_leases_list['BLDG'] == tower].copy()
            row_flag = write_to_excel(new_leases_list_T, sh = sh_new, row_flag = row_flag)

            row_flag = write_to_excel('Tower {tower} Expired Leases During the Period'.format(tower = tower), sh = sh_new, row_flag = row_flag)
            expired_leases_list_T = expired_leases_list.loc[expired_leases_list['BLDG'] == tower].copy()
            row_flag = write_to_excel(expired_leases_list_T, sh = sh_new, row_flag = row_flag)

        Sheet('Sheet1').delete()
        wb.save("Operating Statistics Q{quarter} {year}".format(quarter = quarter, year = year))
        #wb.close()        

        return "OK"
    

        

if __name__ == "__main__":
    pass

