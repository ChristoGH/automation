#!/usr/bin/env python
# coding: utf-8

# C:\Users\christo.strydom\github_repos\automation\CBR\Copy of 20210104_Master_Gap_CBR_LE_CS.xlsx

# In[1]:


import glob

import os, os.path
from os import listdir
from os.path import isfile, join
import pandas as pd
import openpyxl
import win32com.client
from datetime import datetime, timedelta
from dateutil.relativedelta import *
import calendar
import dateutil.parser

def add_months(sourcedate, months):
    month = sourcedate.month - 1 + months
    year = sourcedate.year + month // 12
    month = month % 12 + 1
    day = min(sourcedate.day, calendar.monthrange(year,month)[1])
    return datetime.date(year, month, day)

def last_day_of_month_fn(any_day):
    # this will never fail
    # get close to the end of the month for any day, and add 4 days 'over'
    next_month = any_day.replace(day=28) + timedelta(days=4)
    # subtract the number of remaining 'overage' days to get last day of current month, or said programattically said, the previous day of the first of next month
    return next_month - timedelta(days=next_month.day)



data_path="C:\\Users\\christo.strydom\\github_repos\\automation\\CBR\\data\\"
slices_path="C:\\Users\\christo.strydom\\github_repos\\automation\\CBR\\cbr_data_slices\\"
report_path="""C:\\Users\\christo.strydom\\github_repos\\automation\\CBR\\cbr_reports\\"""
path, dirs, files = next(os.walk(data_path))
file_count = len(files)
# presentation_path='\\klo-jhb-fp01\\\Xelus\\\Data Science\\ChristoStrydom\\Presentations\\'
# presentation_name='cbr_presentation_template'

# path, dirs, files = next(os.walk(data_path))
# file_count = len(files)

onlyfiles = [f for f in listdir(data_path) if isfile(join(data_path, f))]

owls_members='cbr_members_09032021.csv' # 
owls_terminations='cbr_terminations_09032021.csv'
owls_received_claims='cbr_received_claims_09032021.csv'

# owls_claims_df=pd.read_csv(filepath_or_buffer=path+owls_claims)
owls_members_df=pd.read_csv(filepath_or_buffer=data_path+owls_members)
owls_terminations_df=pd.read_csv(filepath_or_buffer=data_path+owls_terminations)
owls_received_claims_df=pd.read_csv(filepath_or_buffer=data_path+owls_received_claims)

owls_members_df['policy_inceptiondate']=pd.to_datetime(owls_members_df['policy_inceptiondate'], format='%Y/%m/%d', errors='ignore')
owls_members_df['policy_createddate']=pd.to_datetime(owls_members_df['policy_createddate'], format='%Y/%m/%d', errors='ignore')
owls_terminations_df['policy_createddate']=pd.to_datetime(owls_terminations_df['policy_createddate'], format='%Y/%m/%d', errors='ignore')
owls_terminations_df['policy_inceptiondate']=pd.to_datetime(owls_terminations_df['policy_inceptiondate'], format='%Y/%m/%d', errors='ignore')
owls_terminations_df['policy_cancellationdate']=pd.to_datetime(owls_terminations_df['policy_cancellationdate'], format='%Y/%m/%d', errors='ignore')
owls_received_claims_df['Notification Date']=pd.to_datetime(owls_received_claims_df['Notification Date'], format='%m/%d/%Y', errors='ignore')
owls_received_claims_df['Policy Inception Date']=pd.to_datetime(owls_received_claims_df['Policy Inception Date'], format='%m/%d/%Y', errors='ignore')
owls_received_claims_df['Notification Date']=pd.to_datetime(owls_received_claims_df['Notification Date'], format='%m/%d/%Y', errors='ignore')
owls_received_claims_df=owls_received_claims_df[~owls_received_claims_df['Notification Date'].isna()].copy()
owls_received_claims_df[['notification_date_1','notification_date_2', 'notification_date_3']]=owls_received_claims_df['Notification Date'].str.split('/',expand=True,)
owls_received_claims_df=owls_received_claims_df[~pd.to_numeric(owls_received_claims_df['notification_date_1'],errors='coerce').isna()].copy()
owls_received_claims_df=owls_received_claims_df[~pd.to_numeric(owls_received_claims_df['notification_date_2'],errors='coerce').isna()].copy()
owls_received_claims_df=owls_received_claims_df[~pd.to_numeric(owls_received_claims_df['notification_date_3'],errors='coerce').isna()].copy()
owls_received_claims_df.loc[owls_received_claims_df['notification_date_1'].str.len()==4,'notification_date_year']=owls_received_claims_df['notification_date_1']
owls_received_claims_df.loc[owls_received_claims_df['notification_date_3'].str.len()==4,'notification_date_year']=owls_received_claims_df['notification_date_3']

datestring=owls_received_claims_df['Notification Date'].iloc[-1]

owls_received_claims_df['notification_date']=owls_received_claims_df.apply(lambda x: dateutil.parser.parse(x['Notification Date']), axis=1)
owls_received_claims_df['som_notification_date']=owls_received_claims_df.apply(lambda x: datetime(x['notification_date'].year,x['notification_date'].month,1), axis=1)

# today = datetime.today()
# months_ago = today + relativedelta(months=-12)
# som_months_ago = datetime(months_ago.year,months_ago.month,1)
# iterate_months=[som_months_ago+relativedelta(months=n) for n in range(12)]


# In[35]:


nmonths=13
today = datetime.today()
months_ago = today + relativedelta(months=-nmonths)
som_months_ago = datetime(months_ago.year,months_ago.month,1)
iterate_months=[som_months_ago+relativedelta(months=n) for n in range(nmonths+1)]

datelist=list(set(owls_terminations_df['policy_inceptiondate']))
datelist.sort()
# datelist.append(add_months(max(datelist),1))
# dates=pd.date_range(min(datelist),add_months(max(datelist),1), freq='1M')-pd.offsets.MonthBegin(1)

f=~owls_terminations_df['productsetup_productsetupname'].isna()
product_list=list(set(owls_terminations_df[f]['productsetup_productsetupname']))


members_live_filter=(owls_members_df['policy_status']=='Live')
ntu_filter=owls_terminations_df['policy_status']=='Not Taken Up'
cancelations_filter=owls_terminations_df['policy_status']=='Cancelled'
owls_members_df=owls_members_df[['policy_policynumber','policygroup_policygroupname','policy_paymentmethod','policy_grosspremium','policy_status','policy_totalpremium','productsetup_productsetupname','policy_inceptiondate']].copy()
owls_terminations_df=owls_terminations_df[['policy_policynumber','policy_paymentmethod','policy_totalpremium','productsetup_productsetupname','policy_inceptiondate','policy_cancellationdate']].copy()
owls_received_claims_df=owls_received_claims_df[['Claim Number','Total Payments','Notification Date','productsetupname','som_notification_date','Original Reserve','notification_date','Progress','Status']].copy()


# In[43]:


owls_members_df['corporate_individual']=''
f=owls_members_df['policy_paymentmethod']=='Debit'
g=(owls_members_df['policy_paymentmethod']!='Debit')&(owls_members_df['policy_paymentmethod']!='')
owls_members_df.loc[f, 'corporate_individual_split']=owls_members_df['productsetup_productsetupname'] + ' - Individual'
owls_members_df.loc[g, 'corporate_individual_split']=owls_members_df['productsetup_productsetupname'] + ' - Corporate'


# In[44]:


owls_terminations_df['corporate_individual']=''
f=owls_terminations_df['policy_paymentmethod']=='Debit'
g=(owls_terminations_df['policy_paymentmethod']!='Debit')&(owls_terminations_df['policy_paymentmethod']!='')
owls_terminations_df.loc[f, 'corporate_individual_split']=owls_terminations_df['productsetup_productsetupname'] + ' - Individual'
owls_terminations_df.loc[g, 'corporate_individual_split']=owls_terminations_df['productsetup_productsetupname'] + ' - Corporate'

owls_received_claims_df[['policy_policynumber_a','policy_policynumber_b']]=owls_received_claims_df['Claim Number'].str.split("/",expand = True)
owls_received_claims_df.policy_policynumber_a = owls_received_claims_df.policy_policynumber_a.str.strip()
df1=owls_received_claims_df.merge(owls_members_df[['policy_policynumber','corporate_individual_split']], left_on= 'policy_policynumber_a', right_on='policy_policynumber', how='inner')
df2=owls_received_claims_df.merge(owls_terminations_df[['policy_policynumber','corporate_individual_split']], left_on= 'policy_policynumber_a', right_on='policy_policynumber', how='inner')
gf=df1.append(df2, ignore_index = True) 


# In[46]:


owls_received_claims_df=gf.copy()
pay_now_filter=owls_received_claims_df['Progress']=='Approve, Pay Now'
reject_claims_filter=owls_received_claims_df['Progress']=='Reject'
outstanding_claims_filter=owls_received_claims_df['Status']=='Outstanding docs received - sent to processing'


# In[47]:


product_list=list(set(owls_members_df['corporate_individual_split']))
product_list = [x for x in product_list if str(x) != 'nan']
owls_terminations_temp_df=owls_terminations_df.copy()


# In[50]:

for product in product_list:
    # product='Sanlam Gap - Individual'
    print('------------------------------------------------------------------------------------------------------------------')
    print('doing product = {product}'.format(product=product))
    max_date=max([owls_members_df['policy_inceptiondate'].max(),owls_terminations_df['policy_inceptiondate'].max(),owls_terminations_df['policy_cancellationdate'].max()])
    owls_members_df=owls_members_df[['policy_grosspremium','policy_status','policy_totalpremium','productsetup_productsetupname','policy_inceptiondate','corporate_individual_split']].copy()
    owls_terminations_df=owls_terminations_df[['policy_totalpremium','productsetup_productsetupname','policy_inceptiondate','policy_cancellationdate','corporate_individual_split']].copy()
    owls_received_claims_df=owls_received_claims_df[['Total Payments','Notification Date','productsetupname','som_notification_date','Original Reserve','notification_date','corporate_individual_split']].copy()

    terminations_product_filter=owls_terminations_df['corporate_individual_split']==product
    members_product_filter=(owls_members_df['corporate_individual_split']==product)
    claims_received_product_filter=(owls_received_claims_df['corporate_individual_split']==product)
    live_df=owls_members_df[members_live_filter&members_product_filter]
    live_policies=live_df.shape[0]
    live_df=owls_members_df[members_live_filter&members_product_filter]
    live_policies_grosspremium=live_df['policy_totalpremium'].sum()

# -------------------------------------------------------------------------------------------------------------------------------------
    live_policies_list=[]
    new_inactive_policies_list=[]
    active_policies_list=[]
    valid_terminations_list=[]
    for i in range(1,len(iterate_months)-1):
        print(i,iterate_months[i])

        members_date_filter=owls_members_df['policy_inceptiondate']>iterate_months[i]
        new_policies_df=owls_members_df[members_date_filter&members_product_filter]
        new_policies=new_policies_df.shape[0]
        terminations_policy_inceptiondate_filter=owls_terminations_df['policy_inceptiondate']<=max_date #iterate_months[i+1]
        terminations_policy_cancellationdate_filter=owls_terminations_df['policy_cancellationdate']>iterate_months[i-1]
        valid_terminations_filter=(terminations_policy_inceptiondate_filter&terminations_policy_cancellationdate_filter&terminations_product_filter)
        valid_terminations_df=owls_terminations_df[valid_terminations_filter]
        valid_terminations=valid_terminations_df.shape[0]
        active_policies=live_policies-new_policies+valid_terminations    
        print('i = ',i,'; month: ', iterate_months[i],owls_terminations_df[terminations_policy_cancellationdate_filter].shape[0],'; active_policies: ',active_policies,'; live_policies: ',live_policies,'; new_policies: ',new_policies, ' valid_terminations: ',valid_terminations)
        live_policies_list.append(live_policies)
        new_inactive_policies_list.append(new_policies)
        valid_terminations_list.append(valid_terminations)
        active_policies_list.append(active_policies)

# -------------------------------------------------------------------------------------------------------------------------------------
    live_policies_grosspremium_list=[]
    new_policies_grosspremium_list=[]
    active_grosspremium_list=[]
    valid_policies_grosspremium_list=[]
    for i in range(1,len(iterate_months)-1):
        print(i,iterate_months[i])
        member_date_filter=owls_members_df['policy_inceptiondate']>iterate_months[i]
    #   ----------------------------------------------------------------------------------------------------------------------------
        new_policies_df=owls_members_df[member_date_filter&members_product_filter]
        new_policies_grosspremium=new_policies_df['policy_totalpremium'].sum()
    #   ----------------------------------------------------------------------------------------------------------------------------
        terminations_policy_inceptiondate_filter=owls_terminations_df['policy_inceptiondate']<=max_date#iterate_months[i+1]
        terminations_policy_cancellationdate_filter=owls_terminations_df['policy_cancellationdate']>iterate_months[i-1]
        valid_terminations_filter=(terminations_policy_inceptiondate_filter&terminations_policy_cancellationdate_filter&terminations_product_filter)
        valid_terminations_df=owls_terminations_df[valid_terminations_filter]
        valid_policies_grosspremium=valid_terminations_df['policy_totalpremium'].sum()
    #   ----------------------------------------------------------------------------------------------------------------------------
        active_policies_grosspremium=live_policies_grosspremium-new_policies_grosspremium+valid_policies_grosspremium    
        print('i = ',i,'; month: ', iterate_months[i],'; iterate_months[i-1]: ',iterate_months[i-1],'; active_policies_grosspremium: ',active_policies_grosspremium,'; live_policies_grosspremium: ',live_policies_grosspremium,'; new_policies_grosspremium: ',new_policies_grosspremium, ' valid_policies_grosspremium: ',valid_policies_grosspremium)
        live_policies_grosspremium_list.append(live_policies_grosspremium)
        new_policies_grosspremium_list.append(new_policies_grosspremium)
        valid_policies_grosspremium_list.append(valid_policies_grosspremium)
        active_grosspremium_list.append(active_policies_grosspremium)

# -------------------------------------------------------------------------------------------------------------------------------------    new_policies_list=[]
    new_policies_list=[]
    for i in range(1,len(iterate_months)-1):
        print(i,iterate_months[i])
        new_policies_date_filter=owls_members_df['policy_inceptiondate']==iterate_months[i]
    #   ----------------------------------------------------------------------------------------------------------------------------
        new_policies_df=owls_members_df[new_policies_date_filter&members_product_filter]
        new_policies=new_policies_df['policy_totalpremium'].shape[0]
        new_policies_list.append(new_policies)

# -------------------------------------------------------------------------------------------------------------------------------------    gwp_new_policies_list=[]
    gwp_new_policies_list=[]
    for i in range(1,len(iterate_months)-1):
    #     print(i,iterate_months[i])
        new_policies_date_filter=owls_members_df['policy_inceptiondate']==iterate_months[i]
    #   ----------------------------------------------------------------------------------------------------------------------------
        new_policies_df=owls_members_df[new_policies_date_filter&members_product_filter]
        gwp_new_policies=new_policies_df['policy_grosspremium'].sum()
        gwp_new_policies_list.append(gwp_new_policies)
        print(i,iterate_months[i],'; ', gwp_new_policies)

# -------------------------------------------------------------------------------------------------------------------------------------    active_grosspremium_list
    active_policies_list
    average_premium_list=[]
    for f, b in zip(active_grosspremium_list, active_policies_list):
        average_premium_list.append(f/b)
        print(f,b,f/b)

# -------------------------------------------------------------------------------------------------------------------------------------    ntu_list=[]
    ntu_list=[]
    for i in range(1,len(iterate_months)-1):
        terminations_policy_inceptiondate_filter=owls_terminations_df['policy_inceptiondate']==iterate_months[i]
        f=terminations_product_filter&ntu_filter&terminations_policy_inceptiondate_filter
        valid_terminations_df=owls_terminations_df[f]
        ntu=valid_terminations_df.shape[0]
        ntu_list.append(ntu)
        print(i,iterate_months[i],'; ', ntu)

# -------------------------------------------------------------------------------------------------------------------------------------    gp_ntu_list=[]
    gp_ntu_list=[]
    for i in range(1,len(iterate_months)-1):
        terminations_policy_inceptiondate_filter=owls_terminations_df['policy_inceptiondate']==iterate_months[i]
        f=terminations_product_filter&ntu_filter&terminations_policy_inceptiondate_filter
        valid_terminations_df=owls_terminations_df[f]
        gp_ntu=valid_terminations_df['policy_totalpremium'].sum()
        gp_ntu_list.append(gp_ntu)
        print(i,iterate_months[i],'; ', gp_ntu,owls_terminations_df[terminations_policy_inceptiondate_filter].shape[0])

# -------------------------------------------------------------------------------------------------------------------------------------
    cancelations_list=[]
    gp_cancelations_list=[]
    for i in range(1,len(iterate_months)-1):
        last_day_of_month=last_day_of_month_fn(iterate_months[i])
        terminations_policy_cancellationdate_filter=owls_terminations_df['policy_cancellationdate']==last_day_of_month
        f=terminations_product_filter&cancelations_filter&terminations_policy_cancellationdate_filter
        valid_terminations_df=owls_terminations_df[f]
        cancelations=valid_terminations_df.shape[0]
        cancelations_list.append(cancelations)
        gp_cancelations=valid_terminations_df['policy_totalpremium'].sum()
        gp_cancelations_list.append(gp_cancelations)    
        print(i,iterate_months[i],'; ',cancelations, gp_cancelations)

# -------------------------------------------------------------------------------------------------------------------------------------
    claims_received_list=[]
    value_claims_received_list=[]
    # gp_cancelations_list=[]
    for i in range(1,len(iterate_months)-1):
    #     last_day_of_month=last_day_of_month_fn(iterate_months[i])
        claims_received_notification_date_filter=owls_received_claims_df['som_notification_date']==iterate_months[i]
        f=claims_received_product_filter&claims_received_notification_date_filter
        df=owls_received_claims_df[f].copy()
        claims_received=df.shape[0]
        claims_received_list.append(claims_received)
        value_claims_received=df['Original Reserve'].sum()
        value_claims_received_list.append(value_claims_received)    
        print(i,iterate_months[i],'; ',claims_received, value_claims_received)

# -------------------------------------------------------------------------------------------------------------------------------------    highest_claim_paid_list=[]
    # value_claims_received_list=[]
    # gp_cancelations_list=[]
    highest_claim_paid_list=[]
    for i in range(1,len(iterate_months)-1):
    #     last_day_of_month=last_day_of_month_fn(iterate_months[i])
        claims_paid_date_filter=owls_received_claims_df['som_notification_date']==iterate_months[i]
        f=claims_received_product_filter&claims_paid_date_filter    
        df=owls_received_claims_df[f].copy()
        highest_claim_paid=df['Total Payments'].max()
        highest_claim_paid_list.append(highest_claim_paid)    
        print(i,iterate_months[i],'; ',highest_claim_paid)
    highest_claim_paid_list=[0 if x != x else x for x in highest_claim_paid_list]

# -------------------------------------------------------------------------------------------------------------------------------------
    number_of_claims_paid_list=[]
    value_of_claims_paid_list=[]
    average_value_of_claims_paid_list=[]
    # value_claims_received_list=[]
    # gp_cancelations_list=[]
    for i in range(1,len(iterate_months)-1):
    #     last_day_of_month=last_day_of_month_fn(iterate_months[i])
        average_value_of_claims_paid=0
        claims_paid_date_filter=owls_received_claims_df['som_notification_date']==iterate_months[i]
        f=claims_received_product_filter&claims_paid_date_filter&pay_now_filter
        df=owls_received_claims_df[f].copy()
        number_of_claims_paid=df.shape[0]
        value_of_claims_paid=df['Original Reserve'].sum()
        number_of_claims_paid_list.append(number_of_claims_paid)
        value_of_claims_paid_list.append(value_of_claims_paid)
        if number_of_claims_paid!=0:
            average_value_of_claims_paid=value_of_claims_paid/number_of_claims_paid
    #         average_value_of_claims_paid_list.append(average_value_of_claims_paid)     
        average_value_of_claims_paid_list.append(average_value_of_claims_paid)
        print(i,iterate_months[i],'; ',average_value_of_claims_paid,number_of_claims_paid,value_of_claims_paid)

# -------------------------------------------------------------------------------------------------------------------------------------
    reject_claims_list=[]
    value_of_reject_claims_list=[]
    average_value_of_reject_claims_list=[]
    # average_value_of_claims_paid_list=[]
    # value_claims_received_list=[]
    # gp_cancelations_list=[]
    for i in range(1,len(iterate_months)-1):
    #     last_day_of_month=last_day_of_month_fn(iterate_months[i])
        average_value_of_reject_claims=0
        claims_paid_date_filter=owls_received_claims_df['som_notification_date']==iterate_months[i]
        f=claims_received_product_filter&claims_paid_date_filter&reject_claims_filter
        df=owls_received_claims_df[f].copy()
        reject_claims=df.shape[0]
        value_of_reject_claims=df['Original Reserve'].sum()
        reject_claims_list.append(reject_claims)
        value_of_reject_claims_list.append(value_of_reject_claims)
        if reject_claims!=0:
            average_value_of_reject_claims=value_of_reject_claims/reject_claims
    #         average_value_of_reject_claims_list.append(average_value_of_reject_claims)     
        average_value_of_reject_claims_list.append(average_value_of_reject_claims)
        print(i,iterate_months[i],'; average_value_of_reject_claims: ',average_value_of_reject_claims,'; reject_claims: ',reject_claims,'; value_of_reject_claims: ',value_of_reject_claims)
# -------------------------------------------------------------------------------------------------------------------------------------
    outstanding_claims_list=[]
    value_of_outstanding_claims_list=[]
    average_value_of_outstanding_claims_list=[]
    # average_value_of_claims_paid_list=[]
    # value_claims_received_list=[]
    # gp_cancelations_list=[]
    for i in range(1,len(iterate_months)-1):
    #     last_day_of_month=last_day_of_month_fn(iterate_months[i])
        average_value_of_outstanding_claims=0
        outstanding_claims_date_filter=owls_received_claims_df['som_notification_date']==iterate_months[i]
        f=claims_received_product_filter&outstanding_claims_date_filter&outstanding_claims_filter
        print(owls_received_claims_df[outstanding_claims_filter].shape[0])
        df=owls_received_claims_df[f].copy()
        outstanding_claims=df.shape[0]
        value_of_outstanding_claims=df['Original Reserve'].sum()
        outstanding_claims_list.append(outstanding_claims)
        value_of_outstanding_claims_list.append(value_of_outstanding_claims)
        if outstanding_claims!=0:
            average_value_of_outstanding_claims=value_of_outstanding_claims/outstanding_claims
    #         average_value_of_outstanding_claims_list.append(average_value_of_outstanding_claims)   
        average_value_of_outstanding_claims_list.append(average_value_of_outstanding_claims)
        print(i,iterate_months[i],'; value_of_outstanding_claims: ',value_of_outstanding_claims,'; outstanding_claims: ',outstanding_claims,'; value_of_reject_claims: ',value_of_reject_claims)
    # Average Cost Per Claim


    # ***

    # ### Construct the report dataframe

    # In[77]:


    index_list=['Active Policies',
                'Live Policies',
                'New Inactive Policies',
                'Active Terminated Policies',
                'Gross Premium',
                'Gross Premium from Live Policies',
                'Gross Premium from New Inactive Policies',
                'Gross Premium from Active Terminated Policies',
                'New Policies',
                'GWP of New Policies',
                'Average Premium',
                'NTU',
                'Gross Premium for NTU',
                'Cancellations',
                'Gross Premium for Cancellations',
                'Number of Claims Received (Incl O/Docs Claims)',
                'Value of Claims Received',
                'Highest Claim Paid',
                'Number of Claims Paid',
                'Value of Claims Paid',
                'Average Cost Per Claim',
                'Repudiated Claims',
                'Value of Repudiated Claims',
                'Average Cost per repudiated claim',
                'Number of Oustanding Claims',
                'Value of Outstanding Claims',
            'Average Value of Outstanding Claims']


    # In[78]:


    data=[]
    data.append(active_policies_list)
    data.append(live_policies_list)
    data.append(new_inactive_policies_list)
    data.append(valid_terminations_list)
    data.append(active_grosspremium_list)
    data.append(live_policies_grosspremium_list)
    data.append(new_policies_grosspremium_list)
    data.append(valid_policies_grosspremium_list)

    data.append(new_policies_list)
    data.append(gwp_new_policies_list)
    data.append(average_premium_list)
    data.append(ntu_list)
    data.append(gp_ntu_list)
    data.append(cancelations_list)
    data.append(gp_cancelations_list)
    data.append(claims_received_list)
    data.append(value_claims_received_list)
    data.append(highest_claim_paid_list)
    data.append(number_of_claims_paid_list)
    data.append(value_of_claims_paid_list)

    data.append(average_value_of_claims_paid_list)
    data.append(reject_claims_list)
    data.append(value_of_reject_claims_list)
    data.append(average_value_of_reject_claims_list)
    data.append(outstanding_claims_list)
    data.append(value_of_outstanding_claims_list)
    data.append(average_value_of_outstanding_claims_list)
    print(data)
    for i,j in zip(index_list,data):
        print(i,len(j))
    df = pd.DataFrame(data,index=index_list, columns=iterate_months[1:-1])
    s=owls_members.split("_")[1].split(".")[0]
    filename="cbr"+'_'+product+"_"+s+".csv"
    df.to_csv(path_or_buf=report_path+filename, sep=',', index=True)


