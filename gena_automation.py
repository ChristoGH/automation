#!/usr/bin/env python
# coding: utf-8

# ### This is the automation script for claims experience reports.

# In[1]:


import pandas as pd
import openpyxl
import win32com.client


# Create a function to populate a range of cells in an excel spreadsheet. 
# - __address_array__ is the array of cellsthat will be populated,
# - __xlsx_sheet__ is the name of the sheet and 
# - __values_array__ is the list of values used to populate the cells.

# In[2]:


def populate_range_fn(address_array,xlsx_sheet,values_array):
    c=0
    for i in range(0,len(xlsx_sheet[address_array[0][1].replace('$', '')])):
        for item in xlsx_sheet[address_array[0][1].replace('$', '')][i]:
    #         print(item.value)
            item.value=values_array[c]
            c+=1


# __add_list_fn__ is a __function__ that adds two lists:

# In[3]:


def add_list_fn(list1,list2):
    zip_object = zip(list1, list2)
    sumlist=[]
    for list1_i, list2_i in zip_object:
        sumlist.append(list1_i+list2_i)
    return sumlist.copy()


# Define three variables:

# In[ ]:


group_name='Overberg Agri'
year='2018'
pgn_list=['Overberg Agri',
'Overberg Agri - Pensioners',
'Overberg Agri Branch: Boltfast',
'Overberg Agri Branch: P & B Limeworks',
'Overberg Agri -Moov Fuel',
'Overberg Agri -Wealth and Risk Managment']


# In[4]:


template_name='automated_loss_ratio_report_template.xlsx'
report_name='{group_name} - care range gap cover - {year} claims experience.pdf'.format(group_name=group_name,year=year)
file_name='Claims vs Premiums {year}.xlsx'.format(year=year)
file_path='C:\\Users\\christo.strydom\\github_repos\\automation\\GenaOosthuizen\\'
path_to_pdf=file_path+report_name


# In[5]:


df_premiums=pd.read_excel(io=file_path+file_name,sheet_name='Premiums')
df_claims=pd.read_excel(io=file_path+file_name,sheet_name='Claims Per Policy')
df_claims_report=pd.read_excel(io=file_path+file_name,sheet_name='Claims Report')


# Establish the name of the report, using the __group_name__ and __year__ attribute

# In[7]:


title='{group_name} - CARE RANGE GAP COVER - {year} CLAIMS EXPERIENCE'.format(group_name=group_name,year=year)


# Slice the premiums dataframe and do some calculations:

# In[39]:


gf_premiums=df_premiums[df_premiums['Policy Group Name'].isin(pgn_list)].fillna(0).copy()
gf_premiums.set_index('Policy Group Name',inplace=True)
gf_premiums.loc['Total Premium',:]=gf_premiums.sum(axis=0)
gf_premiums.loc['Risk Premium',:]=gf_premiums.loc['Total Premium',:]*0.645
risk_premium_values=gf_premiums.loc['Risk Premium',:]
total_premium_values=gf_premiums.loc['Total Premium',:]


# Extract the claims dataframe and perform some calculations:

# In[40]:


gf_claims=df_claims[df_claims['Policy Group Name'].isin(pgn_list)].fillna(0).copy()
gf_claims.set_index('Policy Group Name',inplace=True)
gf_claims.loc['Claims Paid',:]=(-1)*gf_claims.sum(axis=0)


# Calculate the average claim amount:

# In[17]:


gf_claims_report=df_claims_report[df_claims_report['Policy Group Name'].isin(pgn_list)].fillna(0).copy()
average_claim=gf_claims_report['Amount Paid'].mean()


# Calculate the claims ratio, total claims paid to total risk premium for the year:

# In[19]:


claims_ratio=-gf_claims.loc['Claims Paid','Grand Total']/gf_premiums.loc['Risk Premium','TOTAL']


# Calculate claims to total premium, ie claims paid to total premium:

# In[20]:


claims_vs_total_premium=-gf_claims.loc['Claims Paid','Grand Total']/gf_premiums.loc['Total Premium','TOTAL']


# In[21]:


table_headings=list(gf_premiums)


# Open the report template and identify the report sheet:

# In[22]:


file_name='automated_loss_ratio_report_template.xlsx'
automated_loss_ratio_report_template = openpyxl.load_workbook(file_path+file_name) 
summary_sheet = automated_loss_ratio_report_template["Summary"]


# Extract all __named ranges__ necessary for the report:

# In[23]:


claims_vs_total_premium_address = list(automated_loss_ratio_report_template.defined_names['claims_vs_total_premium'].destinations)
claims_ratio_address = list(automated_loss_ratio_report_template.defined_names['claims_ratio'].destinations)
average_claim_address = list(automated_loss_ratio_report_template.defined_names['average_claim'].destinations)
title_cell_address = list(automated_loss_ratio_report_template.defined_names['title_cell'].destinations)
table_heading_address = list(automated_loss_ratio_report_template.defined_names['table_heading'].destinations)
risk_premium_values_address = list(automated_loss_ratio_report_template.defined_names['risk_premium_values'].destinations)
total_premium_values_address = list(automated_loss_ratio_report_template.defined_names['total_premium_values'].destinations)
claims_paid_values_address = list(automated_loss_ratio_report_template.defined_names['claims_paid_values'].destinations)
total_values_address = list(automated_loss_ratio_report_template.defined_names['total_values'].destinations)


# Define arrays for insertion into the report:

# In[24]:


risk_premium_values=list(gf_premiums.loc['Risk Premium',:].values)
total_premium_values=list(gf_premiums.loc['Total Premium',:].values)
claims_paid_values=list(gf_claims.loc['Claims Paid',:].values)
riskpremium_plus_claimspaid=add_list_fn(risk_premium_values,claims_paid_values)


# Populate __named ranges__ in our template with arrays as defined above:

# In[25]:


populate_range_fn(address_array=table_heading_address, xlsx_sheet=summary_sheet, values_array=table_headings)
populate_range_fn(address_array=risk_premium_values_address, xlsx_sheet=summary_sheet, values_array=risk_premium_values)
populate_range_fn(address_array=total_premium_values_address, xlsx_sheet=summary_sheet, values_array=total_premium_values)
populate_range_fn(address_array=claims_paid_values_address, xlsx_sheet=summary_sheet, values_array=claims_paid_values)
populate_range_fn(address_array=total_values_address, xlsx_sheet=summary_sheet, values_array=riskpremium_plus_claimspaid)


# Insert 4 single values (__title__, __average_claim__, __claims_ratio__ and __claims_vs_total_premium__) onto named ranges:

# In[26]:


summary_sheet[claims_vs_total_premium_address[0][1].replace('$', '')] = claims_vs_total_premium
summary_sheet[claims_ratio_address[0][1].replace('$', '')] = claims_ratio
summary_sheet[average_claim_address[0][1].replace('$', '')] = average_claim
summary_sheet[title_cell_address[0][1].replace('$', '')] = title
# summary_sheet[table_heading_address[0][1].replace('$', '')] = table_headings


# Save the template file:

# In[27]:


automated_loss_ratio_report_template.save(file_path+template_name)


# ### Print the pdf

# In[28]:


o = win32com.client.Dispatch("Excel.Application")
o.Visible = False
# open the template file:
wb = o.Workbooks.Open(file_path+template_name)
# find the summary sheet:
ws = wb.Worksheets["Summary"]
# set the print area:
ws.PageSetup.PrintArea = "Print_Area"
# print to pdf:
wb.ActiveSheet.ExportAsFixedFormat(0, path_to_pdf)
# set excel to visible:
o.Visible = True


# In[ ]:




