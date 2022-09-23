#!/usr/bin/env python
# coding: utf-8

# In[13]:


#AD affected


# #  importing pandas package
# 

# In[14]:


import pandas as pd


# # Reading the AD affected in Liang's databse and extracting the list of genes at different sheet

# In[15]:


AD_affected_EC=pd.read_excel('AD affected.xlsx',sheet_name='EC',engine='openpyxl')
AD_affected_HIP=pd.read_excel('AD affected.xlsx',sheet_name='HIP',engine='openpyxl')
AD_affected_PC=pd.read_excel('AD affected.xlsx',sheet_name='PC',engine='openpyxl')
AD_affected_MTG=pd.read_excel('AD affected.xlsx',sheet_name='MTG',engine='openpyxl')
AD_affected_SFG=pd.read_excel('AD affected.xlsx',sheet_name='SFG',engine='openpyxl')
AD_affected_VCX=pd.read_excel('AD affected.xlsx',sheet_name='VCX',engine='openpyxl')
AD_affected_EC=list(AD_affected_EC['symbol'])
AD_affected_HIP=list(AD_affected_HIP['symbol'])
AD_affected_PC=list(AD_affected_PC['symbol'])
AD_affected_MTG=list(AD_affected_MTG['symbol'])
AD_affected_SFG=list(AD_affected_SFG['symbol'])
AD_affected_VCX=list(AD_affected_VCX['symbol'])


# # Reading the MMC7 file with different sheet name. 
# 

# In[16]:


mmc7_Herpesvirus_virus2host=pd.read_excel('mmc7_Herpesvirus.xlsx',sheet_name='virus2host',engine='openpyxl')
mmc7_Herpesvirus_host2virus=pd.read_excel('mmc7_Herpesvirus.xlsx',sheet_name='host2virus',engine='openpyxl')


# # Filtering "MMC7" with sheetname "virus2host" in selecting rows with having the same gene with Liang( AD affected) in a "hostGene" column.

# In[17]:


intresect_virus2host_AD_affected_EC = mmc7_Herpesvirus_virus2host[mmc7_Herpesvirus_virus2host['hostGene'].isin(AD_affected_EC)]
intresect_virus2host_AD_affected_HIP = mmc7_Herpesvirus_virus2host[mmc7_Herpesvirus_virus2host['hostGene'].isin(AD_affected_HIP)]
intresect_virus2host_AD_affected_PC = mmc7_Herpesvirus_virus2host[mmc7_Herpesvirus_virus2host['hostGene'].isin(AD_affected_PC)]
intresect_virus2host_AD_affected_MTG = mmc7_Herpesvirus_virus2host[mmc7_Herpesvirus_virus2host['hostGene'].isin(AD_affected_MTG)]
intresect_virus2host_AD_affected_SFG= mmc7_Herpesvirus_virus2host[mmc7_Herpesvirus_virus2host['hostGene'].isin(AD_affected_SFG)]
intresect_virus2host_AD_affected_VCX= mmc7_Herpesvirus_virus2host[mmc7_Herpesvirus_virus2host['hostGene'].isin(AD_affected_VCX)]


# # Filtering "MMC7" with sheetname "host2virus" in selecting rows with having the same gene with Liang( AD affected) in a "hostGene" column.

# In[18]:


intresect_host2virus_AD_affected_EC = mmc7_Herpesvirus_host2virus[mmc7_Herpesvirus_host2virus['hostGene'].isin(AD_affected_EC)]
intresect_host2virus_AD_affected_HIP = mmc7_Herpesvirus_host2virus[mmc7_Herpesvirus_host2virus['hostGene'].isin(AD_affected_HIP)]
intresect_host2virus_AD_affected_PC = mmc7_Herpesvirus_host2virus[mmc7_Herpesvirus_host2virus['hostGene'].isin(AD_affected_PC)]
intresect_host2virus_AD_affected_MTG = mmc7_Herpesvirus_host2virus[mmc7_Herpesvirus_host2virus['hostGene'].isin(AD_affected_MTG)]
intresect_host2virus_AD_affected_SFG= mmc7_Herpesvirus_host2virus[mmc7_Herpesvirus_host2virus['hostGene'].isin(AD_affected_SFG)]
intresect_host2virus_AD_affected_VCX= mmc7_Herpesvirus_host2virus[mmc7_Herpesvirus_host2virus['hostGene'].isin(AD_affected_VCX)]


# # Saving the filtering MMC7 based on Liang( AD affected) as a CSV file

# In[19]:


writer = pd.ExcelWriter('AD-Affected.xlsx', engine='xlsxwriter')
intresect_virus2host_AD_affected_EC.to_excel(writer, 'virustohost-EC')
intresect_virus2host_AD_affected_SFG.to_excel(writer, 'virustohost-SFG')
intresect_virus2host_AD_affected_HIP.to_excel(writer, 'virustohost-HIP')
intresect_virus2host_AD_affected_PC.to_excel(writer, 'virustohost-PC')
intresect_virus2host_AD_affected_MTG.to_excel(writer, 'virustohost-MTG')
intresect_virus2host_AD_affected_VCX.to_excel(writer, 'virustohost-VCX')
intresect_host2virus_AD_affected_EC.to_excel(writer, 'host2virus-EC')
intresect_host2virus_AD_affected_HIP.to_excel(writer, 'host2virus-HIP')
intresect_host2virus_AD_affected_PC.to_excel(writer, 'host2virus-PC')
intresect_host2virus_AD_affected_MTG.to_excel(writer, 'host2virus-MTG')
intresect_host2virus_AD_affected_SFG.to_excel(writer, 'host2virus-SFG')
intresect_host2virus_AD_affected_VCX.to_excel(writer, 'host2virus-VCX')
writer.save()


# # Reading the Non demanted in Liang's databse and extracting the list of genes at different sheet

# In[20]:


Non_demented_EC=pd.read_excel('Liang-non-demented.xlsx',sheet_name='entorhinal cortex',engine='openpyxl')
Non_demented_HIP=pd.read_excel('Liang-non-demented.xlsx',sheet_name='hippocampus',engine='openpyxl')
Non_demented_PC=pd.read_excel('Liang-non-demented.xlsx',sheet_name='middle temporal gyrus',engine='openpyxl')
Non_demented_MTG=pd.read_excel('Liang-non-demented.xlsx',sheet_name='posterior cingulate corrtex',engine='openpyxl')
Non_demented_SFG=pd.read_excel('Liang-non-demented.xlsx',sheet_name='superior frontal gyrus',engine='openpyxl')
Non_demented_VCX=pd.read_excel('Liang-non-demented.xlsx',sheet_name='primary visual cortex',engine='openpyxl')
Non_demented_EC=list(Non_demented_EC['symbol'])
Non_demented_HIP=list(Non_demented_HIP['symbol'])
Non_demented_PC=list(Non_demented_PC['symbol'])
Non_demented_MTG=list(Non_demented_MTG['symbol'])
Non_demented_SFG=list(Non_demented_SFG['symbol'])
Non_demented_VCX=list(Non_demented_VCX['symbol'])


# # Filtering "MMC7" with sheetname "virus2host" in selecting rows with having the same gene with Liang( Non demented) in a "hostGene" column.

# In[21]:


intresect_virus2host_Non_demented_EC = mmc7_Herpesvirus_virus2host[mmc7_Herpesvirus_virus2host['hostGene'].isin(Non_demented_EC)]
intresect_virus2host_Non_demented_HIP = mmc7_Herpesvirus_virus2host[mmc7_Herpesvirus_virus2host['hostGene'].isin(Non_demented_HIP)]
intresect_virus2host_Non_demented_PC = mmc7_Herpesvirus_virus2host[mmc7_Herpesvirus_virus2host['hostGene'].isin(Non_demented_PC)]
intresect_virus2host_Non_demented_MTG = mmc7_Herpesvirus_virus2host[mmc7_Herpesvirus_virus2host['hostGene'].isin(Non_demented_MTG)]
intresect_virus2host_Non_demented_SFG= mmc7_Herpesvirus_virus2host[mmc7_Herpesvirus_virus2host['hostGene'].isin(Non_demented_SFG)]
intresect_virus2host_Non_demented_VCX= mmc7_Herpesvirus_virus2host[mmc7_Herpesvirus_virus2host['hostGene'].isin(Non_demented_VCX)]


# # Filtering "MMC7" with sheetname "host2virus" in selecting rows with having the same gene with Liang(Non demented) in a "hostGene" column.

# In[22]:


intresect_host2virus_Non_demented_EC = mmc7_Herpesvirus_host2virus[mmc7_Herpesvirus_host2virus['hostGene'].isin(Non_demented_EC)]
intresect_host2virus_Non_demented_HIP = mmc7_Herpesvirus_host2virus[mmc7_Herpesvirus_host2virus['hostGene'].isin(Non_demented_HIP)]
intresect_host2virus_Non_demented_PC = mmc7_Herpesvirus_host2virus[mmc7_Herpesvirus_host2virus['hostGene'].isin(Non_demented_PC)]
intresect_host2virus_Non_demented_MTG = mmc7_Herpesvirus_host2virus[mmc7_Herpesvirus_host2virus['hostGene'].isin(Non_demented_MTG)]
intresect_host2virus_Non_demented_SFG= mmc7_Herpesvirus_host2virus[mmc7_Herpesvirus_host2virus['hostGene'].isin(Non_demented_SFG)]
intresect_host2virus_Non_demented_VCX= mmc7_Herpesvirus_host2virus[mmc7_Herpesvirus_host2virus['hostGene'].isin(Non_demented_VCX)]


# # Saving the filtering MMC7 based on Liang( Non demented) as a CSV file
# 

# In[62]:


writer = pd.ExcelWriter('Non-demented.xlsx', engine='xlsxwriter')
intresect_virus2host_Non_demented_EC.to_excel(writer, 'virustohost-EC')
intresect_virus2host_Non_demented_MTG.to_excel(writer, 'virustohost-MTG')
intresect_virus2host_Non_demented_SFG.to_excel(writer, 'virustohost-SFG')
intresect_host2virus_Non_demented_SFG.to_excel(writer, 'hosttovirus-SFG')
intresect_virus2host_Non_demented_HIP.to_excel(writer, 'virustohost-HIP')
intresect_virus2host_Non_demented_PC.to_excel(writer, 'virustohost-PC')
intresect_virus2host_Non_demented_VCX.to_excel(writer, 'virustohost-VCX')
intresect_host2virus_Non_demented_EC.to_excel(writer, 'hosttovirus-EC')
intresect_host2virus_Non_demented_HIP.to_excel(writer, 'hosttovirus-HIP')
intresect_host2virus_Non_demented_PC.to_excel(writer, 'hosttovirus-PC')
intresect_host2virus_Non_demented_MTG.to_excel(writer, 'hosttovirus-MTG')
intresect_host2virus_Non_demented_VCX.to_excel(writer, 'hosttovirus-VCX')

writer.save()


# # Reading the Normal aged in Liang's databse and extracting the list of genes at different sheet

# In[23]:


Normal_aged_EC=pd.read_excel('Liang-normal-aged.xlsx',sheet_name='EC',engine='openpyxl')
Normal_aged_HIP=pd.read_excel('Liang-normal-aged.xlsx',sheet_name='HIP',engine='openpyxl')
Normal_aged_MTG=pd.read_excel('Liang-normal-aged.xlsx',sheet_name='MTG',engine='openpyxl')
Normal_aged_PC=pd.read_excel('Liang-normal-aged.xlsx',sheet_name='PC',engine='openpyxl')
Normal_aged_SFG=pd.read_excel('Liang-normal-aged.xlsx',sheet_name='SFG',engine='openpyxl')
Normal_aged_VCX=pd.read_excel('Liang-normal-aged.xlsx',sheet_name='VCX',engine='openpyxl')
Normal_aged_EC=list(Normal_aged_EC['symbol'])
Normal_aged_HIP=list(Normal_aged_HIP['symbol'])
Normal_aged_PC=list(Normal_aged_PC['symbol'])
Normal_aged_MTG=list(Normal_aged_MTG['symbol'])
Normal_aged_SFG=list(Normal_aged_SFG['symbol'])
Normal_aged_VCX=list(Normal_aged_VCX['symbol'])


# # Filtering "MMC7" with sheetname "virus2host" in selecting rows with having the same gene with Liang(Normal aged) in a "hostGene" column.

# In[67]:


intresect_virus2host_Normal_aged_EC = mmc7_Herpesvirus_virus2host[mmc7_Herpesvirus_virus2host['hostGene'].isin(Normal_aged_EC)]
intresect_virus2host_Normal_aged_HIP = mmc7_Herpesvirus_virus2host[mmc7_Herpesvirus_virus2host['hostGene'].isin(Normal_aged_HIP)]
intresect_virus2host_Normal_aged_PC = mmc7_Herpesvirus_virus2host[mmc7_Herpesvirus_virus2host['hostGene'].isin(Normal_aged_PC)]
intresect_virus2host_Normal_aged_MTG = mmc7_Herpesvirus_virus2host[mmc7_Herpesvirus_virus2host['hostGene'].isin(Normal_aged_MTG)]
intresect_virus2host_Normal_aged_SFG= mmc7_Herpesvirus_virus2host[mmc7_Herpesvirus_virus2host['hostGene'].isin(Normal_aged_SFG)]
intresect_virus2host_Normal_aged_VCX= mmc7_Herpesvirus_virus2host[mmc7_Herpesvirus_virus2host['hostGene'].isin(Normal_aged_VCX)]


# # Filtering "MMC7" with sheetname "host2virus" in selecting rows with having the same gene with Liang(Normal aged) in a "hostGene" column.

# In[84]:


intresect_host2virus_Normal_aged_EC = mmc7_Herpesvirus_host2virus[mmc7_Herpesvirus_host2virus['hostGene'].isin(Normal_aged_EC)]
intresect_host2virus_Normal_aged_HIP = mmc7_Herpesvirus_host2virus[mmc7_Herpesvirus_host2virus['hostGene'].isin(Normal_aged_HIP)]
intresect_host2virus_Normal_aged_PC = mmc7_Herpesvirus_host2virus[mmc7_Herpesvirus_host2virus['hostGene'].isin(Normal_aged_PC)]
intresect_host2virus_Normal_aged_MTG = mmc7_Herpesvirus_host2virus[mmc7_Herpesvirus_host2virus['hostGene'].isin(Normal_aged_MTG)]
intresect_host2virus_Normal_aged_SFG= mmc7_Herpesvirus_host2virus[mmc7_Herpesvirus_host2virus['hostGene'].isin(Normal_aged_SFG)]
intresect_host2virus_Normal_aged_VCX= mmc7_Herpesvirus_host2virus[mmc7_Herpesvirus_host2virus['hostGene'].isin(Normal_aged_VCX)]


# # Saving the filtering MMC7 based on Liang( Normal aged) as a CSV file
# 

# In[91]:


writer = pd.ExcelWriter('Normal-aged.xlsx', engine='xlsxwriter')
intresect_virus2host_Normal_aged_MTG.to_excel(writer, 'virustohost-MTG')
intresect_host2virus_Normal_aged_MTG.to_excel(writer, 'hosttovirus-MTG')
intresect_virus2host_Normal_aged_EC.to_excel(writer, 'virustohost-EC')
intresect_virus2host_Normal_aged_HIP.to_excel(writer, 'virustohost-HIP')
intresect_virus2host_Normal_aged_PC.to_excel(writer, 'virustohost-PC')
intresect_virus2host_Normal_aged_SFG.to_excel(writer, 'virustohost-SFG')
intresect_virus2host_Normal_aged_VCX.to_excel(writer, 'virustohost-VCX')
intresect_host2virus_Normal_aged_EC.to_excel(writer, 'hosttovirus-EC')
intresect_host2virus_Normal_aged_HIP.to_excel(writer, 'hosttovirus-HIP')
intresect_host2virus_Normal_aged_PC.to_excel(writer, 'hosttovirus-PC')
intresect_host2virus_Normal_aged_SFG.to_excel(writer, 'hosttovirus-SFG')
intresect_host2virus_Normal_aged_VCX.to_excel(writer, 'hosttovirus-VCX')
writer.save()


# In[ ]:




