#!/usr/bin/env python
# coding: utf-8

# #  importing pandas , numpy and matplotlib.pyplot packages

# In[21]:


import pandas as pd
import numpy as np
import matplotlib.pyplot as plt


# # Reading the PDB databse

# In[22]:


ATG_PDB = pd.read_csv('PDB.txt')
ATG_PDB=ATG_PDB['symbol'].values.tolist()
# len(ATG_PDB)


# # Reading the Morishita & Mizushima's database
# 
# 

# In[23]:


ATG_Mizushima = pd.read_csv('Murshita.txt')
ATG_Mizushima=ATG_Mizushima['symbol'].values.tolist()
# len(ATG_Mizushima)


# # Reading the Tanpaku database
# 

# In[24]:


ATG_tanpaku = pd.read_csv('Japaness(tanpaku).txt')
ATG_tanpaku=ATG_tanpaku['symbol'].values.tolist()
# len(ATG_tanpaku)


# # Reading the other papars

# In[25]:


ATG_Isaac = pd.read_excel('Pathogenic Single Nucleotide Polymorphisms on Autophagy-Related Genes.xlsx')
ATG_Isaac=ATG_Isaac['Gene'].values.tolist()
ATG_Isaac =[x.strip(' ') for x in ATG_Isaac]
# len(ATG_Isaac)


# # Reading Tudor I. Oprea's dataset

# In[26]:


ATG_Tudor=pd.read_excel('Tudor Opera.xlsx',sheet_name='Input_Output',engine='openpyxl')
ATG_Tudor=ATG_Tudor['symbol'].values.tolist()
# len(ATG_Tudor)


# # Reading Tudor I. Oprea's + dark genes dataset
# 

# In[27]:


dark_genes = pd.read_csv('dark_genes.txt')
dark_genes=dark_genes['symbol'].values.tolist()
# len(dark_genes)


# # Making a set of the ATG genes (comprehensive list) and removing a NaN value from the list. The total number of comprehensive list is 9812

# In[28]:


comprehensive= ATG_PDB + ATG_Mizushima + ATG_tanpaku + ATG_Isaac + ATG_Tudor + dark_genes
comprehensive= set(comprehensive)
comprehensive= [str(x) for x in comprehensive]
comprehensive = [x for x in comprehensive if x !='nan']
len(comprehensive)


# # Converting the list of comprehensive to dataframe
# # ( https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.html)

# In[29]:


df = pd.DataFrame (comprehensive,columns=['symbol'])
df


# # Saving the comprehensive dataframe as a CSV file

# In[14]:


df.to_csv('comprehensive.csv')


# # Reading the MMC7 file with different sheet name. 

# In[31]:


mmc7_virus2host=pd.read_excel('1-s2.0-S0896627318304215-mmc7.xlsx',sheet_name='virus2host',engine='openpyxl')
mmc7_host2virus=pd.read_excel('1-s2.0-S0896627318304215-mmc7.xlsx',sheet_name='host2virus',engine='openpyxl')


# # Filtering "MMC7" with sheetname "virus2host" in selecting rows with having the comprehensive value in a "hostGene" column.

# In[39]:


intresect_virus2host_comprehensive = mmc7_virus2host[mmc7_virus2host['hostGene'].isin(comprehensive)]
intresect_virus2host_comprehensive


# # Filtering "MMC7" with sheetname "host2virus" in selecting rows with having the comprehensive value in a "hostGene" column.

# In[40]:


intresect_host2virus_comprehensive = mmc7_host2virus[mmc7_host2virus['hostGene'].isin(comprehensive)]
intresect_host2virus_comprehensive


# # Length of ATG genes in MMC7-virus2host

# In[45]:


len(intresect_virus2host_comprehensive['hostGene'])


# # Length of ATG genes in MMC7-host2virus

# In[48]:


len(intresect_host2virus_comprehensive['hostGene'])


# #  Length of unique ATG genes in MMC7-virus2host

# In[49]:


len(set(intresect_virus2host_comprehensive['hostGene']))


# #  Length of unique ATG genes in MMC7-host2virus

# In[50]:


len(set(intresect_host2virus_comprehensive['hostGene']))


# # Filtering based on "Herpesvirus"

# In[58]:


list_herpesvirus_virus2host=intresect_virus2host_comprehensive[intresect_virus2host_comprehensive['virus_name'].str.contains('herpesvirus')]
list_herpesvirus_host2virus=intresect_host2virus_comprehensive[intresect_host2virus_comprehensive['virus_name'].str.contains('herpesvirus')]


# # Length of ATG genes in MMC7-virus2host based on "Herpesvirus"

# In[69]:


len(list_herpesvirus_virus2host['hostGene'])


# #  Length of ATG genes in MMC7-host2virus based on "Herpesvirus"

# In[70]:


len(list_herpesvirus_host2virus['hostGene'])


# # Length of unique ATG genes in MMC7-virus2host based on "Herpesvirus"

# In[72]:


len(set(list_herpesvirus_virus2host['hostGene']))


# #  Length of unique ATG genes in MMC7-host2virus based on "Herpesvirus"

# In[73]:


len(set(list_herpesvirus_host2virus['hostGene']))


# # Saving the filtering MMC7 as a CSV file

# In[16]:


writer = pd.ExcelWriter('mmc7_Herpesvirus.xlsx', engine='xlsxwriter')
list_herpesvirus_virus2host.to_excel(writer, 'virus2host')
list_herpesvirus_host2virus.to_excel(writer, 'host2virus')

writer.save()

