#!/usr/bin/env python
# coding: utf-8

# In[1]:


from credsTest.GoogleSheetsAPI import *;
import numpy as np;
import pandas as pd;
from credsTest.TableauRestAPi import *;


# In[2]:


def sendImport(Url, bi_directory, worksheet_name, view_filters={}):
    my_workbook = list(bi_directory.values())[0]
    my_worksheet = list(bi_directory.values())[1]
    if not view_filters:
        df = tSCRestAPI(my_workbook, my_worksheet)
    elif len(view_filters)==1:
        dim1 = list(view_filters.keys())[0]
        val1 = list(view_filters.values())[0]
        df = tSCRestAPI(my_workbook, my_worksheet, dim1, val1)
    elif len(view_filters)==2:
        dim1 = list(view_filters.keys())[0]
        val1 = list(view_filters.values())[0]
        dim2 = list(view_filters.keys())[1]
        val2 = list(view_filters.values())[1]
        df = tSCRestAPI(my_workbook, my_worksheet, dim1, val1, dim2, val2)
    gsheetId = getSheetId(Url) 
    clearSheets(gsheetId,worksheet_name)
    writeDataToSheetDf(worksheet_name,gsheetId,df)


# In[3]:


# get LA Stock Data and send to spreadsheet
#sendImport(Url, bi_directory, view_filters={}, worksheet_name)
Url = 'https://docs.google.com/spreadsheets/d/1EPo-yeHzap_BKWz0VvdOpFsN1Atw4jL2ZA7aXhl1aps/edit#gid=273865484'
worksheet_name = 'Sheet7'
bi_directory = {'my_workbook': 'test inventory', 'my_worksheet': 'LA Stock'}
view_filters = {'Negative Inventory': 'negative inv'}
sendImport(Url, bi_directory, worksheet_name, view_filters)


# In[ ]:


<table>
  <tr>
    <th>Has Negative Inventory</th>
    <th>Location</th>
    <th>Count</th>
  </tr>
  <tr>
    <td>True</td>
    <td>LA STOCK</td>
    <td>20</td>
  </tr>
</table>

