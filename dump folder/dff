#!/usr/bin/env python
# coding: utf-8

# In[4]:


get_ipython().run_cell_magic('capture', '', '\nimport voila;\nimport numpy as np;\nimport pandas as pd;\nfrom pyFolder.GoogleSheetsAPI import *;\nfrom pyFolder.TableauRestAPi import *;\nfrom pyFolder.gmailAPI import *;\nfrom pretty_html_table import build_table')


# In[ ]:


# !pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib;
# !pip install voila;
# !pip install pandas;
# !pip install numpy;
# !pip install --upgrade tableau-api-lib
# !pip install widgetsnbextension
# !pip install ipywidgets
# !jupyter nbextension enable --py widgetsnbextension --sys-prefix


# In[20]:


# Create email list
recepients = 'paulg@cabadesign.co'


# In[47]:



def mo_src_issue_msg_html(html_table, Url):
    header = 'MOs With Wrong Source Location'
    msg_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
      <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
      <title>html title</title>
      <style type="text/css" media="screen">
    .attribution {{ color: #aaaaaa; font-size: 8pt }}
    .greeting {{ font-size: 14pt; font-styweight: bold}}
    </style></head>
    <body>
    <div>
    <span class="greeting">Hello,</span>
    <p>Your spreadsheet has been updated!
       There are MOs with incorrect source location!
    </p>
    </div>
    <div>
    Below is a preview of the table updated
    {html_table}
    </div>
    <p class="attribution">
    <a href={Url}>
    Link to Spreadsheet
    </a>
    This email was sent by a bot!
    </p>
    </body></html>
    """
    return {'header':header,'body':msg_html}

def shopify_issue_msg_html(html_table, Url):
    header = 'Shopify Orders Not Fulfilled'
    msg_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
      <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
      <title>html title</title>
      <style type="text/css" media="screen">
    .attribution {{ color: #aaaaaa; font-size: 8pt }}
    .greeting {{ font-size: 14pt; font-styweight: bold}}
    </style></head>
    <body>
    <div>
    <span class="greeting">Hello,</span>
    <p>Your spreadsheet has been updated!
       There Orders Not Fulfilled in Shopify!
    </p>
    </div>
    <div>
    Below is a preview of the table updated
    {html_table}
    </div>
    <p class="attribution">
    <a href={Url}>
    Link to Spreadsheet
    </a>
    This email was sent by a bot!
    </p>
    </body></html>
    """
    return {'header':header,'body':msg_html}

def picking_issue_msg_html(html_table, Url):
    header = 'Picking List Missing'
    msg_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
      <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
      <title>html title</title>
      <style type="text/css" media="screen">
    .attribution {{ color: #aaaaaa; font-size: 8pt }}
    .greeting {{ font-size: 14pt; font-styweight: bold}}
    </style></head>
    <body>
    <div>
    <span class="greeting">Hello,</span>
    <p>Your spreadsheet has been updated!</p>
    <b>There are Orders With Picking List Missing!</b>
    </div>
    <div>
    Below is a preview of the table updated
    {html_table}
    </div>
    <p class="attribution">
    <a href={Url}>
    Link to Spreadsheet
    </a>
    This email was sent by a bot!
    </p>
    </body></html>
    """
    return {'header':header,'body':msg_html}


# In[45]:


def sendImport(Url, bi_directory, worksheet_name, view_filters={}):
    my_workbook = list(bi_directory.values())[0]
    my_worksheet = list(bi_directory.values())[1]
    creds=('ivank', 'Prius@12345')
    try:
        for i in range(len(my_worksheet)):
            if not view_filters:
                df = tSCRestAPI(creds, my_workbook, my_worksheet[i])
            elif len(view_filters)==1:
                dim1 = list(view_filters.keys())[0]
                val1 = list(view_filters.values())[0]
                df = tSCRestAPI(creds, my_workbook, my_worksheet[i], dim1, val1)
            elif len(view_filters)==2:
                dim1 = list(view_filters.keys())[0]
                val1 = list(view_filters.values())[0]
                dim2 = list(view_filters.keys())[1]
                val2 = list(view_filters.values())[1]
                df = tSCRestAPI(creds, my_workbook, my_worksheet[i], dim1, val1, dim2, val2)
            gsheetId = getSheetId(Url) 
            clearSheets(gsheetId,worksheet_name[i])
            writeDataToSheetDf(worksheet_name[i],gsheetId,df)
            # build table arguments
            if len(df) >20:
                df_html = df.head()
            else:
                df_html = df
            html_table = build_table(df_html
                        , 'yellow_dark'
                        , font_size='medium'
                        , font_family='Open Sans, sans-serif'
                        , text_align='center'
                        , width='auto'
                        , index=False
                        ,conditions={
                            'Age': {
                                'min': 25,
                                'max': 60,
                                'min_color': 'red',
                                'max_color': 'green',
                            }
                        }
                                     , even_color='black'
                                     , even_bg_color='white')
            #sendEmail(msg_string, recepient, subject, template='html')
            header = list(picking_issue_msg_html(html_table, Url).values())[0]
            msg_html = list(picking_issue_msg_html(html_table, Url).values())[1]
            sendEmail(msg_html, recepients, header)
    except Exception as e:
            print(e)
            




# In[17]:


# get LA Stock Data and send to spreadsheet
#sendImport(Url, bi_directory, view_filters={}, worksheet_name)
Url = 'https://docs.google.com/spreadsheets/d/1xQZbQps8aPDniv8iC8gFfviWppYNaXLBuLRbo5eymZ0/edit#gid=1436299014'
worksheet_name = ['LA Stock Import', 'Remaining to Receive Import', 'PO Received Qty Import']
bi_directory = {'my_workbook': 'test inventory', 'my_worksheet': ['LA Stock', 'Remaining to Receive', 'PO Moves']}
view_filters = {'Negative Inventory': 'negative inv'}
sendImport(Url, bi_directory, worksheet_name)


# In[30]:


# MO Source Location Issue
Url = 'https://docs.google.com/spreadsheets/d/1yVbEtF66ViW-Rk6yG4K1H9JUJAP4eLx0ywF5bpqJj-0/edit#gid=0'
worksheet_name = ['MO Source Location Issue']
bi_directory = {'my_workbook': 'MO Checks', 'my_worksheet': ['MO Source Location Issue']}
view_filters = {'Negative Inventory': 'negative inv'}
sendImport(Url, bi_directory, worksheet_name)


# In[32]:


# Shopify Fulfillment Issue
recepients = 'paulg@cabadesign.co,roberto@cabadesign.co'
Url = 'https://docs.google.com/spreadsheets/d/1yVbEtF66ViW-Rk6yG4K1H9JUJAP4eLx0ywF5bpqJj-0/edit#gid=0'
worksheet_name = ['Shopify Fulfillment Issue']
bi_directory = {'my_workbook': 'SO Status', 'my_worksheet': ['Shopify Fulfillment Issue']}
view_filters = {'Negative Inventory': 'negative inv'}
sendImport(Url, bi_directory, worksheet_name)


# In[49]:


# Picking List Missing
recepients = 'paulg@cabadesign.co,ivank@cabadesign.co'
Url = 'https://docs.google.com/spreadsheets/d/1yVbEtF66ViW-Rk6yG4K1H9JUJAP4eLx0ywF5bpqJj-0/edit#gid=0'
worksheet_name = ['Picking List Missing']
bi_directory = {'my_workbook': 'SO Status', 'my_worksheet': ['Picking List Missing']}
view_filters = {'Negative Inventory': 'negative inv'}
sendImport(Url, bi_directory, worksheet_name)

