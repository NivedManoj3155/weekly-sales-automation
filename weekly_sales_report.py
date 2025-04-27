#!/usr/bin/env python
# coding: utf-8

# In[1]:


conda install pandas openpyxl
pip install gspread google-auth


# In[3]:


get_ipython().system('conda install pandas openpyxl -y')
get_ipython().system('pip install gspread google-auth')


# In[11]:


import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# Step 1: Setup credentials
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
SERVICE_ACCOUNT_FILE = 'service_account_key.json'

credentials = Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
client = gspread.authorize(credentials)

# Step 2: Load Excel files
export_df = pd.read_excel('Export.xlsx')
base_nf_df = pd.read_excel('Base_NF.xlsx')

# Step 3: Merge MRP into base sales file
merged_df = base_nf_df.merge(export_df[['SKU', 'Variant Compare At Price']],
                             how='left', left_on='SKU', right_on='SKU')

# Step 4: Create 'Returned' and 'Offtake' columns
merged_df['Returned'] = merged_df['Net items sold'].apply(lambda x: 'YES' if x < 0 else 'NO')
merged_df['Offtake'] = merged_df['Net items sold'] * merged_df['Variant Compare At Price']

# Step 5: Filter vendors
vendors = ['Barkbulter', 'Truelove']
vendor_df = merged_df[merged_df['Vendor'].isin(vendors)]

# Step 6: Create weekly report (last 7 days)
vendor_df['Order Date'] = pd.to_datetime(vendor_df['Order Date'])  # adjust column name if needed
last_7_days = vendor_df[vendor_df['Order Date'] >= pd.Timestamp.now() - pd.Timedelta(days=7)]

# Step 7: Group data for Dashboard
dashboard_df = last_7_days.groupby(['Product title at time of sale', 'Product type']).agg({
    'Offtake': 'sum',
    'Net items sold': 'mean'
}).reset_index()

dashboard_df.rename(columns={
    'Net items sold': 'Daily Sale Units',
}, inplace=True)

# Step 8: Upload to Google Sheets
dashboard_sheet = client.open('Vendor_Weekly_Dashboard').worksheet('Dashboard')  # change if needed
dashboard_sheet.clear()

dashboard_sheet.update([dashboard_df.columns.values.tolist()] + dashboard_df.values.tolist())

print("Dashboard updated successfully 🚀")


# In[7]:


SERVICE_ACCOUNT_FILE = 'C:/Users/YourName/Downloads/service_account_key.json'


# In[13]:


import os
os.getcwd()


# In[15]:


import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# Step 1: Setup credentials
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
SERVICE_ACCOUNT_FILE = 'service_account_key.json'

credentials = Credentials.from_service_account_file(
    SERVICE_A
CCOUNT_FILE, scopes=SCOPES)
client = gspread.authorize(credentials)

# Step 2: Load Excel files
export_df = pd.read_excel('Export.xlsx')
base_nf_df = pd.read_excel('Base_NF.xlsx')

# Step 3: Merge MRP into base sales file
merged_df = base_nf_df.merge(export_df[['SKU', 'Variant Compare At Price']],
                             how='left', left_on='SKU', right_on='SKU')

# Step 4: Create 'Returned' and 'Offtake' columns
merged_df['Returned'] = merged_df['Net items sold'].apply(lambda x: 'YES' if x < 0 else 'NO')
merged_df['Offtake'] = merged_df['Net items sold'] * merged_df['Variant Compare At Price']

# Step 5: Filter vendors
vendors = ['Barkbulter', 'Truelove']
vendor_df = merged_df[merged_df['Vendor'].isin(vendors)]

# Step 6: Create weekly report (last 7 days)
vendor_df['Order Date'] = pd.to_datetime(vendor_df['Order Date'])  # adjust column name if needed
last_7_days = vendor_df[vendor_df['Order Date'] >= pd.Timestamp.now() - pd.Timedelta(days=7)]

# Step 7: Group data for Dashboard
dashboard_df = last_7_days.groupby(['Product title at time of sale', 'Product type']).agg({
    'Offtake': 'sum',
    'Net items sold': 'mean'
}).reset_index()

dashboard_df.rename(columns={
    'Net items sold': 'Daily Sale Units',
}, inplace=True)

# Step 8: Upload to Google Sheets
dashboard_sheet = client.open('Vendor_Weekly_Dashboard').worksheet('Dashboard')  # change if needed
dashboard_sheet.clear()

dashboard_sheet.update([dashboard_df.columns.values.tolist()] + dashboard_df.values.tolist())

print("Dashboard updated successfully 🚀")


# In[17]:


import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# Step 1: Setup credentials
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
SERVICE_ACCOUNT_FILE = 'service_account_key.json'  # Make sure this file is in same folder

credentials = Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
client = gspread.authorize(credentials)

# Step 2: Load Excel files
export_df = pd.read_excel('Export.xlsx')  # Your export file
base_nf_df = pd.read_excel('Base_NF.xlsx')  # Your base file

# Step 3: Merge MRP into base sales file
merged_df = base_nf_df.merge(export_df[['SKU', 'Variant Compare At Price']],
                             how='left', left_on='SKU', right_on='SKU')

# Step 4: Create 'Returned' and 'Offtake' columns
merged_df['Returned'] = merged_df['Net items sold'].apply(lambda x: 'YES' if x < 0 else 'NO')
merged_df['Offtake'] = merged_df['Net items sold'] * merged_df['Variant Compare At Price']

# Step 5: Filter vendors
vendors = ['Barkbulter', 'Truelove']
vendor_df = merged_df[merged_df['Vendor'].isin(vendors)]

# Step 6: Create weekly report (last 7 days)
vendor_df['Order Date'] = pd.to_datetime(vendor_df['Order Date'])  # Make sure this column exists
last_7_days = vendor_df[vendor_df['Order Date'] >= pd.Timestamp.now() - pd.Timedelta(days=7)]

# Step 7: Group data for Dashboard
dashboard_df = last_7_days.groupby(['Product title at time of sale', 'Product type']).agg({
    'Offtake': 'sum',
    'Net items sold': 'mean'
}).reset_index()

dashboard_df.rename(columns={
    'Net items sold': 'Daily Sale Units',
}, inplace=True)

# Step 8: Upload to Google Sheets
dashboard_sheet = client.open('Vendor_Weekly_Dashboard').worksheet('Dashboard')  # Sheet name must exist
dashboard_sheet.clear()

# Update the sheet
dashboard_sheet.update([dashboard_df.columns.values.tolist()] + dashboard_df.values.tolist())

print("Dashboard updated successfully 🚀")


# In[19]:


credentials = Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
client = gspread.authorize(credentials)


# In[21]:


import os
print(os.getcwd())


# In[49]:


import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# Step 1: Setup credentials
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
SERVICE_ACCOUNT_FILE = 'service_account_key.json'

credentials = Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
client = gspread.authorize(credentials)

# Step 2: Load Excel files
export_df = pd.read_excel(r"C:\Users\Supertails\Downloads\Export_2025-04-21_131843 (5).xlsx")
base_nf_df = pd.read_excel(r"C:\Users\Supertails\Downloads\base_nf.xlsx")




# Step 3: Merge MRP into base sales file
merged_df = base_nf_df.merge(export_df[['SKU', 'Variant Compare At Price']],
                             how='left', left_on='SKU', right_on='SKU')

# Step 4: Create 'Returned' and 'Offtake' columns
merged_df['Returned'] = merged_df['Net items sold'].apply(lambda x: 'YES' if x < 0 else 'NO')
merged_df['Offtake'] = merged_df['Net items sold'] * merged_df['Variant Compare At Price']

# Step 5: Filter vendors
vendors = ['Barkbulter', 'Truelove']
vendor_df = merged_df[merged_df['Vendor'].isin(vendors)]

# Step 6: Create weekly report (last 7 days)
vendor_df['Order Date'] = pd.to_datetime(vendor_df['Order Date'])  # adjust column name if needed
last_7_days = vendor_df[vendor_df['Order Date'] >= pd.Timestamp.now() - pd.Timedelta(days=7)]


# Step 7: Group data for Dashboard
dashboard_df = last_7_days.groupby(['Product title at time of sale', 'Product type']).agg({
    'Offtake': 'sum',
    'Net items sold': 'mean'
}).reset_index()

dashboard_df.rename(columns={
    'Net items sold': 'Daily Sale Units',
}, inplace=True)

# Step 8: Upload to Google Sheets
dashboard_sheet = client.open('Pupscribe weekly  sales dashboard').worksheet('Dashboard')  # change if needed
dashboard_sheet.clear()

dashboard_sheet.update([dashboard_df.columns.values.tolist()] + dashboard_df.values.tolist())

print("Dashboard updated successfully 🚀")


# In[51]:


print(export_df.columns.tolist())


# In[71]:


import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# Step 1: Setup credentials
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
SERVICE_ACCOUNT_FILE = 'service_account_key.json'

credentials = Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
client = gspread.authorize(credentials)

# Step 2: Load Excel files
export_df = pd.read_excel(r"C:\Users\Supertails\Downloads\Export_2025-04-21_131843 (5).xlsx")
base_nf_df = pd.read_excel(r"C:\Users\Supertails\Downloads\base_nf.xlsx")

# Step 3: Merge MRP into base sales file using 'Variant SKU' instead of 'SKU'
merged_df = base_nf_df.merge(export_df[['Variant SKU', 'Variant Compare At Price']],
                             how='left', left_on='Product variant SKU', right_on='Variant SKU')

# Step 4: Create 'Returned' and 'Offtake' columns
merged_df['Returned'] = merged_df['Net items sold'].apply(lambda x: 'YES' if x < 0 else 'NO')
merged_df['Offtake'] = merged_df['Net items sold'] * merged_df['Variant Compare At Price']

# Step 5: Filter vendors
vendors = ['Barkbulter', 'Truelove']
vendor_df = merged_df[merged_df['Product vendor'].isin(vendors)]

# Step 6: Create weekly report (last 7 days)
vendor_df['Day'] = pd.to_datetime(vendor_df['Day'])  # adjust column name if needed
last_7_days = vendor_df[vendor_df['Day'] >= pd.Timestamp.now() - pd.Timedelta(days=7)]

# Step 7: Group data for Dashboard
dashboard_df = last_7_days.groupby(['Product title at time of sale', 'Product type']).agg({
    'Offtake': 'sum',
    'Net items sold': 'mean'
}).reset_index()

dashboard_df.rename(columns={
    'Net items sold': 'Daily Sale Units',
}, inplace=True)

# Step 8: Upload to Google Sheets
dashboard_sheet = client.open('Pupscribe weekly  sales dashboard').worksheet('Dashboard')  # change if needed
dashboard_sheet.clear()

dashboard_sheet.update([dashboard_df.columns.values.tolist()] + dashboard_df.values.tolist())

print("Dashboard updated successfully 🚀")


# In[69]:


# List all worksheets in the spreadsheet
spreadsheet = client.open('Pupscribe weekly  sales dashboard')
worksheets = spreadsheet.worksheets()
for worksheet in worksheets:
    print(worksheet.title)


# In[73]:


import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# Step 1: Setup credentials
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
SERVICE_ACCOUNT_FILE = 'service_account_key.json'

credentials = Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
client = gspread.authorize(credentials)

# Step 2: Load Excel files
export_df = pd.read_excel(r"C:\Users\Supertails\Downloads\Export_2025-04-21_131843 (5).xlsx")
base_nf_df = pd.read_excel(r"C:\Users\Supertails\Downloads\base_nf.xlsx")

# Step 3: Merge MRP into base sales file using 'Product variant SKU' and 'variant SKU'
merged_df = base_nf_df.merge(export_df[['variant SKU', 'Variant Compare At Price']],
                             how='left', left_on='Product variant SKU', right_on='variant SKU')

# Step 4: Create 'Returned' and 'Offtake' columns
merged_df['Returned'] = merged_df['Net items sold'].apply(lambda x: 'YES' if x < 0 else 'NO')
merged_df['Offtake'] = merged_df['Net items sold'] * merged_df['Variant Compare At Price']

# Step 5: Filter vendors
vendors = ['Barkbulter', 'Truelove']
vendor_df = merged_df[merged_df['Vendor'].isin(vendors)]

# Step 6: Filter data for this week (Monday to Sunday)
# Get the current date and calculate the start of the week (Monday)
today = pd.Timestamp.today()
start_of_week = today - pd.Timedelta(days=today.weekday())  # Start of the week (Monday)
end_of_week = start_of_week + pd.Timedelta(days=6)  # End of the week (Sunday)

# Filter data for this week based on 'Order Date'
vendor_df['Order Date'] = pd.to_datetime(vendor_df['Order Date'])  # Ensure it's a datetime column
this_week_sales = vendor_df[(vendor_df['Order Date'] >= start_of_week) & (vendor_df['Order Date'] <= end_of_week)]

# Step 7: Group data for Dashboard (for this week)
dashboard_df = this_week_sales.groupby(['Product title at time of sale', 'Product type']).agg({
    'Offtake': 'sum',
    'Net items sold': 'sum'  # Use sum for this week's sales
}).reset_index()

# Rename 'Net items sold' to 'Daily Sale Units' for clarity
dashboard_df.rename(columns={'Net items sold': 'Daily Sale Units'}, inplace=True)

# Step 8: Upload to Google Sheets
dashboard_sheet = client.open('Pupscribe weekly sales dashboard').worksheet('Dashboard')
dashboard_sheet.clear()  # Clear previous data

# Update the sheet with the new data
dashboard_sheet.update([dashboard_df.columns.values.tolist()] + dashboard_df.values.tolist())

print("This week's sales dashboard updated successfully 🚀")


# In[77]:


import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# Step 1: Setup credentials
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
SERVICE_ACCOUNT_FILE = 'service_account_key.json'

credentials = Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
client = gspread.authorize(credentials)

# Step 2: Load Excel files
export_df = pd.read_excel(r"C:\Users\Supertails\Downloads\Export_2025-04-21_131843 (5).xlsx")
base_nf_df = pd.read_excel(r"C:\Users\Supertails\Downloads\base_nf.xlsx")

# Step 3: Merge MRP into base sales file using 'Product variant SKU' and 'Variant SKU'
merged_df = base_nf_df.merge(export_df[['Variant SKU', 'Variant Compare At Price']],
                             how='left', left_on='Product variant SKU', right_on='Variant SKU')

# Step 4: Create 'Returned' and 'Offtake' columns
merged_df['Returned'] = merged_df['Net items sold'].apply(lambda x: 'YES' if x < 0 else 'NO')
merged_df['Offtake'] = merged_df['Net items sold'] * merged_df['Variant Compare At Price']

# Step 5: Filter vendors
vendors = ['Barkbulter', 'Truelove']
vendor_df = merged_df[merged_df['Product vendor'].isin(vendors)]

# Step 6: Filter data for this week (Monday to Sunday)
today = pd.Timestamp.today()
start_of_week = today - pd.Timedelta(days=today.weekday())  # Start of the week (Monday)
end_of_week = start_of_week + pd.Timedelta(days=6)  # End of the week (Sunday)

# Filter the data to include only this week's sales
vendor_df['Day'] = pd.to_datetime(vendor_df['Day'])  # Ensure 'Day' is a datetime column
this_week_sales = vendor_df[(vendor_df['Day'] >= start_of_week) & (vendor_df['Day'] <= end_of_week)]

# Step 7: Group data for Dashboard (for this week)
dashboard_df = this_week_sales.groupby(['Product title at time of sale', 'Product type']).agg({
    'Offtake': 'sum',  # Sum of total sales
    'Net items sold': 'sum'  # Total sales units for this week
}).reset_index()

# Rename 'Net items sold' to 'Daily Sale Units'
dashboard_df.rename(columns={'Net items sold': 'Daily Sale Units'}, inplace=True)

# Step 8: Upload to Google Sheets
dashboard_sheet = client.open('Pupscribe weekly  sales dashboard').worksheet('Dashboard')  # Adjust the name as needed
dashboard_sheet.clear()  # Clear any previous data

# Update the sheet with the new dashboard data
dashboard_sheet.update([dashboard_df.columns.values.tolist()] + dashboard_df.values.tolist())

print("This week's sales dashboard updated successfully 🚀")


# In[83]:


import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# Step 1: Setup credentials
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
SERVICE_ACCOUNT_FILE = 'service_account_key.json'

credentials = Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
client = gspread.authorize(credentials)

# Step 2: Load Excel files
export_df = pd.read_excel(r"C:\Users\Supertails\Downloads\Export_2025-04-21_131843 (5).xlsx")
base_nf_df = pd.read_excel(r"C:\Users\Supertails\Downloads\base_nf.xlsx")

# Step 3: Merge MRP into base sales file using 'Product variant SKU' and 'Variant SKU'
merged_df = base_nf_df.merge(export_df[['Variant SKU', 'Variant Compare At Price']],
                             how='left', left_on='Product variant SKU', right_on='Variant SKU')

# Step 4: Create 'Returned' and 'Offtake' columns
merged_df['Returned'] = merged_df['Net items sold'].apply(lambda x: 'YES' if x < 0 else 'NO')
merged_df['Offtake'] = merged_df['Net items sold'] * merged_df['Variant Compare At Price']

# Step 5: Filter vendors
vendors = ['Barkbutler', 'True Love']
vendor_df = merged_df[merged_df['Product vendor'].isin(vendors)]

# Step 6: Filter data between 14-04-2025 and 20-04-2025
vendor_df['Day'] = pd.to_datetime(vendor_df['Day'])  # Make sure 'Day' is datetime

start_date = pd.to_datetime('2025-04-14')
end_date = pd.to_datetime('2025-04-20')

# Filter the data for the given date range
date_filtered_sales = vendor_df[(vendor_df['Day'] >= start_date) & (vendor_df['Day'] <= end_date)]

# Step 7: Group data for Dashboard
dashboard_df = date_filtered_sales.groupby(['Product title at time of sale', 'Product type']).agg({
    'Offtake': 'sum',
    'Net items sold': 'sum'  # Sum units sold
}).reset_index()

dashboard_df.rename(columns={
    'Net items sold': 'Total Sale Units',
}, inplace=True)

# Step 8: Upload to Google Sheets
dashboard_sheet = client.open('Pupscribe weekly  sales dashboard').worksheet('Dashboard')  # Adjust the name if needed
dashboard_sheet.clear()

dashboard_sheet.update([dashboard_df.columns.values.tolist()] + dashboard_df.values.tolist())

print("Sales report for 14-Apr-2025 to 20-Apr-2025 updated successfully 🚀")


# In[85]:


import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# Step 1: Setup credentials
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
SERVICE_ACCOUNT_FILE = 'service_account_key.json'

credentials = Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
client = gspread.authorize(credentials)

# Step 2: Load Excel files
export_df = pd.read_excel(r"C:\Users\Supertails\Downloads\Export_2025-04-21_131843 (5).xlsx")
base_nf_df = pd.read_excel(r"C:\Users\Supertails\Downloads\base_nf.xlsx")

# Step 3: Merge MRP into base sales file using 'Product variant SKU' and 'Variant SKU'
merged_df = base_nf_df.merge(export_df[['Variant SKU', 'Variant Compare At Price']],
                             how='left', left_on='Product variant SKU', right_on='Variant SKU')

# Step 4: Create 'Returned' and 'Offtake' columns
merged_df['Returned'] = merged_df['Net items sold'].apply(lambda x: 'YES' if x < 0 else 'NO')
merged_df['Offtake'] = merged_df['Net items sold'] * merged_df['Variant Compare At Price']

# Step 5: Filter vendors
vendors = ['Barkbutler', 'True Love']
vendor_df = merged_df[merged_df['Product vendor'].isin(vendors)].copy()

# Step 6: Filter data between 14-04-2025 and 20-04-2025
vendor_df['Day'] = pd.to_datetime(vendor_df['Day'])  # Make sure 'Day' is datetime

start_date = pd.to_datetime('2025-04-14')
end_date = pd.to_datetime('2025-04-20')

# Filter the data for the given date range
date_filtered_sales = vendor_df[(vendor_df['Day'] >= start_date) & (vendor_df['Day'] <= end_date)]

# Step 7: Prepare Dashboard
dashboard_df = date_filtered_sales[['Day', 'Product title at time of sale', 'Product type', 'Offtake', 'Net items sold']].copy()

dashboard_df.rename(columns={
    'Day': 'Order Date',
    'Net items sold': 'Sale Units',
}, inplace=True)

# Step 8: Upload to Google Sheets
dashboard_sheet = client.open('Pupscribe weekly  sales dashboard').worksheet('Dashboard')  # Adjust the name if needed
dashboard_sheet.clear()

dashboard_sheet.update([dashboard_df.columns.values.tolist()] + dashboard_df.values.tolist())

print("Sales report for 14-Apr-2025 to 20-Apr-2025 updated successfully 🚀")


# In[87]:


# Step 7: Prepare Dashboard
dashboard_df = date_filtered_sales[['Day', 'Product title at time of sale', 'Product type', 'Offtake', 'Net items sold']].copy()

dashboard_df.rename(columns={
    'Day': 'Order Date',
    'Net items sold': 'Sale Units',
}, inplace=True)

# ⭐ Fix: Convert 'Order Date' to string
dashboard_df['Order Date'] = dashboard_df['Order Date'].dt.strftime('%Y-%m-%d')


# In[89]:


import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# Step 1: Setup credentials
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
SERVICE_ACCOUNT_FILE = 'service_account_key.json'

credentials = Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
client = gspread.authorize(credentials)

# Step 2: Load Excel files
export_df = pd.read_excel(r"C:\Users\Supertails\Downloads\Export_2025-04-21_131843 (5).xlsx")
base_nf_df = pd.read_excel(r"C:\Users\Supertails\Downloads\base_nf.xlsx")

# Step 3: Merge MRP into base sales file using 'Product variant SKU' and 'Variant SKU'
merged_df = base_nf_df.merge(export_df[['Variant SKU', 'Variant Compare At Price']],
                             how='left', left_on='Product variant SKU', right_on='Variant SKU')

# Step 4: Create 'Returned' and 'Offtake' columns
merged_df['Returned'] = merged_df['Net items sold'].apply(lambda x: 'YES' if x < 0 else 'NO')
merged_df['Offtake'] = merged_df['Net items sold'] * merged_df['Variant Compare At Price']

# Step 5: Filter vendors
vendors = ['Barkbutler', 'True Love']
vendor_df = merged_df[merged_df['Product vendor'].isin(vendors)].copy()

# Step 6: Filter data between 14-04-2025 and 20-04-2025
vendor_df['Day'] = pd.to_datetime(vendor_df['Day'])  # Make sure 'Day' is datetime

start_date = pd.to_datetime('2025-04-14')
end_date = pd.to_datetime('2025-04-20')

# Filter the data for the given date range
date_filtered_sales = vendor_df[(vendor_df['Day'] >= start_date) & (vendor_df['Day'] <= end_date)]

# Step 7: Prepare Dashboard
dashboard_df = date_filtered_sales[['Day', 'Product title at time of sale', 'Product type', 'Offtake', 'Net items sold']].copy()

dashboard_df.rename(columns={
    'Day': 'Order Date',
    'Net items sold': 'Sale Units',
}, inplace=True)

# Step 8: Upload to Google Sheets
dashboard_sheet = client.open('Pupscribe weekly  sales dashboard').worksheet('Dashboard')  # Adjust the name if needed
dashboard_sheet.clear()

dashboard_sheet.update([dashboard_df.columns.values.tolist()] + dashboard_df.values.tolist())

print("Sales report for 14-Apr-2025 to 20-Apr-2025 updated successfully 🚀")


# In[91]:


dashboard_df = dashboard_df.astype(str)


# In[93]:


dashboard_sheet.update([dashboard_df.columns.values.tolist()] + dashboard_df.values.tolist())


# In[97]:


dashboard_df = dashboard_df.sort_values(by='Offtake', ascending=True)


# In[ ]:


import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# Step 1: Setup credentials
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
SERVICE_ACCOUNT_FILE = 'service_account_key.json'

credentials = Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
client = gspread.authorize(credentials)

# Step 2: Load Excel files
export_df = pd.read_excel(r"C:\Users\Supertails\Downloads\Export_2025-04-21_131843 (5).xlsx")
base_nf_df = pd.read_excel(r"C:\Users\Supertails\Downloads\base_nf.xlsx")

# Step 3: Merge MRP into base sales file using 'Product variant SKU' and 'Variant SKU'
merged_df = base_nf_df.merge(export_df[['Variant SKU', 'Variant Compare At Price']],
                             how='left', left_on='Product variant SKU', right_on='Variant SKU')

# Step 4: Create 'Returned' and 'Offtake' columns
merged_df['Returned'] = merged_df['Net items sold'].apply(lambda x: 'YES' if x < 0 else 'NO')
merged_df['Offtake'] = merged_df['Net items sold'] * merged_df['Variant Compare At Price']

# Step 5: Filter vendors
vendors = ['Barkbutler', 'True Love']
vendor_df = merged_df[merged_df['Product vendor'].isin(vendors)].copy()

# Step 6: Filter data between 14-04-2025 and 20-04-2025
vendor_df['Day'] = pd.to_datetime(vendor_df['Day'])  # Make sure 'Day' is datetime

start_date = pd.to_datetime('2025-04-14')
end_date = pd.to_datetime('2025-04-20')

# Filter the data for the given date range
date_filtered_sales = vendor_df[(vendor_df['Day'] >= start_date) & (vendor_df['Day'] <= end_date)]

# Step 7: Prepare Dashboard
dashboard_df = date_filtered_sales[['Day', 'Product title at time of sale', 'Product type', 'Offtake', 'Net items sold']].copy()

dashboard_df.rename(columns={
    'Day': 'Order Date',
    'Net items sold': 'Sale Units',
}, inplace=True)

# Step 8: Upload to Google Sheets
dashboard_sheet = client.open('Pupscribe weekly  sales dashboard').worksheet('Dashboard')  # Adjust the name if needed
dashboard_sheet.clear()

dashboard_sheet.update([dashboard_df.columns.values.tolist()] + dashboard_df.values.tolist())

print("Sales report for 14-Apr-2025 to 20-Apr-2025 updated successfully 🚀")


# In[ ]:


import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# Step 1: Setup credentials
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
SERVICE_ACCOUNT_FILE = 'service_account_key.json'

credentials = Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
client = gspread.authorize(credentials)

# Step 2: Load Excel files
export_df = pd.read_excel(r"C:\Users\Supertails\Downloads\Export_2025-04-21_131843 (5).xlsx")
base_nf_df = pd.read_excel(r"C:\Users\Supertails\Downloads\base_nf.xlsx")

# Step 3: Merge MRP into base sales file using 'Product variant SKU' and 'Variant SKU'
merged_df = base_nf_df.merge(export_df[['Variant SKU', 'Variant Compare At Price']],
                             how='left', left_on='Product variant SKU', right_on='Variant SKU')

# Step 4: Create 'Returned' and 'Offtake' columns
merged_df['Returned'] = merged_df['Net items sold'].apply(lambda x: 'YES' if x < 0 else 'NO')
merged_df['Offtake'] = merged_df['Net items sold'] * merged_df['Variant Compare At Price']

# Step 5: Filter vendors
vendors = ['Barkbutler', 'True Love']
vendor_df = merged_df[merged_df['Product vendor'].isin(vendors)].copy()

# Step 6: Filter data between 14-04-2025 and 20-04-2025
vendor_df['Day'] = pd.to_datetime(vendor_df['Day'])  # Make sure 'Day' is datetime

start_date = pd.to_datetime('2025-04-14')
end_date = pd.to_datetime('2025-04-20')

# Filter the data for the given date range
date_filtered_sales = vendor_df[(vendor_df['Day'] >= start_date) & (vendor_df['Day'] <= end_date)]

# Step 7: Prepare Dashboard
dashboard_df = date_filtered_sales[['Day', 'Product title at time of sale', 'Product type', 'Offtake', 'Net items sold']].copy()

dashboard_df.rename(columns={
    'Day': 'Order Date',
    'Net items sold': 'Sale Units',
}, inplace=True)

# Step 8: Upload to Google Sheets
dashboard_sheet = client.open('Pupscribe weekly  sales dashboard').worksheet('Dashboard')  # Adjust the name if needed
dashboard_sheet.clear()

dashboard_sheet.update([dashboard_df.columns.values.tolist()] + dashboard_df.values.tolist())

print("Sales report for 14-Apr-2025 to 20-Apr-2025 updated successfully 🚀")


# In[109]:


import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# Step 1: Setup credentials
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
SERVICE_ACCOUNT_FILE = 'service_account_key.json'

credentials = Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
client = gspread.authorize(credentials)

# Step 2: Load Excel files
export_df = pd.read_excel(r"C:\Users\Supertails\Downloads\Export_2025-04-21_131843 (5).xlsx")
base_nf_df = pd.read_excel(r"C:\Users\Supertails\Downloads\base_nf.xlsx")

# Step 3: Merge MRP into base sales file using 'Product variant SKU' and 'Variant SKU'
merged_df = base_nf_df.merge(export_df[['Variant SKU', 'Variant Compare At Price']],
                             how='left', left_on='Product variant SKU', right_on='Variant SKU')

# Step 4: Create 'Returned' and 'Offtake' columns
merged_df['Returned'] = merged_df['Net items sold'].apply(lambda x: 'YES' if x < 0 else 'NO')
merged_df['Offtake'] = merged_df['Net items sold'] * merged_df['Variant Compare At Price']

# Step 5: Filter vendors
vendors = ['Barkbutler', 'True Love']
vendor_df = merged_df[merged_df['Product vendor'].isin(vendors)].copy()

# Step 6: Filter data between 14-04-2025 and 20-04-2025
vendor_df['Day'] = pd.to_datetime(vendor_df['Day'])  # Make sure 'Day' is datetime

start_date = pd.to_datetime('2025-04-14')
end_date = pd.to_datetime('2025-04-20')

# Filter the data for the given date range
date_filtered_sales = vendor_df[(vendor_df['Day'] >= start_date) & (vendor_df['Day'] <= end_date)]

# Step 7: Prepare Dashboard
dashboard_df = date_filtered_sales[['Day', 'Product title at time of sale', 'Product type', 'Offtake', 'Net items sold']].copy()

dashboard_df.rename(columns={
    'Day': 'Order Date',
    'Net items sold': 'Sale Units',
}, inplace=True)

# Step 8: Sort by Offtake in increasing order
dashboard_df = dashboard_df.sort_values(by='Offtake', ascending=False)

# Step 9: Upload to Google Sheets
dashboard_sheet = client.open('Pupscribe weekly  sales dashboard').worksheet('Dashboard')  # Adjust the name if needed
dashboard_sheet.clear()

dashboard_sheet.update([dashboard_df.columns.values.tolist()] + dashboard_df.values.tolist())
dashboard_df = dashboard_df.astype(str)


print("Sales report for 14-Apr-2025 to 20-Apr-2025 updated successfully 🚀")


# In[113]:


import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# Step 1: Setup credentials
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
SERVICE_ACCOUNT_FILE = 'service_account_key.json'

credentials = Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
client = gspread.authorize(credentials)

# Step 2: Load Excel files
export_df = pd.read_excel(r"C:\Users\Supertails\Downloads\Export_2025-04-21_131843 (5).xlsx")
base_nf_df = pd.read_excel(r"C:\Users\Supertails\Downloads\base_nf.xlsx")

# Step 3: Merge MRP into base sales file using 'Product variant SKU' and 'Variant SKU'
merged_df = base_nf_df.merge(export_df[['Variant SKU', 'Variant Compare At Price']],
                             how='left', left_on='Product variant SKU', right_on='Variant SKU')

# Step 4: Create 'Returned' and 'Offtake' columns
merged_df['Returned'] = merged_df['Net items sold'].apply(lambda x: 'YES' if x < 0 else 'NO')
merged_df['Offtake'] = merged_df['Net items sold'] * merged_df['Variant Compare At Price']

# Step 5: Filter vendors
vendors = ['Barkbutler', 'True Love']
vendor_df = merged_df[merged_df['Product vendor'].isin(vendors)].copy()

# Step 6: Filter data between 14-04-2025 and 20-04-2025
vendor_df['Day'] = pd.to_datetime(vendor_df['Day'])  # Ensure 'Day' is datetime

start_date = pd.to_datetime('2025-04-14')
end_date = pd.to_datetime('2025-04-20')

# Filter the data for the given date range
date_filtered_sales = vendor_df[(vendor_df['Day'] >= start_date) & (vendor_df['Day'] <= end_date)]

# Step 7: Prepare Dashboard
dashboard_df = date_filtered_sales[['Day', 'Product title at time of sale', 'Product type', 'Offtake', 'Net items sold']].copy()

dashboard_df.rename(columns={
    'Day': 'Order Date',
    'Net items sold': 'Sale Units',
}, inplace=True)

# Step 8: Sort by Offtake in descending order
dashboard_df = dashboard_df.sort_values(by='Offtake', ascending=False)

# Step 9: Convert any datetime columns to string format
# Convert 'Order Date' column to string format to avoid JSON serialization issues
dashboard_df['Order Date'] = dashboard_df['Order Date'].dt.strftime('%Y-%m-%d')

# Step 10: Upload to Google Sheets
dashboard_sheet = client.open('Pupscribe weekly  sales dashboard').worksheet('Dashboard')  # Adjust the name if needed
dashboard_sheet.clear()

# Update the sheet with the sorted data
dashboard_sheet.update([dashboard_df.columns.values.tolist()] + dashboard_df.values.tolist())

print("Sales report for 14-Apr-2025 to 20-Apr-2025 updated successfully 🚀")


# In[121]:


import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta

# Step 1: Setup credentials
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
SERVICE_ACCOUNT_FILE = 'service_account_key.json'

credentials = Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
client = gspread.authorize(credentials)

# Step 2: Load Excel files
export_df = pd.read_excel(r"C:\Users\Supertails\Downloads\Export_2025-04-21_131843 (5).xlsx")
base_nf_df = pd.read_excel(r"C:\Users\Supertails\Downloads\base_nf.xlsx")

# Step 3: Merge MRP into base sales file using 'Product variant SKU' and 'Variant SKU'
merged_df = base_nf_df.merge(export_df[['Variant SKU', 'Variant Compare At Price']],
                             how='left', left_on='Product variant SKU', right_on='Variant SKU')

# Step 4: Create 'Returned' and 'Offtake' columns
merged_df['Returned'] = merged_df['Net items sold'].apply(lambda x: 'YES' if x < 0 else 'NO')
merged_df['Offtake'] = merged_df['Net items sold'] * merged_df['Variant Compare At Price']

# Step 5: Filter vendors
vendors = ['Barkbutler', 'True Love']
vendor_df = merged_df[merged_df['Product vendor'].isin(vendors)].copy()

# Step 6: Calculate previous week's Monday to Sunday
today = datetime.today()
start_date = today - timedelta(days=today.weekday() + 7)   # Previous Monday
end_date = start_date + timedelta(days=6)                  # Previous Sunday

print(f"Generating report for {start_date.strftime('%d-%b-%Y')} to {end_date.strftime('%d-%b-%Y')}")

# Step 7: Filter data for previous week
vendor_df['Day'] = pd.to_datetime(vendor_df['Day'])  # Ensure 'Day' is datetime

date_filtered_sales = vendor_df[(vendor_df['Day'] >= start_date) & (vendor_df['Day'] <= end_date)]

# Step 8: Prepare Dashboard
dashboard_df = date_filtered_sales[['Day', 'Product title at time of sale', 'Product type', 'Offtake', 'Net items sold']].copy()

dashboard_df.rename(columns={
    'Day': 'Order Date',
    'Net items sold': 'Sale Units',
}, inplace=True)

# Step 9: Sort by Offtake in descending order
dashboard_df = dashboard_df.sort_values(by='Offtake', ascending=False)

# Step 10: Convert any datetime columns to string format
dashboard_df['Order Date'] = dashboard_df['Order Date'].dt.strftime('%Y-%m-%d')

# Step 11: Upload to Google Sheets
dashboard_sheet = client.open('Pupscribe weekly  sales dashboard').worksheet('Dashboard')  # Adjust the name if needed
dashboard_sheet.clear()

# Update the sheet with the sorted data
dashboard_sheet.update([dashboard_df.columns.values.tolist()] + dashboard_df.values.tolist())

print(f"Sales report for {start_date.strftime('%d-%b-%Y')} to {end_date.strftime('%d-%b-%Y')} updated successfully 🚀")


# In[123]:


import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta

# Step 1: Setup credentials
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
SERVICE_ACCOUNT_FILE = 'service_account_key.json'

credentials = Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
client = gspread.authorize(credentials)

# Step 2: Load Excel files
export_df = pd.read_excel(r"C:\Users\Supertails\Downloads\Export_2025-04-21_131843 (5).xlsx")
base_nf_df = pd.read_excel(r"C:\Users\Supertails\Downloads\base_nf.xlsx")

# Step 3: Merge MRP into base sa
les file using 'Product variant SKU' and 'Variant SKU'
merged_df = base_nf_df.merge(export_df[['Variant SKU', 'Variant Compare At Price']],
                             how='left', left_on='Product variant SKU', right_on='Variant SKU')

# Step 4: Create 'Returned' and 'Offtake' columns
merged_df['Returned'] = merged_df['Net items sold'].apply(lambda x: 'YES' if x < 0 else 'NO')
merged_df['Offtake'] = merged_df['Net items sold'] * merged_df['Variant Compare At Price']

# Step 5: Filter vendors
vendors = ['Barkbutler', 'True Love']
vendor_df = merged_df[merged_df['Product vendor'].isin(vendors)].copy()

# Step 6: Calculate previous week's Monday to Sunday properly
today = datetime.today()

# Get the last Monday
last_monday = today - timedelta(days=today.weekday() + 7)
last_sunday = last_monday + timedelta(days=6)

start_date = last_monday.replace(hour=0, minute=0, second=0, microsecond=0)
end_date = last_sunday.replace(hour=23, minute=59, second=59, microsecond=999999)

print(f"Generating report for {start_date.strftime('%d-%b-%Y')} to {end_date.strftime('%d-%b-%Y')}")

# Step 7: Filter data for previous week
vendor_df['Day'] = pd.to_datetime(vendor_df['Day'])  # Ensure 'Day' is datetime

date_filtered_sales = vendor_df[(vendor_df['Day'] >= start_date) & (vendor_df['Day'] <= end_date)]

# Step 8: Prepare Dashboard
dashboard_df = date_filtered_sales[['Day', 'Product title at time of sale', 'Product type', 'Offtake', 'Net items sold']].copy()

dashboard_df.rename(columns={
    'Day': 'Order Date',
    'Net items sold': 'Sale Units',
}, inplace=True)

# Step 9: Sort by Offtake in descending order
dashboard_df = dashboard_df.sort_values(by='Offtake', ascending=False)

# Step 10: Convert any datetime columns to string format
dashboard_df['Order Date'] = dashboard_df['Order Date'].dt.strftime('%Y-%m-%d')

# Step 11: Upload to Google Sheets
dashboard_sheet = client.open('Pupscribe weekly  sales dashboard').worksheet('Dashboard')  # Adjust the name if needed
dashboard_sheet.clear()

# Update the sheet with the sorted data
dashboard_sheet.update([dashboard_df.columns.values.tolist()] + dashboard_df.values.tolist())

print(f"Sales report for {start_date.strftime('%d-%b-%Y')} to {end_date.strftime('%d-%b-%Y')} updated successfully 🚀")


# In[125]:


import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# Step 1: Setup credentials
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
SERVICE_ACCOUNT_FILE = 'service_account_key.json'

credentials = Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
client = gspread.authorize(credentials)

# Step 2: Load Excel files
export_df = pd.read_excel(r"C:\Users\Supertails\Downloads\Export_2025-04-21_131843 (5).xlsx")
base_nf_df = pd.read_excel(r"C:\Users\Supertails\Downloads\base_nf.xlsx")

# Step 3: Merge MRP into base sales file using 'Product variant SKU' and 'Variant SKU'
merged_df = base_nf_df.merge(export_df[['Variant SKU', 'Variant Compare At Price']],
                             how='left', left_on='Product variant SKU', right_on='Variant SKU')

# Step 4: Create 'Returned' and 'Offtake' columns
merged_df['Returned'] = merged_df['Net items sold'].apply(lambda x: 'YES' if x < 0 else 'NO')
merged_df['Offtake'] = merged_df['Net items sold'] * merged_df['Variant Compare At Price']

# Step 5: Filter vendors
vendors = ['Barkbutler', 'True Love']
vendor_df = merged_df[merged_df['Product vendor'].isin(vendors)].copy()

# Step 6: Filter data between 14-04-2025 and 20-04-2025
vendor_df['Day'] = pd.to_datetime(vendor_df['Day'])  # Ensure 'Day' is datetime

start_date = pd.to_datetime('2025-04-14')
end_date = pd.to_datetime('2025-04-20')

# Filter the data for the given date range
date_filtered_sales = vendor_df[(vendor_df['Day'] >= start_date) & (vendor_df['Day'] <= end_date)]

# Step 7: Prepare Dashboard
dashboard_df = date_filtered_sales[['Day', 'Product title at time of sale', 'Product type', 'Offtake', 'Net items sold']].copy()

dashboard_df.rename(columns={
    'Day': 'Order Date',
    'Net items sold': 'Sale Units',
}, inplace=True)

# Step 8: Sort by Offtake in descending order
dashboard_df = dashboard_df.sort_values(by='Offtake', ascending=False)

# Step 9: Convert any datetime columns to string format
# Convert 'Order Date' column to string format to avoid JSON serialization issues
dashboard_df['Order Date'] = dashboard_df['Order Date'].dt.strftime('%Y-%m-%d')

# Step 10: Upload to Google Sheets
dashboard_sheet = client.open('Pupscribe weekly  sales dashboard').worksheet('Dashboard')  # Adjust the name if needed
dashboard_sheet.clear()

# Update the sheet with the sorted data
dashboard_sheet.update([dashboard_df.columns.values.tolist()] + dashboard_df.values.tolist())

print("Sales report for 14-Apr-2025 to 20-Apr-2025 updated successfully 🚀")


# In[131]:


import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# Step 1: Setup credentials
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
SERVICE_ACCOUNT_FILE = 'service_account_key.json'

credentials = Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
client = gspread.authorize(credentials)

# Step 2: Load Excel files
export_df = pd.read_excel(r"C:\Users\Supertails\Downloads\Export_2025-04-21_131843 (5).xlsx")
base_nf_df = pd.read_excel(r"C:\Users\Supertails\Downloads\base_nf.xlsx")

# Step 3: Merge MRP into base sales file using 'Product variant SKU' and 'Variant SKU'
merged_df = base_nf_df.merge(export_df[['Variant SKU', 'Variant Compare At Price']],
                             how='left', left_on='Product variant SKU', right_on='Variant SKU')

# Step 4: Create 'Returned' and 'Offtake' columns
merged_df['Returned'] = merged_df['Net items sold'].apply(lambda x: 'YES' if x < 0 else 'NO')
merged_df['Offtake'] = merged_df['Net items sold'] * merged_df['Variant Compare At Price']

# Step 5: Filter vendors
vendors = ['Barkbutler', 'True Love']
vendor_df = merged_df[merged_df['Product vendor'].isin(vendors)].copy()

# Step 6: Filter data between 14-04-2025 and 20-04-2025
vendor_df['Day'] = pd.to_datetime(vendor_df['Day'])  # Ensure 'Day' is datetime

start_date = pd.to_datetime('2025-04-14')
end_date = pd.to_datetime('2025-04-20')

# Filter the data for the given date range
date_filtered_sales = vendor_df[(vendor_df['Day'] >= start_date) & (vendor_df['Day'] <= end_date)]

# Step 7: Prepare Dashboard
dashboard_df = date_filtered_sales[['Day', 'Product title at time of sale', 'Product type', 'Product variant SKU', 'Offtake', 'Net items sold']].copy()

dashboard_df.rename(columns={
    'Day': 'Order Date',
    'Product variant SKU': 'Variant SKU',  # <-- Rename for clean header
    'Net items sold': 'Sale Units',
}, inplace=True)

# Step 8: Sort by Offtake in descending order
dashboard_df = dashboard_df.sort_values(by='Offtake', ascending=False)

# Step 9: Convert datetime columns to string
dashboard_df['Order Date'] = dashboard_df['Order Date'].dt.strftime('%Y-%m-%d')

# Step 10: Upload to Google Sheets
dashboard_sheet = client.open('Pupscribe weekly  sales dashboard').worksheet('Dashboard')  # Adjust the name if needed
dashboard_sheet.clear()

dashboard_sheet.update([dashboard_df.columns.values.tolist()] + dashboard_df.values.tolist())

print("Sales report updated successfully 🚀")


# In[ ]:




