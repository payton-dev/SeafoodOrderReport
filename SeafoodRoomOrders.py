#!/usr/bin/env python
# coding: utf-8

# In[1]:


import xlsxwriter
import pandas as pd

#import Seafood Codes - SEAFOODCODELIST.csv
#INV>Inv Report>Quick List
#User Define Class 16 only
#Export as CSV to "SEAFOODCODELIST.csv"


#imports codes and preps DF
codelists = pd.read_csv('SEAFOODCODELIST.csv',dtype={'order_no': object})
keep_col = ['product_id','category']
codelist1 = codelists[keep_col]
codelist2 = codelist1.loc[(codelist1.category == 16)]

#codelist2.to_csv('Partcodelist.csv', index=False)
#Add in Frozen Codes to DF
new_row = {'product_id':201178, 'category':15}
new_row1 = {'product_id':201330, 'category':15}

codelist2 = codelist2.append(new_row, ignore_index=True)
codelist3 = codelist2.append(new_row1, ignore_index=True)


# In[2]:


#import Order Demand DEMAND.CSV
#Production Scale > Demand Butto,
#Select Today Date and Tomorrow Date
#Print All > Export as CSV to "DEMAND.CSV"

#Converts columns to proper dtypes and rename
demand = pd.read_csv('DEMAND.csv',dtype={'product_id': object, 'refno': object,'ord_qty': str})
demand['order_no'] = demand['refno']

#Keep columns requred
keep_col = ['order_no','product_id','descript','cust_id','comment','ord_qty','shp_unit']
demands1 = demand[keep_col]


#Merge Code List into Order Demand to get ONLY Orders wanted for seafood room
orderlist = pd.merge(codelist3, demands1, how='left', on="product_id")
orderlist = orderlist.dropna()

#Creates a new DF for Comments to look for rows with a comment
comments = orderlist.loc[orderlist['comment'] == 'T']

#Keep columns required for merging later
keep_col = ['product_id','order_no','descript']
comments1 = comments[keep_col]

#rename and del
comments1['note'] = comments1['descript']
keep_col = ['product_id','order_no','descript']
comment = comments1

#Locate all lines without comment so we can get product desc name as comments get written in said column aswell
#Rows with comments get printed with all other info same thats why two different DF are required one T one F for comments
order_df = orderlist.loc[orderlist['comment'] == 'F']


#Merges commments with main order DF
merged_left = pd.merge(left=order_df, right=comment, how='left', left_on=['order_no','product_id'], right_on=['order_no','product_id'])


# In[3]:


#Imports customer info
location = pd.read_csv('CUSTOMERINFO.csv')

#Merge Customer info With Orders
#Customer name, city, and location type

final_data = pd.merge(left=merged_left, right=location, how='left', left_on=['cust_id'], right_on=['cust_id'])


# REMOVE CODES NOT NEEDED FOR ROOM

# In[4]:


#Removes codes not required
#Other Fresh Seafood Codes

to_sort = final_data[final_data.product_id != '200590']
to_sort1 = to_sort[to_sort.product_id != '200356']
to_sort2 = to_sort1[to_sort1.product_id != '200657']
to_sort3 = to_sort2[to_sort2.product_id != '200656']
to_sort4 = to_sort3[to_sort3.product_id != '202300']
to_sort5 = to_sort4[to_sort4.product_id != '202307']
to_sort6 = to_sort5[to_sort5.product_id != '70578']
to_sort7 = to_sort6[to_sort6.product_id != '70542']
to_sort8 = to_sort7[to_sort7.product_id != '190201']
to_sort9 = to_sort8[to_sort8.product_id != '200047']
to_sort10 = to_sort9[to_sort9.product_id != '200500']
to_sort11 = to_sort10[to_sort10.product_id != '200566']
to_sort12 = to_sort11[to_sort11.product_id != '202301']
to_sort13 = to_sort12[to_sort12.product_id != '202306']
to_sort14 = to_sort13[to_sort13.product_id != '201534']

#to_sort.to_csv('tobesorted.csv', index=False)


# In[5]:


#import xlsxwriter
#import pandas as pd

#Clean up DF of extra columns no longer required
#df = pd.read_csv('tobesorted.csv')

keep_col = ['type','product_id','descript_x','ord_qty','shp_unit','note','company1','city','order_no']
sorted1 = to_sort[keep_col]


#Create a 'typeno' column for each location 'type'
sorted1['typeno'] = sorted1['type']
#Define name and type numbers lowest number = highest priority
typeno1 = {'VAN': 1,'FREIGHT': 2,'UPISLAND': 3,'TOWN': 4,'LOOKUP':5}
#Create number for each type
sorted1.typeno = [typeno1[item] for item in sorted1.typeno] 

#Sort orders
#First sort priority: Product ID
#Then second sort priority: Type
sorted1.groupby(['product_id','type'], as_index=False)

#Sort by ascending numbers for both - Get highest priority type first by each product_id
sorted_id = sorted1.sort_values(['typeno','product_id'], ascending=[True,True])
#Delete Type number as it's only wanted for sorting
del sorted_id['typeno']
#Commented out can make into csv if wanted - Excel will sort into seperate sheets
#sorted_id.to_csv('SeafoodSortedID.csv', index=False)



#Creates a list for saving and naming worksheets inside of workbook by product_id
sorted_pid = sorted_id['product_id'].unique().tolist()

#Writes DF into a excel file and
writer = pd.ExcelWriter("SeafoodOrderReport.xlsx", engine='xlsxwriter')

workbook  = writer.book
workbook.filename = 'SeafoodOrderReport.xlsm'
workbook.add_vba_project('./vbaProject.bin')

workbook1 = workbook.add_worksheet('Master')

# Add a button tied to a macro in the VBA project.
workbook1.insert_button('B3', {'AutofitColumns':   'AutofitColumns',
                               'caption': 'Press Me',
                               'width':   80,
                               'height':  30})


#Writes each product_id to seperate sheet and names the sheet said product_id
for p_id in sorted_pid:
    mydf = sorted_id.loc[sorted_id.product_id==p_id]
    mydf.to_excel(writer, sheet_name=p_id, index=False)

    


writer.save()


# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:




