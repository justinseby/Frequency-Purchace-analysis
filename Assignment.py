#!/usr/bin/env python
# coding: utf-8

# In[5]:


from pandas import read_excel
my_sheet = 'Data'
file_name = 'FileInput.xlsx'
df = read_excel(file_name, sheet_name = my_sheet)

# In[18]:


print(df[['Outlet ID','Brand Name','Sales Value']])

# In[22]:


dups = df.pivot_table(index = ['Outlet ID'], aggfunc ='size').reset_index().    rename(columns={0:'records'})
print(dups)


# In[68]:


Sum=0
count=[]
for i in range(9):
    count.append(0)
sale=[]
for i in range(9):
    sale.append(0)
for i in range(len(dups)) :
    if (dups.loc[i, "records"]) == 1:
        count[0]+=1
        sale[0]+=df.loc[i, "Sales Value"]
    elif (dups.loc[i, "records"]) == 2:
        count[1]+=1
        sale[1]+=df.loc[i, "Sales Value"]
    elif (dups.loc[i, "records"]) == 3:
        count[2]+=1
        sale[2]+=df.loc[i, "Sales Value"]
    elif (dups.loc[i, "records"]) == 4:
        count[3]+=1
        sale[3]+=df.loc[i, "Sales Value"]
    elif (dups.loc[i, "records"]) == 5:
        count[4]+=1
        sale[4]+=df.loc[i, "Sales Value"]
    elif (dups.loc[i, "records"]) == 6:
        count[5]+=1
        sale[5]+=df.loc[i, "Sales Value"]
    elif (dups.loc[i, "records"]) == 7:
        count[6]+=1
        sale[6]+=df.loc[i, "Sales Value"]
    elif (dups.loc[i, "records"]) == 8:
        count[7]+=1
        sale[7]+=df.loc[i, "Sales Value"]
    elif (dups.loc[i, "records"]) == 9:
        count[8]+=1
        sale[8]+=df.loc[i, "Sales Value"]
    else:
        pass
print("Countlist", count)
print("Salelist", sale)
    


# In[69]:


question=['Outlets purchased once','Outlets purchased 2times','Outlets purchased 3times','Outlets purchased 4times','Outlets purchased 5times','Outlets purchased 6times','Outlets purchased 7times','Outlets purchased 8times','Outlets purchased 9times']


# In[74]:


import csv
df1 = pd.DataFrame()
df1["Details"]=question
df1["Number of outlets"]=count
df1["total Sales Values"]=sale
# Converting to excel
df1.to_excel('result.xlsx', index = False)

#with open('Answer.csv', 'w') as f:
#    writer = csv.writer(f)
#    writer.writerows(zip(question, count, sale))


# In[ ]:




