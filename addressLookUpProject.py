#!/usr/bin/env python
# coding: utf-8

# In[25]:


import pandas as pd
import re
#function converts useful file contents to 2D array#
def readAddressFile(scroreVar):
    df=pd.read_csv(scoreVar,usecols=[0,2])
    array=[]
    array=df.values.tolist()
    return array
#This gets the schools address file and puts the addresses into an array
def addressToArray(addressVar):
    df=pd.read_excel(addressVar,"Sheet1")
    array=[]
    array=list(df["Postcode"])
    for x in range (len(array)):
        var=array[x]
        var = re.sub(r"\s+", "", var, flags=re.UNICODE)
        array[x]=var
    return array


print("Enter address of file which contains students addresses")
print("It must be an excel file ")
print("EXAMPLE: C :\ user\ desktop\ file.xlsx")
addressVar=str(input(""))

print("Enter address of file which contains addresses and quintile scores (the file downloaded)")
print("It must be a CSV file ")
print("EXAMPLE: C :\ user\ desktop\ file.CSV")
scoreVar=str(input(""))
addressArray=readAddressFile(scoreVar)
schoolArray=addressToArray(addressVar)






def searchAddress(add,schoArray):
    array=[]
    for x in range(len(schoArray)):
        for y in range(len(add)):
            if schoArray[x]==add[y][0]:
                array.append(add[y][1])
    for x in range (len(array)):
        array[x]=int(array[x])
    return array
results=searchAddress(addressArray,schoolArray)


a=0
def write(cell):
    global results
    global a
    tempVar=results[a]
    a=a+1
    return tempVar
df=pd.read_excel(addressVar,"Sheet1", converters={"https://www.officeforstudents.org.uk/data-and-analysis/young-participation-by-area/search-by-postcode/": write})
df
    


# In[34]:


import pandas as pd
def highLigh_number(number):
    criteria=(number==2)|(number==1)
    return ["background-color: green" if i else "" for i in criteria]
df=df.style.apply(highLigh_number)


# In[35]:


df


# In[36]:


df.to_excel(addressVar)


# In[ ]:




