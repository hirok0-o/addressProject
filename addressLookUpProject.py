#!/usr/bin/env python
# coding: utf-8

# In[25]:

#i didnt write this pip install code, got from: https://stackoverflow.com/questions/44210656/how-to-check-if-a-module-is-installed-in-python-and-if-not-install-it-within-t
import sys
import subprocess
import os


if not 'panda' in sys.modules:
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pandas']) 
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'xlrd']) 

import pandas as pd
import re
#function converts useful file contents to 2D array#
def readAddressFile(scoreVar):
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

print("ALL FILES MUST BE CLOSED WHEN PROGRAM IS RUNNING")
print("")

print("Enter path of folder which stores both files, example: C :\ user \ documents \ addressLookUp")
root=str(input("> "))
root=re.sub(r"\s+", "", root, flags=re.UNICODE)
for root,dirs,files in os.walk(root):
    for file in files:
        if ".csv" in file:
            scoreVar=os.path.join(root,file)
        if ".xlsx" in file:
            addressVar=os.path.join(root,file)


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

print("Task Finished")
print("You can now open your student address file")

