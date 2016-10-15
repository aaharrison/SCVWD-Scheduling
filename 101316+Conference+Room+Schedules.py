
# coding: utf-8

#Update Test

# In[1]:

import pandas as pd
import numpy as np
from pandas import DataFrame, Series
import matplotlib.pyplot as plt
import os
import datetime
import time


# # Crest Conference Rooms

# In[3]:

os.chdir("Conf Rm - C225 (Crest)")


# In[4]:

os.listdir()


# In[20]:

C209=pd.read_excel("Conf Rm - C209 (Crest) RESTRICTED.xls")
C212=pd.read_excel("Conf Rm - C212 RESTRICTED.xls")
C225=pd.read_excel("Conf Rm - C225 (Crest).xls")


# In[21]:

C209["Room"]="C209"
C212["Room"]="C212"
C225["Room"]="C225"


# In[39]:

print(673+467+474)


# In[40]:

Crest=pd.concat([C209,C212,C225])


# In[41]:

Crest.info()


# # Maintainance Confrence Rooms

# In[43]:

os.listdir()


# In[45]:

D102=pd.read_excel("Conf Rm - D102 (MaintReady Room).xls")
D131=pd.read_excel("Conf Rm - D131 (Maint).xls")


# In[46]:

D102["Room"]="D102"
D131["Room"]="D131"


# In[51]:

print(1694+953)


# In[55]:

Maintainance.info()


# In[54]:

Maintainance=pd.concat([D102,D131])


# # Headquarters Confrence Rooms

# In[57]:

os.listdir()


# In[59]:

A124 = pd.read_excel("Conf Rm - A124 (HQ).xls")
A136 = pd.read_excel("Conf Rm - A136 (HQ).xls")
A143 = pd.read_excel("Conf Rm - A143 (HQ).xls")
A168 = pd.read_excel("Conf Rm - A168 (HQ).xls")
A226 = pd.read_excel("Conf Rm - A226 (HQ).xls")
A235 = pd.read_excel("Conf Rm - A235 (HQ).xls")
A236 = pd.read_excel("Conf Rm - A236 (HQ).xls")
A318 = pd.read_excel("Conf Rm - A318 (HQ).xls")
A345 = pd.read_excel("Conf Rm - A345 (HQ).xls")
A346 = pd.read_excel("Conf Rm - A346 (HQ).xls")
Boardroom = pd.read_excel("Conf Rm - Boardroom (HQ).xls")
District = pd.read_excel("Conf Rm - District Counsel (HQ).xls")


# In[60]:

A124["Room"]="A124 (HQ)"
A136["Room"]="A136 (HQ)"
A143["Room"]="A143 (HQ)"
A168["Room"]="A168 (HQ)"
A226["Room"]="A226 (HQ)"
A235["Room"]="A235 (HQ)"
A236["Room"]="A236 (HQ)"
A318["Room"]="A318 (HQ)"
A345["Room"]="A345 (HQ)"
A346["Room"]="A346 (HQ)"
Boardroom["Room"]="Boardroom (HQ)"
District["Room"]="District Counsel (HQ)"


# In[66]:

HQcount=0
for x in [A124,A136,A143,A168,A226,A235,A236,A318,A345,A346,Boardroom,District]:
    HQcount+=len(x)

print(HQcount)


# In[67]:

HQ.info()


# In[62]:

HQ=pd.concat([A124,A136,A143,A168,A226,A235,A236,A318,A345,A346,Boardroom,District])


# # Admin Conference Rooms

# In[69]:

os.chdir("/Users/adeniyiharrison/Desktop/Outlook Scheduling/Conf Rom - B228 (Admin) RESTRICTED")
os.listdir()


# In[71]:

B108 = pd.read_excel("Conf Rm - B108 (Admin) - Old Board Room.xls")
B124 = pd.read_excel("Conf Rm - B124 (Admin).xls")
B207 = pd.read_excel("Conf Rm - B207 (Admin).xls")
B232 = pd.read_excel("Conf Rm - B232 (Admin).xls")
B228 = pd.read_excel("Conf Rom - B228 (Admin) RESTRICTED.xls")


# In[72]:

B108["Room"] = "B108 (Admin)"
B124["Room"] = "B124 (Admin)"
B207["Room"] = "B207 (Admin)"
B232["Room"] = "B232 (Admin)"
B228["Room"] = "B228 (Admin)"


# In[73]:

adminCount=0
for x in [B108,B124,B207,B232,B228]:
    adminCount+= len(x)
    
print(adminCount)


# In[76]:

Admin.info()


# In[75]:

Admin = pd.concat([B108,B124,B207,B232,B228])


# # Maple Conference Rooms

# In[77]:

os.chdir("/Users/adeniyiharrison/Desktop/Outlook Scheduling/Training Room - Maple (BHA)")
os.listdir()


# In[79]:

G111 = pd.read_excel("Conf Rm - G111 (BHA).xls")
G112 = pd.read_excel("Conf Rm - G112 (BHA).xls")
Computers = pd.read_excel("Training Room - Computers (BHA).xls")
Evergreen = pd.read_excel("Training Room - Evergreen (BHA).xls")
Maple = pd.read_excel("Training Room - Maple (BHA) (1).xls")


# In[80]:

G111["Room"] = "G111 (BHA)"
G112["Room"] = "G112 (BHA)"
Computers["Room"] = "Computers (BHA)"
Evergreen["Room"] = "Evergreen (BHA)"
Maple["Room"] = "Maple (BHA)"


# In[81]:

mapleCount = 0
for x in [G111,G112,Computers,Evergreen,Maple]:
    mapleCount+=len(x)
    
print(mapleCount)


# In[83]:

Maple = pd.concat([G111,G112,Computers,Evergreen,Maple])


# # 316 Confrence Room 

# In[84]:

os.chdir("/Users/adeniyiharrison/Desktop/Outlook Scheduling")
os.listdir()


# In[86]:

conf316 = pd.read_excel("Conf. Rm R-316 Rinc.xls")


# In[87]:

conf316["Room"] = "R-316 Rinc"


# In[89]:

conf316.info()


# # ALL COMBINE

# In[91]:

allCount = 0
for x in [Crest,Maintainance,HQ,Admin,Maple,conf316]:
    allCount+=len(x)
    
print(allCount)


# In[93]:

schedulingData.info()


# In[92]:

schedulingData = pd.concat([Crest,Maintainance,HQ,Admin,Maple,conf316])


# In[94]:

schedulingData.to_csv("Outlook Scheduling Data.csv")


# # Date/Time Data Cleaning

# In[95]:

schedulingData.head()


# In[97]:

df=schedulingData.copy()


# In[122]:

datetime.datetime.strptime(df["StartDate"].iloc[0], "%m/%d/%Y").date()


# ### Format start and end dates

# In[124]:

formatStartDate=[]
for x in df["StartDate"]:
    formatStartDate.append(datetime.datetime.strptime(x,"%m/%d/%Y").date())


# In[127]:

formatEndDate=[]
for x in df["EndDate"]:
    formatEndDate.append(datetime.datetime.strptime(x,"%m/%d/%Y").date())


# ### Format start and end times

# In[210]:

formatStartTime = []
for x in df["StartTime"]:
    formatStartTime.append(datetime.datetime.strptime(x, '%I:%M:%S %p').time())


# In[211]:

formatEndTime = []
for x in df["EndTime"]:
    formatEndTime.append(datetime.datetime.strptime(x, '%I:%M:%S %p').time())


# ### Combine dates and times

# In[214]:

formatStartDateTime=[]
for x in range(len(formatStartTime)):
    formatStartDateTime.append(datetime.datetime.combine(formatStartDate[x],formatStartTime[x]))


# In[215]:

formatEndDateTime=[]
for x in range(len(formatEndTime)):
    formatEndDateTime.append(datetime.datetime.combine(formatEndDate[x],formatEndTime[x]))


# ### Add formatted Date and Time to dataframe

# In[216]:

df["Clean Start DateTime"] = formatStartDateTime
df["Clean End DateTime"] = formatEndDateTime


# ### Calculate Meeting Length

# In[220]:

df["Meeting Length"] = df["Clean End DateTime"] - df["Clean Start DateTime"]


# In[234]:

df["Meeting Length"].mean()


# In[237]:

df[df["Meeting Length"] > "8:00:00"]


# In[239]:

df["Clean Start DateTime"].min()


# In[240]:

df["Clean Start DateTime"].max()


# # Calculate Number of Occupants per Meeting

# In[251]:

df[["RequiredAttendees","OptionalAttendees"]].head()


# In[270]:

df["RequiredAttendees"].iloc[1].split(";")


# In[295]:

df["RequiredAttendees"].iloc[0] == 


# In[263]:

type(df["RequiredAttendees"].iloc[1])


# In[302]:

requiredCount=[]
for x in df["RequiredAttendees"]:
    try:
        requiredCount.append(len(x.split(";")))
    except:
        requiredCount.append(np.nan)
        
requiredCount[:5]


# In[303]:

optionalCount = []
for x in df["OptionalAttendees"]:
    try:
        optionalCount.append(len(x.split(";")))
    except:
        optionalCount.append(np.nan)

optionalCount[:5]


# In[304]:

df["Required Count"] = requiredCount
df["Optional Count"] = optionalCount


# In[312]:

df["Total Members"] = df["Total Members"] = df["Required Count"].fillna(0)+df["Optional Count"].fillna(0)


# # Remove and Rearrange Columns

# In[319]:

df=df.drop(["Alldayevent","Reminderonoff","ReminderDate","ReminderTime","MeetingOrganizer","BillingInformation","Categories","Description","Mileage","Showtimeas"],axis=1)

columns = ["Room","Subject","Clean Start DateTime","Clean End DateTime","Meeting Length",
              "Required Count","Optional Count","Total Members","Location","RequiredAttendees","OptionalAttendees",
             "Subject","StartDate","StartTime","EndDate","EndTime","Priority","Private","Sensitivity"]

df=df[columns]


# # Trip data to one year

# In[351]:

df["Clean Start DateTime"].min()


# In[352]:

df["Clean End DateTime"].max()


# ### Date range between 10/13/15 and 10/13/16

# In[364]:

df=df.sort(["Clean Start DateTime","Clean End DateTime"],ascending = [True,False])


# In[366]:

df=df[(df["Clean Start DateTime"] >= "2015-10-13") & (df["Clean End DateTime"] <= "2016-10-13")]


# In[ ]:




# In[ ]:




# In[ ]:



