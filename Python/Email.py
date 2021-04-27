#!/usr/bin/env python
# coding: utf-8

# In[4]:


import win32com.client as win32
outlook = win32.Dispatch("Outlook.application")
message = outlook.Createitem(0)

message.SentOnBehalfOfName = "Levin"
message.To = "Levin@Outlook.com"
message.CC = "Levin@Outlook.com"
message.BCC = "Levin@Outlook.com"
message.Subject = "Levin"
message.Body = "Hi Levin, \n\nMy Name is Levin."
message.Attachments.Add ("C:\\Users\\Hp\\Desktop\\Data.csv")
message.display()


# In[ ]:




