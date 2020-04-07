# -*- coding: utf-8 -*-
"""
Created on Tue Mar 31 20:35:55 2020

@author: Huang
"""


import os,win32com
for root,dirs,files in os.walk('.'):
    for name in dirs:
        for i in range(2,5):
            name
            temp_name = name.replace("1",str(i))
            os.mkdir(temp_name)
            
            path = os.getcwd()+"\\" + temp_name
            with open(os.path.join(path,temp_name+".doc") , 'w') as temp_file:
                temp_file.close()
  