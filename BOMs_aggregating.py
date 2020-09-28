# -*- coding: utf-8 -*-
"""
Created on Wed Sep 16 12:32:25 2020

@author: User
"""
import os
import pandas as pd
from datetime import datetime

start_path = 'E:/!PY/BOM_agreagating/Data'
start_path = "E:/!AD/AD_Projects/RBH"
#start_path = 'E:\\!AD\\AD_Projects'
tree = os.walk(start_path)

files = []
names = []

for t in tree:
    full_path = t[0]
    if t[2] != None:
        for name in t[2]:
            ext = os.path.splitext(name)[-1].lower()
            if ext == '.xlsx':
                files.append(t[0] + '\\' + name)
                names.append(os.path.splitext(name)[0])

data = pd.DataFrame()
cols = ['Designator', 'Comment', 'Quantity', 'HelpURL']

for f, n in zip(files, names):
    try:
        df = pd.read_excel(f)[cols]
        df['Project_name'] = n
        data = data.append(df, ignore_index = True)
    except Exception:
        print(f)
        print("Couldn't read data from file!" )

time =  datetime.now().strftime("_%m_%d_%Y__%H_%M")
save_root_path = 'E:\\!PY/BOM_agreagating\\BOMs'

save_path = save_root_path + '\\' + 'full_BOM' + time + '.xlsx'
print("Saving to path:" + save_path)

data.to_excel(save_path)
