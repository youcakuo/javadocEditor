#!/usr/bin/env python
# coding: utf-8

# In[11]:


import os, subprocess
import shutil
import time, sys

try:
    shutil.copy('\\\\10.204.1.80\\tfb\\T-NBTS 專案_第二階段\\工作區\個人工作區\\Shen\\javadoc\\JavadocEditor.exe', '.\JavadocEditor.exe')
except:
    print('copy javadocEditor fail')
finally:
    print(os.getcwd())
    os.system('JavadocEditor.exe')
    print('done', end='', flush=True)
    for i in range(3):
        time.sleep(0.7)
        print('.', end='', flush = True)


# In[ ]:




