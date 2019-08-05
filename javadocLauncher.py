#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import os, subprocess
import shutil

try:
    shutil.copy('\\\\10.204.1.80\\tfb\\T-NBTS 專案_第二階段\\工作區\個人工作區\\Shen\\javadoc\\JavadocEditor.exe', '.\JavadocEditor.exe')
except:
    print('copy javadocEditor fail')
finally:
    print(os.getcwd())
    os.system('JavadocEditor.exe')
    #winView = './JavadocEditor.exe'
    #subprocess.run(winView, shell=True)
    print('done...')

    


# In[ ]:




