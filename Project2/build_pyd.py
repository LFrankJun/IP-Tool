# -*- coding: utf-8 -*-
"""
Created on Wed Aug 29 13:33:20 2018
@author: Li Zeng hai
"""
 
from distutils.core import setup
from Cython.Build import cythonize


 
setup(
  name = 'any words.....',
  ext_modules = cythonize(["batchFile.py","other.py", "qlyq2.py", "sms.py", "qualityCheck.py"]),
)