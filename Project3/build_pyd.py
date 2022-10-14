# -*- coding: utf-8 -*-
"""
Created on 2022/10/11
@author: Li Yuan JUn
"""
 
from distutils.core import setup
from Cython.Build import cythonize


 
setup(
  name = 'any words.....',
  ext_modules = cythonize(["batchFile.py","other.py", "qlyq2.py", "sms.py", "qualityCheck.py"]),
)