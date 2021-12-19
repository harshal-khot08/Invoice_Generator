from docx import Document
import openpyxl
from docx2pdf import convert
import os
from num2words import num2words
import shutil
from os import path
from os import listdir
from os.path import join, isfile

directory = "//InvoiceGen//input"

parent_dir = "C:"

path = os.path.join(parent_dir, directory)

if(not os.path.exists(path)):
    os.makedirs(path)
    print("Directory '% s' created" % directory)
cuurent = os.getcwd()
src=cuurent+'//input'
print(src)
files=listdir(src)
print(files)


shutil.copyfile(src ,des)
