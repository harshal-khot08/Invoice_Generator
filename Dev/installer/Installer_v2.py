from docx import Document
import openpyxl
from docx2pdf import convert
import os
from num2words import num2words
import shutil
from os import path
from os import listdir
from os.path import join, isfile
import os, winshell
from win32com.client import Dispatch

directory = "appFiles"
outputFolder="output"

list = ['appFiles', 'output']
  

parent = "C:/InvoiceGen/"

path = os.path.join(parent, directory)
outputPath=path = os.path.join(parent, outputFolder)
if(not os.path.exists(path)):
    os.makedirs(parent)


if(not os.path.exists(path)):
    for items in list:
        path = os.path.join(parent, items)
        os.mkdir(path)
        print("Directory '% s' created" % directory)

cuurent = os.getcwd()
src=cuurent+'//input//sample_invoice.docx'
src2=cuurent+'//input//InvoiceGUISnapshot.exe'
des='C://InvoiceGen//appFiles//'
des2='C://InvoiceGen//'
print("App files are copied..")

shutil.copy(src ,des)
shutil.copy(src2 ,des2)




desktop = winshell.desktop()
path = os.path.join(desktop, "InvoiceGenerator.lnk")
target = r"C://InvoiceGen//InvoiceGUISnapshot.exe"
wDir = r"C://InvoiceGen"
shell = Dispatch('WScript.Shell')
icon = r"C://Users//Harshal//Downloads//icon.png"
shortcut = shell.CreateShortCut(path)
shortcut.Targetpath = target
shortcut.WorkingDirectory = wDir
shortcut.IconLocation = icon

shortcut.save()

