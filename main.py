#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Mar 30 11:29:43 2021

@author: aronj23
"""

# TODO: TEST TIMING FUNCTION
import os
from os import path
import shutil
import zipfile as zf
import _thread
import time as t
global Fin
Fin = False
StartTime = t.time()

fileext = [
".xls",".pdf",".csv",".ods",   # Excel data files alongside pdf
".pptx",".pptm",".ppt",".ppsx",".ppsm",".pps",".odp" ]# Powerpoint files
CurPath = r"/Users/aronj23/Desktop/py2app/"


FilesTR = []
# This code thanks to "shadow0359"
for root, dirs, files in os.walk("/Users", topdown=False): # For every directory, file and root in the users directory:
   for name in files: # For every file that has a name:
       FilesTR += [os.path.join(root, name)] # Append the file name and path to the files list 
    
       
       
valid = []
x = (len(FilesTR) -1 ) #  Create a variable that is 1 less than the num. of items in the list of all files
while x >= 0: # While x is less than or equal to 0
    for item in fileext: # For every file extension in the list of wanted ones
        if item in FilesTR[x]: # Take that item and check if it is in the list of all files
            valid.append(FilesTR[x]) # Append it to a new list
    x = x - 1 # subtract 1 from x, to move onto the next item in the list of all files.
    

# Doing this is about 100x faster than going:
# if ".pdf" in FilesTR or ".xls" in FilesTR [AND SO ON]

valid = list(dict.fromkeys(valid)) # Convert the list to a dictionary, as dicts cannot have duplicate values (thanks to W3 schools for this tip)
print(str(len(valid))) 

y = 0
x = 0

for item in valid: # For every item in the valid list.
    try: # Try and do the following
        shutil.copyfile(item, CurPath + r'FindFiles/' + str(x) + item.split("/")[int(len(item.split("/"))-1)])
    except: # Except for if there is an error
        print(valid[x]) # Print the item that caused an error
        y += 1 # Add 1 to the total amount of errors
    x +=1
    
    
    
# Now lets zip! I make this a function so i can start it in a seporate thread.
# This was modeled after a tutorial avalible here: https://www.tutorialspoint.com/python3/python_multithreading.htm 
def ZipIt(threadName):
# I used this documentation here: https://docs.python.org/3/library/zipfile.html 
    FilesTZ = []

    for root, dirs, files in os.walk(CurPath + r"FindFiles/", topdown=False): # For every directory, file and root in the files directory:
        for name in files: # For every file that has a name:
            FilesTZ += [os.path.join(root, name)] # Append the file name and path to the files list 
    FilesTZ = list(dict.fromkeys(FilesTZ))
    with zf.ZipFile(CurPath + 'zipped.zip', "w") as myzip: # open a zipped file, create if it isn't 
        for item in FilesTZ: # For every item in the filestz
            myzip.write(item) # Write it to the zip file.
    print("Done Zipping!")
    _thread.interrupt_main() # This is the only way to check if a thread is done (?)    
# Now to start the thread
_thread.start_new_thread(ZipIt,("Zip-Found-1",))

# Now we start doing our specific files

try:  
# The above function is thanks to stackoverflow user atzz. Question was asked by Daryl Spitzer 
    def copytree(src, dst, symlinks=False, ignore=None):
        for item in os.listdir(src):
            s = os.path.join(src, item)
            d = os.path.join(dst, item)
            if os.path.isdir(s):
                shutil.copytree(s, d, symlinks, ignore)
            else:
                shutil.copy2(s, d)
                
                
    suc = False
    try:
        envi = os.environ['LOGNAME'] # Get the users name
        suc = True
    except:
        try:
            print("Couldn't find the LOGNAME key...\nLet me try using getpass")
            import getpass as gp
            envi = gp.getuser() # This was thanks to user Konstantin Tenzin on stackoverflow
            suc = True
        except:
            print("One last try!")
            try:
                import pwd # This answer thanks to Grzegorz Skibinski on stackoverflow.
                penvi = wd.getpwuid(os.getuid())[0]
                suc = True
            except:
                suc = False
                
    
    if suc == True:
            copytree(r"/Users/" + envi + r"/Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles",CurPath + r"Outlook/")
except KeyboardInterrupt:
    Fin = True
while Fin != True:
    t.sleep(1)
FinTime = t.time()
print("This took: " + str(FinTime - StartTime) + " seconds!")
