# Mapcell-Checker
Easily compare all installed map mods for Project Zomboid
# How it works
This Pythonscript crawls through all of your workshop folders and searches for every occurence of a .lotheader file. These files always contain a coordinate in a form of "X_Y.lotheader". It will then split this filename into different parts: everything after the dot is disregarded and the underscore is used as a divider between coordsX and coordsY.

Then this will get the original foldername, which usually is the modname. The modname gets conjoined with both of the coords and put in a dictonary. If there are coordinates that are in use, the next modnames will be added with a plus.
It then generates a .xls file, which can be opened in excel or the openoffice/libreoffice equivalent.

This file is generated with a x- and y-axis labeled with the x- and y-coordinates. The script checks every cell starting at (0,0) and ending at (65,65) then compares that to the saved coords. Modname at these coordinates get added to the cell. If it detects a plus in the modnames, it will add an exclamation mark and paint the cell background red.
# How to use
There are multiple ways to use this.
I created an .exe file for ease of use on your machines, if you don't want to install anything. Just run it and follow the instructions.
If you want to run the python file itself, you need to install openpyxl as well.


![Folder selection](/assets/folder_select.png?raw=true "Folder selection")

When run, you will be asked about your workshop folder which you need to select. If your Project Zomboid is located on your C:\ drive, it should be opened almost automatically. The name of the folder is "108600".

![Workshop folder](/assets/workshop_folder.png?raw=true "Workshop folder")

My folder is located in here: F:\SteamLibrary\steamapps\workshop\content\108600

![Output path](/assets/output_path.png?raw=true "Output path")

Now you have to select any folder on your pc. This is where your .xls file will be generated. Just put it on your desktop for a quick save.

![Output folder](/assets/output_folder.png?raw=true "Output folder")

![Done](/assets/done.png?raw=true "Finished")

Aaand that's it! When you open up your generated file all used cells will be filled in with the corresponding mod name. If there is more than one mod occupying the same cell, it will be highlighted in red and show all mod names.

# Known issues
If one workshop mod supplies multiple versions of a cell, they will be shown as conflicting. SecretZ is one example of this as it supplies four versions in one, which all occupy the same cells.
# Example Mapcell.xls file
This is a generated file for my subscribed workshop mods, if you want to see how it looks before you try it yourself.
