import os
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from openpyxl import Workbook
from openpyxl.styles import PatternFill

zomboid_path = ''
output_path = ''
root = tk.Tk()
root.withdraw()
workshop_box = messagebox.askokcancel('Workshop folder','Please select your Project Zomboid workshop folder called "108600", which is usually located in "C:\Program Files (x86)\Steam\steamapps\workshop\content"')
if workshop_box:
    zomboid_path = filedialog.askdirectory(initialdir='C:\Program Files (x86)\Steam\steamapps\workshop\content',mustexist=True,title='Please select your Project Zomboid workshop folder: 108600')
    if zomboid_path == '':
        quit()
else:
    quit()
output_box = messagebox.askokcancel('Output path','Please select where you want to save your XLS file')
if output_box:
    output_path = filedialog.askdirectory(mustexist=True,title='Please select your save location')
    if output_path == '':
        quit()
else:
    quit()

def find_files_with_coordinates(root_folder):
    # Create a dictionary to store the folder names
    folders = {}
    # Walk through the root folder and its subfolders
    for foldername, subfolders, filenames in os.walk(root_folder):
        # Loop through the filenames
        for filename in filenames:
            # Check if the filename matches the pattern
            if filename.endswith('.lotheader'):
                # Extract the coordinates from the filename
                coords = filename.split('.')[0].split('_')
                # Get the folder name
                folder = os.path.basename(foldername)
                # Add the folder name to the dictionary
                if (coords[0], coords[1]) in folders:
                    folders[(coords[0], coords[1])] += '+' + folder
                else:
                    folders[(coords[0], coords[1])] = folder
    # Write the folder names to a CSV file
    wb = Workbook()
    ws = wb.active
    ws.title = 'Mapcells'
    ws.cell(row=1, column=2).value = 'X'
    ws.cell(row=2, column=1).value = 'Y'
    for i in range(0, 66):
        ws.cell(row=1, column=i+3).value = i
        ws.cell(row=i+3, column=1).value = i
        for j in range(0, 66):
            cell = folders.get((str(j), str(i)), '')
            if '+' in cell:
                cell = '!' + cell
                ws.cell(row=i+3, column=j+3).fill = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')
            ws.cell(row=i+3, column=j+3).value = cell
    wb.save(output_path + '/Mapcells.xls')

# Call the function with the zomboid folder
find_files_with_coordinates(zomboid_path)
exit_message = messagebox.showinfo('Done','You can find your XLS File here: ' + output_path + '/Mapcells.xls')
