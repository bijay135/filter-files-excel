# Files Filter Using Excel For Windows
- This program will filter and move files based on excel entries to a diffrent folder
- It can be useful to organize files in a meta folder
- The correct file references needs to be in a excel spreadsheet

# Technologies Used
- Powershell Scripting
- Command Prompt

# How to use
- Extract the zip contents, make sure all the files are in same folder
- Copy the files from meta folder into `MetaFolder\`
- Open the run-program.bat file
- Enter the full path to the excel file example `F:\files-filter-excel\excel_file.xslx`
- Enter the specefic sheet name example `sheet1`
- Enter the new folder to move filtered items to example `F:\files-filter-excel\FilteredFolder`
- Program will match the excel entries with filesystem names
- If the names match it will move each files to the specified new folder
- When all files have been moved it will display success message