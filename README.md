# Combining multiple Excel files into one Workbook
### Description:
A very common problem in many companies is the amount of data stored in different files. Files that may contain the same information structure, but from different periods or different regions, for example.

Searching for data from different files in Excel can give us some headaches such as:

- High formulas complexity, making the spreadsheet calculation slower and the file size larger
- Reference problems with the referenced file
- Formula maintenance becomes more difficult especially when there are many files

To perform a complete and more effective analysis in Excel, having all the files in the same table makes it much easier.
### The Problem
To consolidate several tables into a single file, Excel already provides some tools but they are mainly aimed at tables that are in the same file but on separate tabs. A possible solution to consolidate  copying each table from each file and pasting it into the main table, but this can take a lot of time especially if there are many files.

In this way, I decided to create a Python program that could solve this problem, where it is possible to consolidate several files into just one in a few minutes, saving time to dedicate to other tasks.
### The Solution
With the aim of unifying several Excel files into just one, I developed a Python automation program that searches for files in a folder and combines the data into an output file in another folder.

Knowing that each Excel user has a standard way of structuring data differently, I developed the program in a way that the user can identify, when running the program via CLI, which tab, in which line is the table header and which are the columns wanted to copy to the output file. In this way, the program can be used in practically all forms of Excel tables.
### About the script
This Python script is designed to consolidate data from multiple Excel files into a single workbook. It supports both regular Excel files (.xls, .xlsx) and Excel Binary Workbook files (.xlsb). The user can choose specific sheets, headers, and columns to be consolidated into the final output.
### Requirements
- Python 3.x
- Pandas
- Openpyxl
- Pyxlsb
### Usage
1. Place the Excel files you want to consolidate in the **`./input`** directory.
2. Run the script.
```
python project.py
```
3. Follow the on-screen instructions to select sheets, headers, and columns.
4. The consolidated data will be saved in the **`./output`** directory as **`output.xlsx`**.
### Note
- Only Excel files (.xls, .xlsx) and Excel Binary Workbook files (.xlsb) are supported.
- Make sure to provide accurate input during the interactive session to avoid errors.
### Features
- Supports both regular Excel files and Excel Binary Workbook files.
- User-friendly interface for selecting sheets, headers, and columns.
- Consolidates data into a single Excel workbook.
- Outputs the consolidated data to **`output.xlsx`** in the **`./output`** directory.
### Conclusion
In summary, this Python program is really good at working with different kinds of tables in Excel. It can easily fit into all sorts of spreadsheets without needing any changes to the computer code.

The program does the work automatically, saving a lot of time that would otherwise be spent combining files one by one. This efficiency means people can use their time for more important things.

In a time when everyone wants to be more productive, this program is super helpful. It's like a valuable tool that gives you a lot in return. Whether you work with numbers, health info, or anything else, this program is great at putting your data together quickly and easily.
### Contact
[Linkedin](https://www.linkedin.com/in/camposerick/)

[Twitter](https://twitter.com/camposerick_)
