# TRUMPF-Export-Dump Data Extraction Program Doc

## Introduction:
The TRUMPF-Export-Dump Data Extraction program is a Python script that reads data from files in a directory and its sub-directories and writes the data to an Excel file. The program is designed to read specific data from files with the .lst extension and write the data to an Excel file extension. The program is intended to be used by individuals who need to extract specific data from the Trumpf export dump.

## Installation:
The program requires Python 3.x and the following Python packages: os, re, xlsxwriter. These packages can be installed using pip. To install the packages, open a command prompt or terminal window and enter the following commands:

```
pip install os
pip install re
pip install xlsxwriter
```

## Usage:
To use the program, follow these steps:

1. Create a file named "Maschinenzeiten_TF.cfg" in the same directory as the Python script.
2. In the "Maschinenzeiten_TF.cfg" file, enter the directory path where the .lst files are located.
3. Run the Python script.

The program will read all .lst files in the specified directory and its sub-directories and write the data to an Excel file named "Maschinenzeiten_TF.xlsx". The Excel file will be created in the same directory as the Python script.

## Code Structure:
The program consists of three functions: main(), loop_through_directories(), and write_to_excel(). The main() function is responsible for initializing variables, calling other functions, extracting the required data and opening the Excel file once it has been written. The loop_through_directories() function is responsible for looping through the directories and get the data from .lst files. The write_to_excel() function is responsible for writing data to an Excel file.

## Example output:
![Excel Screenshot](https://user-images.githubusercontent.com/91200978/234282029-77f17d06-a039-4a3f-95f0-94dd337d4be0.png)

## Conclusion:
The program is a simple yet powerful tool for extracting specific data from .lst files and storing it in an Excel file. The program is easy to use and can save users a significant amount of time when working with large amounts of data.
