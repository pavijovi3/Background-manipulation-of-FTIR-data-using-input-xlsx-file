# Background-manipulation-of-FTIR-data-using-input-xlsx-file
Background manipulation of FTIR data using input xlsx file

Dependencies: The tool requires the following Python libraries to be installed:
tkinter: This library is usually included with standard Python installations.
pandas: You can install it using the command: pip install pandas.
Download the Script: Download the provided Python script (filename: background_reprocessing_tool.py) and save it to a directory of your choice.

Usage
Running the Tool: Open a terminal or command prompt, navigate to the directory where you saved the script, and run the script using the following command:
Copy code: python background_reprocessing_tool.py
Graphical User Interface (GUI): The tool will open a graphical interface window with the title "Background Reprocessing Tool."

Instructions:
The tool's interface displays a label with a brief description of its purpose. Click the "Reprocess background now" button to start using the tool.
The "Reprocess background now" button triggers the data processing and manipulation procedure.

Data Processing:
The tool will prompt you to select an input Excel (XLSX) file (with extension .xlsx) containing the FTIR data series.
After selecting the input file, you will be prompted to choose a specific column from the data series for background manipulation.

Output File:
You will be prompted to select an output directory where the processed data will be saved.
The tool will create a new XLSX file containing the processed data, where the values of the chosen column will be subtracted from all other columns.
Completion Message: After the data processing is completed, a message box will appear with information about the location of the saved output file.

Important Notes
The tool is designed to manipulate FTIR data series stored in Excel files with .xlsx format.
Make sure to have both the input and output directories accessible.
The tool provides a user-friendly interface for choosing the column and directories, enhancing ease of use.
