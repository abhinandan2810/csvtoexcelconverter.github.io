# csvtoexcelconverter.github.io

Description:
The CSV to Excel Converter is a simple and intuitive tool that allows users to convert one or multiple CSV files into an Excel workbook. Each CSV file is stored as a separate sheet within the Excel file. This tool is designed to streamline the process of converting raw data from CSV format into a more structured Excel format, making it easier to analyze, manage, and share the data. With a user-friendly interface, the app features a progress bar to track the conversion status and provides success or error messages.

Requirements:
Python (version 3.6 or higher) installed on your system.
Libraries:
pandas: Used for reading and writing CSV and Excel files.
tkinter: Used for the graphical user interface (GUI).
xlsxwriter: Used by pandas for writing to Excel files with additional functionality (like styling).
Operating System: This tool works on Windows, macOS, and Linux.

Instructions:
Install Required Libraries: Ensure you have the required libraries installed. You can install them using the following commands:
pip install pandas
pip install openpyxl
pip install xlsxwriter
Running the Application:
Step 1: Launch the program by running the Python script csv_to_excel_converter.py (or whatever the script is named).
Step 2: The application window will open.
Selecting CSV Files:
Click the "Select CSV Files and Convert" button to open a file dialog where you can select one or more CSV files from your computer.
The selected CSV files will be added to the conversion queue.
Choosing the Excel File Location:
After selecting the CSV files, a save dialog will appear asking you to choose the location and name of the output Excel file.
You can specify the file name and save it with the .xlsx extension.
Conversion Process:
As the CSV files are being processed, a progress bar will display the percentage of the conversion that has been completed.
Each CSV file will be added to the Excel file as a separate sheet. The sheet names will correspond to the original filenames of the CSV files (without the .csv extension).
Completion and Feedback:
Once the conversion is complete, a message box will appear notifying you of the success of the operation.
If an error occurs at any point, an error message will be displayed detailing the issue.
Exit the Application:
After the conversion is finished, you can close the application window.
