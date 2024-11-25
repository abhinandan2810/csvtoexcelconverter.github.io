import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Progressbar, Style
import os

# Function to convert CSV to Excel
def convert_to_excel():
    # Open file dialog to select one or more CSV files
    csv_files = filedialog.askopenfilenames(filetypes=[("CSV Files", "*.csv")])
    if not csv_files:
        return  # If no files are selected, exit function

    try:
        # Open save dialog to choose location for the Excel file
        excel_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if not excel_file:
            return  # If no location is chosen, exit function

        # Create a new Excel writer
        with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
            total_rows = sum([len(pd.read_csv(csv_file)) for csv_file in csv_files])
            processed_rows = 0

            # Initialize progress bar
            progress_bar['maximum'] = total_rows
            progress_bar['value'] = 0
            root.update_idletasks()

            # Iterate over selected CSV files
            for idx, csv_file in enumerate(csv_files):
                # Read the CSV file
                df = pd.read_csv(csv_file)

                # Write DataFrame to Excel file in a new sheet
                sheet_name = os.path.splitext(os.path.basename(csv_file))[0]
                df.to_excel(writer, sheet_name=sheet_name, index=False)

                # Update progress bar based on rows processed
                processed_rows += len(df)
                progress_bar['value'] = processed_rows
                root.update_idletasks()

            # Update status
            status_label.config(text="Conversion Successful!", fg="#4CAF50")  # Green color
            messagebox.showinfo("Success", "CSV files have been successfully converted to Excel!")

    except Exception as e:
        # Handle any error that occurs during the conversion
        messagebox.showerror("Error", f"An error occurred: {e}")
        status_label.config(text="Conversion Failed", fg="#F44336")  # Red color
        progress_bar['value'] = 0

# GUI Setup
root = tk.Tk()
root.title("CSV to Excel Converter")
root.geometry("550x350")
root.resizable(False, False)
root.configure(bg="#f0f0f0")

# Apply a style for ttk elements
style = Style()
style.configure("TButton", font=("Arial", 10), padding=6, relief="flat", background="#008CBA", foreground="#000000")
style.configure("TLabel", font=("Arial", 10), background="#f0f0f0")

# Header Label
header_label = tk.Label(root, text="CSV to Excel Converter", font=("Arial", 16, "bold"), fg="#008CBA", bg="#f0f0f0")
header_label.pack(pady=10)

# Instruction Label
instruction_label = tk.Label(root, text="Select one or more CSV files to convert into an Excel file.", bg="#f0f0f0", fg="#000000")
instruction_label.pack(pady=5)

# Button to trigger conversion
convert_button = tk.Button(root, text="Select CSV Files and Convert", command=convert_to_excel, bg="#008CBA", fg="#000000", font=("Arial", 12), relief="raised", bd=2)
convert_button.pack(pady=20)

# Progress bar for conversion
progress_bar = Progressbar(root, length=400, mode='determinate', style="TProgressbar")
progress_bar.pack(pady=10)

# Status label to display messages
status_label = tk.Label(root, text="", font=("Arial", 10, "italic"), bg="#f0f0f0")
status_label.pack(pady=5)

# Footer Label
footer_label = tk.Label(root, text="Developed by Abhinandan | CSV to Excel Converter", font=("Arial", 8), fg="#666666", bg="#f0f0f0")
footer_label.pack(side="bottom", pady=10)

# Style for progress bar
style.configure("TProgressbar", thickness=20, troughcolor="#e0e0e0", background="#4CAF50")

root.mainloop()
