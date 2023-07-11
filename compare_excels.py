import tkinter as tk
import tkinter.filedialog as filedialog
import pandas as pd
import datacompy
import json
from tkinter import messagebox
import os

# JSON file to store the user inputs
SETTINGS_FILE = "settings.json"

def save_and_close():
    save_settings()
    root.quit()
    

def save_settings():
    settings = {
        "file1": entry_file1.get(),
        "sheet1": entry_sheet1.get(),
        "file2": entry_file2.get(),
        "sheet2": entry_sheet2.get(),
        "join_column": entry_join.get()
    }
    with open(SETTINGS_FILE, "w") as file:
        json.dump(settings, file)

def load_settings():
    try:
        with open(SETTINGS_FILE, "r") as file:
            settings = json.load(file)
            entry_file1.insert(tk.END, settings["file1"])
            entry_sheet1.insert(tk.END, settings["sheet1"])
            entry_file2.insert(tk.END, settings["file2"])
            entry_sheet2.insert(tk.END, settings["sheet2"])
            entry_join.insert(tk.END, settings["join_column"])
    except FileNotFoundError:
        pass

def select_file(entry):
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    entry.delete(0, tk.END)
    entry.insert(tk.END, file_path)

def compare_excel_files():
    file_path1 = entry_file1.get()
    file_path2 = entry_file2.get()
    sheet_name1 = entry_sheet1.get()
    sheet_name2 = entry_sheet2.get()
    join_column = entry_join.get()

    try:
        df1 = pd.read_excel(file_path1, sheet_name=sheet_name1)
        df2 = pd.read_excel(file_path2, sheet_name=sheet_name2)

        file_name1 = os.path.basename(file_path1)
        file_name2 = os.path.basename(file_path2)

        compare = datacompy.Compare(
            df1,
            df2,
            join_columns=join_column,
            abs_tol=0.0001,
            rel_tol=0,
            df1_name=file_name1,
            df2_name=file_name2
        )

        report_window = tk.Toplevel(root)
        report_window.title("Comparison Report")
        report_window.geometry("1000x1000")  # Set the window size to 1000x1000 pixels

        report_text = tk.Text(report_window, font=("Courier New", 12))  # Set a monospaced font
        report_text.pack(fill=tk.BOTH, expand=True)

        # Insert the comparison report text
        report_text.insert(tk.END, str(compare.report()))

        join_column_lowered = join_column.lower()

        # Find the index of "catalog_id" in the text
        index = report_text.search(join_column_lowered, 1.0, tk.END)

        report_text.tag_configure("red", foreground="red")

        while index:
            # Configure the tag for red color
            

            # Get the next word after "catalog_id"
            next_word_index = report_text.search(r"[a-z|A-Z]", index + f"+{len(join_column)}c", stopindex=tk.END, regexp=True)
            if next_word_index:
                next_word_end = report_text.search(r"\s", next_word_index, stopindex=tk.END, regexp=True)

                # Apply the red color tag to the next word
                report_text.tag_add("red", next_word_index, next_word_end)

            # Find the next occurrence of "catalog_id"
            index = report_text.search(join_column_lowered, next_word_end, tk.END)
        
        

        def close_report_window():
            report_window.destroy()

        report_window.protocol("WM_DELETE_WINDOW", close_report_window)
    except Exception as e:
        messagebox.showerror("Error", str(e))

root = tk.Tk()
root.title("Excel File Comparator")

# File 1
label_file1 = tk.Label(root, text="Excel File 1:")
label_file1.grid(row=0, column=0)
entry_file1 = tk.Entry(root, width=30)
entry_file1.grid(row=0, column=1)
button_file1 = tk.Button(root, text="Browse", command=lambda: select_file(entry_file1))
button_file1.grid(row=0, column=2)

# Sheet 1
label_sheet1 = tk.Label(root, text="Sheet 1:")
label_sheet1.grid(row=1, column=0)
entry_sheet1 = tk.Entry(root, width=30)
entry_sheet1.grid(row=1, column=1)

# File 2
label_file2 = tk.Label(root, text="Excel File 2:")
label_file2.grid(row=2, column=0)
entry_file2 = tk.Entry(root, width=30)
entry_file2.grid(row=2, column=1)
button_file2 = tk.Button(root, text="Browse", command=lambda: select_file(entry_file2))
button_file2.grid(row=2, column=2)

# Sheet 2
label_sheet2 = tk.Label(root, text="Sheet 2:")
label_sheet2.grid(row=3, column=0)
entry_sheet2 = tk.Entry(root, width=30)
entry_sheet2.grid(row=3, column=1)

# Join Column
label_join = tk.Label(root, text="Join Column:")
label_join.grid(row=4, column=0)
entry_join = tk.Entry(root, width=30)
entry_join.grid(row=4, column=1)

# Compare Button
compare_button = tk.Button(root, text="Compare", command=compare_excel_files)
compare_button.grid(row=5, column=0, columnspan=3, pady=10)

# Load previous settings if available
load_settings()

# Save settings on close
root.protocol("WM_DELETE_WINDOW", save_and_close)

root.mainloop()
