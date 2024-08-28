import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from PIL import Image, ImageTk
import pandas as pd
import os
import re
# from pathlib import Path

# Placeholder for your dat module
import dat
model = dat.Model("glove.840B.300d.txt", "words.txt")

def upload_file():
    """Function to upload an Excel file and let the user select the relevant columns."""
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        global df
        df = pd.read_excel(file_path)
        file_label.config(text=os.path.basename(file_path))
        column_options = list(df.columns)

        # Enable dropdowns for column selection
        id_column_selector['values'] = column_options
        words_column_selector['values'] = column_options

        # Set default column selections
        id_column_selector.current(0)  # Set to the 0th column (first column)
        words_column_selector.current(1)  # Set to the 1st column (second column)
        id_column_selector.config(state='readonly')
        words_column_selector.config(state='readonly')

def process_file():
    """Function to process the Excel file based on user-selected columns."""
    try:
        id_column = id_column_selector.get()
        words_column = words_column_selector.get()

        # Check if the user has selected columns
        if id_column == 'Select Email Column' or words_column == 'Select Words Column':
            messagebox.showerror("Error", "Please select both ID and Words columns.")
            return
        
        # Initialize an empty list to store the results
        results = []

        # Iterate over each row in the DataFrame
        for index, row in df.iterrows():
            # print(row)
            word_list = row[words_column].split(',')
            cleaned_words = [''.join(re.findall(r'[a-zA-Z]+', word)).lower() for word in word_list]
            cleaned_words = [word for word in cleaned_words if word]
            validated_words = [model.validate(word) for word in cleaned_words if model.validate(word) is not None]
            # word_list = [word.strip() for word in word_list]  # Remove any extra whitespace
            print(validated_words)

            # Compute the DAT score using your model
            # Replace the next line with your model's DAT function
            score = model.dat(validated_words,len(validated_words))  # Example: model.dat(["cat", "dog"], 2)
            
            # Append the id and score to the results list
            results.append({'Email': row[id_column], 'No of Valid Words':len(validated_words),'DAT Score': score})
        
        # Create a new DataFrame with the results
        result_df = pd.DataFrame(results)
        
        # Save result to a new Excel file
        output_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                        filetypes=[("Excel files", "*.xlsx;*.xls")])
        if output_file_path:
            result_df.to_excel(output_file_path, index=False)
            messagebox.showinfo("Success", "File processed and saved successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# Main window setup
root = tk.Tk()
root.title("DAT Software")

# Set fixed size for window
root.geometry("600x600")  # Adjusted height for additional UI elements
root.resizable(False, False)

# Configure button and label styles for a professional look
style = ttk.Style()
style.configure("TButton", font=("Helvetica", 10, "bold"), padding=10)
style.configure("TLabel", font=("Helvetica", 9), padding=10)
style.configure("TCombobox", font=("Helvetica", 9), padding=10)

# Logo at the top using relative path
# current_dir = Path(__file__).parent
# logo_path = current_dir / 'assets' / 'logo.png'

logo_img = Image.open('iiitdm_logo.png')
# logo_img = logo_img.resize((100, 100))  # Adjust size as needed
logo = ImageTk.PhotoImage(logo_img)
logo_label = tk.Label(root, image=logo)
logo_label.pack(pady=10)

# Upload button and file label
upload_btn = ttk.Button(root, text="Upload Excel File", command=upload_file, style="TButton")
upload_btn.pack(pady=10)

file_label = ttk.Label(root, text="No file uploaded", foreground="gray", style="TLabel")
file_label.pack(pady=5)

# Frame for horizontal column selectors
selector_frame = tk.Frame(root)
selector_frame.pack(pady=15)

# ID Column Selector Label and Combobox within the frame
id_label = ttk.Label(selector_frame, text="ID Column:", style="TLabel")
id_label.grid(row=0, column=0, padx=10)

id_column_selector = ttk.Combobox(selector_frame, state='disabled', style="TCombobox", width=15)
id_column_selector.grid(row=1, column=0, padx=10)

# Words Column Selector Label and Combobox within the frame
words_label = ttk.Label(selector_frame, text="Words Column:", style="TLabel")
words_label.grid(row=0, column=1, padx=10)

words_column_selector = ttk.Combobox(selector_frame, state='disabled', style="TCombobox", width=15)
words_column_selector.grid(row=1, column=1, padx=10)

# Process button
process_btn = ttk.Button(root, text="Process and Download", command=process_file, style="TButton")
process_btn.pack(pady=20)

# Powered by label
powered_by_label = ttk.Label(root, text="Designed by Kameleon", foreground="gray", style="TLabel")
powered_by_label.pack(side=tk.BOTTOM, pady=10)

root.mainloop()
