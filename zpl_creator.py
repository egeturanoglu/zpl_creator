import tkinter as tk
from tkinter import messagebox, scrolledtext, filedialog
import win32print
import os
import re

def create_zpl_code(number, template):
    # Replace FD placeholders in the template with the actual number
    updated_template = re.sub(r'\^FD\d+\^FS', f'^FD{number}^FS', template)
    if "^FD" not in updated_template:
        updated_template += f"^FD{number}^FS"
        
    return updated_template

def print_zpl_code(zpl_code):
    printer_name = win32print.GetDefaultPrinter()
    
    hPrinter = win32print.OpenPrinter(printer_name)
    try:
        hJob = win32print.StartDocPrinter(hPrinter, 1, ("ZPL Label", None, "RAW"))
        win32print.StartPagePrinter(hPrinter)
        
        win32print.WritePrinter(hPrinter, zpl_code.encode('utf-8'))
        
        win32print.EndPagePrinter(hPrinter)
        win32print.EndDocPrinter(hPrinter)
    finally:
        win32print.ClosePrinter(hPrinter)

def create_text_file(number, zpl_code, output_path):
    try:
        with open(output_path, 'w') as file:
            file.write(f"Label Number: {number}\n")
            file.write(zpl_code)
        print(f"Created text file: {output_path}")
    except Exception as e:
        print(f"Error creating text file for {number}: {e}")

def delete_previous_text_files(directory):
    for filename in os.listdir(directory):
        if filename.endswith(".txt"):
            os.remove(os.path.join(directory, filename))

def select_directory():
    directory = filedialog.askdirectory()
    directory_entry.delete(0, tk.END)
    directory_entry.insert(0, directory)

def generate_labels():
    try:
        start_point = int(start_point_entry.get())
        end_point = int(end_point_entry.get())
    except ValueError:
        messagebox.showerror("Invalid input", "Start point and End point must be integers.")
        return

    if end_point <= start_point:
        messagebox.showerror("Invalid input", "End point must be greater than Start point.")
        return

    template = template_text.get("1.0", tk.END).strip()
    if not template:
        messagebox.showerror("Invalid input", "ZPL template cannot be empty.")
        return

    output_directory = directory_entry.get()
    if not output_directory:
        messagebox.showerror("Invalid input", "Output directory cannot be empty.")
        return

    # Ensure the directory exists
    os.makedirs(output_directory, exist_ok=True)

    # Delete previous text files
    delete_previous_text_files(output_directory)

    # Generate and print new labels
    for i in range(start_point, end_point + 1):
        zpl_code = create_zpl_code(i, template)
        print_zpl_code(zpl_code)
        text_file_path = os.path.join(output_directory, f"label_{i}.txt")
        create_text_file(i, zpl_code, text_file_path)
        print(f'Printed label and created text file for {i}')
    
    messagebox.showinfo("Success", f"Labels generated and printed from {start_point} to {end_point}.")

# Set up the GUI
root = tk.Tk()
root.title("ZPL Label Generator")

tk.Label(root, text="Start Point:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
start_point_entry = tk.Entry(root)
start_point_entry.grid(row=0, column=1, padx=5, pady=5)

tk.Label(root, text="End Point:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
end_point_entry = tk.Entry(root)
end_point_entry.grid(row=1, column=1, padx=5, pady=5)

tk.Label(root, text="ZPL Template:").grid(row=2, column=0, padx=5, pady=5, sticky="ne")
template_text = scrolledtext.ScrolledText(root, width=40, height=10)
template_text.grid(row=2, column=1, padx=5, pady=5)

tk.Label(root, text="Output Directory:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
directory_entry = tk.Entry(root, width=30)
directory_entry.grid(row=3, column=1, padx=5, pady=5)
browse_button = tk.Button(root, text="Browse", command=select_directory)
browse_button.grid(row=3, column=2, padx=5, pady=5)

generate_button = tk.Button(root, text="Generate Labels", command=generate_labels)
generate_button.grid(row=4, column=0, columnspan=3, pady=10)

root.mainloop()
