import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import random
import os

# Function to load the regular timetable
def load_timetable(file_path):
    df = pd.read_excel(file_path)
    return df

# Function to generate the remedial timetable
def generate_remedial_timetable(regular_timetable):
    subjects = ['Sub1', 'Sub2', 'Sub3', 'Sub4', 'Sub5']
    remedial_classes = {}
    
    for subject in subjects:
        remedial_classes[subject] = 0  # Track the number of classes scheduled for each subject
    
    # Create a copy of the regular timetable to fill in remedial classes
    remedial_timetable = regular_timetable.copy()
    
    # Iterate over each time slot and day to fill in empty slots
    for i in range(len(remedial_timetable)):
        for j in range(1, len(remedial_timetable.columns)):
            if pd.isna(remedial_timetable.iat[i, j]):  # Check for empty slots
                subject = random.choice(subjects)  # Randomly choose a subject for the remedial class
                remedial_classes[subject] += 1  # Increment the count for the chosen subject
                remedial_timetable.iat[i, j] = "Rem_" + subject  # Fill the empty slot

    return remedial_timetable, remedial_classes

# Function to save the remedial timetable to Excel
def save_timetable(remedial_timetable):
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                               filetypes=[("Excel files", ".xlsx"), ("All files", ".*")])
    if save_path:
        remedial_timetable.to_excel(save_path, index=False)
        messagebox.showinfo("Success", "Remedial timetable saved successfully!")

# Function to handle the import button click
def import_timetable():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", ".xlsx"), ("All files", ".*")])
    if file_path:
        try:
            regular_timetable = load_timetable(file_path)
            remedial_timetable, remedial_classes = generate_remedial_timetable(regular_timetable)
            display_remedial_timetable(remedial_timetable)
            save_button.config(state=tk.NORMAL)  # Enable save button
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

# Function to display the remedial timetable in a new window
def display_remedial_timetable(remedial_timetable):
    # Create a new window
    new_window = tk.Toplevel(root)
    new_window.title("Remedial Timetable")
    
    # Create a text widget to display the timetable
    text_widget = tk.Text(new_window, wrap=tk.WORD)
    text_widget.pack(expand=True, fill=tk.BOTH)
    
    # Convert the timetable DataFrame to string and insert it into the text widget
    timetable_str = remedial_timetable.to_string(index=False)
    text_widget.insert(tk.END, timetable_str)

    # Add a save button to download the timetable
    save_button = tk.Button(new_window, text="Download Timetable", command=lambda: save_timetable(remedial_timetable))
    save_button.pack(pady=10)

# Main application window
root = tk.Tk()
root.title("Remedial Class Timetable Generator")

# Create and place the import button
import_button = tk.Button(root, text="Import Regular Timetable", command=import_timetable)
import_button.pack(pady=20)

# Create the save button (initially disabled)
save_button = tk.Button(root, text="Save Remedial Timetable", command=lambda: None)
save_button.pack(pady=10)
save_button.config(state=tk.DISABLED)

# Start the Tkinter event loop
root.mainloop()