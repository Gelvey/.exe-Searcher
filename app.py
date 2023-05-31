import os
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk

def scan_folder():
    # Get the selected folder for scanning
    folder_path = filedialog.askdirectory(initialdir='/', title='Select Folder for Scanning')
    if not folder_path:
        return

    # Get the selected directory to save the shortcuts
    save_dir = filedialog.askdirectory(initialdir='~/Desktop', title='Select Save Directory')
    if not save_dir:
        return

    # Create the ".exe files" directory if it doesn't exist
    exe_dir = os.path.join(save_dir, '.exe files')
    os.makedirs(exe_dir, exist_ok=True)

    # Calculate the total number of files to scan
    total_files = sum([len(files) for _, _, files in os.walk(folder_path) if any(file.endswith('.exe') for file in files)])

    # Create the progress bar
    progress_bar['maximum'] = total_files
    progress_bar['value'] = 0

    # Scan the specified folder for .exe files and create shortcuts
    for root, _, files in os.walk(folder_path):
        for file in files:
            if file.endswith('.exe'):
                file_path = os.path.join(root, file)
                shortcut_path = os.path.join(exe_dir, f"{file}.lnk")
                create_shortcut(file_path, shortcut_path)

                # Update the progress bar
                progress_bar['value'] += 1
                window.update_idletasks()  # Update the GUI

    # Reset the progress bar after completion
    progress_bar['value'] = 0

    # Show a message box upon completion
    window.after(500, show_completion_message)  # Delayed message box prompt

def show_completion_message():
    messagebox.showinfo("Scan Complete", "Scan and shortcut creation complete!")

def create_shortcut(target_path, shortcut_path):
    from win32com.client import Dispatch

    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(shortcut_path)
    shortcut.Targetpath = target_path
    shortcut.save()

# Create the main Tkinter window
window = tk.Tk()
window.title("Executable Scanner")

# Set the window size
window.geometry("400x200")

# Create a button to select a folder for scanning
scan_button = tk.Button(window, text="Select Folder", command=scan_folder)
scan_button.pack(pady=10)

# Create a progress bar to track the scanning process
progress_bar = ttk.Progressbar(window, orient='horizontal', mode='determinate')
progress_bar.pack(pady=10)

# Start the Tkinter event loop
window.mainloop()