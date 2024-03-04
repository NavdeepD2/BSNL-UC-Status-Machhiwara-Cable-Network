import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from datetime import datetime
import paramiko
import os

def merge_files():
    files = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
    if not files:
        return

    merged_df = pd.DataFrame()
    for file in files:
        df = pd.read_excel(file)
        merged_df = pd.concat([merged_df, df])

    # Apply condition to exclude rows
    merged_df = merged_df.loc[~(
            (merged_df['UC Received'] == 'Y') &
            ((merged_df['BBC Approved'] == 'Y') | (merged_df['BBC Approved'] == 'Pending'))
    )]

    # Get current date and time
    date_time = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")

    # Define the output file name
    output_file = f"BSNL-UCs-LDH-RPR-{date_time}.xlsx"

    # Write the merged dataframe to excel
    merged_df.to_excel(output_file, index=False)

    # Save a copy of the merged file locally
    local_copy = f"BSNL-UCs-LDH-RPR-LocalCopy-{date_time}.xlsx"
    merged_df.to_excel(local_copy, index=False)

    messagebox.showinfo("Merge Complete", "Merge is Done")

    # Upload file to remote server
    try:
        ssh = paramiko.SSHClient()
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        ssh.connect('1.2.3.4', port=22, username='root', password='PASSWORD_HERE')
        sftp = ssh.open_sftp()
        sftp.put(output_file, '/var/www/public_html/BSNL-UCs-LDH-RPR.xlsx')
        sftp.close()
        ssh.close()
        messagebox.showinfo("Upload Complete", "File upload successful")
    except Exception as e:
        messagebox.showerror("Upload Failed", f"File upload failed: {str(e)}")

    # Remove local merged file
    os.remove(output_file)

# Create Tkinter window
root = tk.Tk()
root.title("Merge Excel Files")

# Set window width and height
window_width = 400
window_height = 200

# Calculate x and y coordinates for the window to be centered
x = (root.winfo_screenwidth() // 2) - (window_width // 2)
y = (root.winfo_screenheight() // 2) - (window_height // 2)

# Set the position of the window
root.geometry(f'{window_width}x{window_height}+{x}+{y}')

# Create button to trigger file selection and merging
merge_button = tk.Button(root, text="Merge Files", command=merge_files, width=20, height=3)
merge_button.pack(pady=20)

root.mainloop()
