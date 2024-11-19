import os
import pandas as pd
import win32com.client
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog

# Function to send emails using Outlook, with optional attachments
def send_email(to_address, cc_address, subject, body, attachments=None):
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = to_address
        mail.CC = cc_address
        mail.Subject = subject
        mail.Body = body
        
        # Add attachments if provided
        if attachments:
            for attachment in attachments:
                mail.Attachments.Add(attachment)
        
        mail.Display()  # Opens the email for review instead of sending
        print(f"Email opened for {to_address}")
    except Exception as e:
        print(f"Error opening email for {to_address}: {e}")

# Function to read Excel file into DataFrame
def read_excel(file_path):
    try:
        df = pd.read_excel(file_path)
        return df
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None

# Function to concatenate columns in the DataFrame
def concatenate_columns():
    if df is None:
        messagebox.showerror("Error", "Please load an Excel file first.")
        return

    columns = simpledialog.askstring("Columns", "Enter columns to concatenate (comma separated):")
    separator = simpledialog.askstring("Separator", "Enter the separator:")
    new_column_name = simpledialog.askstring("New Column Name", "Enter the name for the new column:")

    if columns and separator and new_column_name:
        columns_list = [col.strip() for col in columns.split(",")]
        
        # Check if all specified columns are in the DataFrame
        if all(col in df.columns for col in columns_list):
            df[new_column_name] = df[columns_list].astype(str).agg(separator.join, axis=1)
            messagebox.showinfo("Success", f"New column '{new_column_name}' created successfully.")
        else:
            messagebox.showerror("Error", "One or more specified columns do not exist in the DataFrame.")
    else:
        messagebox.showerror("Error", "Please fill all fields.")

# Function to generate and send personalized emails, with optional grouping and attachments
def generate_emails():
    template = email_body_text.get("1.0", tk.END).strip()
    subject_template = subject_entry.get()
    to_template = to_entry.get()
    cc_template = cc_entry.get()

    if df is None:
        messagebox.showerror("Error", "Please load an Excel file first.")
        return

    if not attachments_folder or not files_list:
        pass

    # Check if the user wants to group by a specific column
    if group_column:
        grouped = df.groupby(group_column)

        for group_name, group_df in grouped:
            try:
                first_row = group_df.iloc[0]
                email_subject = subject_template.format(**first_row.to_dict())
                email_to = to_template.format(**first_row.to_dict())
                email_cc = cc_template.format(**first_row.to_dict())

                # Create the email body for the group
                group_body = template

                # Collect values for each placeholder marked with %
                for col in group_df.columns:
                    if f"%{col}%" in group_body:
                        # Join values for this column separated by two line breaks
                        values = "\n\n".join(group_df[col].astype(str))
                        group_body = group_body.replace(f"%{col}%", values)

                # Replace {placeholders} in the email body
                group_body = group_body.format(**first_row.to_dict())

                # Extract keywords from the email subject
                keywords = email_subject.split()

                # Attach files whose names contain any of the keywords
                attachments = [f"{attachments_folder}/{file}" for file in files_list if any(keyword.lower() in file.lower() for keyword in keywords)]
                send_email(email_to, email_cc, email_subject, group_body, attachments)

            except Exception as e:
                print(f"Error generating email for group {group_name}: {e}")
                messagebox.showerror("Error", f"Error generating email for group {group_name}: {e}")
    else:
        # No grouping, send individual emails for each row
        for index, row in df.iterrows():
            try:
                email_body = template
                
                # Replace placeholders for this row
                for col in df.columns:
                    if f"%{col}%" in email_body:
                        email_body = email_body.replace(f"%{col}%", str(row[col]))
                
                # Replace {placeholders} in the email subject, To, and CC fields
                email_subject = subject_template.format(**row.to_dict())
                email_to = to_template.format(**row.to_dict())
                email_cc = cc_template.format(**row.to_dict())

                # Extract keywords from the email subject
                keywords = email_subject.split()

                # Attach files whose names contain any of the keywords
                attachments = [f"{attachments_folder}/{file}" for file in files_list if any(keyword.lower() in file.lower() for keyword in keywords)]
                
                # Send the email with the attachments
                send_email(email_to, email_cc, email_subject, email_body, attachments)

            except Exception as e:
                print(f"Error generating email for row {index}: {e}")
                messagebox.showerror("Error", f"Error generating email for row {index}: {e}")

# Function to load Excel file
def load_file():
    global df, group_column
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        df = read_excel(file_path)
        if df is not None:
            messagebox.showinfo("Success", "Excel file loaded successfully.")
            group_column = simpledialog.askstring("Group by Column", "Enter the column name to group by (or leave blank to skip):")
            if group_column and group_column not in df.columns:
                messagebox.showerror("Error", f"'{group_column}' is not a valid column. No grouping will be applied.")
                group_column = None
        else:
            messagebox.showerror("Error", "Failed to load Excel file.")

# Function to select the attachments folder
def select_attachments_folder():
    global attachments_folder, files_list
    attachments_folder = filedialog.askdirectory()
    if attachments_folder:
        try:
            files_list = os.listdir(attachments_folder)  # Get the list of files in the selected folder
            messagebox.showinfo("Success", f"Selected attachments folder: {attachments_folder}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not read attachments folder: {e}")

# Initialize global variables
df = None
group_column = None
attachments_folder = None
files_list = []

# Set up the GUI
root = tk.Tk()
root.title("Email Sender")

# Frame for email subject
subject_frame = tk.Frame(root)
subject_frame.pack(pady=10)

subject_label = tk.Label(subject_frame, text="Email Subject (Use {} for dynamic fields):")
subject_label.pack(side=tk.LEFT)

subject_entry = tk.Entry(subject_frame, width=50)
subject_entry.pack(side=tk.LEFT)

# Frame for email To field
to_frame = tk.Frame(root)
to_frame.pack(pady=10)

to_label = tk.Label(to_frame, text="To Email (Use {} for dynamic fields):")
to_label.pack(side=tk.LEFT)

to_entry = tk.Entry(to_frame, width=50)
to_entry.pack(side=tk.LEFT)

# Frame for email CC field
cc_frame = tk.Frame(root)
cc_frame.pack(pady=10)

cc_label = tk.Label(cc_frame, text="CC Email (Use {} for dynamic fields):")
cc_label.pack(side=tk.LEFT)

cc_entry = tk.Entry(cc_frame, width=50)
cc_entry.pack(side=tk.LEFT)

# Frame for email body template
body_frame = tk.Frame(root)
body_frame.pack(pady=10)

body_label = tk.Label(body_frame, text="Email Body Template (Use {}/%% for dynamic fields):")
body_label.pack()

email_body_text = tk.Text(body_frame, width=60, height=20)
email_body_text.pack()

# Example text template with placeholders
email_body_text.insert(tk.END, 
"""Hello {example_name},

Your invoice balance in the amount of {example_amount}, is due on {example_date}.

Here is a list of your artworks: 

%example_artworks%

Thank you very much, 

Jack 
"""
                       )

# Button to load Excel file
load_button = tk.Button(root, text="Load Excel File", command=load_file)
load_button.pack(pady=5)

# Button to select attachments folder
attachments_button = tk.Button(root, text="Select Attachments Folder", command=select_attachments_folder)
attachments_button.pack(pady=5)

# Button to generate and send emails
send_button = tk.Button(root, text="Send Emails", command=generate_emails)
send_button.pack(pady=10)

# Button to concatenate columns
concat_button = tk.Button(root, text="Concatenate Columns", command=concatenate_columns)
concat_button.pack(pady=10)

# Run the GUI
root.mainloop()