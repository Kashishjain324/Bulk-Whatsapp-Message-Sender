import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import pywhatkit as pwk
import threading

def browse_file():
    filepath = filedialog.askopenfilename(
        filetypes=[("Excel files", ".xlsx"), ("All files", ".*")]
    )
    file_entry.delete(0, tk.END)
    file_entry.insert(0, filepath)

def send_whatsapp_message(recipient_number, message_body):
    try:
        # Send WhatsApp message using pywhatkit
        pwk.sendwhatmsg_instantly(recipient_number, message_body)
    except Exception as e:
        print(f"Failed to send message to {recipient_number}. Error: {str(e)}")
def send_bulk_whatsapp():
    message_body = message_entry.get("1.0", tk.END).strip()
    excel_file = file_entry.get()

    if not excel_file:
        messagebox.showerror("Error", "Please select an Excel file.")
        return

    if not message_body:
        messagebox.showerror("Error", "Please enter a message to send.")
        return

    try:
        # Read the Excel file
        df = pd.read_excel(excel_file)

        if 'WhatsApp' not in df.columns:
            messagebox.showerror("Error", "'WhatsApp' column not found in Excel file.")
            return

        recipient_numbers = df['WhatsApp'].dropna().tolist()

        for recipient_number in recipient_numbers:
            recipient_number = str(recipient_number).strip()  # Convert to string and strip whitespace

            # Check if the number starts with '+' and add country code if missing
            if not recipient_number.startswith("+"):
                country_code = "+91"  # Replace '91' with your desired country code or make it dynamic
                recipient_number = f"{country_code}{recipient_number}"

            # Validate if the number is in the correct format
            if len(recipient_number) <= 10 or not recipient_number[1:].isdigit():
                messagebox.showwarning("Warning", f"Invalid number format: {recipient_number}")
                continue

            # Send each message in a separate thread
            threading.Thread(target=send_whatsapp_message, args=(recipient_number, message_body)).start()

        messagebox.showinfo("Success", "WhatsApp messages are being sent!")

    except Exception as e:
        messagebox.showerror("Error", f"Failed to send WhatsApp messages. Error: {str(e)}")
# GUI setup
root = tk.Tk()
root.title("Bulk WhatsApp Sender from Excel")

tk.Label(root, text="Message:").grid(row=0, column=0, padx=10, pady=10)
message_entry = tk.Text(root, width=50, height=10)
message_entry.grid(row=0, column=1, padx=10, pady=10)

tk.Label(root, text="Excel File:").grid(row=1, column=0, padx=10, pady=10)
file_entry = tk.Entry(root, width=50)

file_entry.grid(row=1, column=1, padx=10, pady=10)
browse_button = tk.Button(root, text="Browse", command=browse_file)
browse_button.grid(row=1, column=2, padx=10, pady=10)

send_button = tk.Button(root, text="Send WhatsApp Messages", command=send_bulk_whatsapp)
send_button.grid(row=2, column=1, padx=10, pady=20)

root.mainloop()