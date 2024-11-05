import os
import pandas as pd
import datetime
from plyer import notification
from tkinter import Tk, Label, Entry, Button, messagebox

# Path to your Excel file
file_path = "reminders.xlsx"

def load_or_create_excel():
    # Check if the file exists
    if os.path.exists(file_path):
        try:
            # Try reading the file
            return pd.read_excel(file_path, engine='openpyxl')
        except (ValueError, pd.errors.EmptyDataError, FileNotFoundError):
            # If there's an error, delete the file and recreate it
            os.remove(file_path)
    
    # If file doesn't exist or is corrupted, create a new DataFrame
    return pd.DataFrame(columns=["Date", "Time", "Message"])

def save_to_excel(df):
    try:
        df.to_excel(file_path, index=False, engine='openpyxl')
    except PermissionError:
        messagebox.showerror("File Error", f"Permission denied: Unable to write to '{file_path}'. Make sure the file is closed and accessible.")
        return False
    return True

def add_reminder():
    date = date_entry.get()
    time = time_entry.get()
    message = message_entry.get()

    if not date or not time or not message:
        messagebox.showwarning("Input Error", "All fields are required!")
        return
    
    try:
        # Check if date and time are in the correct format
        datetime.datetime.strptime(date, "%Y-%m-%d")
        datetime.datetime.strptime(time, "%H:%M:%S")
    except ValueError:
        messagebox.showerror("Format Error", "Date must be YYYY-MM-DD and Time HH:MM:SS")
        return

    # Load or create the DataFrame
    df = load_or_create_excel()
    
    # Add new reminder to the DataFrame
    new_reminder = pd.DataFrame([[date, time, message]], columns=["Date", "Time", "Message"])
    df = pd.concat([df, new_reminder], ignore_index=True)
    
    # Save to Excel and notify the user if successful
    if save_to_excel(df):
        messagebox.showinfo("Success", "Reminder added successfully!")
    
    # Clear the input fields
    date_entry.delete(0, 'end')
    time_entry.delete(0, 'end')
    message_entry.delete(0, 'end')
def check_and_notify():
    # Load or create the DataFrame
    df = load_or_create_excel()
    current_time = datetime.datetime.now()

    for index, row in df.iterrows():
        try:
            # Parse Date and Time into datetime objects
            reminder_date = pd.to_datetime(row['Date']).date()  # Convert to date
            reminder_time = pd.to_datetime(row['Time']).time()  # Convert to time
            
            # Combine date and time into a single datetime object
            reminder_datetime = datetime.datetime.combine(reminder_date, reminder_time)
        
            # Check if current time matches or is after the reminder time
            if current_time >= reminder_datetime:
                # Send notification
                notification.notify(
                    title="Reminder",
                    message=row['Message'],
                    timeout=10
                )
                # Remove the reminder after it's triggered
                df.drop(index, inplace=True)
        except Exception as e:
            print(f"Error processing row {index}: {e}")  # Print error message for debugging

    # Save the updated DataFrame back to Excel
    save_to_excel(df)
    
    # Schedule the next check in 60 seconds
    root.after(60000, check_and_notify)


# Set up the GUI
root = Tk()
root.title("Reminder App")

Label(root, text="Date (YYYY-MM-DD):").grid(row=0, column=0)
Label(root, text="Time (HH:MM:SS):").grid(row=1, column=0)
Label(root, text="Message:").grid(row=2, column=0)

date_entry = Entry(root)
date_entry.grid(row=0, column=1)
time_entry = Entry(root)
time_entry.grid(row=1, column=1)
message_entry = Entry(root)
message_entry.grid(row=2, column=1)

add_button = Button(root, text="Add Reminder", command=add_reminder)
add_button.grid(row=3, column=0, columnspan=2)

# Start checking reminders
root.after(60000, check_and_notify)
root.mainloop()
