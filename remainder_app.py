import pandas as pd
import datetime
import time
from plyer import notification

# Path to your Excel file
file_path = "reminders.xlsx"

def check_and_notify():
    # Load the Excel file
    df = pd.read_excel(file_path)
    
    # Get the current date and time
    current_time = datetime.datetime.now()
    
    # Iterate through each row in the Excel sheet
    for index, row in df.iterrows():
        # Get reminder time from Excel
        reminder_time = datetime.datetime.combine(row['Date'], row['Time'])
        
        # Check if the current time matches the reminder time
        if current_time >= reminder_time:
            # Send notification
            notification.notify(
                title="Reminder",
                message=row['Message'],
                timeout=10  # Duration in seconds
            )
            # Optionally remove or update the row after the reminder is triggered
            df.drop(index, inplace=True)
    
    # Save updated data back to Excel
    df.to_excel(file_path, index=False)

# Run the application in intervals to keep checking
while True:
    check_and_notify()
    # Check every minute
    time.sleep(60)
