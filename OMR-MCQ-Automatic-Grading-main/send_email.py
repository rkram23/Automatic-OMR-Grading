import os
import pandas as pd
import smtplib
from email.message import EmailMessage
from getpass import getpass
import datetime
import time

# Input your sender email and password
sender_email = input("Enter Your Email ID: ")
sender_pass = getpass("Enter Your Password: ")

# Read data from the Excel file
excel_file = "E:\OMR Mcq Automatic Grading\grades.xlsx"
df = pd.read_excel(excel_file)

# Define the email subject
subject = "Your Test Score"

# Prompt the user to choose whether to send emails immediately or schedule them
while True:
    choice = input("Do you want to send emails immediately (I) or schedule them for later (S)? ").strip().lower()
    
    if choice == "i":
        # Send emails immediately
        for index, row in df.iterrows():
            name = row["Name"]
            email = row["Email"]
            score = row["Score"]
            
            msg = EmailMessage()
            msg['From'] = sender_email
            msg['To'] = email
            msg['Subject'] = subject

            # Compose the email message with the correct variable name "name"
            message_body = f"Hello {name},\n\nYour test score is: {score}%"
            msg.set_content(message_body)

            # Send the email immediately
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(sender_email, sender_pass)
                smtp.send_message(msg)
            print(f"Email to {name} sent immediately.")

        print("All emails sent immediately.")
        break
    elif choice == "s":
        # Prompt the user to specify the scheduled time for all students
        time_str = input("Enter the scheduled time (hh:mm AM/PM): ").strip()
        date_str = input("Enter the date (YYYY-MM-DD): ").strip()

        # Function to convert 12-hour time to 24-hour time
        def convert_to_24_hour(time_str):
            try:
                return datetime.datetime.strptime(time_str, "%I:%M %p").strftime("%H:%M")
            except ValueError:
                return None

        # Convert the specified time to 24-hour format and validate
        date_time_str_24_hour = convert_to_24_hour(time_str)
        if date_time_str_24_hour is None:
            print(f"Invalid 12-hour time format: {time_str}")
            continue

        # Combine date and time and convert to 24-hour format
        full_date_time_str = f"{date_str} {date_time_str_24_hour}"

        try:
            scheduled_time = datetime.datetime.strptime(full_date_time_str, '%Y-%m-%d %H:%M')
            current_time = datetime.datetime.now()
            
            if scheduled_time <= current_time:
                print("Invalid date and time. The specified time has already passed.")
                continue
        except ValueError:
            print(f"Invalid date and time format: {full_date_time_str}")
            continue

        # Iterate through each student's data and schedule the emails
        for index, row in df.iterrows():
            name = row["Name"]
            email = row["Email"]
            score = row["Score"]
            
            # Calculate the time delay until the scheduled time
            delta = scheduled_time - current_time
            seconds_until_send = delta.total_seconds()
            print(f"Email to {name} scheduled for {scheduled_time}.")

            # Sleep for the specified duration before sending the email
            time.sleep(seconds_until_send)

            # Send the email after the specified delay
            msg = EmailMessage()
            msg['From'] = sender_email
            msg['To'] = email
            msg['Subject'] = subject

            # Compose the email message with the correct variable name "name"
            message_body = f"Hello {name},\n\nYour test score is: {score}%"
            msg.set_content(message_body)

            # Send the email
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(sender_email, sender_pass)
                smtp.send_message(msg)
            print(f"Email to {name} sent as scheduled.")

        print("All emails scheduled and will be sent at the specified time.")
        break
    else:
        print("Invalid choice. Please enter 'I' for immediate send or 'S' for scheduling.")
