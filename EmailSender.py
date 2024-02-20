# Import the modules
import pandas as pd
import smtplib

# Set your email address and password
your_email = "your email"
your_password = "app password"

# Create a server object and log in
server = smtplib.SMTP("smtp.gmail.com", 587)
server.starttls()
server.login(your_email, your_password)

# Open the excel file and get the data frame
path = "TestData.xlsx"
email_list = pd.read_excel(path)

# Get the name and email ids from the data frame
name_index = 1 # Change this to match your column index
email_index = 2 # Change this to match your column index
names = email_list.iloc[:, name_index]
emails = email_list.iloc[:, email_index]

# Loop through the names and email ids and send the email
for name, email in zip(names, emails):
    # Set the recipient, subject, and message 
    recipient = email
    subject = "Test Subject"
    message = f"""
Hello {name},

Here comes the test content
"""


    # Format the email
    email = f"Subject: {subject}\n\n{message}"

    # Send the email
    server.sendmail(your_email, recipient, email)

# Close the server
server.quit()
