# Import the modules
import pandas as pd
import smtplib

# Set your email address and password
your_email = "mayanksharma61099@gmail.com"
your_password = "gqsa gftv olbj vvlm"

# Create a server object and log in
server = smtplib.SMTP("smtp.gmail.com", 587)
server.starttls()
server.login(your_email, your_password)

# Open the excel file and get the data frame
path = "ProdData.xlsx"
email_list = pd.read_excel(path)

# Get the name and email ids from the data frame
name_index = 5 # Change this to match your column index
email_index = 6 # Change this to match your column index
names = email_list.iloc[:, name_index]
emails = email_list.iloc[:, email_index]

# Loop through the names and email ids and send the email
for name, email in zip(names, emails):
    # Set the recipient, subject, and message 
    recipient = email
    subject = "Thank you for your response."
    message = f"""
Hello {name},

Thank you for giving your valuable inputs in Leveraging Customer Experience using AI with Embedded Finance form.

This will really help me analyzing user experience and study the market for Embedded Finance.

I appreciate your time and feedback.

Best regards,
Mayank Sharma
"""


    # Format the email
    email = f"Subject: {subject}\n\n{message}"

    # Send the email
    server.sendmail(your_email, recipient, email)

# Close the server
server.quit()