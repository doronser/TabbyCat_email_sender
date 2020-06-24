import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


# Read the Excel with all the data
email_list = pd.read_excel("emails.xlsx")

###Change these to your email user-name and password
### IMPORTANT! if you're using gmail, you must first allow access by turning ON the button here:
### https://myaccount.google.com/lesssecureapps
your_email = email_list['Username'][0]
your_password = email_list['Password'][0]
html_temp = email_list['Message'][0]
subject =  email_list['Subject'][0]

try:
    server = smtplib.SMTP('smtp.gmail.com:587')
    server.ehlo()
    server.starttls()
    server.login(your_email, your_password)
except Exception as e:
    print("could not connect to email server :(")
    print(e)
    exit(0)


# Get all the Names, Email Addreses, Subjects and Messages
all_names = email_list['Name']
all_emails = email_list['Email']
all_urls = email_list['URL']

message = MIMEMultipart("alternative")
message["Subject"] = subject
message["From"] = your_email


### Create an HTML version of your message:
# 1. in the message itself write {0} where you want the tabbycat link
# 2. In order to convert your email to HTML format, use this website:
#       https://www.textfixer.com/html/convert-text-html.php
#       make sure to choose "Use paragraph and line break tags"
text = ""
part1 = MIMEText(text, "plain")

# Loop through the emails
for idx in range(len(all_emails)):

    # Get each records name, email, subject and message
    name = all_names[idx]
    email = all_emails[idx]
    url = all_urls[idx]

    html = html_temp.format(url)
    part2 = MIMEText(html, "html")

    message["To"] = email
    message.attach(part1)
    message.attach(part2)

    try:
        # server.sendmail(your_email, [email], full_email)
        server.sendmail(your_email, [email], message.as_string() )
        print('Email to {} successfully sent!'.format(email))
    except Exception as e:
        print('Email to {} could not be sent :( because {}\n'.format(email, str(e)))

# Close the smtp server
server.close()