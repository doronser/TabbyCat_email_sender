# TabbyCat_email_sender
Quick and dirty python app to automate sending of TabbyCat private URLs.

1. Allow the script to send emails on your behalf:
https://myaccount.google.com/lesssecureapps

2. Update the Excel to match the following format:
  - column A = Participant Name
  - Column B = Email Address
  - Column C = TabbyCat Private URL
3. Update the first row to include:
  - your own user name and password (currently only supports gmail accounts)
  - the subject of the email you wish to send
4. Create the email message
  - Write the email message and add {0} where you want the URL to be added
  - convert message to HTML using this website (make sure to choose Use paragraph and line break tags(
  https://www.textfixer.com/html/convert-text-html.php
5. Download and run sendEmails.exe (from the same folder as the Excel):
https://drive.google.com/file/d/1-qYTm56sx94nAo0wvWbHPktq8bkEyhYR/view?usp=sharing
6. Profit :)
