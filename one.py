import xlrd
xlrd.__version__
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Path to the Excel file containing client information
path = "helooooo.xlsx"
openFile = xlrd.open_workbook(path)
sheet = openFile.sheet_by_name("helooooo")

# Lists to store email addresses, amounts, and client names
mail_list = []
name = []

# Extract relevant data from the Excel sheet
for k in range(sheet.nrows - 1):
    client = sheet.cell_value(k + 1, 0)
    email = sheet.cell_value(k + 1, 1)
    mail_list.append(email)
    name.append(client)

# Gmail credentials
email = "fyivitc@gmail.com"  # Your Gmail address
password = "FYIJOD@VITC"  # Your Gmail password

# Connect to Gmail server
server = smtplib.SMTP("smtp.gmail.com", 587)
server.starttls()
server.login(email, password)

# Send personalized emails to clients who haven't paid
for mail_to in mail_list:
    find_des = mail_list.index(mail_to)
    clientName = name[find_des]
    subject = f"{clientName}, you have a new email"
    message = f"Dear {clientName},\n\n" \
              f"We inform you that you owe ${amount[find_des]}.\n\n" \
              "Best Regards"

    msg = MIMEMultipart()
    msg["From"] = email
    msg["Subject"] = subject
    msg.attach(MIMEText(message, "plain"))

    word_file_path = "helooooo.docx"

    # Attach the Word document
    with open(word_file_path, "rb") as file:
        attachment = MIMEApplication(file.read(), Name=os.path.basename(word_file_path))
        attachment["Content-Disposition"] = f"attachment; filename={os.path.basename(word_file_path)}"
        msg.attach(attachment)

    print(f"Sending email to {clientName}...")
    server.sendmail(email, mail_to, msg.as_string())

# Close the server connection
server.quit()
print("Process is finished!")
