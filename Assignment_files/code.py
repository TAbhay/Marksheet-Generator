import smtplib
from email.message import EmailMessage
msg = EmailMessage()
msg['Subject'] = "Mail"
msg['To'] = 'victorlawrence61@gmail.com'
msg.set_content("Hii \n\n This is a test mail")
with open ("Output.xlsx ", "rb") as f:
   file_data=f.read ()
   print("File data in binary", file_data)
   file_name=f.name
   print ("File name is", file_name)
   msg.add_attachment(file_data, maintype="application", subtype= "xlsx", filename=file_name)
with smtplib.SMTP_SSL('smtp.gmail.com',465) as server:
    server.login("victorratreanurag@gmail.com","victor1234569876")
    server.send_message(msg)
    server.quit()



