from flask import Flask, render_template
import pandas as pd
import datetime
import smtplib
import os
import micropip
app = Flask(__name__)

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

@app.route('/')
def index():
    return render_template("index.html")

@app.route('/send')
def emailsend():
    #your gmail credentials here
    GMAIL_ID = ' yourmail@gmail.com'

    GMAIL_PWD = 'mail_generated_pass' 


    if __name__ == "__main__":

        df = pd.read_excel("data.xlsx")
        today = datetime.datetime.now().strftime("%d-%m")
        yearNow = datetime.datetime.now().strftime("%Y")
        writeInd = []

        for index, item in df.iterrows():
            bday = item['Birthday'].strftime("%d-%m")
            #If Excel date matches with todays date then follows this code
            #IF we want to send email only once in day irespective of number of clicks in sigle day then use this line
            # if(today == bday) and yearNow not in str(item['Year']):
            if(today == bday):

                msg = MIMEMultipart()
                # set message content
                message = "D.Y. Patil wishes you a great birhday and memorable Year. From all of us."
                msg.attach(MIMEText(message))
                
                # set image file path
                image_path = 'D:\KOKP\Programs\Flask (Working)\HBD.png'

                with open(image_path, 'rb') as f:
                    img_data = f.read()
                image = MIMEImage(img_data, name=os.path.basename(image_path))
                msg.attach(image)

                # set email subject and recipient
                subject = "Happy Birthday!"
                recipient_email = item['Email']
                msg['Subject'] = subject
                msg['From'] = GMAIL_ID
                msg['To'] = recipient_email
                
                # send email
                s = smtplib.SMTP('smtp.gmail.com', 587)
                s.starttls()
                s.login(GMAIL_ID, GMAIL_PWD)
                s.sendmail(GMAIL_ID, recipient_email, msg.as_string())
                s.quit()
                print('First Mail send')
        
        for i in writeInd:
            yr = df.loc[i, 'Year']
            df.loc[i, 'Year'] = str(yr) + ', ' + str(yearNow)

        df.to_excel('data.xlsx', index=False) 
        return render_template("confirmation.html")
    
if __name__ == "__main__":
    app.run(debug=True)
