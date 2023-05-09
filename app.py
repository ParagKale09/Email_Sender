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
    GMAIL_ID = ' paragkale49@gmail.com'
    GMAIL_PWD = 'uhpwzibvihaezxbs'


    # def sendEmail(to, sub, msg):
    #     print(f"Email to {to} sent with subject: {sub} and message {msg}" )
    #     s = smtplib.SMTP('smtp.gmail.com', 587)
    #     s.starttls()
    #     s.login(GMAIL_ID, GMAIL_PWD)
    #     s.sendmail(GMAIL_ID, to, f"Subject: {sub}\n\n{msg}")
    #     s.quit()

    if __name__ == "__main__":
        #just for testing
        # sendEmail(GMAIL_ID, "subject", "test message")
        # exit()

        df = pd.read_excel("data.xlsx")
        #print(df)
        today = datetime.datetime.now().strftime("%d-%m")
        yearNow = datetime.datetime.now().strftime("%Y")
        #print(type(today))
        writeInd = []

        for index, item in df.iterrows():
            #print(index, item['Birthday'])
            bday = item['Birthday'].strftime("%d-%m")
            if(today == bday):

                msg = MIMEMultipart()
        
                # set message content
                message = "D.Y. Patil wishes you a great birhday and memorable Year. From all of us."
                msg.attach(MIMEText(message))
                
                # set image file path
                image_path = 'D:\KOKP\Programs\Flask\HBD.png'

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

                # sendEmail(item['Email'], "Happy Birthday", item['Dialogue'])
                # writeInd.append(index)
        
        for i in writeInd:
            yr = df.loc[i, 'Year']
            df.loc[i, 'Year'] = str(yr) + ', ' + str(yearNow)

        df.to_excel('data.xlsx', index=False) 
        return render_template("confirmation.html")
    
if __name__ == "__main__":
    app.run(debug=False,host='0.0.0.0')
