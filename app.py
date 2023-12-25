from flask import Flask, render_template, request, send_file, flash

import pandas as pd
import datetime
import smtplib
import os
import micropip
import uuid

from werkzeug.utils import secure_filename
app = Flask(__name__)

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

#Added
# Specify the upload folder and allowed extensions
UPLOAD_FOLDER = ''
ALLOWED_EXTENSIONS = {'xlsx', 'xls','png'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

new_filename = str(uuid.uuid4()) + '.xlsx'
new_imgname = str(uuid.uuid4()) + '.png'

# Check if the uploaded file has a valid extension
def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS
#end added

@app.route('/')
def index():
    return render_template("index.html")

#added
@app.route('/upload', methods=['POST'])
def upload_file():
    
    global users_message
    users_message = request.form.get('message')  # Get user message

    # # Check if a file is provided in the request # #
    # if 'file' not in request.files:
    #     return "No file part"
    
    file = request.files['file']
    file_image = request.files['image']
    
    # # Check if the user submitted an empty file # #
    if file.filename == '':
        return render_template("index.html")
    
    # Check if the file has a valid extension
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        image_filename = secure_filename(file_image.filename)
        #new_filename = str(uuid.uuid4()) + '_' + filename  # Generate a new filename
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], new_filename)
        image_filepath = os.path.join(app.config['UPLOAD_FOLDER'], new_imgname)
        
        # Save the uploaded file to the upload folder
        file.save(filepath)
        file_image.save(image_filepath)
        
        return render_template("sendmail.html")
    
    

@app.route('/send')
def emailsend():
    #your gmail credentials here
    GMAIL_ID = 'your_mail'
    #GMAIL_PWD = 'uhpwzibvihaezxbs'
    GMAIL_PWD = 'your_password' 


    if __name__ == "__main__":

        #first Change
        df = pd.read_excel(new_filename)
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
                message = users_message
                msg.attach(MIMEText(message))
                
                # set image file path
                image_path = new_imgname

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

        df.to_excel(new_filename, index=False) 
        return render_template("confirmation.html")


    
if __name__ == "__main__":
    app.run(debug=True)
