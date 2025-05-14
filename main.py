from flask import Flask, render_template, request, redirect, url_for, flash
import smtplib
from email.mime.text import MIMEText
import pandas as pd
import os
from datetime import datetime, timedelta
from apscheduler.schedulers.background import BackgroundScheduler
import time

app = Flask(__name__)
app.secret_key = "API_KEY"

EMAIL = "sreenandansivadas@gmail.com"
APP_PASSWORD = "APP_PASSWORD"
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'csv'}

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Initialize APScheduler
scheduler = BackgroundScheduler()
scheduler.start()

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Function to send emails
def send_scheduled_email(email, mail_title, mail_body):
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(EMAIL, APP_PASSWORD)
            msg = MIMEText(mail_body)
            msg['Subject'] = mail_title
            msg['From'] = EMAIL
            msg['To'] = email
            smtp.sendmail(EMAIL, email, msg.as_string())
        print(f"Email sent to {email}")
    except Exception as e:
        print(f"Failed to send email to {email}: {str(e)}")

@app.route('/')
def home():
    return render_template("home.html")

@app.route('/workspace')
def workspace():
    return render_template("workspace.html")

@app.route('/deadline')
def deadline():
    return render_template("deadline.html")

@app.route('/send_mail', methods=['POST'])
def send_mail():
    mail_title = request.form.get('mail_title')
    mail_body = request.form.get('mail_body')
    file = request.files.get('data_sheet')

    if not mail_title or not mail_body or not file:
        flash("All fields and file upload are required.")
        return redirect(url_for('workspace'))

    if not allowed_file(file.filename):
        flash("Only .xlsx or .csv files are allowed.")
        return redirect(url_for('workspace'))

    filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(filepath)

    if file.filename.endswith('.xlsx'):
        df = pd.read_excel(filepath)
    else:
        df = pd.read_csv(filepath)

    if 'Email' not in df.columns:
        flash("The uploaded file must contain an 'Email' column.")
        return redirect(url_for('workspace'))

    for _, row in df.iterrows():
        email = row['Email']

        if pd.isnull(email):
            continue

        # Send email immediately
        send_scheduled_email(email, mail_title, mail_body)
        flash(f"Email sent to {email}")

    return redirect(url_for('workspace'))


@app.route('/send_deadline_notifications', methods=['POST'])
def send_deadline_notifications():
    mail_title = request.form.get('mail_title')
    mail_body = request.form.get('mail_body')
    file = request.files.get('data_sheet')

    if not mail_title or not mail_body or not file:
        flash("All fields and file upload are required.")
        return redirect(url_for('deadline'))

    if not allowed_file(file.filename):
        flash("Only .xlsx or .csv files are allowed.")
        return redirect(url_for('deadline'))

    filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(filepath)

    if file.filename.endswith('.xlsx'):
        df = pd.read_excel(filepath)
    else:
        df = pd.read_csv(filepath)

    if 'Email' not in df.columns or 'Deadline date' not in df.columns or 'Deadline time' not in df.columns:
        flash("The uploaded file must contain 'Email', 'Deadline date', and 'Deadline time' columns.")
        return redirect(url_for('deadline'))

    try:
        # Process the 'Deadline date' and 'Deadline time' columns
        df['Deadline date'] = pd.to_datetime(df['Deadline date'], errors='coerce')
        df['Deadline time'] = pd.to_datetime(df['Deadline time'], format='%H:%M:%S', errors='coerce').dt.time

        # Replace missing or invalid times with a default time (e.g., 23:59:59)
        df['Deadline time'].fillna(datetime.strptime('23:59:59', '%H:%M:%S').time(), inplace=True)

        # Combine 'Deadline date' and 'Deadline time' into a single 'Deadline' column
        df['Deadline'] = df.apply(lambda row: datetime.combine(row['Deadline date'], row['Deadline time']), axis=1)
    except Exception as e:
        flash(f"Failed to process Deadline dates and times: {str(e)}")
        return redirect(url_for('deadline'))

    current_time = datetime.now()

    for _, row in df.iterrows():
        deadline = row['Deadline']
        email = row['Email']

        if pd.isnull(deadline) or pd.isnull(email):
            continue

        # Calculate the time to send the email (1 minute before the deadline)
        send_time = deadline - timedelta(minutes=1)
        if send_time > current_time:
            # Schedule the email to be sent at the calculated time
            scheduler.add_job(send_scheduled_email, 'date', run_date=send_time, args=[email, mail_title, mail_body])
            flash(f"Email scheduled for {email} at {send_time}")
        else:
            flash(f"Skipped scheduling for {email}, deadline has passed or is less than 1 minute away.")

    return redirect(url_for('deadline'))


@app.route('/birthday')
def birthday():
    return render_template("birthday.html")

@app.route('/send_birthday_wishes', methods=['POST'])
def send_birthday_wishes():
    # Retrieve the form data
    mail_title = request.form.get('mail_title')
    mail_body = request.form.get('mail_body')
    file = request.files.get('data_sheet')

    if not mail_title or not mail_body or not file:
        flash("All fields and file upload are required.")
        return redirect(url_for('birthday'))

    if not allowed_file(file.filename):
        flash("Only .xlsx or .csv files are allowed.")
        return redirect(url_for('birthday'))

    # Save the uploaded file
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(filepath)

    # Read the file (xlsx or csv)
    if file.filename.endswith('.xlsx'):
        df = pd.read_excel(filepath)
    else:
        df = pd.read_csv(filepath)

    # Check if the file has the required 'Email', 'Birthday', and 'Name' columns
    if 'Email' not in df.columns or 'Birthday' not in df.columns or 'Name' not in df.columns:
        flash("The uploaded file must contain 'Email', 'Birthday', and 'Name' columns.")
        return redirect(url_for('birthday'))

    # Convert the 'Birthday' column to datetime, handling Excel date formats and string dates
    try:
        df['Birthday'] = pd.to_datetime(df['Birthday'], errors='coerce')
    except Exception as e:
        flash(f"Failed to process Birthday dates: {str(e)}")
        return redirect(url_for('birthday'))

    # Get the current date
    current_date = datetime.now()

    # Filter emails where the birthday matches the current date (month and day)
    birthday_df = df[(df['Birthday'].dt.month == current_date.month) & (df['Birthday'].dt.day == current_date.day)]

    if birthday_df.empty:
        flash("No birthdays today.")
        return redirect(url_for('birthday'))

    # Extract the emails and names
    email_list = birthday_df[['Email', 'Name']].dropna()  # Drop any missing values

    # Try sending the email to each address individually
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(EMAIL, APP_PASSWORD)
            for index, row in email_list.iterrows():
                email = row['Email']
                name = row['Name']

                # Replace '{Name}' or '{{Name}}' in the mail body with the actual name
                personalized_body = mail_body.replace("{ Name }", name)

                # Create a new message for each recipient
                msg = MIMEText(personalized_body)
                msg['Subject'] = mail_title
                msg['From'] = EMAIL
                msg['To'] = email  # Set the 'To' field only once

                smtp.sendmail(EMAIL, email, msg.as_string())
                time.sleep(2)  # Introduce a 2-second delay between emails to avoid spam flags
        flash("Birthday wishes sent successfully.")
    except Exception as e:
        flash(f"Failed to send birthday wishes: {str(e)}")
    
    return redirect(url_for('birthday'))

if __name__ == "__main__":
    app.run(debug=True)
