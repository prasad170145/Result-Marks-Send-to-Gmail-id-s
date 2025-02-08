import os
import pandas as pd
import smtplib
from email.message import EmailMessage
import csv
from twilio.rest import Client

# Read the CSV file from the same directory
data = pd.read_csv('marks.csv')  # Ensure file exists in the same directory

data.columns = data.columns.str.strip()  # Remove leading and trailing spaces

# Email configuration
smtp_server = "smtp.gmail.com"
smtp_port = 587
smtp_user = "npdevandla@gmail.com"  # Replace with your actual email
smtp_pass = "zdqp jsda gxvj lbgn"  # Use environment variable for security

# Create the SMTP connection
server = smtplib.SMTP(smtp_server, smtp_port)
server.starttls()
server.login(smtp_user, smtp_pass)

# Iterate over each row in the CSV and send emails
for index, row in data.iterrows():
    try:
        to_email = row['email']
        subject = "Your Results"

        # Determine overall result
        marks = [row['acd'], row['dccn'], row['ml'], row['bc'], row['eh'], row['acd lab'], row['dccn lab'], row['ml lab'], row['mini project']]
        overall_result = "Fail" if any(m < 14 for m in marks) else "Pass"

        # Calculate percentage and rating
        percentage = (row['total'] / 360) * 100
        if overall_result == "Fail":
            rating = "*"
        elif percentage > 90:
            rating = "*"
        elif percentage > 80:
            rating = ""
        elif percentage > 70:
            rating = "*"
        elif percentage > 60:
            rating = ""
        else:
            rating = "*"

        # Create the email body
        body = f"""
        <html>
        <head>
            <style>
                table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
                th, td {{ border: 1px solid #ddd; padding: 10px; text-align: center; }}
                th {{ background-color: #007bff; color: white; }}
                tr.low-marks {{ background-color: #f8d7da; color: #721c24; }} /* Red row for marks < 14 */
                tr.high-marks {{ background-color: #d4edda; color: #155724; }} /* Green row for marks >= 14 */
                p {{ font-size: 16px; }}
            </style>
        </head>
        <body>
            <div class="container">
                <p>Dear Student,</p>
                <p><i>MID {row['mid']} marks are as follows:</i></p>
                <p><b>Name:</b> {row['name']}</p>
                <p><b>Rollno:</b> {row['rollno']}</p>
                <p><b>Year/Semester:</b> {row['year/semester']}</p>
                <p><b>Branch:</b> {row['branch']}</p>

                <table>
                    <tr>
                        <th>Subject</th>
                        <th>Marks</th>
                    </tr>
                    <tr class="{ 'low-marks' if row['acd'] < 14 else 'high-marks' }"><td>ACD</td><td>{row['acd']}</td></tr>
                    <tr class="{ 'low-marks' if row['dccn'] < 14 else 'high-marks' }"><td>DCCN</td><td>{row['dccn']}</td></tr>
                    <tr class="{ 'low-marks' if row['ml'] < 14 else 'high-marks' }"><td>ML</td><td>{row['ml']}</td></tr>
                    <tr class="{ 'low-marks' if row['bc'] < 14 else 'high-marks' }"><td>BC</td><td>{row['bc']}</td></tr>
                    <tr class="{ 'low-marks' if row['eh'] < 14 else 'high-marks' }"><td>EH</td><td>{row['eh']}</td></tr>
                    <tr class="{ 'low-marks' if row['acd lab'] < 14 else 'high-marks' }"><td>ACD LAB</td><td>{row['acd lab']}</td></tr>
                    <tr class="{ 'low-marks' if row['dccn lab'] < 14 else 'high-marks' }"><td>DCCN LAB</td><td>{row['dccn lab']}</td></tr>
                    <tr class="{ 'low-marks' if row['ml lab'] < 14 else 'high-marks' }"><td>ML LAB</td><td>{row['ml lab']}</td></tr>
                    <tr class="{ 'low-marks' if row['mini project'] < 14 else 'high-marks' }"><td>MINI PROJECT</td><td>{row['mini project']}</td></tr>
                    <tr><td>TOTAL</td><td>{row['total']}/360</td></tr>
                </table>
                
                <p><b>Overall Result:</b> {overall_result}</p>
                <p><b>Rating:</b> {rating}</p>
                <p>Best regards,<br>VBIT</p>
            </div>
        </body>
        </html>
        """

        # Create the email message
        msg = EmailMessage()
        msg['From'] = smtp_user
        msg['To'] = to_email
        msg['Subject'] = subject
        msg.set_content("This email contains HTML content. Please enable HTML viewing.")
        msg.add_alternative(body, subtype='html')

        # Send the email
        server.send_message(msg)
        print(f"Email sent to {to_email}")
    except Exception as e:
        print(f"Failed to send email to {row.get('email', 'Unknown')} due to: {e}")

# Logout and close the connection
server.quit()