import pandas as pd
from fpdf import FPDF
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import os

# Your Gmail details
EMAIL_USER = 'ekagonah@gmail.com'
EMAIL_PASS = 'borl vjyk zsrd muzq'  # App Password from Google (if you have 2-Step Verification enabled)

file_path = "employees.xlsx" 
df = pd.read_excel(file_path)
print(df)

# Create a folder for the payslips if it doesn't exist
if not os.path.exists('payslips'):
    os.makedirs('payslips')

for index, row in df.iterrows():
    emp_id = row['Employee ID']
    name = row['Name']
    email = row['Email']
    basic_salary = row['Basic Salary']
    allowances = row['Allowances']
    deductions = row['Deductions']
    net_salary = basic_salary + allowances - deductions

for index, row in df.iterrows():
    print(f"Employee ID: {row['Employee ID']}")
    print(f"Name: {row['Name']}")
    print(f"Email: {row['Email']}")
    print(f"Basic Salary: {row['Basic Salary']}")
    print(f"Allowances: {row['Allowances']}")
    print(f"Deductions: {row['Deductions']}")
    print("-" * 30)
    
    print(f"Net Salary: {net_salary}")
    print("=" * 30)

    # Create the PDF payslip
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    pdf.cell(200, 10, txt="Payslip", ln=True, align='C')
    pdf.cell(200, 10, txt=f"Employee ID: {row['Employee ID']}", ln=True)
    pdf.cell(200, 10, txt=f"Name: {row['Name']}", ln=True)
    pdf.cell(200, 10, txt=f"Email: {row['Email']}", ln=True)

    # Save the PDF to the 'payslips' folder
    filename = f"payslips/{emp_id}.pdf"
    pdf.output(filename)

    # Prepare the email
    msg = MIMEMultipart()
    msg['From'] = EMAIL_USER
    msg['To'] = email
    msg['Subject'] = 'Your Payslip for This Month'

    body = f"Hello {name},\n\nPlease find your payslip attached.\n\nBest regards,\nHR Team"
    msg.attach(MIMEText(body, 'plain'))

    # Attach the payslip PDF
    with open(filename, 'rb') as attachment:
        part = MIMEApplication(attachment.read(), _subtype="pdf")
        part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(filename))
        msg.attach(part)

    # Send the email using Gmail's SMTP server
    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()  # Start TLS encryption
            server.login(EMAIL_USER, EMAIL_PASS)  # Log in with your email and app password
            server.sendmail(EMAIL_USER, email, msg.as_string())  # Send email
            print(f"Payslip sent to {email} successfully!")
    except Exception as e:
        print(f"Error sending email to {email}: {e}")

print("Payslips generated and emailed successfully!")
