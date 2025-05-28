import pandas as pd
import os as os

# ANSI escape sequences for colored output
bold = "\033[1m"
reset = "\033[0m"

# Define the path to the Excel file
file_path = os.path.join(os.getcwd(), 'data.xlsx')

# Check if the file exists
if not os.path.exists(file_path):
    raise FileNotFoundError(f"The file {file_path} does not exist.")


# Read the Excel file
file = pd.read_excel(file_path, sheet_name=None)


for sheet_name, df in file.items():
    df.columns = df.columns.str.strip()  # Remove leading/trailing spaces
    if 'Quantidade' in df.columns and 'Valor Unitário (R$)' in df.columns:
        df['Quantidade'] = pd.to_numeric(df['Quantidade'], errors='coerce')
        df['Valor Unitário (R$)'] = pd.to_numeric(df['Valor Unitário (R$)'], errors='coerce')
        df['Valor Total'] = df['Quantidade'] * df['Valor Unitário (R$)']

output_file_path = os.path.join(os.getcwd(), 'data_modified.xlsx')

try:
    if os.path.exists(output_file_path):
        print(f"{bold}\nbRemoving existing file:{reset} {output_file_path}\n")
        os.remove(output_file_path)
    # Save the modified data to a new Excel file
    with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
        for sheet_name, df in file.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
except Exception as e: 
    print(f"An error occurred while saving the file: {e}")

print(f"{bold}Modified data saved to{reset} {output_file_path}\n")

import smtplib
from email.message import EmailMessage

from dotenv import load_dotenv
load_dotenv()
import os

EMAIL_ADDRESS = os.getenv('EMAIL_ADDRESS')
EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD')
TO_ADDRESS = os.getenv('TO_ADDRESS')

# Create the email
msg = EmailMessage()
msg['Subject'] = 'Arquivo Excel Formatado'
msg['From'] = EMAIL_ADDRESS
msg['To'] = TO_ADDRESS
msg.set_content('Segue em anexo o arquivo Excel formatado.')

# Attach the Excel file
with open(output_file_path, 'rb') as f:
    file_data = f.read()
    file_name = os.path.basename(output_file_path)
    msg.add_attachment(file_data, maintype='application', subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=file_name)

# Send the email
with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
    smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
    smtp.send_message(msg)

print(f"{bold}Email sent successfully!\n")