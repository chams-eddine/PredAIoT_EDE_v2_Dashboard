import pandas as pd
import matplotlib.pyplot as plt
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import sys
sys.path.append('../ede_v1')  # Import EDE
from ede_core import make_decision

# Paths
DATA_DIR = '../data/'
EXCEL_FILE = os.path.join(DATA_DIR, 'Yield report_Overall report_2015-2025.xlsx')
CREDENTIALS_FILE = os.path.join(DATA_DIR, 'predaiot-oman_credentials.csv')
CSV_OUTPUT = os.path.join(DATA_DIR, 'comparison_report.csv')
PNG_OUTPUT = os.path.join(DATA_DIR, 'yield_comparison.png')

# Step 1: Read Excel (Sheet "Yield report")
df = pd.read_excel(EXCEL_FILE, sheet_name='Yield report', engine='openpyxl', skiprows=1)  # Skip header row if needed
df = df.dropna()  # Clean data

# Step 2: Simulate differences
# Without PredAIoT: Raw data
without_yield = df['Total yield(kWh)']
without_revenue = df['Total revenue'].str.extract('(\d+\.?\d*)').astype(float)[0]  # Extract numeric part

# With PredAIoT: Run EDE per plant (assume maintenance_cost = 10% of revenue, loss_without = 20% of revenue; adjust as needed)
# Boost yield/revenue by 15% based on EDE savings
with_yield = []
with_revenue = []
for index, row in df.iterrows():
    data = {
        "maintenance_cost": row['Total revenue'].str.extract('(\d+\.?\d*)').astype(float)[0] * 0.1,  # Assume 10% cost
        "financial_loss_without": row['Total revenue'].str.extract('(\d+\.?\d*)').astype(float)[0] * 0.2  # Assume 20% loss without
    }
    decision = make_decision(data)
    boost_factor = 1 + (decision['savings'] / data['financial_loss_without']) * 0.15 if decision['savings'] > 0 else 1
    with_yield.append(row['Total yield(kWh)'] * boost_factor)
    with_revenue.append(without_revenue[index] * boost_factor)

df['Without_Yield'] = without_yield
df['With_Yield'] = with_yield
df['Difference_Yield'] = df['With_Yield'] - df['Without_Yield']
df['Without_Revenue'] = without_revenue
df['With_Revenue'] = with_revenue
df['Difference_Revenue'] = df['With_Revenue'] - df['Without_Revenue']

# Step 3: Generate CSV
df[['Plant name', 'Without_Yield', 'With_Yield', 'Difference_Yield', 'Without_Revenue', 'With_Revenue', 'Difference_Revenue']].to_csv(CSV_OUTPUT, index=False)

# Step 4: Generate Chart PNG (Bar chart)
plants = df['Plant name']
fig, ax = plt.subplots(figsize=(10, 6))
ax.bar(plants, df['Difference_Yield'], label='Yield Difference (kWh)')
ax.bar(plants, df['Difference_Revenue'], label='Revenue Difference (OMR)', alpha=0.7)
ax.set_xlabel('Plant')
ax.set_ylabel('Difference')
ax.set_title('Difference With vs Without PredAIoT')
ax.legend()
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig(PNG_OUTPUT)

# Step 5: Send Email
# Read credentials
creds = pd.read_csv(CREDENTIALS_FILE)
smtp_user = creds['Nom d\'utilisateur SMTP'].iloc[0]
smtp_pass = creds['Mot de passe SMTP'].iloc[0]

sender = 'al.shams.invest@gmail.com'
recipients = ['chamsvanbuurn@gmail.com', 'prediaotworks@gmail.com']
subject = 'PredAIoT Yield Comparison Report'
body = 'Attached: Comparison CSV and Chart PNG showing differences with/without PredAIoT.'

msg = MIMEMultipart()
msg['From'] = sender
msg['To'] = ', '.join(recipients)
msg['Subject'] = subject
msg.attach(MIMEText(body, 'plain'))

# Attach CSV
with open(CSV_OUTPUT, 'rb') as f:
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(CSV_OUTPUT)}')
    msg.attach(part)

# Attach PNG
with open(PNG_OUTPUT, 'rb') as f:
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(PNG_OUTPUT)}')
    msg.attach(part)

# Send via AWS SES SMTP
server = smtplib.SMTP('email-smtp.us-east-1.amazonaws.com', 587)  # AWS SES endpoint (adjust region if needed)
server.starttls()
server.login(smtp_user, smtp_pass)
server.sendmail(sender, recipients, msg.as_string())
server.quit()

print("Email sent successfully!")