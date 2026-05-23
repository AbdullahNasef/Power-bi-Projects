import smtplib 
import time 
from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText 
import pandas as pd 

# 1...Server setup and email data
SMTP_SERVER = "smtp.gmail.com"  
SMTP_PORT = 587
SENDER_EMAIL = "@gmail.com" 
SENDER_PASSWORD = ""  

# 2.....Read data file

file_path = "real_em.xlsx"
df = pd.read_excel(file_path)

# 3...email server connection
try:
    server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    server.starttls()  
    server.login(SENDER_EMAIL, SENDER_PASSWORD)
    print("Successfully connected to the SMTP server. Starting to send emails...\n")

    # 4...Going through the customer row row
    for index, row in df.iterrows():
        customer_name = row["Customer_Name"]
        customer_email = row["Customer_Email"]
        segment = row["Customer_Segment"]

        # Message text
        subject = f"We miss you, {customer_name}! Special offer inside"

        body = f"""Hi {customer_name},

We noticed it's been a while since your last order with us. As one of our valued customers, 
we truly miss having you around.

We would love to welcome you back! Use the coupon code 'WELCOMEBACK10' to get 10% off on your next purchase.

Best regards,
Your Company Team
"""

        ##Email structure
        msg = MIMEMultipart()
        msg["From"] = SENDER_EMAIL
        msg["To"] = customer_email
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "plain"))

        #Send email✔️🚀🚀🚀👌😊
        server.sendmail(SENDER_EMAIL, customer_email, msg.as_string())
        print(f"[{index + 1}/{len(df)}] Email sent successfully to {customer_name} ({customer_email})")

        
        time.sleep(3)

except Exception as e:
    print(f"An error occurred: {e}")

finally:
    
    server.quit()
    print("\nAll emails processed. Connection closed cleanly.")
