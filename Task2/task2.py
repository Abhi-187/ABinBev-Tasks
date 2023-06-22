##TO automate the script every three hours we can use crontab
## Wrie the below command in terminal
## crontab -e
## 0 */3 * * * /Users/abhipatel/Desktop/ABinBev/Task2/automate.py


import time
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from getpass import getpass
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import hashlib

# Set the path to the Chrome WebDriver executable
webdriver_path = "/Users/abhipatel/Desktop/ABinBev/chromedriver"
chrome_options = Options()


def initialize_driver():
    # Initialize the Chrome WebDriver with the specified path
    driver = webdriver.Chrome(
        service=Service(executable_path=webdriver_path), options=chrome_options
    )
    return driver


def login(driver, username, password):
    target_url = "https://www.linkedin.com/?original_referer="

    while driver.current_url != target_url:
        driver.get(target_url)
        time.sleep(1)

    wait = WebDriverWait(driver, 20)
    username_input = wait.until(
        EC.visibility_of_element_located((By.XPATH, "//input[@name='session_key']"))
    )
    password_input = driver.find_element(By.XPATH, "//input[@name='session_password']")

    username_input.send_keys(username)
    password_input.send_keys(password)

    submit_button = driver.find_element(By.XPATH, "//button[@type='submit']")
    submit_button.click()

    time.sleep(2)


def get_notification_counts(driver):
    notification_badges = WebDriverWait(driver, 20).until(
        EC.presence_of_all_elements_located(
            (By.XPATH, "//span[contains(@class, 'notification-badge__count')]")
        )
    )

    if len(notification_badges) == 4:
        messaging_count = notification_badges[1].text
        notification_count = notification_badges[2].text
    else:
        notification_count = notification_badges[2].text
        messaging_count = notification_badges[3].text

    return messaging_count, notification_count


def load_existing_data(excel_filename):
    try:
        df = pd.read_excel(excel_filename)
    except FileNotFoundError:
        df = pd.DataFrame(columns=["User", "DateTime", "Message", "Notification"])

    return df


def calculate_notification_diff(df, username, messaging_count, notification_count):
    filtered_df = df[df["User"] == username]

    if not filtered_df.empty and "Message" in filtered_df.columns:
        last_record = filtered_df.iloc[-1]
        message_diff = int(messaging_count) - int(last_record["Message"])
        notification_diff = int(notification_count) - int(last_record["Notification"])
    else:
        message_diff = 0
        notification_diff = 0

    return message_diff, notification_diff


def append_data_to_dataframe(
    df,
    username,
    current_time,
    messaging_count,
    notification_count,
    message_diff,
    notification_diff,
):
    new_row = pd.DataFrame(
        {
            "User": [username],
            "DateTime": [current_time],
            "Message": [messaging_count],
            "Notification": [notification_count],
            "Message Difference": [message_diff],
            "Notification Difference": [notification_diff],
        }
    )
    df = pd.concat([df, new_row], ignore_index=True)
    return df


def write_dataframe_to_excel(df, excel_filename):
    df.to_excel(excel_filename, index=False)


def send_email(
    username, messaging_count, notification_count, message_diff, notification_diff
):
    # Email configuration
    sender_email = "your_email@gmail.com"
    sender_password = "your_password"
    receiver_email = username

    # Create the email body
    email_body = create_email_body(
        username, messaging_count, notification_count, message_diff, notification_diff
    )

    # Create the email message
    message = MIMEMultipart("alternative")
    message["Subject"] = "Your Email Report"
    message["From"] = sender_email
    message["To"] = receiver_email

    # Attach the HTML email body to the message
    email_content = MIMEText(email_body, "html")
    message.attach(email_content)

    # Connect to the email server and send the message
    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login("test2abinbev@gmail.com", "qeapjpolqeqnshao")
        server.sendmail(sender_email, receiver_email, message.as_string())


def create_email_body(
    username, messaging_count, notification_count, message_diff, notification_diff
):
    # Define the HTML email body
    email_body = f"""
    <html>
    <head>
        <style>
            body {{
                font-family: Arial, sans-serif;
                background-color: #f1f1f1;
                margin: 0;
                padding: 20px;
            }}
            .container {{
                background-color: #fff;
                padding: 20px;
                border-radius: 5px;
                box-shadow: 0 2px 5px rgba(0, 0, 0, 0.1);
            }}
            h1 {{
                color: #333;
                font-size: 24px;
                margin-top: 0;
            }}
            p {{
                color: #666;
                font-size: 16px;
                margin-bottom: 20px;
            }}
            table {{
                border-collapse: collapse;
                width: 100%;
            }}
            th, td {{
                padding: 8px;
                text-align: left;
                border-bottom: 1px solid #ddd;
            }}
            th {{
                font-weight: bold;
            }}
            .footer {{
                color: #999;
                font-size: 14px;
                text-align: center;
                margin-top: 20px;
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <h1>Email Report</h1>
            <p>Hello {username},</p>
            <p>Here is your email report:</p>
            <table>
                <tr>
                    <th>Metrics</th>
                    <th>Count</th>
                </tr>
                <tr>
                    <td>Number of Unread Messages</td>
                    <td>{messaging_count}</td>
                </tr>
                <tr>
                    <td>Number of Unread Notifications</td>
                    <td>{notification_count}</td>
                </tr>
                <tr>
                    <td>Message Difference</td>
                    <td>{message_diff}</td>
                </tr>
                <tr>
                    <td>Notification Difference</td>
                    <td>{notification_diff}</td>
                </tr>
            </table>
            <p>Thank you for using our service!</p>
            <p class="footer">Developed by Abhi Mukeshkumar Patel</p>
        </div>
    </body>
    </html>
    """
    return email_body


def encrypt_password(password):
    # Encrypt the password using a secure hashing algorithm (e.g., SHA-256)
    hashed_password = hashlib.sha256(password.encode()).hexdigest()
    return hashed_password


# Main execution flow
def main():
    # Get user credentials
    # username = "zealshah017@gmail.com"
    # password = "@Kmr0507"

    username = "marubackup187@gmail.com"
    password = "Abhi@187"

    # Initialize driver
    driver = initialize_driver()

    # Log in
    login(driver, username, password)

    # Get notification counts
    messaging_count, notification_count = get_notification_counts(driver)

    # Load existing data
    excel_filename = "/Users/abhipatel/Desktop/ABinBev/Task2/notification_data.xlsx"
    df = load_existing_data(excel_filename)

    # Get current time
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Calculate notification differences
    message_diff, notification_diff = calculate_notification_diff(
        df, username, messaging_count, notification_count
    )

    # Append data to the DataFrame
    df = append_data_to_dataframe(
        df,
        username,
        current_time,
        messaging_count,
        notification_count,
        message_diff,
        notification_diff,
    )

    # Write DataFrame to the Excel file
    write_dataframe_to_excel(df, excel_filename)

    # Send email with the report
    send_email(
        username, messaging_count, notification_count, message_diff, notification_diff
    )

    # Close the Chrome WebDriver
    driver.quit()


if __name__ == "__main__":
    main()
