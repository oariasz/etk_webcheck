import os
import time
import openpyxl
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

class WebChecker:
    def __init__(self, excel_file, max_checks_per_day):
        '''
        Initializes the WebChecker object.

        :param excel_file: Path to the Excel file containing the list of URLs.
        :param sender_email_env: Environment variable name for the sender's email account.
        :param recipient_email_env: Environment variable name for the recipient's email account.
        :param password_env: Environment variable name for the sender email account password.
        :param max_checks_per_day: Maximum number of URL checks allowed in a day.
        '''
        
        self.excel_file = excel_file
        self.sender_email = os.getenv('SENDER_EMAIL_ENV_VAR')
        self.recipient_email = os.getenv('RECIPIENT_EMAIL_ENV_VAR')
        self.password_env = 'EMAIL_PASSWORD_ENV_VAR'
        self.max_checks_per_day = max_checks_per_day
        self.checked_urls = set()

    def read_password(self):
        '''
        Reads the sender email account password from the environment variable.

        :return: Password as a string.
        '''
        return os.getenv(self.password_env)

    def send_email(self, subject, body):
        '''
        Sends an email with the specified subject and body.

        :param subject: Subject of the email.
        :param body: Body content of the email.
        '''
        password = self.read_password()

        msg = MIMEMultipart()
        msg['From'] = self.sender_email
        msg['To'] = self.recipient_email
        msg['Subject'] = subject

        msg.attach(MIMEText(body, 'plain'))

        with smtplib.SMTP('smtp.example.com', 587) as server:
            server.starttls()
            server.login(self.sender_email, password)
            server.sendmail(self.sender_email, self.recipient_email, msg.as_string())

    def check_urls(self):
        '''
        Checks the list of URLs from the Excel file for updates.

        :return: List of updated URLs with titles.
        '''
        # Open Excel file and get sheet
        workbook = openpyxl.load_workbook(self.excel_file)
        sheet = workbook.active

        updated_urls = []

        for row in sheet.iter_rows(min_row=2, values_only=True):
            url, last_checked, title = row
            if url not in self.checked_urls:
                # Print the URL being checked
                print(f'Checking URL: {url}')

                # Perform URL checking logic here to determine if it has been updated
                # If updated, add to the list
                updated_urls.append({'url': url, 'title': title})
                self.checked_urls.add(url)

        workbook.save(self.excel_file)
        return updated_urls


    def run(self, checks_per_day):
        '''
        Runs the WebChecker application, performing URL checks and sending emails.

        :param checks_per_day: Number of times to check URLs per day.
        '''
        elapsed_time = 0
        start_time = time.time()

        try:
            # Perform the initial check
            self.check_urls()

            while True:
                print('WHILE')
                if checks_per_day > self.max_checks_per_day:
                    raise ValueError('Exceeded maximum checks per day.')

                updated_urls = self.check_urls()

                if updated_urls:
                    subject = 'Web Update Summary'
                    body = '\n'.join([f"{url['title']}: {url['url']}" for url in updated_urls])
                    self.send_email(subject, body)

                elapsed_time = time.time() - start_time
                
                # TODO: Remove this line, uncomment the next one
                time.sleep(120)
                # time.sleep(86400 / checks_per_day)  # 86400 seconds in a day

        except (KeyboardInterrupt, ValueError) as e:
            print(f'\nProgram terminated. Elapsed time: {elapsed_time:.2f} seconds.')
            print(f'Emails sent: {len(self.checked_urls)}')


if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(description='Web Checker App')
    parser.add_argument('checks_per_day', type=int, help='Number of times to check URLs per day.')
    args = parser.parse_args()

    try:
        web_checker = WebChecker(
            excel_file='urls.xlsx',
            max_checks_per_day=24
        )
        web_checker.run(args.checks_per_day)

    except Exception as e:
        print(f'Error: {e}')
