import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import smtplib
import ssl
from email.message import EmailMessage
import os


def initialize_gsc_service(service_account_json_key, scopes):
    credentials = service_account.Credentials.from_service_account_file(
        filename=service_account_json_key,
        scopes=scopes
    )
    service = build('webmasters', 'v3', credentials=credentials)
    return service


def get_top_pages(service, site_url, start_date, end_date, dimensions, row_limit):
    request = {
        'startDate': start_date,
        'endDate': end_date,
        'dimensions': dimensions,
        'rowLimit': row_limit
    }
    response = service.searchanalytics().query(siteUrl=site_url, body=request).execute()
    return response


def process_response(response):
    rows = response.get('rows', [])
    data = []
    for row in rows:
        keys = row.get('keys', [])
        metrics = row.get('clicks', 0), row.get('impressions', 0)
        data.append(keys + list(metrics))
    columns = ['page', 'clicks', 'impressions']
    df = pd.DataFrame(data, columns=columns)
    return df


def create_html_content(df):
    html = """
    <html>
    <body>
        <h2>Top Performing Pages</h2>
        <table border="1" style="width:100%; border-collapse: collapse;">
            <tr>
                <th>Page</th>
                <th>Clicks</th>
                <th>Impressions</th>
            </tr>
    """
    for index, row in df.iterrows():
        html += f"""
            <tr>
                <td><a href="{row['page']}">{row['page']}</a></td>
                <td>{row['clicks']}</td>
                <td>{row['impressions']}</td>
            </tr>
        """
    html += """
        </table>
    </body>
    </html>
    """
    return html


def send_email(sender, receiver, subject, html_content, password):
    em = EmailMessage()
    em['From'] = sender
    em['To'] = receiver
    em['Subject'] = subject
    em.set_content(html_content, subtype='html')

    context = ssl.create_default_context()

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
            smtp.login(sender, password)
            smtp.sendmail(sender, receiver, em.as_string())
            print('Email was sent successfully!')
    except Exception as e:
        print(f"Error occurred: {e}")


if __name__ == "__main__":
    # Google Search Console parameters
    service_account_json_key = 'botpresso-mail-automate-service-key.json'
    scopes = ['https://www.googleapis.com/auth/webmasters.readonly']
    site_url = 'sc-domain:botpresso.com'  # or 'sc-domain:botpresso.com'
    start_date = '2023-01-01'
    end_date = '2023-01-31'
    dimensions = ['page']
    row_limit = 5  # Fetch the top 5 pages

    try:
        service = initialize_gsc_service(service_account_json_key, scopes)
        response = get_top_pages(service, site_url, start_date, end_date, dimensions, row_limit)
        df_top_pages = process_response(response)

        # Create HTML content for email
        html_content = create_html_content(df_top_pages)

        # Email parameters
        email_sender = 'harryraj1413@gmail.com'
        email_password = 'tvgy rtda iqnp qdwi'  # Retrieve from environment variable
        email_receiver = 'harry.raj@learner.manipal.edu'
        subject = 'Top Performing Pages for Your Website'

        # Send email
        send_email(email_sender, email_receiver, subject, html_content, email_password)
    except HttpError as e:
        error_content = e.content.decode('utf-8')
        print(f"HTTP error occurred: {error_content}")
    except Exception as e:
        print(f"An error occurred: {e}")
