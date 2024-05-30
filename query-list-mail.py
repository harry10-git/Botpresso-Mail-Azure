import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from email.message import EmailMessage
import smtplib
import ssl
from datetime import datetime, timedelta


def initialize_gsc_service(service_account_json_key, scopes):
    credentials = service_account.Credentials.from_service_account_file(
        filename=service_account_json_key,
        scopes=scopes
    )
    service = build('webmasters', 'v3', credentials=credentials)
    return service


def calculate_average_clicks(service, site_url, query, start_date, end_date):
    request = {
        'startDate': start_date,
        'endDate': end_date,
        'dimensions': ['date'],
        'dimensionFilterGroups': [{
            'filters': [{
                'dimension': 'query',
                'expression': query
            }]
        }]
    }
    response = service.searchanalytics().query(siteUrl=site_url, body=request).execute()
    rows = response.get('rows', [])
    total_clicks = sum(row.get('clicks', 0) for row in rows)
    average_clicks = total_clicks / len(rows) if len(rows) > 0 else 0
    return average_clicks


def get_today_clicks(service, site_url, query, today_date):
    request = {
        'startDate': today_date,
        'endDate': today_date,
        'dimensions': ['query']
    }
    response = service.searchanalytics().query(siteUrl=site_url, body=request).execute()
    rows = response.get('rows', [])
    for row in rows:
        if row.get('keys', [])[0] == query:
            return row.get('clicks', 0)
    return 0


def get_last_available_date(service, site_url):
    end_date = datetime.now().strftime('%Y-%m-%d')
    while True:
        request = {
            'startDate': end_date,
            'endDate': end_date,
            'dimensions': ['date']
        }
        response = service.searchanalytics().query(siteUrl=site_url, body=request).execute()
        if 'rows' in response and response['rows']:
            return end_date
        end_date = (datetime.strptime(end_date, '%Y-%m-%d') - timedelta(days=1)).strftime('%Y-%m-%d')


def create_html_content(data):
    html = """
    <html>
    <body>
        <h2>Queries with Significant Decrease in Clicks</h2>
        <table border="1" style="width:100%; border-collapse: collapse;">
            <tr>
                <th>Query</th>
                <th>Average Clicks (Last 3 Months)</th>
                <th>Clicks (3 Days Ago)</th>
                <th>Deviation (%)</th>
            </tr>
    """
    for item in data:
        html += f"""
            <tr>
                <td>{item['Query']}</td>
                <td>{item['Average Clicks']:.2f}</td>
                <td>{item['Today Clicks']}</td>
                <td>{item['Deviation']:.2f}%</td>
            </tr>
        """
    html += """
        </table>
    </body>
    </html>
    """
    return html


def send_email(sender, receivers, subject, html_content, password):
    em = EmailMessage()
    em['From'] = sender
    em['To'] = ', '.join(receivers)
    em['Subject'] = subject
    em.set_content(html_content, subtype='html')

    context = ssl.create_default_context()

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
            smtp.login(sender, password)
            smtp.sendmail(sender, receivers, em.as_string())
            print('Email was sent successfully!')
    except Exception as e:
        print(f"Error occurred: {e}")



if __name__ == "__main__":
    # Google Search Console parameters
    service_account_json_key = '/home/azureuser/Botpresso-Mail-Azure/botpresso-mail-automate-service-key.json'
    scopes = ['https://www.googleapis.com/auth/webmasters.readonly']
    site_url = 'sc-domain:botpresso.com'  # or 'sc-domain:botpresso.com'

    queries = ['botpresso', 'botpresso extension', 'botpresso seo extension' , 'saas seo checklist']

    try:
        service = initialize_gsc_service(service_account_json_key, scopes)

        # Get the last available date
        last_available_date = get_last_available_date(service, site_url)
        start_date = (datetime.strptime(last_available_date, '%Y-%m-%d') - timedelta(days=90)).strftime('%Y-%m-%d')
        today_date = (datetime.strptime(last_available_date, '%Y-%m-%d') - timedelta(days=3)).strftime('%Y-%m-%d')

        query_data = []

        for query in queries:
            average_clicks = calculate_average_clicks(service, site_url, query, start_date, last_available_date)
            today_clicks = get_today_clicks(service, site_url, query, today_date)

            if today_clicks < 0.7 * average_clicks:
                deviation_percentage = ((average_clicks - today_clicks) / average_clicks) * 100
                query_data.append({
                    'Query': query,
                    'Average Clicks': average_clicks,
                    'Today Clicks': today_clicks,
                    'Deviation': deviation_percentage
                })

        if query_data:
            # Create HTML content for email
            html_content = create_html_content(query_data)

            # Email parameters
            email_sender = 'harryraj1413@gmail.com'
            email_password = 'tvgy rtda iqnp qdwi'  # Retrieve from environment variable
            email_receivers = ['harry.raj@learner.manipal.edu', 'harryraj1413@gmail.com', 'kunjal@botpresso.com']
            subject = 'Queries with Significant Decrease in Clicks'

            # Send email
            send_email(email_sender, email_receivers, subject, html_content, email_password)
        else:
            print("No significant decrease in clicks for the specified queries.")
    except HttpError as e:
        error_content = e.content.decode('utf-8')
        print(f"HTTP error occurred: {error_content}")
    except Exception as e:
        print(f"An error occurred: {e}")
