import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from email.message import EmailMessage
import smtplib, ssl
from datetime import datetime, timedelta

def initialize_gsc_service(service_account_json_key, scopes):
    credentials = service_account.Credentials.from_service_account_file(
        filename=service_account_json_key,
        scopes=scopes
    )
    service = build('webmasters', 'v3', credentials=credentials)
    return service

def get_query_data(service, site_url, query, start_date, end_date):
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
    return rows

def calculate_average_clicks(rows):
    total_clicks = sum(row.get('clicks', 0) for row in rows)
    average_clicks = total_clicks / len(rows) if rows else 0
    return average_clicks

def get_clicks_for_date(rows, target_date):
    for row in rows:
        if row.get('keys', [])[0] == target_date:
            return row.get('clicks', 0)
    return 0

def create_html_content(data):
    html = """
    <html>
    <body>
        <h2>Query Performance</h2>
        <table border="1" style="width:100%; border-collapse: collapse;">
            <tr>
                <th>Query</th>
                <th>Clicks on {today}</th>
                <th>Clicks on {last_day}</th>
                <th>Average Clicks (Last 3 Months)</th>
                <th>Deviation Percentage</th>
            </tr>
    """.format(today=data[0]['today_date_formatted'], last_day=data[0]['last_day_date_formatted'])
    for item in data:
        html += f"""
            <tr>
                <td>{item['query']}</td>
                <td>{item['today_clicks']}</td>
                <td>{item['last_day_clicks']}</td>
                <td>{item['average_clicks']}</td>
                <td>{item['deviation']}</td>
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

def get_previous_weekday(date):
    date = datetime.strptime(date, '%Y-%m-%d')
    prev_weekday = date - timedelta(days=7)
    return prev_weekday.strftime('%Y-%m-%d'), prev_weekday.strftime('%A')

if __name__ == "__main__":
    service_account_json_key = 'botpresso-mail-automate-service-key.json'
    scopes = ['https://www.googleapis.com/auth/webmasters.readonly']
    site_url = 'sc-domain:botpresso.com'
    queries = ['botpresso', 'botpresso extension', 'botpresso seo extension', 'saas seo checklist']

    service = initialize_gsc_service(service_account_json_key, scopes)
    last_available_date = get_last_available_date(service, site_url)
    start_date = (datetime.strptime(last_available_date, '%Y-%m-%d') - timedelta(days=90)).strftime('%Y-%m-%d')
    today_date = (datetime.strptime(last_available_date, '%Y-%m-%d') - timedelta(days=1)).strftime('%Y-%m-%d')
    today_date_name = datetime.strptime(today_date, '%Y-%m-%d').strftime('%A')
    last_week_date, last_week_date_name = get_previous_weekday(today_date)

    data = []

    for query in queries:
        rows = get_query_data(service, site_url, query, start_date, last_available_date)
        average_clicks = calculate_average_clicks(rows)
        today_clicks = get_clicks_for_date(rows, today_date)
        last_week_clicks = get_clicks_for_date(rows, last_week_date)
        deviation = (1 - today_clicks / average_clicks) * 100 if average_clicks > 0 else 0

        if today_clicks < 0.7 * average_clicks:
            data.append({
                'query': query,
                'today_date_formatted': f"{today_date_name} ({today_date})",
                'last_day_date_formatted': f"{last_week_date_name} ({last_week_date})",
                'today_clicks': today_clicks,
                'last_day_clicks': last_week_clicks,
                'average_clicks': average_clicks,
                'deviation': f"{deviation:.2f}% down"
            })

    if data:
        html_content = create_html_content(data)
        email_sender = 'harryraj1413@gmail.com'
        email_password = 'tvgy rtda iqnp qdwi'
        email_receivers = ['harryraj1413@gmail.com', 'kunjal@botpresso.com']
        subject = 'Query Performance Alert'

        send_email(email_sender, email_receivers, subject, html_content, email_password)
    else:
        print("No significant deviations found in the queries.")
