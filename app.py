import requests
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
import validators

# Enter the API key from Google Cloud Console (PageSpeed Insights)
API_KEY = 'YOUR_API_KEY_HERE'

# Authentication and initialization of Google Sheets API
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
client = gspread.authorize(creds)

# Open the spreadsheet and select the worksheet
spreadsheet_url = 'YOUR_SPREADSHEET_URL_HERE'
spreadsheet = client.open_by_url(spreadsheet_url)
worksheet = spreadsheet.sheet1

# Read the link from Google Sheets
links = worksheet.col_values(1)  # Retrieve all values from the first column

# Read the results from the worksheet if they already exist
try:
    result_worksheet = spreadsheet.worksheet("PageSpeed Results")
    existing_data = result_worksheet.get_all_records()
    existing_df = pd.DataFrame(existing_data)
except gspread.exceptions.WorksheetNotFound:
    existing_df = pd.DataFrame()  # Worksheet does not exist yet, so there is no data

# Function to retrieve performance from PageSpeed Insights
def get_page_speed(link, strategy='mobile'):
    url = f"https://www.googleapis.com/pagespeedonline/v5/runPagespeed?url={link}&key={API_KEY}&strategy={strategy}"
    response = requests.get(url)
    if response.status_code == 200:
        return response.json()
    else:
        response.raise_for_status()

# List to store the results
results = []

# Loop through each link and retrieve the results if they do not already exist
for link in links[1:]:  # Skip header
    if not validators.url(link):
        print(f"Invalid URL: {link}")
        results.append({
            'Link': link,
            'Score Mobile': 'Invalid URL',
            'FCP Mobile (s)': '',
            'LCP Mobile (s)': '',
            'TBT Mobile (ms)': '',
            'CLS Mobile': '',
            'SI Mobile (s)': '',
            'Score Desktop': 'Invalid URL',
            'FCP Desktop (s)': '',
            'LCP Desktop (s)': '',
            'TBT Desktop (ms)': '',
            'CLS Desktop': '',
            'SI Desktop (s)': '',
            'PageSpeed Link': ''
        })
        continue

    # Check if the link has already been measured
    if not existing_df.empty and link in existing_df['Link'].values:
        print(f"Skipping {link}, already measured.")
        continue

    try:
        mobile_result = get_page_speed(link, 'mobile')
        desktop_result = get_page_speed(link, 'desktop')

        # Extract metrics for mobile
        mobile_lighthouse = mobile_result['lighthouseResult']
        mobile_score = mobile_lighthouse['categories']['performance']['score'] * 100
        mobile_metrics = mobile_lighthouse['audits']
        fcp_mobile = mobile_metrics['first-contentful-paint']['displayValue']
        lcp_mobile = mobile_metrics['largest-contentful-paint']['displayValue']
        tbt_mobile = mobile_metrics['total-blocking-time']['displayValue']
        cls_mobile = mobile_metrics['cumulative-layout-shift']['displayValue']
        si_mobile = mobile_metrics['speed-index']['displayValue']

        # Extract metrics for desktop
        desktop_lighthouse = desktop_result['lighthouseResult']
        desktop_score = desktop_lighthouse['categories']['performance']['score'] * 100
        desktop_metrics = desktop_lighthouse['audits']
        fcp_desktop = desktop_metrics['first-contentful-paint']['displayValue']
        lcp_desktop = desktop_metrics['largest-contentful-paint']['displayValue']
        tbt_desktop = desktop_metrics['total-blocking-time']['displayValue']
        cls_desktop = desktop_metrics['cumulative-layout-shift']['displayValue']
        si_desktop = desktop_metrics['speed-index']['displayValue']

        page_speed_link = f"https://developers.google.com/speed/pagespeed/insights/?url={link}"

        results.append({
            'Link': link,
            'Score Mobile': mobile_score,
            'FCP Mobile (s)': fcp_mobile,
            'LCP Mobile (s)': lcp_mobile,
            'TBT Mobile (ms)': tbt_mobile,
            'CLS Mobile': cls_mobile,
            'SI Mobile (s)': si_mobile,
            'Score Desktop': desktop_score,
            'FCP Desktop (s)': fcp_desktop,
            'LCP Desktop (s)': lcp_desktop,
            'TBT Desktop (ms)': tbt_desktop,
            'CLS Desktop': cls_desktop,
            'SI Desktop (s)': si_desktop,
            'PageSpeed Link': page_speed_link
        })
        
        # Output to terminal
        print(f"Processed {link}: Mobile Score = {mobile_score}, Desktop Score = {desktop_score}")
    except requests.exceptions.RequestException as e:
        print(f"Error processing {link}: {e}")
        results.append({
            'Link': link,
            'Score Mobile': 'Error',
            'FCP Mobile (s)': '',
            'LCP Mobile (s)': '',
            'TBT Mobile (ms)': '',
            'CLS Mobile': '',
            'SI Mobile (s)': '',
            'Score Desktop': 'Error',
            'FCP Desktop (s)': '',
            'LCP Desktop (s)': '',
            'TBT Desktop (ms)': '',
            'CLS Desktop': '',
            'SI Desktop (s)': '',
            'PageSpeed Link': ''
        })
    except KeyError as e:
        print(f"Unexpected response structure for {link}: {e}")
        results.append({
            'Link': link,
            'Score Mobile': 'Error',
            'FCP Mobile (s)': '',
            'LCP Mobile (s)': '',
            'TBT Mobile (ms)': '',
            'CLS Mobile': '',
            'SI Mobile (s)': '',
            'Score Desktop': 'Error',
            'FCP Desktop (s)': '',
            'LCP Desktop (s)': '',
            'TBT Desktop (ms)': '',
            'CLS Desktop': '',
            'SI Desktop (s)': '',
            'PageSpeed Link': ''
        })
    time.sleep(2)  # Pause for 2 seconds between requests to avoid rate limiting

# Convert the results to a DataFrame and upload to Google Sheets
result_df = pd.DataFrame(results)

# Add the results to a new worksheet or update the existing worksheet
try:
    if existing_df.empty:
        result_worksheet = spreadsheet.add_worksheet(title="PageSpeed Results", rows="100", cols="15")
    else:
        result_worksheet = spreadsheet.worksheet("PageSpeed Results")

    # Append new results to the existing DataFrame
    if not existing_df.empty:
        result_df = pd.concat([existing_df, result_df], ignore_index=True)
except gspread.exceptions.WorksheetNotFound:
    result_worksheet = spreadsheet.add_worksheet(title="PageSpeed Results", rows="100", cols="15")

# Clear the worksheet before writing new results
result_worksheet.clear()

# Write header and results
result_worksheet.update([result_df.columns.values.tolist()] + result_df.values.tolist())

print("Process completed, results saved in the 'PageSpeed Results' worksheet in Google Sheets")
