import pandas as pd
import requests
import time
from bs4 import BeautifulSoup
import random

# Add a startup warning message about the cookie
print("Warning: Ensure that the cookie value in 'cookie.txt' is up-to-date.")

# Function to get active modules from the HTML response
def get_active_modules(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    modules = []
    module_mapping = {
        "Prospect": "PROSPECTS",
        "Applicant": "APPLICATIONS",
        "Enrolled": "ENROLLMENTS",
        "Financial Aid": "AID"
    }
    rows = soup.find_all('tr')
    for row in rows:
        cols = row.find_all('td')
        if len(cols) >= 2 and cols[1].text.strip() == 'Active':
            module_name = cols[0].text.strip()
            if module_name in module_mapping:
                modules.append(str(module_mapping[module_name]))
    return modules

# Function to check for 'Update Successful' in the response
def is_update_successful(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    update_message = soup.find("span", class_="rvt-inline-alert__message")
    return update_message is not None and "Update Successful" in update_message.text


# Read Employee IDs from Excel file
excel_path = 'ID.xlsx'  # Replace with your Excel file path
df = pd.read_excel(excel_path)
empl_ids = df['ID'].tolist()  # Replace with your column name

# cookie_value = str(input("Please enter the Cookie value: "))
# print(cookie_value)
with open('cookie.txt', 'r') as file:
    cookie_value = file.read().strip()


# URL and headers for the GET request
url_template = "https://sisjee-stage.iu.edu/essweb-stg/web/sisjee/crm/load.html?emplid={}"
headers = {
    # Your headers here
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
    "Accept-Encoding": "gzip, deflate, br, zstd",
    "Accept-Language": "en-US,en;q=0.9",
    "Cache-Control": "no-cache",
    "Connection": "keep-alive",
    "Cookie": cookie_value,
    "Host": "sisjee-stage.iu.edu",
    "Pragma": "no-cache",
    "Referer": "https://sisjee-stage.iu.edu/essweb-stg/web/sisjee/crm/load.html",
    "Sec-Fetch-Dest": "document",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "same-origin",
    "Sec-Fetch-User": "?1",
    "Upgrade-Insecure-Requests": "1",
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
    "sec-ch-ua": "\"Google Chrome\";v=\"123\", \"Not:A-Brand\";v=\"8\", \"Chromium\";v=\"123\"",
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": "\"macOS\""
}

# URL for the POST request
post_url = "https://sisjee-stage.iu.edu/essweb-stg/web/sisjee/crm/load.html"

# Confirmation before processing
print(f"There are {len(empl_ids)} employee IDs to process.")
confirmation = input("Do you want to proceed? (yes/no): ")
if confirmation.lower() != 'yes':
    print("Processing aborted.")
    exit()

# Process each Employee ID
for empl_id in empl_ids:
    # GET request to get HTML content
    get_url = url_template.format(empl_id)
    get_response = requests.get(get_url, headers=headers)

    if get_response.status_code == 200:
        active_modules = get_active_modules(get_response.text)
        payload = {
            "emplid": str(empl_id),
            "modules": active_modules
        }
        print("Payload for emplid", empl_id, ":", payload)

        # POST request with the payload
        post_response = requests.post(post_url, headers=headers, json=payload)
        if is_update_successful(post_response.text):
            print("Response for emplid", empl_id, ": Update Successful")
        else:
            print("Response for emplid", empl_id, ": Update not successful or not found in response")


        delay = random.randint(10, 15)
        print(f"Waiting for {delay} seconds before next request.")
        time.sleep(delay)

    else:
        print(f"Failed to fetch data for emplid {empl_id}")
        choice = input("Do you want to update the cookie and retry? (yes/no): ")
        if choice.lower() == 'yes':
            with open('cookie.txt', 'r') as file:
                cookie_value = file.read().strip()
            headers["Cookie"] = cookie_value
        else:
            try_again = False
