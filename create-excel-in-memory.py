import os
import calendar
import requests
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta, date
from shareplum import Site
from shareplum import Office365
from shareplum.site import Version

# Function to get the total count of issues from Jira based on a JQL query
def get_total_issues_count(jira_url, jql_query, token):
    api_url = f"{jira_url}/rest/api/2/search"
    headers = {'Authorization': f"Bearer {token}"}
    params = {'jql': jql_query, 'maxResults': 0}

    response = requests.get(api_url, params=params, headers=headers)

    if response.status_code == 200:
        return response.json().get('total', 0)
    else:
        print(f"Error: {response.status_code}, {response.text}")
        return None

# Helper function to get the start and end dates for a month
def get_month_range(year, month_name):
    month_number = list(calendar.month_name).index(month_name)
    start_date = date(year, month_number, 1).strftime("%Y-%m-%d")
    end_date = (date(year, month_number, calendar.monthrange(year, month_number)[1]) + timedelta(days=1)).strftime("%Y-%m-%d")
    return start_date, end_date

# Function to generate and upload an Excel file to SharePoint
def generate_and_upload_excel(site, folder_path, jira_url, jql_queries, token):
    current_year = datetime.now().year
    months = [calendar.month_name[i] for i in range(1, 13)]
    data = []

    for jql_query in jql_queries:
        row = {'JQL Query': jql_query}
        for month in months:
            start_date, end_date = get_month_range(current_year, month)
            modified_jql = f"{jql_query} AND created >= '{start_date}' AND created < '{end_date}'"
            count = get_total_issues_count(jira_url, modified_jql, token)
            row[month] = count
        data.append(row)

    # Creating a DataFrame and converting it to an Excel file
    df = pd.DataFrame(data)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)

    # Uploading the Excel file to SharePoint
    output.seek(0)
    file_content = output.read()
    file_name = f"Jira_Data_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx"
    folder = site.Folder(folder_path)
    folder.upload_file(file_name, file_content)

if __name__ == "__main__":
    # Setting up environment variables
    jira_url = os.getenv("JIRA_URL", "jira_url_here")
    jql_queries = [
        'first_jql_query_here',
        'second_jql_query_here',
        # Add more JQL queries as needed
    ]
    token = os.getenv("jiratoken", "jira_api_token_here")
    sharepoint_site_url = "sharepoint_site_url_here"
    sharepoint_folder_path = "Shared Documents/folder_here"  # Path to marsya SharePoint folder

    # SharePoint authentication
    username = "sharepoint_username_here"
    password = "sharepoint_password_here"
    authcookie = Office365(sharepoint_site_url, username=username, password=password).GetCookies()
    site = Site(sharepoint_site_url, version=Version.v365, authcookie=authcookie)

    # Generate and upload Excel file to SharePoint
    generate_and_upload_excel(site, sharepoint_folder_path, jira_url, jql_queries, token)
