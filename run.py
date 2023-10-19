import requests
import openpyxl
from openpyxl.styles import Font
from bs4 import BeautifulSoup
from tqdm import tqdm  # Import tqdm for the progress bar

# Function to check if a link is alive
def is_link_alive(url):
    try:
        response = requests.head(url)
        return response.status_code == 200
    except requests.ConnectionError:
        return False

# Function to extract links from a page
def get_links_from_page(url):
    try:
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')
        links = []
        for link in soup.find_all('a'):
            href = link.get('href')
            if href and href.startswith('http'):
                links.append(href)
        return links
    except requests.ConnectionError:
        return []

# URL of the website you want to check
website_url = 'https://example.com'

# Create a new Excel workbook and add a worksheet
workbook = openpyxl.Workbook()
worksheet = workbook.active

# Set the headers for the Excel file
worksheet['A1'] = 'Element'
worksheet['B1'] = 'Link'
worksheet['C1'] = 'Status'

# Style the header row
font = Font(bold=True)
for cell in worksheet['1:1']:
    cell.font = font

# Initialize row counter
row = 2

# Check links on the website
links_to_check = get_links_from_page(website_url)

# Use tqdm to create a progress bar
with tqdm(total=len(links_to_check), unit='link') as pbar:
    for link in links_to_check:
        if is_link_alive(link):
            status = '✔️'
        else:
            status = '❌'
        worksheet.cell(row=row, column=1, value='Link')
        worksheet.cell(row=row, column=2, value=link)
        worksheet.cell(row=row, column=3, value=status)
        row += 1
        pbar.update(1)  # Update the progress bar

# Save the Excel file
workbook.save('dead_links_report.xlsx')

print("Dead links report generated in 'dead_links_report.xlsx'")
