import requests
from bs4 import BeautifulSoup
import os
import pandas as pd
import xlwings as xw

# Constants
last_id = 0
AMOUNT = 1000  # The amount of Thingiverse models to check
OUTPUT_FOLDER = 'Thing Files'
COLUMNS = ['ID', 'Title', 'URL', 'Image', 'Description']
EXCEL_FILE_PATH = os.path.join(OUTPUT_FOLDER, 'Thingiverse.xlsx')
URL_TEMPLATE = 'https://www.thingiverse.com/thing:{}'
rows_to_append = []

def update_hyperlinks():
    # Open the Excel file with xlwings
    app = xw.App()
    wb = xw.Book(EXCEL_FILE_PATH)
    sheet = wb.sheets['Sheet1']

    # Iterate through the "URL" column and add hyperlinks
    for row_num in range(2, len(df) + 2):  # Start from the second row (index 2)
        url_cell = sheet.range(f'C{row_num}')  # Assuming "URL" column is column C
        url_cell.add_hyperlink(url_cell.value, text_to_display=url_cell.value)

        image_cell = sheet.range(f'D{row_num}')  # Assuming "Image" column is column D
        image_cell.add_hyperlink(image_cell.value, text_to_display=image_cell.value)

    # Save the updated Excel file with hyperlinks
    wb.save()
    wb.close()
    app.quit()
    

# Create the output folder if it doesn't exist
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Load or create the DataFrame
if os.path.exists(EXCEL_FILE_PATH):
    df = pd.read_excel(EXCEL_FILE_PATH)
    last_id = df['ID'].max()
else:
    df = pd.DataFrame(columns=COLUMNS)
    last_id = 0

# Get the new range of IDs to check
fromRange = last_id + 1 if last_id else 0
toRange = fromRange + AMOUNT

for number in range(fromRange, toRange):
    url = URL_TEMPLATE.format(number)
    print(f'Checking {url}')

    # Fetch the webpage
    try:
        r = requests.get(url)
        r.raise_for_status()
    except requests.exceptions.RequestException as e:
        print(f"Failed to fetch {url}: {e}")
        continue
    
    soup = BeautifulSoup(r.text, 'html.parser')

    # Get the title of the model
    title = soup.find('title').text

    # Get the content of the property="og:description"
    description = soup.find('meta', property="og:description")

    # Get the content of the property="og:image"
    image = soup.find('meta', property="og:image")

    # Create a dictionary for the new row
    new_row = {'ID': number, 'Title': title, 'URL': url, 'Image': image['content'], 'Description': description['content']}
    rows_to_append.append(new_row)

# Append the new rows to the DataFrame using pd.concat
if rows_to_append:
    df = pd.concat([df, pd.DataFrame(rows_to_append)], ignore_index=True)

# Save the DataFrame to the Excel file without hyperlinks for now
df.to_excel(EXCEL_FILE_PATH, index=False)

update_hyperlinks()
