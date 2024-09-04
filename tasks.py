import os
import re
import requests
from openpyxl import Workbook
from robocorp.tasks import task
from RPA.Browser.Selenium import Selenium
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from RPA.Robocorp.WorkItems import WorkItems

browser = Selenium()
workitems = WorkItems()

search_phrase = "Babar Azam"

images_dir = "output/news_images"
os.makedirs(images_dir, exist_ok=True)
excel_dir = "output/excel_files"
os.makedirs(excel_dir, exist_ok=True)

 # Data storage
data = []

@task
def extract_news():

    open_website_and_search_phrase()

    get_news_data()

    save_news_data_in_excel()

    # Close the browser
    browser.close_browser()


def open_website_and_search_phrase():
     # Open the website
    browser.open_available_browser("https://apnews.com/")   
    
    # Wait until the search button is visible and clickable
    browser.wait_until_element_is_visible("//button[contains(@class, 'SearchOverlay-search-button')]")
    browser.click_element("//button[contains(@class, 'SearchOverlay-search-button')]")
    
    # Input the search term into the search bar
    browser.input_text_when_element_is_visible("//input[contains(@class, 'SearchOverlay-search-input')]", search_phrase)
    
    # Click the search submit button
    browser.click_element("//button[contains(@class, 'SearchOverlay-search-submit')]")
    
    # Wait for the search results to load
    browser.wait_until_page_contains(search_phrase)

def get_news_data():
    # Use WebDriverWait to ensure the articles container is present
    try:
        search_phrase, news_category, months = extract_parameters_from_workitem()
        if search_phrase == '':
            search_phrase = search_phrase

        wait = WebDriverWait(browser.driver, 10)
        articles_container = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".PageListStandardD .PageList-items")))

        # Extract the news articles
        articles = articles_container.find_elements(By.CSS_SELECTOR, ".PageList-items-item")
        
        # Pattern for detecting money amounts
        money_pattern = re.compile(r'\$\d+(?:,\d{3})*(?:\.\d+)?|\d+(?:,\d{3})*(?:\.\d+)?\s?(?:dollars|USD)', re.IGNORECASE)
        
        for article in articles:
            try:
                title_element = article.find_element(By.CSS_SELECTOR, ".PagePromo-title .PagePromoContentIcons-text")
                description_element = article.find_element(By.CSS_SELECTOR, ".PagePromo-description .PagePromoContentIcons-text")
                date_element = article.find_element(By.CSS_SELECTOR, ".PagePromo-date .Timestamp-template")
                
                title = title_element.text
                description = description_element.text
                date = date_element.text
                
                # Try to find the image element, handle if it's missing
                image_url = None
                try:
                    image_file = article.find_element(By.CSS_SELECTOR, ".PagePromo-media .Image")
                    image_url = image_file.get_attribute("src")
                except Exception as e:
                    print(f"Image not found for article: {title}. Skipping image download.")

                # Count occurrences of search phrase
                title_count = title.lower().count(search_phrase.lower())
                description_count = description.lower().count(search_phrase.lower())
                
                # Check for money in text
                contains_money = bool(money_pattern.search(title + ' ' + description))
                
                # Download image if URL was found
                image_filename = "Not Available"
                if image_url:
                    image_filename = f"{title[:15]}.jpg".replace(" ", "_")  # Use a truncated title as filename
                    image_path = os.path.join(images_dir, image_filename)
                    response = requests.get(image_url)
                    if response.status_code == 200:
                        with open(image_path, 'wb') as file:
                            file.write(response.content)
                    else:
                        image_filename = "Not Available"  # Default if image download fails
                
                data.append({
                    "Title": title,
                    "Date": date,
                    "Description": description,
                    "Image Filename": image_filename,
                    "Search Phrase Count": title_count + description_count,
                    "Contains Money": contains_money
                })
            
            except Exception as e:
                print(f"Error extracting details from article: {e}")
    
    except Exception as e:
        print(f"Error locating articles: {e}")


def save_news_data_in_excel():
     # Save to Excel
    wb = Workbook()
    ws = wb.active
    
    # Ensure the worksheet title is not empty
    if search_phrase.strip():
        ws.title = search_phrase
    else:
        ws.title = "News Data"

    # Write headers
    headers = ["Title", "Date", "Description", "Image Filename", "Search Phrase Count", "Contains Money"]
    ws.append(headers)

    # Write data
    for item in data:
        ws.append([
            item["Title"],
            item["Date"],
            item["Description"],
            item["Image Filename"],
            item["Search Phrase Count"],
            item["Contains Money"]
        ])

    # Save the Excel file
    timestamp = time.time()

    # Separate the fractional part after the decimal point
    fractional_part = int(timestamp)
    file_name = excel_dir + "/news_file_"+ str(fractional_part) +".xlsx"
    wb.save(file_name)

def extract_parameters_from_workitem():

 # Attempt to get the input work item
    try:
        workitems.get_input_work_item()
        
        # Fetch parameters from the work item
        search_phrase = workitems.get_work_item_variable("search_phrase", "")
        news_category = workitems.get_work_item_variable("news_category", "")
        months = int(workitems.get_work_item_variable("months", 0))

        return search_phrase, news_category, months
        
    except RuntimeError as e:
        # If no active work item, use default values for testing
        print(f"No active work item found. Using default values. Error: {e}")
        search_phrase = "Babar Azam"
        news_category = "sports"
        months = 1

        return search_phrase, news_category, months
