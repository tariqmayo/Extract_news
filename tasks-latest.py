import os
import re
import requests
from openpyxl import Workbook
from robocorp.tasks import task
from RPA.Browser.Selenium import Selenium
from RPA.Robocorp.WorkItems import WorkItems
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta

browser = Selenium()
workitems = WorkItems()


save_dir = "output/news_images"
os.makedirs(save_dir, exist_ok=True)

@task
def extract_news():
    # Example usage
    search_phrase, news_category, months = extract_parameters_from_workitem()
    if search_phrase == '':
        search_phrase = "cricket news"

    if months == '':
        months = 1

    # Prepare date range based on months
    start_date, end_date = get_month_range(months)
    start_date = start_date.strftime("%m-%d")
    end_date = end_date.strftime("%m-%d")
    print("Search Phrase : "+search_phrase)
    print("Date : "+str(start_date))
    
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

    # Data storage
    data = []

    # Use WebDriverWait to ensure the articles container is present
    try:
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
                # image_file = article.find_element(By.CSS_SELECTOR, ".PagePromo-media .Image")
                
                title = title_element.text
                description = description_element.text
                
                # Assuming the date string is in the format "Month Day" (e.g., "September 3")
                date_str = date_element.text.strip()  # Strip any extra spaces
                
                # Append the current year to the date string
                # current_year = datetime.now().year
                # date_str_with_year = f"{date_str} {current_year}"
                
                # Parse and format date
                date_obj = datetime.strptime(date_str, "%B %d")
                news_date = date_obj.strftime("%m-%d")
                
                

                 # Check if the article falls within the date range
                if start_date <= news_date <= end_date:

                    # image_url = image_file.get_attribute("src")
                    
                    # Count occurrences of search phrase
                    title_count = title.lower().count(search_phrase.lower())
                    description_count = description.lower().count(search_phrase.lower())
                    
                    # Check for money in text
                    contains_money = bool(money_pattern.search(title + ' ' + description))
                    
                    # Download image
                    # image_filename = f"{title[:30]}.jpg".replace(" ", "_")  # Use a truncated title as filename
                    # image_path = os.path.join(save_dir, image_filename)
                    # response = requests.get(image_url)
                    # if response.status_code == 200:
                        # with open(image_path, 'wb') as file:
                            # file.write(response.content)
                    # else:
                        # image_filename = "Not Available"  # Default if image download fails
                    
                    data.append({
                            "Title": title,
                            "Date": news_date,
                            "Description": description,
                            # "Image Filename": image_filename,
                            "Search Phrase Count": title_count + description_count,
                            "Contains Money": contains_money
                        })
            
            except Exception as e:
                print(f"Error extracting details from article: {e}")
    
    except Exception as e:
        print(f"Error locating articles: {e}")

    # Save to Excel
    wb = Workbook()
    ws = wb.active
    ws.title = search_phrase

    # Write headers
    headers = ["Title", "Date", "Description",  "Search Phrase Count", "Contains Money"]
    ws.append(headers)

    # Write data
    for item in data:
        ws.append([
            item["Title"],
            item["Date"],
            item["Description"],
            # item["Image Filename"],
            item["Search Phrase Count"],
            item["Contains Money"]
        ])

    # Save the Excel file
    wb.save("news_articles.xlsx")

    # Close the browser
    browser.close_browser()



# extract workitem parameters
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
        search_phrase = "cricket news"
        news_category = "sports"
        months = 1

        return search_phrase, news_category, months



#   0 or 1 - only the current month, 2 - current and previous month, 3 - current and two previous months, and so on
def get_month_range(months_back: int):
    today = datetime.today()
    
    # Start of the current month
    start_of_current_month = today.replace(day=1)
    
    # Calculate the start date based on the input
    start_date = start_of_current_month - relativedelta(months=months_back-1)
    
    # End date is the last day of the current month
    end_date = today
    
    return start_date.date(), end_date.date()