import time
import re
from robocorp import browser
from robocorp import workitems
from datetime import datetime
from robocorp.tasks import task
from RPA.Excel.Files import Files

@task
def scrap_lanews():
    """Save in excel each news item that meets the parameters"""
    try:
        with workitems.inputs.reserve() as input_work_item:
            if input_work_item is not None:
                month_number = input_work_item.payload.get(
                    'month_number', 3)

                browser.configure(
                    slowmo=500,
                )
                open_lanews_website()
                search_phrase_and_set_parameters()
                create_excel_file(month_number)
            else:
                print("No input work item available.")
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        print("Cleanup actions (if any) go here.")

def open_lanews_website():
    """Navigates to the given URL"""
    browser.goto("https://www.latimes.com/")
    time.sleep(2)

def search_phrase_and_set_parameters():
    """Parameters to filter the search"""
    page = browser.page()
    page.click("span:text('Show Search')")
    page.fill("span:text('Search Query')", "liberty") #pending to change to Work Item input
    page.click("span:text('Submit Search')")
    page.click("span:text('California')") #pending to change to Work Item input
    page.select_option('select[name="s"]', "1")
    
def extract_found_items():
    """Saves all components of the news item in the item list if it meets the criteria """
    page = browser.page()
    items = []
    li = page.locator("div.promo-wrapper")
    for i in range(li.count()):
        title = li.nth(i).locator("h3.promo-title > a.link").text_content()
        date = li.nth(i).locator("p.promo-timestamp").text_content()
        desc = li.nth(i).locator("p.promo-description").text_content()
        image_src = li.nth(i).locator("img.image").get_attribute('src')
        if money_value(title) or money_value(desc):
            money_amount = True
        else:
            money_amount = False
        items.append((
            title, date, desc, 
            image_src, money_amount
            ))

    return items

def money_value(text):
    """Validates if the title or description contains money values"""
    money_regex = r"\$\d{1,3}(,\d{3})*(\.\d{1,2})?|\d+\s(dollars|USD)"
    return re.search(money_regex, text)

def filter_items_by_date(items, months_back):
    """
    Validates that each date saved in the list
    matches the number of months previous to saving
    """
    current_date = datetime.now()
    filtered_items = []

    for item in items:
        title, date, desc, image_src, money_amount = item
        date_formats = ["%b %d, %Y", "%B %d, %Y", "%b. %d, %Y"]
        for fmt in date_formats:
            try:
                news_date = datetime.strptime(date, fmt)
                break
            except ValueError:
                continue
        delta_months = ((current_date.year 
                        - news_date.year) 
                        * 12 + current_date.month 
                        - news_date.month
                        )
        if delta_months <= months_back:
            filtered_items.append(item)
    
    return filtered_items

def create_excel_file(months_back):
    """Create the excel file with the filtered news"""
    excel = Files()
    excel.create_workbook("output/news_data.xlsx",sheet_name="data")
    headers=[
        "Title", "Date", "Description", 
        "Image filename", "Money amount"
        ]
    excel.append_rows_to_worksheet(
        [headers], 
        header=False)
    items = extract_found_items()
    filtered_items = filter_items_by_date(items, months_back)
    time.sleep(2)
    excel.append_rows_to_worksheet(filtered_items, header=False)
    excel.save_workbook()
    excel.close_workbook()
