from robocorp.tasks import task
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.action_chains import ActionChains
from datetime import datetime
from dateutil.relativedelta import relativedelta
import time, openpyxl, requests, re, sys

def input_variables():
    search = input("Enter the search term: ")
    months = input("Enter the number of months to search: ")
    return search, months

@task
def create_excel():
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet["A1"] = "Title"
    sheet["B1"] = "Description"
    sheet["C1"] = "Date"
    sheet["D1"] = "Image"
    sheet["E1"] = "Dolar"
    sheet["F1"] = "Results"
    browser(wb, sheet)

def browser(wb, sheet):
    search, months = input_variables()
    page = webdriver.Chrome()
    page.get("https://www.latimes.com/")
    page.find_element(By.XPATH, "/html/body/ps-header/header/div[2]/button").click()
    search_element = page.find_element(By.XPATH, "/html/body/ps-header/header/div[2]/div[2]/form/label/input")
    search_element.send_keys(search)
    search_element.send_keys(Keys.RETURN)
    time.sleep(3)
    
    if page.find_element(By.NAME, "s"): 
        select_element = page.find_element(By.NAME, "s")
        select_element.click()
    else:
        sys.exit()
    time.sleep(3)
    select = Select(page.find_element(By.NAME, "s"))
    select.select_by_visible_text("Newest")
    
    title_list = []
    description_list = []
    date_list = []
    image_list = []
    dolar_list = [] 
    hv_mounth = True
    month_str = get_month_numbers(int(months))
    time.sleep(10)
    page_count = 2
    index_image = 0
    while hv_mounth:
        index = 1
        print(f"Page: {page_count}")
        for i in range(1, 11):
            title_element = page.find_element(By.CSS_SELECTOR, f"body > div.page-content > ps-search-results-module > form > div.search-results-module-ajax > ps-search-filters > div > main > ul > li:nth-child({index}) > ps-promo > div > div.promo-content > div > h3 > a")
            title_url = title_element.get_attribute("href")
            date_text = page.find_element(By.CSS_SELECTOR, f"body > div.page-content > ps-search-results-module > form > div.search-results-module-ajax > ps-search-filters > div > main > ul > li:nth-child({index}) > ps-promo > div > div.promo-content > p.promo-timestamp").text
            hv_date, month_url = check_date(title_url)
            if hv_date:
                month_url = month_url[6:8]
                if int(month_url) not in month_str:
                    hv_mounth = False
                    break
            elif hv_date == False:
                if 'ago' in date_text or 'hour' in date_text or 'minute' in date_text:
                    pass
                else:
                    index += 1
                    continue
            
            title_text = title_element.text
            results = page.find_element(By.CSS_SELECTOR, "body > div.page-content > ps-search-results-module > form > div.search-results-module-ajax > ps-search-filters > div > main > div.search-results-module-results-header > div.search-results-module-count > span.search-results-module-count-desktop").text
            results = just_numbers(results)
            descripition_text = page.find_element(By.CSS_SELECTOR, f"body > div.page-content > ps-search-results-module > form > div.search-results-module-ajax > ps-search-filters > div > main > ul > li:nth-child({index}) > ps-promo > div > div.promo-content > p.promo-description").text
            present_dolar_list = ["$", "dollar", "USD"]
            for dolar in present_dolar_list:
                if dolar in descripition_text or dolar in title_text:
                    dolar_list.append("True")
                    break
                else:
                    dolar_list.append("False")
            date_text = page.find_element(By.CSS_SELECTOR, f"body > div.page-content > ps-search-results-module > form > div.search-results-module-ajax > ps-search-filters > div > main > ul > li:nth-child({index}) > ps-promo > div > div.promo-content > p.promo-timestamp").text
            url_image = page.find_element(By.CSS_SELECTOR, f"body > div.page-content > ps-search-results-module > form > div.search-results-module-ajax > ps-search-filters > div > main > ul > li:nth-child({index}) > ps-promo > div > div.promo-media > a > picture > img").get_attribute("src")
            response = requests.get(url_image)
            image_path = f"output/img/image{index_image}.png"
            with open(image_path, "wb") as file:
                file.write(response.content)
            title_list.append(title_text)
            description_list.append(descripition_text)
            date_list.append(date_text)
            image_list.append(image_path)
            index_image += 1
            index += 1
        if page_count == 10:
            hv_mounth = False
        page.get(f"https://www.latimes.com/search?q={search}&s=1&p={page_count}")
        time.sleep(3)
        page_count += 1
    page.quit()
    update_excel(wb, sheet, title_list, description_list, date_list, image_list, dolar_list, results)
    

def update_excel(wb, sheet, title_list, description_list, date_list, image_list, dolar_list, results):

    index = 2
    for i in range(len(title_list)):
        sheet[f"A{index}"] = title_list[i]
        sheet[f"B{index}"] = description_list[i]
        sheet[f"C{index}"] = date_list[i]
        sheet[f"D{index}"] = image_list[i]
        sheet[f"E{index}"] = dolar_list[i]
        sheet[f"F{index}"] = results
        index += 1

    wb.save("output/news.xlsx")


def just_numbers(text):

    match = re.search(r'\b\d{1,3}(?:,\d{3})*\b', text)
    return match.group()

def get_month_numbers(n):

    current_date = datetime.now()
    month_numbers = []

    for i in range(n):
        month_date = current_date - relativedelta(months=i)
        month_number = month_date.month
        month_numbers.append(month_number)

    return month_numbers

def check_date(date):
    hv_date = False
    match = re.search(r'/(\d{4})-(\d{2})-(\d{2})/', date)
    if match:
        hv_date = True
        return hv_date, match.group()
    else:
        return hv_date, None