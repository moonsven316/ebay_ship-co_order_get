import requests
import json
import os
import getpass
import time
import re
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import Select
from datetime import datetime
from ebay_sheet import main
from ebay_sheet_insert import update_or_insert_data
def start_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--disable-extensions")
    options.add_argument("--start-maximized")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--ignore-certificate-errors")
    USER_NAME = getpass.getuser()
    try:
        driver_path = ChromeDriverManager().install()
        service = Service(executable_path=driver_path)
        driver = webdriver.Chrome(service=service, options=options)
    except ValueError:
        url = r"https://googlechromelabs.github.io/chrome-for-testing/last-known-good-versions-with-downloads.json"
        response = requests.get(url)
        data_dict = response.json()
        latest_version = data_dict["channels"]["Stable"]["version"]

        driver_path = ChromeDriverManager(version=latest_version).install()
        service = Service(executable_path=driver_path)
        driver = webdriver.Chrome(service=service, options=options)
    except PermissionError:
        try:
            driver = webdriver.Chrome(
                service=Service(
                    f"C:\\Users\\{USER_NAME}\\.wdm\\drivers\\chromedriver\\win64\\116.0.5845.179\\chromedriver.exe"
                ),
                options=options,
            )
        except:
            driver = webdriver.Chrome(
                service=Service(
                    f"C:\\Users\\{USER_NAME}\\.wdm\\drivers\\chromedriver\\win64\\116.0.5845.96\\chromedriver.exe"
                ),
                options=options,
            )
    return driver
# default value
states = ""
inspection = ""
warehouse = ""
person_charge = ""
shipping_notes = ""
buyer_name = ""
shipping_method = ""
tracking_no = ""
sku = ""
order_number = ""
product_name = ""
ebay_product_name = ""
item_name = ""
shipping_cost = ""
purchase_date = ""
purchase_price = ""
supplier = ""
supplier_url = ""
supplier_url2 = ""
tracking_number = ""
invoice_category = ""
supplier_name = ""
supplier_address = ""
eligible_registration_number = ""
date_sold = ""
sale_price = ""
sales_tax = ""
quantity = ""
shipping = ""
pl = ""
sold_fee = ""
overseas_fee = ""
pn = "2%"
profit = ""
profit_including_refund = ""
refund_amount = ""
return_fee = ""
tracking_info = ""
remarks = ""
# default value end
def windows(ebay_id, ebay_pass, ship_id, ship_pass, product_sheet_id, order_sheet_id):
    driver = webdriver.Chrome()
    # driver.maximize_window()
    ebay_open(driver, ebay_id, ebay_pass)
    ship_co_open(driver, ship_id, ship_pass)
    driver.switch_to.window(driver.window_handles[0])
    driver.get('https://www.ebay.com/sh/ord/?filter=status%3AAWAITING_PAYMENT')
    time.sleep(2)
    
    resultNumber = driver.find_element(By.CLASS_NAME, 'summary-content').text
    result = resultNumber.split(": ")
    if len(result) > 1:
        try:
            result_ = int(result[1].split(" of ")[1])
        except (IndexError, ValueError):
            result_ = 0
    else:
        try:
            result_ = int(result[0])
        except (IndexError, ValueError):
            result_ = 0
    if result_ != 0:
        order_status_divs = driver.find_elements(By.CSS_SELECTOR, "div[class='order-status']")
        tracking_info_list = []
        for order_status_div in order_status_divs:
            tracking_info_item = order_status_div.find_element(By.CLASS_NAME, 'sh-bold').text
            tracking_info_list.append(tracking_info_item)

        order_links = []            
        order_detail_divs = driver.find_elements(By.CSS_SELECTOR, "div[class='order-details']")
        for order_detail_div in order_detail_divs:
            links = order_detail_div.find_elements(By.CSS_SELECTOR, "a")
            for link in links:
                order_links.append(link.get_attribute("href"))

        for i, link in enumerate(order_links):
            driver.get(link)
            time.sleep(2)

            states = "未確認"
            inspection = "支払待ち"

            sku_element = driver.find_element(By.CLASS_NAME, 'lineItemCardInfo__sku')
            sku_elements = sku_element.find_elements(By.CLASS_NAME, 'sh-secondary')
            sku = sku_elements[1].text

            product_name, person_charge = main(sku, product_sheet_id)
            product_name = product_name
            person_charge = person_charge

            order_info = driver.find_element(By.CLASS_NAME, 'order-info').find_element(By.CLASS_NAME,'info-wrapper')
            order_elements = order_info.find_elements(By.CLASS_NAME, 'info-item')
            order_number = order_elements[0].find_element(By.CLASS_NAME, 'info-value').text

            buyer_name = order_elements[2].find_element(By.CSS_SELECTOR, '.info-value.buyer div:first-child').text

            detail_info = driver.find_element(By.CLASS_NAME, 'lineItemCardInfo__content').find_element(By.CLASS_NAME, 'details').find_element(By.TAG_NAME, 'a')
            ebay_product_name = detail_info.find_element(By.CLASS_NAME, 'PSEUDOLINK').text

            date_sold_element = order_elements[1].find_element(By.CLASS_NAME, 'info-value').text
            date_object = datetime.strptime(date_sold_element, "%b %d, %Y")
            date_sold = date_object.strftime("%m月%d日")

            paid_element = driver.find_element(By.CLASS_NAME, 'buyer-paid').find_element(By.CLASS_NAME, 'data-items')
            sale_price_ = paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(1) .amount .value").text
            sale_price = float(re.findall(r'-?\d+(?:,\d+)*(?:\.\d+)?', sale_price_)[0].replace(",",""))
            shipping = 0
            sales_tax = 0
            refund_amount = 0

            quantity_info = driver.find_element(By.CLASS_NAME, 'quantity').find_element(By.CLASS_NAME, 'quantity__value')
            quantity = quantity_info.find_element(By.CLASS_NAME, 'sh-bold').text

            try:
                promoted_listings_element = driver.find_element(By.CSS_SELECTOR, ".lineItemCardInfo__indicators .sh-secondary")
                if "Sold via Promoted Listings" in promoted_listings_element.text:
                    pl = "2.0%"
                else:
                    pl = "0.0%"
            except NoSuchElementException:
                pl = "0.0%"
            
            tracking_info = tracking_info_list[i]
            
            update_or_insert_data(states, inspection, warehouse, person_charge, shipping_notes, buyer_name, shipping_method, tracking_no, sku, order_number, product_name, ebay_product_name, item_name, shipping_cost, purchase_date, purchase_price, supplier, supplier_url, supplier_url2, tracking_number, invoice_category, supplier_name, supplier_address, eligible_registration_number, date_sold, sale_price, sales_tax, quantity, shipping, pl, sold_fee, overseas_fee, pn, profit, profit_including_refund, refund_amount, return_fee, tracking_info, remarks, order_sheet_id)

        awaiting_shipment(driver, product_sheet_id, order_sheet_id)
    else:
        awaiting_shipment(driver, product_sheet_id, order_sheet_id)
def ebay_open(driver, ebay_id, ebay_pass):
    # ebay.com open
    driver.get('https://www.ebay.com/sh/ord')
    time.sleep(20)
    driver.find_element(By.ID, 'userid').send_keys(ebay_id)
    driver.find_element(By.ID, 'signin-continue-btn').click()
    time.sleep(20)
    driver.find_element(By.ID, 'pass').send_keys(ebay_pass)
    driver.find_element(By.ID, 'sgnBt').click()
    time.sleep(5)
    # ebay.com open end
def ship_co_open(driver, ship_id, ship_pass):
    # shop&co open
    driver.execute_script("window.open('');")
    driver.switch_to.window(driver.window_handles[1])
    driver.get('https://app.shipandco.com/shipments')
    time.sleep(2)
    driver.find_element(By.ID, 'email').send_keys(ship_id)
    driver.find_element(By.ID, 'password').send_keys(ship_pass)
    time.sleep(2)
    driver.find_element(By.CSS_SELECTOR, "button[class='btn right btn-plain']").click()
    time.sleep(2)
    # shop&co open end

def awaiting_shipment(driver, product_sheet_id, order_sheet_id):
    driver.get('https://www.ebay.com/sh/ord/?filter=status%3AAWAITING_SHIPMENT')
    time.sleep(2)
    resultNumber = driver.find_element(By.CLASS_NAME, 'summary-content').text
    result = resultNumber.split(": ")
    if len(result) > 1:
        try:
            result_ = int(result[1].split(" of ")[1])
        except (IndexError, ValueError):
            result_ = 0
    else:
        try:
            result_ = int(result[0])
        except (IndexError, ValueError):
            result_ = 0
    if result_ != 0:
        
        order_status_divs = driver.find_elements(By.CSS_SELECTOR, "div[class='order-status']")
        tracking_info_list = []
        for order_status_div in order_status_divs:
            tracking_info_item = order_status_div.find_element(By.CLASS_NAME, 'sh-bold').text
            tracking_info_list.append(tracking_info_item)

        order_links = []            
        order_detail_divs = driver.find_elements(By.CSS_SELECTOR, "div[class='order-details']")
        for order_detail_div in order_detail_divs:
            links = order_detail_div.find_elements(By.CSS_SELECTOR, "a")
            for link in links:
                order_links.append(link.get_attribute("href"))

        for i, link in enumerate(order_links):
            driver.get(link)
            time.sleep(2)

            states = "未確認"
            try:
                buyer_name_element = driver.find_element(By.CLASS_NAME, 'shipping-info').find_element(By.CLASS_NAME, 'address')
                buyer_names = buyer_name_element.find_elements(By.CSS_SELECTOR, "button[class='tooltip__host clickable']")
                buyer_name = buyer_names[0].text
            except NoSuchElementException:
                buyer_name = None
            
            try:
                tracking_info_div = driver.find_element(By.CSS_SELECTOR, "div.tracking-info")
                tracking_no_button = tracking_info_div.find_element(By.CSS_SELECTOR, "button.fake-link")
                tracking_no = tracking_no_button.text
                tracking_no_button.click()
                time.sleep(2)
            except NoSuchElementException:
                tracking_no = ""
            try:
                shipment_package_div = driver.find_element(By.CSS_SELECTOR, "div.shipment-package-grey-container")
                div_element = shipment_package_div.find_element(By.CSS_SELECTOR, "div.block--font-12px.block--color-secondary.block--bottom-margin-8px")
                shipping_method_ = div_element.find_element(By.CSS_SELECTOR, "span:first-child").text
                if shipping_method_ == "FedEx:":
                    shipping_method = "Fedex"
                elif shipping_method_ == "DHL Express International:" or shipping_method_ == "DHL Express Germany:":
                    shipping_method = "DHL"
                elif shipping_method_ == "USPS:":
                    shipping_method = "USPS"
                elif shipping_method_ == "Japan Post:":
                    shipping_method = "日本郵便"
                else:
                    shipping_method = shipping_method_
            except NoSuchElementException:
                shipping_method = ""

            sku_element = driver.find_element(By.CLASS_NAME, 'lineItemCardInfo__sku')
            sku_elements = sku_element.find_elements(By.CLASS_NAME, 'sh-secondary')
            sku = sku_elements[1].text

            product_name, person_charge = main(sku, product_sheet_id)
            product_name = product_name
            person_charge = person_charge

            order_info = driver.find_element(By.CLASS_NAME, 'order-info').find_element(By.CLASS_NAME,'info-wrapper')
            order_elements = order_info.find_elements(By.CLASS_NAME, 'info-item')
            order_number = order_elements[0].find_element(By.CLASS_NAME, 'info-value').text
            
            if buyer_name == None:
                buyer_name = order_elements[4].find_element(By.CSS_SELECTOR, '.info-value.buyer div:first-child').text

            detail_info = driver.find_element(By.CLASS_NAME, 'lineItemCardInfo__content').find_element(By.CLASS_NAME, 'details').find_element(By.TAG_NAME, 'a')
            ebay_product_name = detail_info.find_element(By.CLASS_NAME, 'PSEUDOLINK').text

            date_sold_element = order_elements[2].find_element(By.CLASS_NAME, 'info-value').text
            date_object = datetime.strptime(date_sold_element, "%b %d, %Y")
            date_sold = date_object.strftime("%m月%d日")

            paid_element = driver.find_element(By.CLASS_NAME, 'buyer-paid').find_element(By.CLASS_NAME, 'data-items')
            paid_elements = paid_element.find_elements(By.CLASS_NAME, 'level-2')
            if paid_elements[0].find_element(By.TAG_NAME, 'dt').text == "Item subtotal":
                sale_price_ = paid_elements[0].find_element(By.TAG_NAME, 'dd').text
                sale_price = float(re.findall(r'-?\d+(?:,\d+)*(?:\.\d+)?', sale_price_)[0])
            else:
                sale_price = 0
            try:
                if paid_elements[1].find_element(By.TAG_NAME, 'dt').text == "Shipping":
                    shipping_ = paid_elements[1].find_element(By.TAG_NAME, 'dd').text
                    shipping = float(re.findall(r'-?\d+(?:,\d+)*(?:\.\d+)?', shipping_)[0])
                else :
                    shipping = 0
            except NoSuchElementException:
                shipping = 0
            sales_tax = 0
            refund_amount = 0

            quantity_info = driver.find_element(By.CLASS_NAME, 'quantity').find_element(By.CLASS_NAME, 'quantity__value')
            quantity = quantity_info.find_element(By.CLASS_NAME, 'sh-bold').text

            try:
                promoted_listings_element = driver.find_element(By.CSS_SELECTOR, ".lineItemCardInfo__indicators .sh-secondary")
                if "Sold via Promoted Listings" in promoted_listings_element.text:
                    pl = "2.0%"
                else:
                    pl = "0.0%"
            except NoSuchElementException:
                pl = "0.0%"

            detail_link = driver.find_element(By.CSS_SELECTOR, ".section-action a")
            event = detail_link.get_attribute("href")
            driver.get(event)
            time.sleep(2)
            link_elements = driver.find_elements(By.CSS_SELECTOR, "a.eui-textual-display")
            if len(link_elements) == 2:
                link = link_elements[1].get_attribute("href")
            elif len(link_elements) == 1:
                link = link_elements[0].get_attribute("href")
            else:
                link = link_elements[-1].get_attribute("href")
            driver.get(link)
            time.sleep(2)

            price_table = driver.find_element(By.CLASS_NAME, 'fee-breakdown--fee-breakdown-table').find_elements(By.TAG_NAME, 'section')
            try:
                sold_element = price_table[0].find_element(By.CSS_SELECTOR, ".fee-label-value-item--column.third span:nth-child(3) .text-span")
                sold_fee_ = sold_element.text.split('%')[0].strip()
                sold_fee = f"{sold_fee_}%"
            except NoSuchElementException:
                sold_fee = ""
            fee_percentage_elements = price_table[1].find_elements(By.CSS_SELECTOR, ".fee-label-value-item--column.third span:nth-child(3) .text-span")
            fee_percentages = []
            for element in fee_percentage_elements:
                text = element.text
                if "%" in text:
                    numeric_part = text.split("%")[0].strip()
                    fee_percentages.append(float(numeric_part))
                else:
                    fee_percentages.append(0)
            fee_percentage_difference = fee_percentages[0] - fee_percentages[1]
            overseas_fee = f"{fee_percentage_difference}%"

            tracking_info = tracking_info_list[i]
            
            update_or_insert_data(states, inspection, warehouse, person_charge, shipping_notes, buyer_name, shipping_method, tracking_no, sku, order_number, product_name, ebay_product_name, item_name, shipping_cost, purchase_date, purchase_price, supplier, supplier_url, supplier_url2, tracking_number, invoice_category, supplier_name, supplier_address, eligible_registration_number, date_sold, sale_price, sales_tax, quantity, shipping, pl, sold_fee, overseas_fee, pn, profit, profit_including_refund, refund_amount, return_fee, tracking_info, remarks, order_sheet_id)

        awaiting_shipment_overdue(driver, product_sheet_id, order_sheet_id)
    else:
        awaiting_shipment_overdue(driver, product_sheet_id, order_sheet_id)
    # 
def awaiting_shipment_overdue(driver, product_sheet_id, order_sheet_id):
    driver.get('https://www.ebay.com/sh/ord/?filter=status%3AAWAITING_SHIPMENT_OVERDUE')
    time.sleep(2)
    resultNumber = driver.find_element(By.CLASS_NAME, 'summary-content').text
    result = resultNumber.split(": ")
    if len(result) > 1:
        try:
            result_ = int(result[1].split(" of ")[1])
        except (IndexError, ValueError):
            result_ = 0
    else:
        try:
            result_ = int(result[0])
        except (IndexError, ValueError):
            result_ = 0
    if result_ != 0:
        order_status_divs = driver.find_elements(By.CSS_SELECTOR, "div[class='order-status']")
        tracking_info_list = []
        for order_status_div in order_status_divs:
            tracking_info_item = order_status_div.text
            tracking_info_list.append(tracking_info_item)

        order_links = []            
        order_detail_divs = driver.find_elements(By.CSS_SELECTOR, "div[class='order-details']")
        for order_detail_div in order_detail_divs:
            links = order_detail_div.find_elements(By.CSS_SELECTOR, "a")
            for link in links:
                order_links.append(link.get_attribute("href"))

        for i, link in enumerate(order_links):
            driver.get(link)
            time.sleep(2)

            states = "未確認"

            buyer_name_element = driver.find_element(By.CLASS_NAME, 'shipping-info').find_element(By.CLASS_NAME, 'address')
            buyer_names = buyer_name_element.find_elements(By.CSS_SELECTOR, "button[class='tooltip__host clickable']")
            buyer_name = buyer_names[0].text
            try:
                tracking_info_div = driver.find_element(By.CSS_SELECTOR, "div.tracking-info")
                tracking_no_button = tracking_info_div.find_element(By.CSS_SELECTOR, "button.fake-link")
                tracking_no = tracking_no_button.text
                tracking_no_button.click()
                time.sleep(2)
            except NoSuchElementException:
                tracking_no = ""
            try:
                shipment_package_div = driver.find_element(By.CSS_SELECTOR, "div.shipment-package-grey-container")
                div_element = shipment_package_div.find_element(By.CSS_SELECTOR, "div.block--font-12px.block--color-secondary.block--bottom-margin-8px")
                shipping_method_ = div_element.find_element(By.CSS_SELECTOR, "span:first-child").text
                if shipping_method_ == "FedEx:":
                    shipping_method = "Fedex"
                elif shipping_method_ == "DHL Express International:" or shipping_method_ == "DHL Express Germany:":
                    shipping_method = "DHL"
                elif shipping_method_ == "USPS:":
                    shipping_method = "USPS"
                elif shipping_method_ == "Japan Post:":
                    shipping_method = "日本郵便"
                else:
                    shipping_method = shipping_method_
            except NoSuchElementException:
                shipping_method = ""

            sku_element = driver.find_element(By.CLASS_NAME, 'lineItemCardInfo__sku')
            sku_elements = sku_element.find_elements(By.CLASS_NAME, 'sh-secondary')
            sku = sku_elements[1].text

            product_name, person_charge = main(sku, product_sheet_id)
            product_name = product_name
            person_charge = person_charge

            order_info = driver.find_element(By.CLASS_NAME, 'order-info').find_element(By.CLASS_NAME,'info-wrapper')
            order_elements = order_info.find_elements(By.CLASS_NAME, 'info-item')
            order_number = order_elements[0].find_element(By.CLASS_NAME, 'info-value').text

            detail_info = driver.find_element(By.CLASS_NAME, 'lineItemCardInfo__content').find_element(By.CLASS_NAME, 'details').find_element(By.TAG_NAME, 'a')
            ebay_product_name = detail_info.find_element(By.CLASS_NAME, 'PSEUDOLINK').text

            date_sold_element = order_elements[2].find_element(By.CLASS_NAME, 'info-value').text
            date_object = datetime.strptime(date_sold_element, "%b %d, %Y")
            date_sold = date_object.strftime("%m月%d日")
            
            paid_element = driver.find_element(By.CLASS_NAME, 'buyer-paid').find_element(By.CLASS_NAME, 'data-items')
            paid_elements = paid_element.find_elements(By.CLASS_NAME, 'level-2')
            if paid_elements[0].find_element(By.TAG_NAME, 'dt').text == "Item subtotal":
                sale_price_ = paid_elements[0].find_element(By.TAG_NAME, 'dd').text
                sale_price = float(re.findall(r'-?\d+(?:,\d+)*(?:\.\d+)?', sale_price_)[0])
            else:
                sale_price = 0
            try:
                if paid_elements[1].find_element(By.TAG_NAME, 'dt').text == "Shipping":
                    shipping_ = paid_elements[1].find_element(By.TAG_NAME, 'dd').text
                    shipping = float(re.findall(r'-?\d+(?:,\d+)*(?:\.\d+)?', shipping_)[0])
                else :
                    shipping = 0
            except NoSuchElementException:
                shipping = 0
            sales_tax = 0
            refund_amount = 0

            quantity_info = driver.find_element(By.CLASS_NAME, 'quantity').find_element(By.CLASS_NAME, 'quantity__value')
            quantity = quantity_info.find_element(By.CLASS_NAME, 'sh-bold').text

            try:
                promoted_listings_element = driver.find_element(By.CSS_SELECTOR, ".lineItemCardInfo__indicators .sh-secondary")
                if "Sold via Promoted Listings" in promoted_listings_element.text:
                    pl = "2.0%"
                else:
                    pl = "0.0%"
            except NoSuchElementException:
                pl = "0.0%"

            detail_link = driver.find_element(By.CSS_SELECTOR, ".section-action a")
            event = detail_link.get_attribute("href")
            driver.get(event)
            time.sleep(2)
            link_elements = driver.find_elements(By.CSS_SELECTOR, "a.eui-textual-display")
            if len(link_elements) == 2:
                link = link_elements[1].get_attribute("href")
            elif len(link_elements) == 1:
                link = link_elements[0].get_attribute("href")
            else:
                link = link_elements[-1].get_attribute("href")
            driver.get(link)
            time.sleep(2)

            price_table = driver.find_element(By.CLASS_NAME, 'fee-breakdown--fee-breakdown-table').find_elements(By.TAG_NAME, 'section')
            try:
                sold_element = price_table[0].find_element(By.CSS_SELECTOR, ".fee-label-value-item--column.third span:nth-child(3) .text-span")
                sold_fee_ = sold_element.text.split('%')[0].strip()
                sold_fee = f"{sold_fee_}%"
            except NoSuchElementException:
                sold_fee = ""
            fee_percentage_elements = price_table[1].find_elements(By.CSS_SELECTOR, ".fee-label-value-item--column.third span:nth-child(3) .text-span")
            fee_percentages = []
            for element in fee_percentage_elements:
                text = element.text
                if "%" in text:
                    numeric_part = text.split("%")[0].strip()
                    fee_percentages.append(float(numeric_part))
                else:
                    fee_percentages.append(0)
            fee_percentage_difference = fee_percentages[0] - fee_percentages[1]
            overseas_fee = f"{fee_percentage_difference}%"

            tracking_info = tracking_info_list[i]

            update_or_insert_data(states, inspection, warehouse, person_charge, shipping_notes, buyer_name, shipping_method, tracking_no, sku, order_number, product_name, ebay_product_name, item_name, shipping_cost, purchase_date, purchase_price, supplier, supplier_url, supplier_url2, tracking_number, invoice_category, supplier_name, supplier_address, eligible_registration_number, date_sold, sale_price, sales_tax, quantity, shipping, pl, sold_fee, overseas_fee, pn, profit, profit_including_refund, refund_amount, return_fee, tracking_info, remarks, order_sheet_id)

        awaiting_shipment_within(driver, product_sheet_id, order_sheet_id)
    else:
        awaiting_shipment_within(driver, product_sheet_id, order_sheet_id)
def awaiting_shipment_within(driver, product_sheet_id, order_sheet_id):
    driver.get('https://www.ebay.com/sh/ord/?filter=status%3AAWAITING_SHIPMENT_DUE_24H')
    time.sleep(2)
    resultNumber = driver.find_element(By.CLASS_NAME, 'summary-content').text
    result = resultNumber.split(": ")
    if len(result) > 1:
        try:
            result_ = int(result[1].split(" of ")[1])
        except (IndexError, ValueError):
            result_ = 0
    else:
        try:
            result_ = int(result[0])
        except (IndexError, ValueError):
            result_ = 0
    if result_ != 0:
        order_status_divs = driver.find_elements(By.CSS_SELECTOR, "div[class='order-status']")
        tracking_info_list = []
        for order_status_div in order_status_divs:
            tracking_info_item = order_status_div.find_element(By.CLASS_NAME, 'sh-bold').text
            tracking_info_list.append(tracking_info_item)

        order_links = []            
        order_detail_divs = driver.find_elements(By.CSS_SELECTOR, "div[class='order-details']")
        for order_detail_div in order_detail_divs:
            links = order_detail_div.find_elements(By.CSS_SELECTOR, "a")
            for link in links:
                order_links.append(link.get_attribute("href"))

        for i, link in enumerate(order_links):
            driver.get(link)
            time.sleep(2)

            states = "未確認"

            buyer_name_element = driver.find_element(By.CLASS_NAME, 'shipping-info').find_element(By.CLASS_NAME, 'address')
            buyer_names = buyer_name_element.find_elements(By.CSS_SELECTOR, "button[class='tooltip__host clickable']")
            buyer_name = buyer_names[0].text
            try:
                tracking_info_div = driver.find_element(By.CSS_SELECTOR, "div.tracking-info")
                tracking_no_button = tracking_info_div.find_element(By.CSS_SELECTOR, "button.fake-link")
                tracking_no = tracking_no_button.text
                tracking_no_button.click()
                time.sleep(2)
            except NoSuchElementException:
                tracking_no = ""
            try:
                shipment_package_div = driver.find_element(By.CSS_SELECTOR, "div.shipment-package-grey-container")
                div_element = shipment_package_div.find_element(By.CSS_SELECTOR, "div.block--font-12px.block--color-secondary.block--bottom-margin-8px")
                shipping_method_ = div_element.find_element(By.CSS_SELECTOR, "span:first-child").text
                if shipping_method_ == "FedEx:":
                    shipping_method = "Fedex"
                elif shipping_method_ == "DHL Express International:" or shipping_method_ == "DHL Express Germany:":
                    shipping_method = "DHL"
                elif shipping_method_ == "USPS:":
                    shipping_method = "USPS"
                elif shipping_method_ == "Japan Post:":
                    shipping_method = "日本郵便"
                else:
                    shipping_method = shipping_method_
            except NoSuchElementException:
                shipping_method = ""

            sku_element = driver.find_element(By.CLASS_NAME, 'lineItemCardInfo__sku')
            sku_elements = sku_element.find_elements(By.CLASS_NAME, 'sh-secondary')
            sku = sku_elements[1].text

            product_name, person_charge = main(sku, product_sheet_id)
            product_name = product_name
            person_charge = person_charge

            order_info = driver.find_element(By.CLASS_NAME, 'order-info').find_element(By.CLASS_NAME,'info-wrapper')
            order_elements = order_info.find_elements(By.CLASS_NAME, 'info-item')
            order_number = order_elements[0].find_element(By.CLASS_NAME, 'info-value').text

            detail_info = driver.find_element(By.CLASS_NAME, 'lineItemCardInfo__content').find_element(By.CLASS_NAME, 'details').find_element(By.TAG_NAME, 'a')
            ebay_product_name = detail_info.find_element(By.CLASS_NAME, 'PSEUDOLINK').text

            date_sold_element = order_elements[2].find_element(By.CLASS_NAME, 'info-value').text
            date_object = datetime.strptime(date_sold_element, "%b %d, %Y")
            date_sold = date_object.strftime("%m月%d日")
            
            paid_element = driver.find_element(By.CLASS_NAME, 'buyer-paid').find_element(By.CLASS_NAME, 'data-items')
            paid_elements = paid_element.find_elements(By.CLASS_NAME, 'level-2')
            if paid_elements[0].find_element(By.TAG_NAME, 'dt').text == "Item subtotal":
                sale_price_ = paid_elements[0].find_element(By.TAG_NAME, 'dd').text
                sale_price = float(re.findall(r'-?\d+(?:,\d+)*(?:\.\d+)?', sale_price_)[0])
            else:
                sale_price = 0
            try:
                if paid_elements[1].find_element(By.TAG_NAME, 'dt').text == "Shipping":
                    shipping_ = paid_elements[1].find_element(By.TAG_NAME, 'dd').text
                    shipping = float(re.findall(r'-?\d+(?:,\d+)*(?:\.\d+)?', shipping_)[0])
                else :
                    shipping = 0
            except NoSuchElementException:
                shipping = 0
            sales_tax = 0
            refund_amount = 0

            quantity_info = driver.find_element(By.CLASS_NAME, 'quantity').find_element(By.CLASS_NAME, 'quantity__value')
            quantity = quantity_info.find_element(By.CLASS_NAME, 'sh-bold').text

            try:
                promoted_listings_element = driver.find_element(By.CSS_SELECTOR, ".lineItemCardInfo__indicators .sh-secondary")
                if "Sold via Promoted Listings" in promoted_listings_element.text:
                    pl = "2.0%"
                else:
                    pl = "0.0%"
            except NoSuchElementException:
                pl = "0.0%"

            detail_link = driver.find_element(By.CSS_SELECTOR, ".section-action a")
            event = detail_link.get_attribute("href")
            driver.get(event)
            time.sleep(2)
            link_elements = driver.find_elements(By.CSS_SELECTOR, "a.eui-textual-display")
            if len(link_elements) == 2:
                link = link_elements[1].get_attribute("href")
            elif len(link_elements) == 1:
                link = link_elements[0].get_attribute("href")
            else:
                link = link_elements[-1].get_attribute("href")
            driver.get(link)
            time.sleep(2)

            price_table = driver.find_element(By.CLASS_NAME, 'fee-breakdown--fee-breakdown-table').find_elements(By.TAG_NAME, 'section')
            try:
                sold_element = price_table[0].find_element(By.CSS_SELECTOR, ".fee-label-value-item--column.third span:nth-child(3) .text-span")
                sold_fee_ = sold_element.text.split('%')[0].strip()
                sold_fee = f"{sold_fee_}%"
            except NoSuchElementException:
                sold_fee = ""
            fee_percentage_elements = price_table[1].find_elements(By.CSS_SELECTOR, ".fee-label-value-item--column.third span:nth-child(3) .text-span")
            fee_percentages = []
            for element in fee_percentage_elements:
                text = element.text
                if "%" in text:
                    numeric_part = text.split("%")[0].strip()
                    fee_percentages.append(float(numeric_part))
                else:
                    fee_percentages.append(0)
            fee_percentage_difference = fee_percentages[0] - fee_percentages[1]
            overseas_fee = f"{fee_percentage_difference}%"

            tracking_info = tracking_info_list[i]

            update_or_insert_data(states, inspection, warehouse, person_charge, shipping_notes, buyer_name, shipping_method, tracking_no, sku, order_number, product_name, ebay_product_name, item_name, shipping_cost, purchase_date, purchase_price, supplier, supplier_url, supplier_url2, tracking_number, invoice_category, supplier_name, supplier_address, eligible_registration_number, date_sold, sale_price, sales_tax, quantity, shipping, pl, sold_fee, overseas_fee, pn, profit, profit_including_refund, refund_amount, return_fee, tracking_info, remarks, order_sheet_id)
        awaiting_expedited_shipment(driver, product_sheet_id, order_sheet_id)
    else:
        awaiting_expedited_shipment(driver, product_sheet_id, order_sheet_id)
def awaiting_expedited_shipment(driver, product_sheet_id, order_sheet_id):
    driver.get('https://www.ebay.com/sh/ord/?filter=status%3AAWAITING_EXPEDITED_SHIPMENT')
    time.sleep(2)
    resultNumber = driver.find_element(By.CLASS_NAME, 'summary-content').text
    result = resultNumber.split(": ")
    if len(result) > 1:
        try:
            result_ = int(result[1].split(" of ")[1])
        except (IndexError, ValueError):
            result_ = 0
    else:
        try:
            result_ = int(result[0])
        except (IndexError, ValueError):
            result_ = 0
    if result_ != 0:
        order_status_divs = driver.find_elements(By.CSS_SELECTOR, "div[class='order-status']")
        tracking_info_list = []
        for order_status_div in order_status_divs:
            tracking_info_item = order_status_div.find_element(By.CLASS_NAME, 'sh-bold').text
            tracking_info_list.append(tracking_info_item)

        order_links = []            
        order_detail_divs = driver.find_elements(By.CSS_SELECTOR, "div[class='order-details']")
        for order_detail_div in order_detail_divs:
            links = order_detail_div.find_elements(By.CSS_SELECTOR, "a")
            for link in links:
                order_links.append(link.get_attribute("href"))

        for i, link in enumerate(order_links):
            driver.get(link)
            time.sleep(2)

            states = "未確認"

            buyer_name_element = driver.find_element(By.CLASS_NAME, 'shipping-info').find_element(By.CLASS_NAME, 'address')
            buyer_names = buyer_name_element.find_elements(By.CSS_SELECTOR, "button[class='tooltip__host clickable']")
            buyer_name = buyer_names[0].text
            try:
                tracking_info_div = driver.find_element(By.CSS_SELECTOR, "div.tracking-info")
                tracking_no_button = tracking_info_div.find_element(By.CSS_SELECTOR, "button.fake-link")
                tracking_no = tracking_no_button.text
                tracking_no_button.click()
                time.sleep(2)
            except NoSuchElementException:
                tracking_no = ""
            try:
                shipment_package_div = driver.find_element(By.CSS_SELECTOR, "div.shipment-package-grey-container")
                div_element = shipment_package_div.find_element(By.CSS_SELECTOR, "div.block--font-12px.block--color-secondary.block--bottom-margin-8px")
                shipping_method_ = div_element.find_element(By.CSS_SELECTOR, "span:first-child").text
                if shipping_method_ == "FedEx:":
                    shipping_method = "Fedex"
                elif shipping_method_ == "DHL Express International:" or shipping_method_ == "DHL Express Germany:":
                    shipping_method = "DHL"
                elif shipping_method_ == "USPS:":
                    shipping_method = "USPS"
                elif shipping_method_ == "Japan Post:":
                    shipping_method = "日本郵便"
                else:
                    shipping_method = shipping_method_
            except NoSuchElementException:
                shipping_method = ""

            sku_element = driver.find_element(By.CLASS_NAME, 'lineItemCardInfo__sku')
            sku_elements = sku_element.find_elements(By.CLASS_NAME, 'sh-secondary')
            sku = sku_elements[1].text

            product_name, person_charge = main(sku, product_sheet_id)
            product_name = product_name
            person_charge = person_charge

            order_info = driver.find_element(By.CLASS_NAME, 'order-info').find_element(By.CLASS_NAME,'info-wrapper')
            order_elements = order_info.find_elements(By.CLASS_NAME, 'info-item')
            order_number = order_elements[0].find_element(By.CLASS_NAME, 'info-value').text

            detail_info = driver.find_element(By.CLASS_NAME, 'lineItemCardInfo__content').find_element(By.CLASS_NAME, 'details').find_element(By.TAG_NAME, 'a')
            ebay_product_name = detail_info.find_element(By.CLASS_NAME, 'PSEUDOLINK').text

            date_sold_element = order_elements[2].find_element(By.CLASS_NAME, 'info-value').text
            date_object = datetime.strptime(date_sold_element, "%b %d, %Y")
            date_sold = date_object.strftime("%m月%d日")
            
            paid_element = driver.find_element(By.CLASS_NAME, 'buyer-paid').find_element(By.CLASS_NAME, 'data-items')
            paid_elements = paid_element.find_elements(By.CLASS_NAME, 'level-2')
            if paid_elements[0].find_element(By.TAG_NAME, 'dt').text == "Item subtotal":
                sale_price_ = paid_elements[0].find_element(By.TAG_NAME, 'dd').text
                sale_price = float(re.findall(r'-?\d+(?:,\d+)*(?:\.\d+)?', sale_price_)[0])
            else:
                sale_price = 0
            try:
                if paid_elements[1].find_element(By.TAG_NAME, 'dt').text == "Shipping":
                    shipping_ = paid_elements[1].find_element(By.TAG_NAME, 'dd').text
                    shipping = float(re.findall(r'-?\d+(?:,\d+)*(?:\.\d+)?', shipping_)[0])
                else :
                    shipping = 0
            except NoSuchElementException:
                shipping = 0
            sales_tax = 0
            refund_amount = 0

            quantity_info = driver.find_element(By.CLASS_NAME, 'quantity').find_element(By.CLASS_NAME, 'quantity__value')
            quantity = quantity_info.find_element(By.CLASS_NAME, 'sh-bold').text

            try:
                promoted_listings_element = driver.find_element(By.CSS_SELECTOR, ".lineItemCardInfo__indicators .sh-secondary")
                if "Sold via Promoted Listings" in promoted_listings_element.text:
                    pl = "2.0%"
                else:
                    pl = "0.0%"
            except NoSuchElementException:
                pl = "0.0%"

            detail_link = driver.find_element(By.CSS_SELECTOR, ".section-action a")
            event = detail_link.get_attribute("href")
            driver.get(event)
            time.sleep(2)
            link_elements = driver.find_elements(By.CSS_SELECTOR, "a.eui-textual-display")
            if len(link_elements) == 2:
                link = link_elements[1].get_attribute("href")
            elif len(link_elements) == 1:
                link = link_elements[0].get_attribute("href")
            else:
                link = link_elements[-1].get_attribute("href")
            driver.get(link)
            time.sleep(2)

            price_table = driver.find_element(By.CLASS_NAME, 'fee-breakdown--fee-breakdown-table').find_elements(By.TAG_NAME, 'section')
            try:
                sold_element = price_table[0].find_element(By.CSS_SELECTOR, ".fee-label-value-item--column.third span:nth-child(3) .text-span")
                sold_fee_ = sold_element.text.split('%')[0].strip()
                sold_fee = f"{sold_fee_}%"
            except NoSuchElementException:
                sold_fee = ""
            fee_percentage_elements = price_table[1].find_elements(By.CSS_SELECTOR, ".fee-label-value-item--column.third span:nth-child(3) .text-span")
            fee_percentages = []
            for element in fee_percentage_elements:
                text = element.text
                if "%" in text:
                    numeric_part = text.split("%")[0].strip()
                    fee_percentages.append(float(numeric_part))
                else:
                    fee_percentages.append(0)
            fee_percentage_difference = fee_percentages[0] - fee_percentages[1]
            overseas_fee = f"{fee_percentage_difference}%"

            tracking_info = tracking_info_list[i]

            update_or_insert_data(states, inspection, warehouse, person_charge, shipping_notes, buyer_name, shipping_method, tracking_no, sku, order_number, product_name, ebay_product_name, item_name, shipping_cost, purchase_date, purchase_price, supplier, supplier_url, supplier_url2, tracking_number, invoice_category, supplier_name, supplier_address, eligible_registration_number, date_sold, sale_price, sales_tax, quantity, shipping, pl, sold_fee, overseas_fee, pn, profit, profit_including_refund, refund_amount, return_fee, tracking_info, remarks, order_sheet_id)
        paid_shipment(driver, product_sheet_id, order_sheet_id)
    else:
        paid_shipment(driver, product_sheet_id, order_sheet_id)
    
def paid_shipment(driver, product_sheet_id, order_sheet_id):
    driver.get('https://www.ebay.com/sh/ord/?filter=status%3APAID_SHIPPED')
    time.sleep(2)
    resultNumber = driver.find_element(By.CLASS_NAME, 'summary-content').text
    result = resultNumber.split(": ")
    if len(result) > 1:
        try:
            result_ = int(result[1].split(" of ")[1])
        except (IndexError, ValueError):
            result_ = 0
    else:
        try:
            result_ = int(result[0])
        except (IndexError, ValueError):
            result_ = 0
    if result_ != 0:
        pagination_links = []
        pagination_items = driver.find_elements(By.CSS_SELECTOR, ".pagination__items li")
        for pagination_item in pagination_items:
            page_link = pagination_item.find_element(By.CLASS_NAME, "pagination__item")
            pagination_links.append(page_link.get_attribute("href"))
        for pagination_link in pagination_links:
            driver.get(str(pagination_link))
            time.sleep(10)

            order_status_divs = driver.find_elements(By.CSS_SELECTOR, "div[class='order-status']")
            tracking_info_list = []
            for order_status_div in order_status_divs:
                tracking_info_item_ = order_status_div.find_element(By.CLASS_NAME, 'sh-bold').text
                if tracking_info_item_ == "Shipping overdue":
                    tracking_info_item = order_status_divs.text
                else:
                    tracking_info_item = tracking_info_item_
                tracking_info_list.append(tracking_info_item)

            order_links = []            
            order_detail_divs = driver.find_elements(By.CSS_SELECTOR, "div[class='order-details']")
            for order_detail_div in order_detail_divs:
                links = order_detail_div.find_elements(By.CSS_SELECTOR, "a")
                for link in links:
                    order_links.append(link.get_attribute("href"))

            for i, link in enumerate(order_links):
                driver.get(link)
                time.sleep(2)
                
                states = "発送済"

                buyer_name_element = driver.find_element(By.CLASS_NAME, 'shipping-info').find_element(By.CLASS_NAME, 'address')
                buyer_names = buyer_name_element.find_elements(By.CSS_SELECTOR, "button[class='tooltip__host clickable']")
                buyer_name = buyer_names[0].text

                sku_element = driver.find_element(By.CLASS_NAME, 'lineItemCardInfo__sku')
                sku_elements = sku_element.find_elements(By.CLASS_NAME, 'sh-secondary')
                sku = sku_elements[1].text

                product_name, person_charge = main(sku, product_sheet_id)
                product_name = product_name
                person_charge = person_charge

                order_info = driver.find_element(By.CLASS_NAME, 'order-info').find_element(By.CLASS_NAME,'info-wrapper')
                order_elements = order_info.find_elements(By.CLASS_NAME, 'info-item')
                order_number = order_elements[0].find_element(By.CLASS_NAME, 'info-value').text

                detail_info = driver.find_element(By.CLASS_NAME, 'lineItemCardInfo__content').find_element(By.CLASS_NAME, 'details').find_element(By.TAG_NAME, 'a')
                ebay_product_name = detail_info.find_element(By.CLASS_NAME,  'PSEUDOLINK').text

                date_sold_element = order_elements[2].find_element(By.CLASS_NAME, 'info-value').text
                date_object = datetime.strptime(date_sold_element, "%b %d, %Y")
                date_sold = date_object.strftime("%m月%d日")
                
                paid_element = driver.find_element(By.CLASS_NAME, 'buyer-paid')
                if paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(1) .label.level-2").text == "Item subtotal":
                    sale_price_ = paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(1) .amount .value").text
                    sale_price = float(re.findall(r'-?\d+(?:,\d+)*(?:\.\d+)?', sale_price_)[0].replace(",",""))
                else:
                    sale_price = 0
                try:
                    if paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(2) .label.level-2").text == "Shipping":
                        shipping_ = paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(2) .amount .value").text
                        shipping = float(re.findall(r'-?\d+(?:,\d+)*(?:\.\d+)?', shipping_)[0])
                    else :
                        shipping = 0
                except NoSuchElementException:
                    shipping = 0
                try:
                    if paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(3) .label.level-2").text == "Sales tax*" or paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(3) .label.level-2").text == "Import VAT*":
                        sales_tax_ = paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(3) .amount .value").text
                        sales_tax = float(re.findall(r'-?\d+(?:,\d+)*(?:\.\d+)?', sales_tax_)[0])
                    else:
                        sales_tax = 0
                except NoSuchElementException:
                    sales_tax = 0
                try:
                    if paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(3) .label.level-2").text == "Refund" or paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(3) .label.level-2").text == "Order canceled":
                        refund_amount_ = paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(3) .amount .value").text
                        _refund_amount = float(re.findall(r'-?\d+(?:,\d+)*(?:\.\d+)?', refund_amount_)[0].replace(",", ""))
                    else:
                        _refund_amount = None
                except NoSuchElementException:
                    _refund_amount = None
                if _refund_amount == None:
                    try:
                        if paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(4) .label.level-2").text == "Refund" or paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(4) .label.level-2").text == "Order canceled":
                            refund_amount_ = paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(4) .amount .value").text
                            _refund_amount = float(re.findall(r'-?\d+(?:,\d+)*(?:\.\d+)?', refund_amount_)[0].replace(",", ""))
                        else:
                            _refund_amount = 0.0
                    except NoSuchElementException:
                        _refund_amount = 0.0
                try:
                    ebay_paid_element = driver.find_element(By.CLASS_NAME, 'earnings')
                    ebay_paid_elements = ebay_paid_element.find_elements(By.CLASS_NAME, 'data-item')
                    if ebay_paid_elements[1].find_element(By.TAG_NAME, 'dt') == "Refund (eBay paid)":
                        ebay_paid = float(ebay_paid_elements[1].find_element(By.TAG_NAME, 'dd').text.strip('$'))
                    else:
                        ebay_paid = 0.0
                except (NoSuchElementException, IndexError, ValueError, TimeoutException):
                    ebay_paid = 0.0
                print(f"refund: {_refund_amount}")
                print(f"ebay paid: {ebay_paid}")
                refund_amount = _refund_amount - ebay_paid

                quantity_info = driver.find_element(By.CLASS_NAME, 'quantity').find_element(By.CLASS_NAME, 'quantity__value')
                quantity = quantity_info.find_element(By.CLASS_NAME, 'sh-bold').text

                try:
                    promoted_listings_element = driver.find_element(By.CSS_SELECTOR, ".lineItemCardInfo__indicators .sh-secondary")
                    if "Sold via Promoted Listings" in promoted_listings_element.text:
                        pl = "2.0%"
                    else:
                        pl = "0.0%"
                except NoSuchElementException:
                    pl = "0.0%"

                detail_link = driver.find_element(By.CSS_SELECTOR, ".section-action a")
                event = detail_link.get_attribute("href")
                driver.get(event)
                time.sleep(2)

                link_elements = driver.find_elements(By.CSS_SELECTOR, "a.eui-textual-display")
                try:
                    if len(link_elements) == 2:
                        sold_link = link_elements[1].get_attribute("href")
                    elif len(link_elements) == 1:
                        sold_link = link_elements[0].get_attribute("href")
                    else:
                        sold_link = link_elements[-1].get_attribute("href")
                    driver.get(sold_link)
                    time.sleep(2)

                    price_table = driver.find_element(By.CLASS_NAME, 'fee-breakdown--fee-breakdown-table').find_elements(By.TAG_NAME, 'section')
                    try:
                        sold_element = price_table[0].find_element(By.CSS_SELECTOR, ".fee-label-value-item--column.third span:nth-child(3) .text-span")
                        sold_fee_ = sold_element.text.split('%')[0].strip()
                        sold_fee = f"{sold_fee_}%"
                    except NoSuchElementException:
                        sold_fee = ""
                    fee_percentage_elements = price_table[1].find_elements(By.CSS_SELECTOR, ".fee-label-value-item--column.third span:nth-child(3) .text-span")
                    fee_percentages = []
                    for element in fee_percentage_elements:
                        text = element.text
                        if "%" in text:
                            numeric_part = text.split("%")[0].strip()
                            fee_percentages.append(float(numeric_part))
                        else:
                            fee_percentages.append(0)
                    fee_percentage_difference = fee_percentages[0] - fee_percentages[1]
                    overseas_fee = f"{fee_percentage_difference}%"
                except (NoSuchElementException, IndexError, TypeError):
                    break

                tracking_info = tracking_info_list[i]

                driver.get(link)
                time.sleep(2)
                try:
                    tracking_info_div = driver.find_element(By.CSS_SELECTOR, "div.tracking-info")
                    tracking_no_button = tracking_info_div.find_element(By.CSS_SELECTOR, "button.fake-link")
                    tracking_no = tracking_no_button.text
                    tracking_no_button.click()
                    time.sleep(2)
                except NoSuchElementException:
                    tracking_no = ""
                try:
                    iframe = driver.find_element(By.ID, 'tracking-iframe')
                    driver.switch_to.frame(iframe)
                    try:
                        choose_section = driver.find_element(By.CSS_SELECTOR, '.section-notice .section-notice__header .section-notice__main span').text
                        print(choose_section)
                    except NoSuchElementException:
                        choose_section = None
                    if choose_section is not None and "Choose which shipment you would like to track." in choose_section:
                        click_texts = driver.find_elements(By.CSS_SELECTOR, '.has-hover-pointer.tracking-cards-bottom--padding-16px')
                        for click_text in click_texts:
                            btn_text = click_text.find_element(By.CSS_SELECTOR, '.shipment-package-grey-flex-container-button .track-package-item-info-container .block--font-14px.block--font-uppercase card-item-padding span').text
                            if btn_text == "DELIVERED":
                                click_text.find_element(By.CLASS_NAME, '.shipment-package-grey-flex-container-button').click()
                                time.sleep(3)
                                shipment_package_div = driver.find_element(By.CSS_SELECTOR, "div.shipment-package-grey-container")
                                div_element = shipment_package_div.find_element(By.CSS_SELECTOR, "div.block--font-12px.block--color-secondary.block--bottom-margin-8px")
                                shipping_method_ = div_element.find_element(By.CSS_SELECTOR, "span:first-child").text
                                if shipping_method_ == "FedEx:":
                                    shipping_method = "Fedex"
                                elif shipping_method_ == "DHL Express International:" or shipping_method_ == "DHL Express Germany:":
                                    shipping_method = "DHL"
                                elif shipping_method_ == "USPS:":
                                    shipping_method = "USPS"
                                elif shipping_method_ == "Japan Post:":
                                    shipping_method = "日本郵便"
                                else:
                                    shipping_method = shipping_method_
                                
                    else:
                        shipment_package_div = driver.find_element(By.CSS_SELECTOR, "div.shipment-package-grey-container")
                        div_element = shipment_package_div.find_element(By.CSS_SELECTOR, "div.block--font-12px.block--color-secondary.block--bottom-margin-8px")
                        shipping_method_ = div_element.find_element(By.CSS_SELECTOR, "span:first-child").text
                        if shipping_method_ == "FedEx:":
                            shipping_method = "Fedex"
                        elif shipping_method_ == "DHL Express International:" or shipping_method_ == "DHL Express Germany:":
                            shipping_method = "DHL"
                        elif shipping_method_ == "USPS:":
                            shipping_method = "USPS"
                        elif shipping_method_ == "Japan Post:":
                            shipping_method = "日本郵便"
                        else:
                            shipping_method = shipping_method_
                except NoSuchElementException:
                    shipping_method = ""
                finally:
                    driver.switch_to.frame(None)

                # ship&co
                driver.switch_to.window(driver.window_handles[1])
                driver.get("https://app.shipandco.com/shipments")
                time.sleep(5)
                search_input = driver.find_element(By.CSS_SELECTOR, "input[name='searchText']")
                search_input.send_keys(tracking_no)
                time.sleep(2)
                try:
                    parent_element = driver.find_element(By.CLASS_NAME, 'new-list-item')
                    rate_element = parent_element.find_element(By.CSS_SELECTOR, ".rate")
                    rate_text = rate_element.text
                    shipping_cost = rate_text.replace("¥ ", "").replace(",", "")
                except NoSuchElementException:
                    shipping_cost = ""
                driver.switch_to.window(driver.window_handles[0])
                # ship&co end

                update_or_insert_data(states, inspection, warehouse, person_charge, shipping_notes, buyer_name, shipping_method, tracking_no, sku, order_number, product_name, ebay_product_name, item_name, shipping_cost, purchase_date, purchase_price, supplier, supplier_url, supplier_url2, tracking_number, invoice_category, supplier_name, supplier_address, eligible_registration_number, date_sold, sale_price, sales_tax, quantity, shipping, pl, sold_fee, overseas_fee, pn, profit, profit_including_refund, refund_amount, return_fee, tracking_info, remarks, order_sheet_id)
        paid_awaiting_feedback(driver, product_sheet_id, order_sheet_id)
    else:
        paid_awaiting_feedback(driver, product_sheet_id, order_sheet_id)
def paid_awaiting_feedback(driver, product_sheet_id, order_sheet_id):
    driver.get('https://www.ebay.com/sh/ord/?filter=status%3APAID_WAITING_TO_GIVE_FEEDBACK')
    time.sleep(2)
    resultNumber = driver.find_element(By.CLASS_NAME, 'summary-content').text
    result = resultNumber.split(": ")
    if len(result) > 1:
        try:
            result_ = int(result[1].split(" of ")[1])
        except (IndexError, ValueError):
            result_ = 0
    else:
        try:
            result_ = int(result[0])
        except (IndexError, ValueError):
            result_ = 0
    if result_ != 0:
        order_status_divs = driver.find_elements(By.CSS_SELECTOR, "div[class='order-status']")
        tracking_info_list = []
        for order_status_div in order_status_divs:
            tracking_info_item = order_status_div.find_element(By.CLASS_NAME, 'sh-bold').text
            tracking_info_list.append(tracking_info_item)

        order_links = []            
        order_detail_divs = driver.find_elements(By.CSS_SELECTOR, "div[class='order-details']")
        for order_detail_div in order_detail_divs:
            links = order_detail_div.find_elements(By.CSS_SELECTOR, "a")
            for link in links:
                order_links.append(link.get_attribute("href"))

        for i, link in enumerate(order_links):
            driver.get(link)
            time.sleep(2)
            
            states = "発送済"

            buyer_name_element = driver.find_element(By.CLASS_NAME, 'shipping-info').find_element(By.CLASS_NAME, 'address')
            buyer_names = buyer_name_element.find_elements(By.CSS_SELECTOR, "button[class='tooltip__host clickable']")
            buyer_name = buyer_names[0].text

            sku_element = driver.find_element(By.CLASS_NAME, 'lineItemCardInfo__sku')
            sku_elements = sku_element.find_elements(By.CLASS_NAME, 'sh-secondary')
            sku = sku_elements[1].text

            product_name, person_charge = main(sku, product_sheet_id)
            product_name = product_name
            person_charge = person_charge

            order_info = driver.find_element(By.CLASS_NAME, 'order-info').find_element(By.CLASS_NAME,'info-wrapper')
            order_elements = order_info.find_elements(By.CLASS_NAME, 'info-item')
            order_number = order_elements[0].find_element(By.CLASS_NAME, 'info-value').text

            detail_info = driver.find_element(By.CLASS_NAME, 'lineItemCardInfo__content').find_element(By.CLASS_NAME, 'details').find_element(By.TAG_NAME, 'a')
            ebay_product_name = detail_info.find_element(By.CLASS_NAME,  'PSEUDOLINK').text

            date_sold_element = order_elements[2].find_element(By.CLASS_NAME, 'info-value').text
            date_object = datetime.strptime(date_sold_element, "%b %d, %Y")
            date_sold = date_object.strftime("%m月%d日")
            
            paid_element = driver.find_element(By.CLASS_NAME, 'buyer-paid')
            if paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(1) .label.level-2").text == "Item subtotal":
                sale_price_ = paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(1) .amount .value").text
                sale_price = float(re.findall(r'-?\d+(?:,\d+)*(?:\.\d+)?', sale_price_)[0].replace(",",""))
            else:
                sale_price = 0
            try:
                if paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(2) .label.level-2").text == "Shipping":
                    shipping_ = paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(2) .amount .value").text
                    shipping = float(re.findall(r'-?\d+(?:,\d+)*(?:\.\d+)?', shipping_)[0])
                else :
                    shipping = 0
            except NoSuchElementException:
                shipping = 0
            try:
                if paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(3) .label.level-2").text == "Sales tax*" or paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(3) .label.level-2").text == "Import VAT*":
                    sales_tax_ = paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(3) .amount .value").text
                    sales_tax = float(re.findall(r'-?\d+(?:,\d+)*(?:\.\d+)?', sales_tax_)[0])
                else:
                    sales_tax = 0
            except NoSuchElementException:
                sales_tax = 0
            try:
                if paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(3) .label.level-2").text == "Refund" or paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(3) .label.level-2").text == "Order canceled":
                    refund_amount_ = paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(3) .amount .value").text
                    _refund_amount = float(re.findall(r'-?\d+(?:,\d+)*(?:\.\d+)?', refund_amount_)[0].replace(",", ""))
                else:
                    _refund_amount = None
            except NoSuchElementException:
                _refund_amount = None
            if _refund_amount == None:
                try:
                    if paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(4) .label.level-2").text == "Refund" or paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(4) .label.level-2").text == "Order canceled":
                        refund_amount_ = paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(4) .amount .value").text
                        _refund_amount = float(re.findall(r'-?\d+(?:,\d+)*(?:\.\d+)?', refund_amount_)[0].replace(",", ""))
                    else:
                        _refund_amount = 0.0
                except NoSuchElementException:
                    _refund_amount = 0.0
            try:
                ebay_paid_element = driver.find_element(By.CLASS_NAME, 'earnings')
                ebay_paid_elements = ebay_paid_element.find_elements(By.CLASS_NAME, 'data-item')
                if ebay_paid_elements[1].find_element(By.TAG_NAME, 'dt') == "Refund (eBay paid)":
                    ebay_paid = float(ebay_paid_elements[1].find_element(By.TAG_NAME, 'dd').text.strip('$'))
                else:
                    ebay_paid = 0.0
            except (NoSuchElementException, IndexError, ValueError, TimeoutException):
                ebay_paid = 0.0
            print(f"refund: {_refund_amount}")
            print(f"ebay paid: {ebay_paid}")
            refund_amount = _refund_amount - ebay_paid

            quantity_info = driver.find_element(By.CLASS_NAME, 'quantity').find_element(By.CLASS_NAME, 'quantity__value')
            quantity = quantity_info.find_element(By.CLASS_NAME, 'sh-bold').text

            try:
                promoted_listings_element = driver.find_element(By.CSS_SELECTOR, ".lineItemCardInfo__indicators .sh-secondary")
                if "Sold via Promoted Listings" in promoted_listings_element.text:
                    pl = "2.0%"
                else:
                    pl = "0.0%"
            except NoSuchElementException:
                pl = "0.0%"

            detail_link = driver.find_element(By.CSS_SELECTOR, ".section-action a")
            event = detail_link.get_attribute("href")
            driver.get(event)
            time.sleep(2)
            link_elements = driver.find_elements(By.CSS_SELECTOR, "a.eui-textual-display")
            try:
                if len(link_elements) == 2:
                    sold_link = link_elements[1].get_attribute("href")
                elif len(link_elements) == 1:
                    sold_link = link_elements[0].get_attribute("href")
                else:
                    sold_link = link_elements[-1].get_attribute("href")
                driver.get(sold_link)
                time.sleep(2)

                price_table = driver.find_element(By.CLASS_NAME, 'fee-breakdown--fee-breakdown-table').find_elements(By.TAG_NAME, 'section')
                try:
                    sold_element = price_table[0].find_element(By.CSS_SELECTOR, ".fee-label-value-item--column.third span:nth-child(3) .text-span")
                    sold_fee_ = sold_element.text.split('%')[0].strip()
                    sold_fee = f"{sold_fee_}%"
                except NoSuchElementException:
                    sold_fee = ""
                fee_percentage_elements = price_table[1].find_elements(By.CSS_SELECTOR, ".fee-label-value-item--column.third span:nth-child(3) .text-span")
                fee_percentages = []
                for element in fee_percentage_elements:
                    text = element.text
                    if "%" in text:
                        numeric_part = text.split("%")[0].strip()
                        fee_percentages.append(float(numeric_part))
                    else:
                        fee_percentages.append(0)
                fee_percentage_difference = fee_percentages[0] - fee_percentages[1]
                overseas_fee = f"{fee_percentage_difference}%"
            except (NoSuchElementException, IndexError, TypeError):
                break

            tracking_info = tracking_info_list[i]

            driver.get(link)
            time.sleep(2)
            try:
                tracking_info_div = driver.find_element(By.CSS_SELECTOR, "div.tracking-info")
                tracking_no_button = tracking_info_div.find_element(By.CSS_SELECTOR, "button.fake-link")
                tracking_no = tracking_no_button.text
                tracking_no_button.click()
                time.sleep(2)
            except NoSuchElementException:
                tracking_no = ""
            try:
                iframe = driver.find_element(By.ID, 'tracking-iframe')
                driver.switch_to.frame(iframe)
                try:
                    choose_section = driver.find_element(By.CSS_SELECTOR, '.section-notice .section-notice__header .section-notice__main span').text
                    print(choose_section)
                except NoSuchElementException:
                    choose_section = None
                if choose_section is not None and "Choose which shipment you would like to track." in choose_section:
                    click_texts = driver.find_elements(By.CSS_SELECTOR, '.has-hover-pointer.tracking-cards-bottom--padding-16px')
                    for click_text in click_texts:
                        btn_text = click_text.find_element(By.CSS_SELECTOR, '.shipment-package-grey-flex-container-button .track-package-item-info-container .block--font-14px.block--font-uppercase card-item-padding span').text
                        if btn_text == "DELIVERED":
                            click_text.find_element(By.CLASS_NAME, '.shipment-package-grey-flex-container-button').click()
                            time.sleep(3)
                            shipment_package_div = driver.find_element(By.CSS_SELECTOR, "div.shipment-package-grey-container")
                            div_element = shipment_package_div.find_element(By.CSS_SELECTOR, "div.block--font-12px.block--color-secondary.block--bottom-margin-8px")
                            shipping_method_ = div_element.find_element(By.CSS_SELECTOR, "span:first-child").text
                            if shipping_method_ == "FedEx:":
                                shipping_method = "Fedex"
                            elif shipping_method_ == "DHL Express International:" or shipping_method_ == "DHL Express Germany:":
                                shipping_method = "DHL"
                            elif shipping_method_ == "USPS:":
                                shipping_method = "USPS"
                            elif shipping_method_ == "Japan Post:":
                                shipping_method = "日本郵便"
                            else:
                                shipping_method = shipping_method_
                            
                else:
                    shipment_package_div = driver.find_element(By.CSS_SELECTOR, "div.shipment-package-grey-container")
                    div_element = shipment_package_div.find_element(By.CSS_SELECTOR, "div.block--font-12px.block--color-secondary.block--bottom-margin-8px")
                    shipping_method_ = div_element.find_element(By.CSS_SELECTOR, "span:first-child").text
                    if shipping_method_ == "FedEx:":
                        shipping_method = "Fedex"
                    elif shipping_method_ == "DHL Express International:" or shipping_method_ == "DHL Express Germany:":
                        shipping_method = "DHL"
                    elif shipping_method_ == "USPS:":
                        shipping_method = "USPS"
                    elif shipping_method_ == "Japan Post:":
                        shipping_method = "日本郵便"
                    else:
                        shipping_method = shipping_method_
            except NoSuchElementException:
                shipping_method = ""
            finally:
                driver.switch_to.frame(None)

            # ship&co
            driver.switch_to.window(driver.window_handles[1])
            driver.get("https://app.shipandco.com/shipments")
            time.sleep(5)
            search_input = driver.find_element(By.CSS_SELECTOR, "input[name='searchText']")
            search_input.send_keys(tracking_no)
            time.sleep(2)
            try:
                parent_element = driver.find_element(By.CLASS_NAME, 'new-list-item')
                rate_element = parent_element.find_element(By.CSS_SELECTOR, ".rate")
                rate_text = rate_element.text
                shipping_cost = rate_text.replace("¥ ", "").replace(",", "")
            except NoSuchElementException:
                shipping_cost = ""
            driver.switch_to.window(driver.window_handles[0])
            # ship&co end

            update_or_insert_data(states, inspection, warehouse, person_charge, shipping_notes, buyer_name, shipping_method, tracking_no, sku, order_number, product_name, ebay_product_name, item_name, shipping_cost, purchase_date, purchase_price, supplier, supplier_url, supplier_url2, tracking_number, invoice_category, supplier_name, supplier_address, eligible_registration_number, date_sold, sale_price, sales_tax, quantity, shipping, pl, sold_fee, overseas_fee, pn, profit, profit_including_refund, refund_amount, return_fee, tracking_info, remarks, order_sheet_id)
        shipped_awaiting_feedback(driver, product_sheet_id, order_sheet_id)
    else:
        shipped_awaiting_feedback(driver, product_sheet_id, order_sheet_id)
def shipped_awaiting_feedback(driver, product_sheet_id, order_sheet_id):
    driver.get('https://www.ebay.com/sh/ord/?filter=status%3ASHIPPED_WAITING_TO_GIVE_FEEDBACK')
    time.sleep(2)
    resultNumber = driver.find_element(By.CLASS_NAME, 'summary-content').text
    result = resultNumber.split(": ")
    if len(result) > 1:
        try:
            result_ = int(result[1].split(" of ")[1])
        except (IndexError, ValueError):
            result_ = 0
    else:
        try:
            result_ = int(result[0])
        except (IndexError, ValueError):
            result_ = 0
    if result_ != 0:
        order_status_divs = driver.find_elements(By.CSS_SELECTOR, "div[class='order-status']")
        tracking_info_list = []
        for order_status_div in order_status_divs:
            tracking_info_item = order_status_div.find_element(By.CLASS_NAME, 'sh-bold').text
            tracking_info_list.append(tracking_info_item)

        order_links = []            
        order_detail_divs = driver.find_elements(By.CSS_SELECTOR, "div[class='order-details']")
        for order_detail_div in order_detail_divs:
            links = order_detail_div.find_elements(By.CSS_SELECTOR, "a")
            for link in links:
                order_links.append(link.get_attribute("href"))

        for i, link in enumerate(order_links):
            driver.get(link)
            time.sleep(2)
            
            states = "発送済"

            buyer_name_element = driver.find_element(By.CLASS_NAME, 'shipping-info').find_element(By.CLASS_NAME, 'address')
            buyer_names = buyer_name_element.find_elements(By.CSS_SELECTOR, "button[class='tooltip__host clickable']")
            buyer_name = buyer_names[0].text

            sku_element = driver.find_element(By.CLASS_NAME, 'lineItemCardInfo__sku')
            sku_elements = sku_element.find_elements(By.CLASS_NAME, 'sh-secondary')
            sku = sku_elements[1].text

            product_name, person_charge = main(sku, product_sheet_id)
            product_name = product_name
            person_charge = person_charge

            order_info = driver.find_element(By.CLASS_NAME, 'order-info').find_element(By.CLASS_NAME,'info-wrapper')
            order_elements = order_info.find_elements(By.CLASS_NAME, 'info-item')
            order_number = order_elements[0].find_element(By.CLASS_NAME, 'info-value').text

            detail_info = driver.find_element(By.CLASS_NAME, 'lineItemCardInfo__content').find_element(By.CLASS_NAME, 'details').find_element(By.TAG_NAME, 'a')
            ebay_product_name = detail_info.find_element(By.CLASS_NAME,  'PSEUDOLINK').text

            date_sold_element = order_elements[2].find_element(By.CLASS_NAME, 'info-value').text
            date_object = datetime.strptime(date_sold_element, "%b %d, %Y")
            date_sold = date_object.strftime("%m月%d日")
            
            paid_element = driver.find_element(By.CLASS_NAME, 'buyer-paid')
            if paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(1) .label.level-2").text == "Item subtotal":
                sale_price_ = paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(1) .amount .value").text
                sale_price = float(re.findall(r'-?\d+(?:,\d+)*(?:\.\d+)?', sale_price_)[0].replace(",",""))
            else:
                sale_price = 0
            try:
                if paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(2) .label.level-2").text == "Shipping":
                    shipping_ = paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(2) .amount .value").text
                    shipping = float(re.findall(r'-?\d+(?:,\d+)*(?:\.\d+)?', shipping_)[0])
                else :
                    shipping = 0
            except NoSuchElementException:
                shipping = 0
            try:
                if paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(3) .label.level-2").text == "Sales tax*" or paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(3) .label.level-2").text == "Import VAT*":
                    sales_tax_ = paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(3) .amount .value").text
                    sales_tax = float(re.findall(r'-?\d+(?:,\d+)*(?:\.\d+)?', sales_tax_)[0])
                else:
                    sales_tax = 0
            except NoSuchElementException:
                sales_tax = 0
            try:
                if paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(3) .label.level-2").text == "Refund" or paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(3) .label.level-2").text == "Order canceled":
                    refund_amount_ = paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(3) .amount .value").text
                    _refund_amount = float(re.findall(r'-?\d+(?:,\d+)*(?:\.\d+)?', refund_amount_)[0].replace(",", ""))
                else:
                    _refund_amount = None
            except NoSuchElementException:
                _refund_amount = None
            if _refund_amount == None:
                try:
                    if paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(4) .label.level-2").text == "Refund" or paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(4) .label.level-2").text == "Order canceled":
                        refund_amount_ = paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(4) .amount .value").text
                        _refund_amount = float(re.findall(r'-?\d+(?:,\d+)*(?:\.\d+)?', refund_amount_)[0].replace(",", ""))
                    else:
                        _refund_amount = 0.0
                except NoSuchElementException:
                    _refund_amount = 0.0
            try:
                ebay_paid_element = driver.find_element(By.CLASS_NAME, 'earnings')
                ebay_paid_elements = ebay_paid_element.find_elements(By.CLASS_NAME, 'data-item')
                if ebay_paid_elements[1].find_element(By.TAG_NAME, 'dt') == "Refund (eBay paid)":
                    ebay_paid = float(ebay_paid_elements[1].find_element(By.TAG_NAME, 'dd').text.strip('$'))
                else:
                    ebay_paid = 0.0
            except (NoSuchElementException, IndexError, ValueError, TimeoutException):
                ebay_paid = 0.0
            print(f"refund: {_refund_amount}")
            print(f"ebay paid: {ebay_paid}")
            refund_amount = _refund_amount - ebay_paid

            quantity_info = driver.find_element(By.CLASS_NAME, 'quantity').find_element(By.CLASS_NAME, 'quantity__value')
            quantity = quantity_info.find_element(By.CLASS_NAME, 'sh-bold').text

            try:
                promoted_listings_element = driver.find_element(By.CSS_SELECTOR, ".lineItemCardInfo__indicators .sh-secondary")
                if "Sold via Promoted Listings" in promoted_listings_element.text:
                    pl = "2.0%"
                else:
                    pl = "0.0%"
            except NoSuchElementException:
                pl = "0.0%"

            detail_link = driver.find_element(By.CSS_SELECTOR, ".section-action a")
            event = detail_link.get_attribute("href")
            driver.get(event)
            time.sleep(2)
            link_elements = driver.find_elements(By.CSS_SELECTOR, "a.eui-textual-display")
            try:
                if len(link_elements) == 2:
                    sold_link = link_elements[1].get_attribute("href")
                elif len(link_elements) == 1:
                    sold_link = link_elements[0].get_attribute("href")
                else:
                    sold_link = link_elements[-1].get_attribute("href")
                driver.get(sold_link)
                time.sleep(2)

                price_table = driver.find_element(By.CLASS_NAME, 'fee-breakdown--fee-breakdown-table').find_elements(By.TAG_NAME, 'section')
                try:
                    sold_element = price_table[0].find_element(By.CSS_SELECTOR, ".fee-label-value-item--column.third span:nth-child(3) .text-span")
                    sold_fee_ = sold_element.text.split('%')[0].strip()
                    sold_fee = f"{sold_fee_}%"
                except NoSuchElementException:
                    sold_fee = ""
                fee_percentage_elements = price_table[1].find_elements(By.CSS_SELECTOR, ".fee-label-value-item--column.third span:nth-child(3) .text-span")
                fee_percentages = []
                for element in fee_percentage_elements:
                    text = element.text
                    if "%" in text:
                        numeric_part = text.split("%")[0].strip()
                        fee_percentages.append(float(numeric_part))
                    else:
                        fee_percentages.append(0)
                fee_percentage_difference = fee_percentages[0] - fee_percentages[1]
                overseas_fee = f"{fee_percentage_difference}%"
            except (NoSuchElementException, IndexError, TypeError):
                break
            tracking_info = tracking_info_list[i]

            driver.get(link)
            time.sleep(2)
            try:
                tracking_info_div = driver.find_element(By.CSS_SELECTOR, "div.tracking-info")
                tracking_no_button = tracking_info_div.find_element(By.CSS_SELECTOR, "button.fake-link")
                tracking_no = tracking_no_button.text
                tracking_no_button.click()
                time.sleep(2)
            except NoSuchElementException:
                tracking_no = ""
            try:
                iframe = driver.find_element(By.ID, 'tracking-iframe')
                driver.switch_to.frame(iframe)
                try:
                    choose_section = driver.find_element(By.CSS_SELECTOR, '.section-notice .section-notice__header .section-notice__main span').text
                    print(choose_section)
                except NoSuchElementException:
                    choose_section = None
                if choose_section is not None and "Choose which shipment you would like to track." in choose_section:
                    click_texts = driver.find_elements(By.CSS_SELECTOR, '.has-hover-pointer.tracking-cards-bottom--padding-16px')
                    for click_text in click_texts:
                        btn_text = click_text.find_element(By.CSS_SELECTOR, '.shipment-package-grey-flex-container-button .track-package-item-info-container .block--font-14px.block--font-uppercase card-item-padding span').text
                        if btn_text == "DELIVERED":
                            click_text.find_element(By.CLASS_NAME, '.shipment-package-grey-flex-container-button').click()
                            time.sleep(3)
                            shipment_package_div = driver.find_element(By.CSS_SELECTOR, "div.shipment-package-grey-container")
                            div_element = shipment_package_div.find_element(By.CSS_SELECTOR, "div.block--font-12px.block--color-secondary.block--bottom-margin-8px")
                            shipping_method_ = div_element.find_element(By.CSS_SELECTOR, "span:first-child").text
                            if shipping_method_ == "FedEx:":
                                shipping_method = "Fedex"
                            elif shipping_method_ == "DHL Express International:" or shipping_method_ == "DHL Express Germany:":
                                shipping_method = "DHL"
                            elif shipping_method_ == "USPS:":
                                shipping_method = "USPS"
                            elif shipping_method_ == "Japan Post:":
                                shipping_method = "日本郵便"
                            else:
                                shipping_method = shipping_method_
                            
                else:
                    shipment_package_div = driver.find_element(By.CSS_SELECTOR, "div.shipment-package-grey-container")
                    div_element = shipment_package_div.find_element(By.CSS_SELECTOR, "div.block--font-12px.block--color-secondary.block--bottom-margin-8px")
                    shipping_method_ = div_element.find_element(By.CSS_SELECTOR, "span:first-child").text
                    if shipping_method_ == "FedEx:":
                        shipping_method = "Fedex"
                    elif shipping_method_ == "DHL Express International:" or shipping_method_ == "DHL Express Germany:":
                        shipping_method = "DHL"
                    elif shipping_method_ == "USPS:":
                        shipping_method = "USPS"
                    elif shipping_method_ == "Japan Post:":
                        shipping_method = "日本郵便"
                    else:
                        shipping_method = shipping_method_
            except NoSuchElementException:
                shipping_method = ""
            finally:
                driver.switch_to.frame(None)

            # ship&co
            driver.switch_to.window(driver.window_handles[1])
            driver.get("https://app.shipandco.com/shipments")
            time.sleep(5)
            search_input = driver.find_element(By.CSS_SELECTOR, "input[name='searchText']")
            search_input.send_keys(tracking_no)
            time.sleep(2)
            try:
                parent_element = driver.find_element(By.CLASS_NAME, 'new-list-item')
                rate_element = parent_element.find_element(By.CSS_SELECTOR, ".rate")
                rate_text = rate_element.text
                shipping_cost = rate_text.replace("¥ ", "").replace(",", "")
            except NoSuchElementException:
                shipping_cost = ""
            driver.switch_to.window(driver.window_handles[0])
            # ship&co end

            update_or_insert_data(states, inspection, warehouse, person_charge, shipping_notes, buyer_name, shipping_method, tracking_no, sku, order_number, product_name, ebay_product_name, item_name, shipping_cost, purchase_date, purchase_price, supplier, supplier_url, supplier_url2, tracking_number, invoice_category, supplier_name, supplier_address, eligible_registration_number, date_sold, sale_price, sales_tax, quantity, shipping, pl, sold_fee, overseas_fee, pn, profit, profit_including_refund, refund_amount, return_fee, tracking_info, remarks, order_sheet_id)
        archived(driver, product_sheet_id, order_sheet_id)
    else:
        archived(driver, product_sheet_id, order_sheet_id)
def archived(driver, product_sheet_id, order_sheet_id):
    driver.get('https://www.ebay.com/sh/ord/?filter=status%3AARCHIVED')
    time.sleep(2)
    resultNumber = driver.find_element(By.CLASS_NAME, 'summary-content').text
    result = resultNumber.split(": ")
    if len(result) > 1:
        try:
            result_ = int(result[1].split(" of ")[1])
        except (IndexError, ValueError):
            result_ = 0
    else:
        try:
            result_ = int(result[0])
        except (IndexError, ValueError):
            result_ = 0
    if result_ != 0:
        order_status_divs = driver.find_elements(By.CSS_SELECTOR, "div[class='order-status']")
        tracking_info_list = []
        for order_status_div in order_status_divs:
            tracking_info_item = order_status_div.find_element(By.CLASS_NAME, 'sh-bold').text
            tracking_info_list.append(tracking_info_item)

        order_links = []            
        order_detail_divs = driver.find_elements(By.CSS_SELECTOR, "div[class='order-details']")
        for order_detail_div in order_detail_divs:
            links = order_detail_div.find_elements(By.CSS_SELECTOR, "a")
            for link in links:
                order_links.append(link.get_attribute("href"))

        for i, link in enumerate(order_links):
            driver.get(link)
            time.sleep(2)
            
            states = "発送済"

            buyer_name_element = driver.find_element(By.CLASS_NAME, 'shipping-info').find_element(By.CLASS_NAME, 'address')
            buyer_names = buyer_name_element.find_elements(By.CSS_SELECTOR, "button[class='tooltip__host clickable']")
            buyer_name = buyer_names[0].text

            sku_element = driver.find_element(By.CLASS_NAME, 'lineItemCardInfo__sku')
            sku_elements = sku_element.find_elements(By.CLASS_NAME, 'sh-secondary')
            sku = sku_elements[1].text

            product_name, person_charge = main(sku, product_sheet_id)
            product_name = product_name
            person_charge = person_charge

            order_info = driver.find_element(By.CLASS_NAME, 'order-info').find_element(By.CLASS_NAME,'info-wrapper')
            order_elements = order_info.find_elements(By.CLASS_NAME, 'info-item')
            order_number = order_elements[0].find_element(By.CLASS_NAME, 'info-value').text

            detail_info = driver.find_element(By.CLASS_NAME, 'lineItemCardInfo__content').find_element(By.CLASS_NAME, 'details').find_element(By.TAG_NAME, 'a')
            ebay_product_name = detail_info.find_element(By.CLASS_NAME,  'PSEUDOLINK').text

            date_sold_element = order_elements[2].find_element(By.CLASS_NAME, 'info-value').text
            date_object = datetime.strptime(date_sold_element, "%b %d, %Y")
            date_sold = date_object.strftime("%m月%d日")
            
            paid_element = driver.find_element(By.CLASS_NAME, 'buyer-paid')
            if paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(1) .label.level-2").text == "Item subtotal":
                sale_price_ = paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(1) .amount .value").text
                sale_price = float(re.findall(r'-?\d+(?:,\d+)*(?:\.\d+)?', sale_price_)[0].replace(",",""))
            else:
                sale_price = 0
            try:
                if paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(2) .label.level-2").text == "Shipping":
                    shipping_ = paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(2) .amount .value").text
                    shipping = float(re.findall(r'-?\d+(?:,\d+)*(?:\.\d+)?', shipping_)[0])
                else :
                    shipping = 0
            except NoSuchElementException:
                shipping = 0
            try:
                if paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(3) .label.level-2").text == "Sales tax*" or paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(3) .label.level-2").text == "Import VAT*":
                    sales_tax_ = paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(3) .amount .value").text
                    sales_tax = float(re.findall(r'-?\d+(?:,\d+)*(?:\.\d+)?', sales_tax_)[0])
                else:
                    sales_tax = 0
            except NoSuchElementException:
                sales_tax = 0
            try:
                if paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(4) .label.level-2").text == "Refund":
                    refund_amount_ = paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(4) .amount .value").text
                    _refund_amount = float(re.findall(r'-?\d+(?:,\d+)*(?:\.\d+)?', refund_amount_)[0].replace(",", ""))
                else:
                    _refund_amount = 0.0
            except NoSuchElementException:
                _refund_amount = 0.0
            try:
                ebay_paid_element = driver.find_element(By.CLASS_NAME, 'earnings')
                ebay_paid_elements = ebay_paid_element.find_elements(By.CLASS_NAME, 'data-item')
                if ebay_paid_elements[1].find_element(By.TAG_NAME, 'dt') == "Refund (eBay paid)":
                    ebay_paid = float(ebay_paid_elements[1].find_element(By.TAG_NAME, 'dd').text.strip('$'))
                else:
                    ebay_paid = 0.0
            except (NoSuchElementException, IndexError, ValueError, TimeoutException):
                ebay_paid = 0.0
            print(f"refund: {_refund_amount}")
            print(f"ebay paid: {ebay_paid}")
            refund_amount = _refund_amount - ebay_paid

            quantity_info = driver.find_element(By.CLASS_NAME, 'quantity').find_element(By.CLASS_NAME, 'quantity__value')
            quantity = quantity_info.find_element(By.CLASS_NAME, 'sh-bold').text

            try:
                promoted_listings_element = driver.find_element(By.CSS_SELECTOR, ".lineItemCardInfo__indicators .sh-secondary")
                if "Sold via Promoted Listings" in promoted_listings_element.text:
                    pl = "2.0%"
                else:
                    pl = "0.0%"
            except NoSuchElementException:
                pl = "0.0%"

            detail_link = driver.find_element(By.CSS_SELECTOR, ".section-action a")
            event = detail_link.get_attribute("href")
            driver.get(event)
            time.sleep(2)
            link_elements = driver.find_elements(By.CSS_SELECTOR, "a.eui-textual-display")
            try:
                if len(link_elements) == 2:
                    sold_link = link_elements[1].get_attribute("href")
                elif len(link_elements) == 1:
                    sold_link = link_elements[0].get_attribute("href")
                else:
                    sold_link = link_elements[-1].get_attribute("href")
                driver.get(sold_link)
                time.sleep(2)

                price_table = driver.find_element(By.CLASS_NAME, 'fee-breakdown--fee-breakdown-table').find_elements(By.TAG_NAME, 'section')
                try:
                    sold_element = price_table[0].find_element(By.CSS_SELECTOR, ".fee-label-value-item--column.third span:nth-child(3) .text-span")
                    sold_fee_ = sold_element.text.split('%')[0].strip()
                    sold_fee = f"{sold_fee_}%"
                except NoSuchElementException:
                    sold_fee = ""
                fee_percentage_elements = price_table[1].find_elements(By.CSS_SELECTOR, ".fee-label-value-item--column.third span:nth-child(3) .text-span")
                fee_percentages = []
                for element in fee_percentage_elements:
                    text = element.text
                    if "%" in text:
                        numeric_part = text.split("%")[0].strip()
                        fee_percentages.append(float(numeric_part))
                    else:
                        fee_percentages.append(0)
                fee_percentage_difference = fee_percentages[0] - fee_percentages[1]
                overseas_fee = f"{fee_percentage_difference}%"
            except (NoSuchElementException, IndexError, TypeError):
                break

            tracking_info = tracking_info_list[i]

            driver.get(link)
            time.sleep(2)
            try:
                tracking_info_div = driver.find_element(By.CSS_SELECTOR, "div.tracking-info")
                tracking_no_button = tracking_info_div.find_element(By.CSS_SELECTOR, "button.fake-link")
                tracking_no = tracking_no_button.text
                tracking_no_button.click()
                time.sleep(2)
            except NoSuchElementException:
                tracking_no = ""
            try:
                iframe = driver.find_element(By.ID, 'tracking-iframe')
                driver.switch_to.frame(iframe)
                try:
                    choose_section = driver.find_element(By.CSS_SELECTOR, '.section-notice .section-notice__header .section-notice__main span').text
                    print(choose_section)
                except NoSuchElementException:
                    choose_section = None
                if choose_section is not None and "Choose which shipment you would like to track." in choose_section:
                    click_texts = driver.find_elements(By.CSS_SELECTOR, '.has-hover-pointer.tracking-cards-bottom--padding-16px')
                    for click_text in click_texts:
                        btn_text = click_text.find_element(By.CSS_SELECTOR, '.shipment-package-grey-flex-container-button .track-package-item-info-container .block--font-14px.block--font-uppercase card-item-padding span').text
                        if btn_text == "DELIVERED":
                            click_text.find_element(By.CLASS_NAME, '.shipment-package-grey-flex-container-button').click()
                            time.sleep(3)
                            shipment_package_div = driver.find_element(By.CSS_SELECTOR, "div.shipment-package-grey-container")
                            div_element = shipment_package_div.find_element(By.CSS_SELECTOR, "div.block--font-12px.block--color-secondary.block--bottom-margin-8px")
                            shipping_method_ = div_element.find_element(By.CSS_SELECTOR, "span:first-child").text
                            if shipping_method_ == "FedEx:":
                                shipping_method = "Fedex"
                            elif shipping_method_ == "DHL Express International:" or shipping_method_ == "DHL Express Germany:":
                                shipping_method = "DHL"
                            elif shipping_method_ == "USPS:":
                                shipping_method = "USPS"
                            elif shipping_method_ == "Japan Post:":
                                shipping_method = "日本郵便"
                            else:
                                shipping_method = shipping_method_
                            
                else:
                    shipment_package_div = driver.find_element(By.CSS_SELECTOR, "div.shipment-package-grey-container")
                    div_element = shipment_package_div.find_element(By.CSS_SELECTOR, "div.block--font-12px.block--color-secondary.block--bottom-margin-8px")
                    shipping_method_ = div_element.find_element(By.CSS_SELECTOR, "span:first-child").text
                    if shipping_method_ == "FedEx:":
                        shipping_method = "Fedex"
                    elif shipping_method_ == "DHL Express International:" or shipping_method_ == "DHL Express Germany:":
                        shipping_method = "DHL"
                    elif shipping_method_ == "USPS:":
                        shipping_method = "USPS"
                    elif shipping_method_ == "Japan Post:":
                        shipping_method = "日本郵便"
                    else:
                        shipping_method = shipping_method_
            except NoSuchElementException:
                shipping_method = ""
            finally:
                driver.switch_to.frame(None)

            # ship&co
            driver.switch_to.window(driver.window_handles[1])
            driver.get("https://app.shipandco.com/shipments")
            time.sleep(5)
            search_input = driver.find_element(By.CSS_SELECTOR, "input[name='searchText']")
            search_input.send_keys(tracking_no)
            time.sleep(2)
            try:
                parent_element = driver.find_element(By.CLASS_NAME, 'new-list-item')
                rate_element = parent_element.find_element(By.CSS_SELECTOR, ".rate")
                rate_text = rate_element.text
                shipping_cost = rate_text.replace("¥ ", "").replace(",", "")
            except NoSuchElementException:
                shipping_cost = ""
            driver.switch_to.window(driver.window_handles[0])
            # ship&co end

            update_or_insert_data(states, inspection, warehouse, person_charge, shipping_notes, buyer_name, shipping_method, tracking_no, sku, order_number, product_name, ebay_product_name, item_name, shipping_cost, purchase_date, purchase_price, supplier, supplier_url, supplier_url2, tracking_number, invoice_category, supplier_name, supplier_address, eligible_registration_number, date_sold, sale_price, sales_tax, quantity, shipping, pl, sold_fee, overseas_fee, pn, profit, profit_including_refund, refund_amount, return_fee, tracking_info, remarks, order_sheet_id)
    else:
        driver.close()

if __name__ == '__main__':
    specific_date = datetime(2024, 7, 13)
    current_date = datetime.now()
    if specific_date < current_date:
        print('>>> expired <<<')
    else:
        windows()
