import requests
import json
import os
import getpass
import time
import re
import sys
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
def get_ebay_data(ebay_id, ebay_pass, product_sheet_id, order_sheet_id, checkbox_val):
    driver = webdriver.Chrome()
    ebay_open(driver, ebay_id, ebay_pass)
    driver.switch_to.window(driver.window_handles[0])
    driver.get('https://www.ebay.com/sh/ord?filter=status%3AALL_ORDERS')
    time.sleep(2)
    if checkbox_val == 0:
        total_get_data(product_sheet_id, order_sheet_id, driver)
    if checkbox_val == 1:
        recent_get_data(product_sheet_id, order_sheet_id, driver)
    
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
def total_get_data(product_sheet_id, order_sheet_id, driver):
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
            if pagination_link == None:
                continue
            driver.get(pagination_link)
            time.sleep(10)
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
                time.sleep(3)
                
                sku_element = driver.find_element(By.CLASS_NAME, 'lineItemCardInfo__sku')
                sku_elements = sku_element.find_elements(By.CLASS_NAME, 'sh-secondary')
                sku = sku_elements[1].text

                product_name, person_charge = main(sku, product_sheet_id)
                product_name = product_name
                person_charge = person_charge

                order_info = driver.find_element(By.CLASS_NAME, 'order-info').find_element(By.CLASS_NAME,'info-wrapper')
                order_elements = order_info.find_elements(By.CLASS_NAME, 'info-item')
                order_number = order_elements[0].find_element(By.CLASS_NAME, 'info-value').text

                try:
                    buyer_name_element = driver.find_element(By.CLASS_NAME, 'shipping-info').find_element(By.CLASS_NAME, 'address')
                    buyer_names = buyer_name_element.find_elements(By.CSS_SELECTOR, "button[class='tooltip__host clickable']")
                    buyer_name = buyer_names[0].text
                except NoSuchElementException:
                    buyer_name = order_elements[2].find_element(By.CSS_SELECTOR, '.info-value.buyer div:first-child').text

                detail_info = driver.find_element(By.CLASS_NAME, 'lineItemCardInfo__content').find_element(By.CLASS_NAME, 'details').find_element(By.TAG_NAME, 'a')
                ebay_product_name = detail_info.find_element(By.CLASS_NAME,  'PSEUDOLINK').text

                try:
                    date_sold_element = order_elements[2].find_element(By.CLASS_NAME, 'info-value').text
                    date_object = datetime.strptime(date_sold_element, "%b %d, %Y")
                    date_sold = date_object.strftime("%m月%d日")
                except (NoSuchElementException, TypeError, IndexError ,ValueError):
                    date_sold_element = order_elements[1].find_element(By.CLASS_NAME, 'info-value').text
                    date_object = datetime.strptime(date_sold_element, "%b %d, %Y")
                    date_sold = date_object.strftime("%m月%d日")

                try:
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
                        if paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(3) .label.level-2").text == "Sales tax*" or paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(3) .label.level-2").text == "Import VAT*" or paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(3) .label.level-2").text == "VAT*":
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
                    def get_ebay_paid_refund(ebay_paid_element):
                        try:
                            refund_element = next((item for item in ebay_paid_element.find_elements(By.CLASS_NAME, 'data-item')
                                                if item.find_element(By.TAG_NAME, 'dt').text == "Refund (eBay paid)"), None)
                            if refund_element:
                                return float(refund_element.find_element(By.TAG_NAME, 'dd').text.strip('$'))
                            else:
                                return 0.0
                        except (NoSuchElementException, ValueError):
                            return 0.0
                        
                    ebay_paid_element = driver.find_element(By.CLASS_NAME, 'earnings')
                    ebay_paid = get_ebay_paid_refund(ebay_paid_element)
                    
                    print(f"refund: {_refund_amount}")
                    print(f"ebay paid: {ebay_paid}")
                    refund_amount = _refund_amount - ebay_paid
                except (NoSuchElementException, TypeError, IndexError):
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
                try:
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
                except (NoSuchElementException, TypeError, IndexError):
                    sold_fee = ""
                    overseas_fee = ""
                    shipping_method = ""
                    tracking_no = ""
                
                tracking_info = tracking_info_list[i]
                shipping_cost = ""
                update_or_insert_data(states, inspection, warehouse, person_charge, shipping_notes, buyer_name, shipping_method, tracking_no, sku, order_number, product_name, ebay_product_name, item_name, shipping_cost, purchase_date, purchase_price, supplier, supplier_url, supplier_url2, tracking_number, invoice_category, supplier_name, supplier_address, eligible_registration_number, date_sold, sale_price, sales_tax, quantity, shipping, pl, sold_fee, overseas_fee, pn, profit, profit_including_refund, refund_amount, return_fee, tracking_info, remarks, order_sheet_id)
            # driver.close()
        else:
            driver.close()
def recent_get_data(product_sheet_id, order_sheet_id, driver):
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
            if pagination_link == None:
                continue
            driver.get(pagination_link)
            time.sleep(10)
            order_status_divs = driver.find_elements(By.CSS_SELECTOR, "div[class='order-status']")
            tracking_info_list = []
            for order_status_div in order_status_divs:
                tracking_info_item = order_status_div.find_element(By.CLASS_NAME, 'sh-bold').text
                tracking_info_list.append(tracking_info_item)

            date_colums = driver.find_elements(By.CSS_SELECTOR, '.order-default-cell.date-column .sh-default')
            date_sold_lists = []
            for date_colum in date_colums:
                date_solded = date_colum.text.split(" at ")[0]
                date_solded_ = datetime.strptime(date_solded, "%b %d, %Y")
                date_sold_lists.append(date_solded_.strftime("%Y/%m/%d"))
            
            order_links = []            
            order_detail_divs = driver.find_elements(By.CSS_SELECTOR, "div[class='order-details']")
            for order_detail_div in order_detail_divs:
                links = order_detail_div.find_elements(By.CSS_SELECTOR, "a")
                for link in links:
                    order_links.append(link.get_attribute("href"))
            for i, link in enumerate(order_links):
                driver.get(link)
                time.sleep(3)
                
                sku_element = driver.find_element(By.CLASS_NAME, 'lineItemCardInfo__sku')
                sku_elements = sku_element.find_elements(By.CLASS_NAME, 'sh-secondary')
                sku = sku_elements[1].text

                product_name, person_charge = main(sku, product_sheet_id)
                product_name = product_name
                person_charge = person_charge

                order_info = driver.find_element(By.CLASS_NAME, 'order-info').find_element(By.CLASS_NAME,'info-wrapper')
                order_elements = order_info.find_elements(By.CLASS_NAME, 'info-item')
                order_number = order_elements[0].find_element(By.CLASS_NAME, 'info-value').text

                try:
                    buyer_name_element = driver.find_element(By.CLASS_NAME, 'shipping-info').find_element(By.CLASS_NAME, 'address')
                    buyer_names = buyer_name_element.find_elements(By.CSS_SELECTOR, "button[class='tooltip__host clickable']")
                    buyer_name = buyer_names[0].text
                except NoSuchElementException:
                    buyer_name = order_elements[2].find_element(By.CSS_SELECTOR, '.info-value.buyer div:first-child').text

                detail_info = driver.find_element(By.CLASS_NAME, 'lineItemCardInfo__content').find_element(By.CLASS_NAME, 'details').find_element(By.TAG_NAME, 'a')
                ebay_product_name = detail_info.find_element(By.CLASS_NAME,  'PSEUDOLINK').text

                try:
                    date_sold_element = order_elements[2].find_element(By.CLASS_NAME, 'info-value').text
                    date_object = datetime.strptime(date_sold_element, "%b %d, %Y")
                    date_sold_get = date_object.strftime("%Y/%m/%d")
                    date_sold = date_object.strftime("%m月%d日")
                except (NoSuchElementException, TypeError, IndexError ,ValueError):
                    date_sold_element = order_elements[1].find_element(By.CLASS_NAME, 'info-value').text
                    date_object = datetime.strptime(date_sold_element, "%b %d, %Y")
                    date_sold_get = date_object.strftime("%Y/%m/%d")
                    date_sold = date_object.strftime("%m月%d日")

                try:
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
                        if paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(3) .label.level-2").text == "Sales tax*" or paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(3) .label.level-2").text == "Import VAT*" or paid_element.find_element(By.CSS_SELECTOR, ".data-item .level-2:nth-child(3) .label.level-2").text == "VAT*":
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
                    def get_ebay_paid_refund(ebay_paid_element):
                        try:
                            refund_element = next((item for item in ebay_paid_element.find_elements(By.CLASS_NAME, 'data-item')
                                                if item.find_element(By.TAG_NAME, 'dt').text == "Refund (eBay paid)"), None)
                            if refund_element:
                                return float(refund_element.find_element(By.TAG_NAME, 'dd').text.strip('$'))
                            else:
                                return 0.0
                        except (NoSuchElementException, ValueError):
                            return 0.0
                        
                    ebay_paid_element = driver.find_element(By.CLASS_NAME, 'earnings')
                    ebay_paid = get_ebay_paid_refund(ebay_paid_element)
                    print(f"refund: {_refund_amount}")
                    print(f"ebay paid: {ebay_paid}")
                    refund_amount = _refund_amount - ebay_paid
                except (NoSuchElementException, TypeError, IndexError):
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
                try:
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
                except (NoSuchElementException, TypeError, IndexError):
                    sold_fee = ""
                    overseas_fee = ""
                    shipping_method = ""
                    tracking_no = ""
                
                tracking_info = tracking_info_list[i]
                shipping_cost = ""
                current_date = datetime.now().strftime("%Y/%m/%d")
                
                if current_date.split("/")[1] == date_sold_get.split("/")[1]:
                    update_or_insert_data(states, inspection, warehouse, person_charge, shipping_notes, buyer_name, shipping_method, tracking_no, sku, order_number, product_name, ebay_product_name, item_name, shipping_cost, purchase_date, purchase_price, supplier, supplier_url, supplier_url2, tracking_number, invoice_category, supplier_name, supplier_address, eligible_registration_number, date_sold, sale_price, sales_tax, quantity, shipping, pl, sold_fee, overseas_fee, pn, profit, profit_including_refund, refund_amount, return_fee, tracking_info, remarks, order_sheet_id)
                else:
                    sys.exit()
        else:
            sys.exit()
if __name__ == '__main__':
    specific_date = datetime(2024, 8, 15)
    current_date = datetime.now()
    if specific_date < current_date:
        print('>>> expired <<<')
    else:
        get_ebay_data()
