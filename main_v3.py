##
#
# Script to automate the customs form creation on the post.at website
# Start with "python main_v3.py"
# Should work with https://www.post.at/en/n/f/customs-form 
# 01.2023 peter@traunmueller.net
#
##




# importing packages
from selenium import webdriver
from selenium.webdriver.chrome import options
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from time import sleep
import os
import csv
from datetime import datetime
from selenium.common import exceptions
import warnings
from selenium.webdriver.support.ui import Select

url = 'https://www.post.at/en/n/f/customs-form'
input_file_csv = './address.csv'
download_directory = './pdf'


# adding chrome options
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('disable-infobars')
chrome_options.add_argument('--disable-notifications')
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
chrome_options.add_experimental_option('useAutomationExtension', False)
chrome_options.add_argument('--disable-popup-blocking')
prefs = {"credentials_enable_service": False,
         "profile.password_manager_enabled": False,
         "download.default_directory" : download_directory,
         "download.prompt_for_download": False,
         "download.directory_upgrade": True,
         "plugins.plugins_disabled": "Chrome PDF Viewer",
         "plugins.always_open_pdf_externally": True}
chrome_options.add_experimental_option("prefs", prefs)
chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])
chrome_options.add_argument('--disable-blink-features=AutomationControlled')
chrome_options.add_argument('--disable-dev-shm-usage')
chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
warnings.filterwarnings("ignore", category=DeprecationWarning)


# finding elements
def elem(x_path,driver):
    if driver.find_elements(by=By.XPATH, value=x_path):
        element = driver.find_element(by=By.XPATH, value=x_path)
    else:
        element = "Not Found"
    return element





# defining driver
driver = webdriver.Chrome(options = chrome_options, executable_path='./chromedriver')
driver.maximize_window()

print("Loading Website...")
driver.get(url)

try:
    WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, '//button[@id="onetrust-accept-btn-handler"]')))
    accept_all_cookies = elem('//button[@id="onetrust-accept-btn-handler"]',driver)
    accept_all_cookies.click()
except Exception as ex:
    print(ex)

sleep(5)

#counting rows
with open(input_file_csv, encoding='UTF-8') as csv_count_file:
    csv_count_reader = csv.reader(csv_count_file, delimiter=',')
    recipient_number_total = len(list(csv_count_reader)) - 1
recipient_number_now = 1

# getting data
with open(input_file_csv,  encoding='UTF-8') as csvfile:
    csvReader = csv.reader(csvfile, delimiter=',')
    next(csvReader)
    for row in csvReader:
        sender = row[0]
        sender_address_Company = row[1]
        sender_address_first_and_last_name = row[2]
        sender_address_street = row[3]
        sender_address_number = row[4]
        sender_address_postcode = row[5]
        sender_address_city = row[6]
        sender_address_country = row[7]
        sender_address_email_address = row[8]
        sender_address_country_code = row[9]
        sender_address_phone_number = row[10]
        recipient_company = row[11]
        recipient_first_and_last_name = row[12]
        recipient_street = row[13]
        recipient_number = row[14]
        recipient_city = row[15]
        recipient_country = row[16]
        recipient_region = row[17]
        recipient_email_address = row[18]
        recipient_country_code = row[19]
        recipient_phone_number = row[20]
        item_category = row[21]
        currency = row[22]
        description_of_content = row[23]
        quantity = row[24]
        net_weight_per_kg = row[25]
        value_in_eur_per_unit = row[26]
        tariff_number = row[27]
        country_of_origin_of_merchandise = row[28]
        gross_weight_in_kg = row[29]
        information_about_importer = row[30]
        approval_number = row[31]
        certificate_number = row[32]
        invoice_number = row[33]
        special_remarks = row[34]
        if_undelivered = row[35]

        print(recipient_first_and_last_name + ": " + str(recipient_number_now) + "/" + str(recipient_number_total))
        #################################################### Item information ##############################################################
        # if Business
        if "Business" in sender:
            try:
                return_to_sender = elem('(//input[@name="formType"])[2]',driver)
                driver.execute_script("arguments[0].click();", return_to_sender)
            except Exception as ex:
                print(ex)

        #################################################### sender details ##############################################################
        # clicking on enter sender address for all details
        try:
            WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, '//a[@title="Enter sender address"]')))
            enter_sender_address = elem('//a[@title="Enter sender address"]',driver)
            enter_sender_address.click()
        except Exception as ex:
            print(ex)
        # sending first and last name
        try:
            WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, '//input[@id="name"]')))
            first_and_last_name_field = elem('//input[@id="name"]',driver)
            first_and_last_name_field.send_keys(sender_address_first_and_last_name)
        except Exception as ex:
            print(ex)

        # sending sender business
        try:
            sending_sender_street = elem('//input[@id="businessName"]',driver)
            sending_sender_street.send_keys(sender_address_Company)
        except Exception as ex:
            print(ex)

        # sending sender Street
        try:
            sending_sender_street = elem('//input[@id="street"]',driver)
            sending_sender_street.send_keys(sender_address_street)
        except Exception as ex:
            print(ex)

        # sending house number
        try:
            sending_sender_house_number = elem('//input[@id="houseNumber"]',driver)
            sending_sender_house_number.send_keys(sender_address_number)
        except Exception as ex:
            print(ex)

        # sending post code
        try:
            sending_sender_post_code = elem('//input[@id="postalCode"]',driver)
            sending_sender_post_code.send_keys(sender_address_postcode)
        except Exception as ex:
            print(ex)

        # sending city
        try:
            sending_sender_city = elem('//input[@id="city"]',driver)
            sending_sender_city.send_keys(sender_address_city)
        except Exception as ex:
            print(ex)

        # sending country in select/.
        try:
            select = Select(driver.find_element(by=By.XPATH, value = '*//select[@id="country"]'))
            sleep(.5)
            # select by visible text
            select.select_by_visible_text(f'{sender_address_country}')
        except Exception as ex:
            print(ex)

        # sending email_address
        try:
            sending_sender_email = elem('//input[@id="email"]',driver)
            sending_sender_email.send_keys(sender_address_email_address)
        except Exception as ex:
            print(ex)

        # sedning country code
        try:
            select_1 = Select(driver.find_element(by=By.XPATH, value = '*//select[@id="tel.country"]'))
            sleep(.5)
            # select by visible text
            select_1.select_by_visible_text(f'{sender_address_country_code}')
        except Exception as ex:
            print(ex)

        # sending phone number
        try:
            sending_sender_phone_number = elem('//input[@id="tel.number"]',driver)
            sending_sender_phone_number.send_keys(sender_address_phone_number)
        except Exception as ex:
            print(ex)

        # clicking on sender apply
        try:
            sender_apply = elem('//a[@title="Apply"]',driver)
            sender_apply.click()
        except Exception as ex:
            print(ex)

        ####################################################### RECIPIENT* #########################################################
        # clicking on Enter Recipient address button
        try:
            WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, '//a[@title="Enter recipient address"]')))
            enter_recipient_address = elem('//a[@title="Enter recipient address"]',driver)
            enter_recipient_address.click()
        except Exception as ex:
            print(ex)

        # sending company
        try:
            WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, '//input[@id="businessName"]')))
            enter_recipient_company = elem('//input[@id="businessName"]',driver)
            enter_recipient_company.send_keys(recipient_company)
        except Exception as ex:
            print(ex)

        # sending first and last name  of recipient
        try:
            recipient_first_and_last_name_field = elem('//input[@id="name"]',driver)
            recipient_first_and_last_name_field.send_keys(recipient_first_and_last_name)
        except Exception as ex:
            print(ex)

        # sending street  of recipient
        try:
            recipient_street_sending = elem('//input[@id="street"]',driver)
            recipient_street_sending.send_keys(recipient_street)
        except Exception as ex:
            print(ex)

        # sending Number house  of recipient
        try:
            recipient_house_number = elem('//input[@id="houseNumber"]',driver)
            recipient_house_number.send_keys(recipient_number)
        except Exception as ex:
            print(ex)

        # sending postal code  of recipient*************************************************** remaining
        # try:
        #     recipient_postacode_sending = elem('//input[@id="postalCode"]',driver)
        #     recipient_postacode_sending.send_keys("2536")
        # except Exception as ex:
        #     print(ex)

        # sending city  of recipient
        try:
            recipient_city_sending = elem('//input[@id="city"]',driver)
            recipient_city_sending.send_keys(recipient_city)
        except Exception as ex:
            print(ex)

        # sending country in select/.
        try:
            select_2 = Select(driver.find_element(by=By.XPATH, value = '*//select[@id="country"]'))
            sleep(.5)
            # select by visible text
            select_2.select_by_visible_text(f'{recipient_country}')
        except Exception as ex:
            print(ex)

        # sending region  of recipient
        try:
            recipient_region_sending = elem('//input[@id="region"]',driver)
            recipient_region_sending.send_keys(recipient_region)
        except Exception as ex:
            print(ex)

        # sending email of recipient
        try:
            recipient_email_sending = elem('//input[@id="email"]',driver)
            recipient_email_sending.send_keys(recipient_email_address)
        except Exception as ex:
            print(ex)

        # sedning country code
        try:
            select_3 = Select(driver.find_element(by=By.XPATH, value = '*//select[@id="tel.country"]'))
            sleep(.5)
            # select by visible text
            select_3.select_by_visible_text(f'{recipient_country_code}')
        except Exception as ex:
            print(ex)

        # sending phone number of recipient
        try:
            recipient_phone_number_sending = elem('//input[@id="tel.number"]',driver)
            recipient_phone_number_sending.send_keys(recipient_phone_number)
        except Exception as ex:
            print(ex)

        # clicking on recipient apply
        try:
            recipient_apply = elem('//a[@title="Apply"]',driver)
            recipient_apply.click()
        except Exception as ex:
            print(ex)

        ################################################## item category etc ########################################################
        # item category
        try:
            WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, '//select[@name="deliveryClassification"]')))
            select_4 = Select(driver.find_element(by=By.XPATH, value = '//select[@name="deliveryClassification"]'))
            sleep(.5)
            # select by visible text
            select_4.select_by_visible_text(f'{item_category}')

        except Exception as ex:
            print(ex)

        # currency
        try:
            select_5 = Select(driver.find_element(by=By.XPATH, value = '//select[@name="currency"]'))
            sleep(.5)
            # select by visible text
            select_5.select_by_visible_text(f'{currency}')

        except Exception as ex:
            print(ex)


        ###################################### description of content ##########################################
        # clicking on description of content
        try:
            WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, '//a[@title="Add content"]')))
            description_of_content_button = elem('//a[@title="Add content"]',driver)
            description_of_content_button.click()
        except Exception as ex:
            print(ex)

        # description of content field
        try:
            WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, '//textarea[@id="contentsummary"]')))
            description_of_content_textarea = elem('//textarea[@id="contentsummary"]',driver)
            description_of_content_textarea.send_keys(description_of_content)
        except Exception as ex:
            print(ex)


         # quantity
        try:
            quantity_field = elem('//input[@id="count"]',driver)
            quantity_field.send_keys(quantity)
        except Exception as ex:
            print(ex)

        # net weight per kg per unit
        try:
            net_weight_per_kg_per_unit_field = elem('//input[@id="nettoweight"]',driver)
            net_weight_per_kg_per_unit_field.send_keys(net_weight_per_kg)
        except Exception as ex:
            print(ex)

        # value in eur per unit
        try:
            value_in_eur_per_unit_field = elem('//input[@id="value"]',driver)
            value_in_eur_per_unit_field.send_keys(value_in_eur_per_unit)
        except Exception as ex:
            print(ex)

        # Tariff number
        try:
            tariff_number_field = elem('//input[@id="ratenumber"]',driver)
            tariff_number_field.send_keys(tariff_number)
        except Exception as ex:
            print(ex)

        # Country of origin
        try:
            WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, '//select[@name="origincountry"]')))
            select_4 = Select(driver.find_element(by=By.XPATH, value = '//select[@name="origincountry"]'))
            sleep(.5)
            # select by visible text
            select_4.select_by_visible_text(f'{country_of_origin_of_merchandise}')
        except Exception as ex:
            print(ex)


        # clicking on  apply
        try:
           add_content_apply = elem('//a[@title="Apply"]',driver)
           add_content_apply.click()
        except Exception as ex:
            print(ex)

        #################################################### additional information #####################################################
        driver.execute_script("window.scrollTo(0, window.scrollY + 700)")
        sleep(1.5)
        # gross weight in kg
        try:
            WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, '//input[@id="zoll_bruttoWeight"]')))

            gross_weight_per_kg = elem('//input[@id="zoll_bruttoWeight"]',driver)
            gross_weight_per_kg.send_keys(gross_weight_in_kg)
        except Exception as ex:
            print(ex)

        # infromation about the importer
        try:
            information_about_importer_field = elem('//input[@id="zoll_importeur"]',driver)
            information_about_importer_field.send_keys(information_about_importer)
        except Exception as ex:
            print(ex)

        # approval number
        try:
            approval_number_field = elem('//input[@id="zoll_approvalNumber"]',driver)
            approval_number_field.send_keys(approval_number)
        except Exception as ex:
            print(ex)

         # certificate number
        try:
            certificate_number_field = elem('//input[@id="zoll_certificateNumber"]',driver)
            certificate_number_field.send_keys(certificate_number)
        except Exception as ex:
            print(ex)

         # invoice number
        try:
            invoice_number_field = elem('//input[@id="zoll_receiptNumber"]',driver)
            invoice_number_field.send_keys(invoice_number)
        except Exception as ex:
            print(ex)


        # special remarks
        try:
            special_remarks_field = elem('//textarea[@id="zoll_specialNote"]',driver)
            special_remarks_field.send_keys(special_remarks)
        except Exception as ex:
            print(ex)

        # if Unavailable
        #if "Return to sender" in if_undelivered:
        #    try:
        #        return_to_sender = elem('(//input[@name="strategyUndeliverability"])[1]',driver)
        #        driver.execute_script("arguments[0].click();", return_to_sender)
        #    except Exception as ex:
        #        print(ex)
        #else:
        #    try:
        #        no_return = elem('(//input[@name="strategyUndeliverability"])[2]',driver)
        #        driver.execute_script("arguments[0].click();", no_return)
        #    except Exception as ex:
        #        print(ex)


        # # clicking on continue
        try:
            continue_btn = elem('//a[@title="Continue"]',driver)
            continue_btn.click()
        except Exception as ex:
            print(ex)


        # create custom form now
        try:
            WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, '//a[@title="Create customs form now"]')))
            sleep(1)
            custom_form_btn = elem('//a[@title="Create customs form now"]',driver)
            custom_form_btn.click()
        except Exception as ex:
            print(ex)

        # download custom form
        try:
            WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, '//a[@title="Download customs form"]')))
            sleep(1)
            custom_form_btn = elem('//a[@title="Download customs form"]',driver)
            custom_form_btn.click()
        except Exception as ex:
            print(ex)

        recipient_number_now = recipient_number_now + 1
        driver.get(url)
        sleep(.5)

#driver.quit()
print("Finished.")
