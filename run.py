#!/usr/bin/env python3
from selenium import webdriver
import selenium
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
import imaplib, email
import configparser
import os
import re
from datetime import datetime
from pprint import pprint
from random import *


scriptdir = os.path.dirname(os.path.abspath(__file__))
os.chdir(scriptdir)

CONFIG = configparser.ConfigParser()


def get_emails():
    return []


# User ID: HLN 32588 26053
# PASSWORD: 22370 26027
def get_survey_user(email_text=str):
    return "HLN 32588 26053"


def get_survey_password(email_text=str):
    return "22370 26027"


# returns the authenticated gmail connection
def setup_gmail_connection():
    con = imaplib.IMAP4_SSL(CONFIG.get("DEFAULT", "IMAP_URL"))
    username = CONFIG.get("DEFAULT", "GMAIL_USER")
    password = CONFIG.get("DEFAULT", "GMAIL_PASSWORD")

    con.login(username, password)
    return con


# Python imaplib requires mailbox names with spaces to be surrounded by apostrophes. So imap.select("INBOX") is fine, but with spaces you'd need imap.select("\"" + "Label Name" + "\"").
# https://stackoverflow.com/questions/65692144/python-imaplib-cant-select-custom-gmail-labels
def run_search_with_label(con):
    con.select('"' + "Home Depot" + '"')
    return get_emails_bytes(
        search("FROM", CONFIG.get("DEFAULT", "HOME_DEPOT_FROM"), con),
        con,
    )


def search(key, value, con):
    result, data = con.search(None, key, '"{}"'.format(value))
    return data


def get_body(msg):
    if msg.is_multipart():
        return get_body(msg.get_payload(0))
    else:
        return msg.get_payload(None, True)


def get_emails_bytes(result_bytes, connection):
    msgs = []  # all the email data are pushed inside an array
    for num in result_bytes[0].split():
        typ, data = connection.fetch(num, "(RFC822)")
        msgs.append(data)

    return msgs


def find_survey_data(email_body=str):
    # regex to extract required strings
    user_expression = 'padding:0px">User ID: (.*?)<\/pre>'
    password_expression = 'padding:0px">PASSWORD: (.*?)<\/pre>'

    username_res = re.findall(user_expression, email_body)
    password_res = re.findall(password_expression, email_body)

    if len(username_res) and len(password_res):
        return str(username_res[:1][0]), str(password_res[:1][0])
    else:
        return None, None


def do_survey(username=str, password=str):
    service = Service("C:\Program Files (x86)\chromedriver.exe")
    options = webdriver.ChromeOptions()
    options.add_experimental_option("excludeSwitches", ["enable-logging"])

    driver = webdriver.Chrome(service=service, options=options)
    driver.get(CONFIG.get("DEFAULT", "SURVEY_URL"))

    try:
        driver.implicitly_wait(10)

        # begin_button = driver.find_element(By.ID, "beginButton")
        # begin_button.click()

        next_button = driver.find_element(By.ID, "nextButton")
        next_button.click()

        zipcode_textbox = driver.find_element(
            By.NAME, "spl_q_thd_postalcode_entry_text"
        )
        zipcode_textbox.send_keys(CONFIG.get("DEFAULT", "ZIPCODE"))

        driver.implicitly_wait(2)

        # If you complete the survey, you will have the opportunity to be entered into a drawing for a $5,000 Home Depot gift card.
        next_button = driver.find_element(By.ID, "nextButton")
        next_button.click()

        driver.implicitly_wait(2)

        next_button.click()

        # HLN 32588 26053
        username_textbox = driver.find_element(
            By.NAME, "spl_q_thd_receiptcode_id_entry_text"
        )
        username_textbox.send_keys(username)

        # 22370 26027
        password_textbox = driver.find_element(
            By.NAME, "spl_q_thd_receiptcode_password_entry_text"
        )
        password_textbox.send_keys(password)

        driver.implicitly_wait(2)
        next_button.click()

        # occupation
        occupation_button = driver.find_element(
            By.ID, "i_onf_q_thd_pro_classification_handyman_yn_1"
        )
        occupation_button.click()
        driver.implicitly_wait(2)
        next_button.click()

        # likely
        likely_button = driver.find_element(By.NAME, "onf_q_thd_shop_likely_radio")
        likely_button.click()
        driver.implicitly_wait(2)
        next_button.click()

        # experience
        # <input type="radio" name="onf_q_thd_shop_comparison_radio" value="5" title="5" class="radio">
        experience_button = driver.find_element_by_css_selector(
            "input[type='radio'][name=onf_q_thd_shop_comparison_radio][value='"
            + randint(4, 5)
            + "']"
        )
        experience_button.click()
        driver.implicitly_wait(2)
        next_button.click()

        # <textarea class="textareas" name="spl_q_thd_shop_experience_text" id="spl_q_thd_shop_experience_text_input" maxlength="500"></textarea>
        # Why did you say that this shopping experience was much better compared to other home improvement stores?
        textarea_for_experience = driver.find_element(
            By.ID, "spl_q_thd_shop_experience_text"
        )
        textarea_for_experience.send_keys(
            "I have had so much more luck finding what I need at Home Depot than at other stores."
        )
        driver.implicitly_wait(2)
        next_button.click()

        # How satisfied were you with each of these areas during this visit to The Home Depot? Please use the full scale below.
        # TWO QUESTIONS ON THIS PAGE
        # onf_q_thd_satisfied_checkout_process_radio
        # onf_q_thd_satisfied_employees_experience_radio

        satisfied_checkout = driver.find_element_by_css_selector(
            "input[type='radio'][name=onf_q_thd_satisfied_checkout_process_radio][value='"
            + randint(4, 5)
            + "']"
        )
        satisfied_checkout.click()

        employees_experience = driver.find_element_by_css_selector(
            "input[type='radio'][name=onf_q_thd_satisfied_employees_experience_radio][value='"
            + randint(4, 5)
            + "']"
        )
        employees_experience.click()

        driver.implicitly_wait(2)
        next_button.click()

        # The store was generally neat and clean
        # onf_q_thd_perceptions_neat_clean_radio

        # Had sufficient quantities of product in-stock on the shelves
        # onf_q_thd_perceptions_in_stock_radio

        # The cashier was friendly
        # onf_q_thd_perceptions_cashier_friendly_radio

        # Employees were consistently friendly throughout the store
        # onf_q_thd_perceptions_employee_friendly_radio
        questions_exp = [
            "onf_q_thd_perceptions_neat_clean_radio",
            "onf_q_thd_perceptions_in_stock_radio",
            "onf_q_thd_perceptions_cashier_friendly_radio",
            "onf_q_thd_perceptions_employee_friendly_radio",
        ]
        for question in questions_exp:
            radiobutton = driver.find_element_by_css_selector(
                "input[type='radio'][name="
                + str(question)
                + "][value='"
                + randint(4, 5)
                + "']"
            )
            radiobutton.click()
        next_button.click()

        # i_onf_q_thd_get_employee_greeting_yn_1
        # <input id="i_onf_q_thd_get_employee_greeting_yn_1" class="radio" type="radio" name="onf_q_thd_get_employee_greeting_yn" value="1">
        questions_exp = [
            "onf_q_thd_get_employee_greeting_yn",
            "onf_q_thd_get_employee_interest_addressing_needs_yn",
            "onf_q_thd_get_employee_thanked_you_yn",
        ]
        for question in questions_exp:
            radiobutton = driver.find_element_by_css_selector(
                "input[type='radio'][name="
                + str(question)
                + "][value='"
                + randint(0, 1)
                + "']"
            )
            radiobutton.click()
        next_button.click()

        # Is there anything else you would like to tell us about your visit or is there anything we could do to improve your shopping experience?
        # Nothing to add
        # onf_q_thd_catchall_oe_optout_yn
        nothing_to_add = driver.find_element_by_css_selector(
            "input[type='radio'][name=onf_q_thd_catchall_oe_optout_yn][value='1']"
        )
        nothing_to_add.click()
        next_button.click()

        # In this time of COVID-19 how comfortable did you feel shopping in the store?
        # onf_i_question_1
        covid_radio_button = driver.find_element_by_css_selector(
            "input[type='radio'][name=onf_i_question_1][value='4']"
        )
        covid_radio_button.click()

        # There is a textarea here that can be filled out or skipped

        next_button.click()

        # We would like to enter your name into a drawing for a $5,000 Home Depot gift card!

        # first name
        first_name = driver.find_element(By.NAME, "spl_q_thd_contact_first_name_text")
        first_name.send_keys(CONFIG.get("DEFAULT", "NAME_FIRST"))

        # last name
        last_name = driver.find_element(By.NAME, "spl_q_thd_contact_last_name_text")
        last_name.send_keys(CONFIG.get("DEFAULT", "NAME_LAST"))

        # email
        email_field = driver.find_element(
            By.NAME, "spl_q_thd_contact_email_sweeps_text"
        )
        email_field.send_keys(CONFIG.get("DEFAULT", "EMAIL"))

        # phone
        phone_field = driver.find_element(
            By.NAME, "spl_q_thd_contact_phone_sweeps_text"
        )
        phone_field.send_keys(CONFIG.get("DEFAULT", "PHONE"))

        done_button = driver.find_element(By.ID, "finishButton")
        done_button.click()

        # You have been entered into the sweepstakes for a $5,000 Home Depot Gift Card.
        # Winners will be selected in random drawings on or about the date listed in the Sweepstakes Rules and will be notified by email or phone.
        driver.find_elements_by_xpath(
            "(//*[contains(text(), '"
            + "You have been entered into the sweepstakes for a $5,000 Home Depot Gift Card."
            + "')] | //*[@value='"
            + "You have been entered into the sweepstakes for a $5,000 Home Depot Gift Card."
            + "'])"
        )

    except Exception as ex:
        print("JoeyException: " + str(ex))
        driver.quit()


if __name__ == "__main__":
    print("Pulling data from config.ini")
    CONFIG.read("config.ini")

    # CONFIG.get("DEFAULT", "MAILGUN_FROM")
    print("Making connection to GMAIL")
    connection_to_gmail = setup_gmail_connection()

    # print("Pulling data")
    # emails = get_emails(connection_to_gmail)

    # print("Emails pulled with no error")

    # printing them by the order they are displayed in your gmail
    # print("Looping emails")
    custom = True
    has_run = False

    # H8B 98019 97498
    # 22368 9748 5

    do_survey("H8B 98019 97498", "22368 97485")

    search_results = []
    # run_search_with_label(connection_to_gmail)
    for email in search_results:
        raw_body = email[0][1]
        # email variable is a list
        # print(raw_body)
        print(
            "========================================================================================"
        )
        # print(type(email), email)

        survey_username, survey_password = find_survey_data(str(raw_body))
        if survey_username is not None and survey_password is not None:
            # print(survey_username, survey_password)
            do_survey(survey_username, survey_password)
        else:
            print("Not found")
