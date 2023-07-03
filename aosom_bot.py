import time
import warnings
import pandas as pd
import telebot
from selenium.webdriver import Chrome
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
from html_el_dict import *
from sql_scripts import *
from to_excel import *
from config import *


class Objects:
    def __init__(self, wait):
        self.wait = wait


    def check_for_general_info(self, general_selector):
        try:
            self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, general_selector)))
            return True
        except:
            return False


    def instock_status(self, buy_btn):
        try:
            self.wait.until(EC.element_to_be_clickable((By.XPATH, buy_btn)))
            return True
        except:
            return False


    def parse_info_from_page(self, browser, right_menu_class, title_class, price_class, alt_price_class):
        html = browser.page_source
        soup = BeautifulSoup(html, 'html.parser')
        right_menu = soup.find('div', class_=right_menu_class)
        try:
            title = right_menu.find('h1', class_=title_class).get_text()
            title = " ".join([i for i in [i.replace('\n', '') for i in [i for i in title.split(" ") if len(i) > 0]] if len(i) > 0])
        except:
            title = 'None'
        try:
            try:
                price = right_menu.find('div', class_=price_class).get_text()
            except:
                price = right_menu.find('div', class_=alt_price_class).get_text()
            price = price.replace('CA$', '')
            price = price.replace(',', '')
            price = float(price)
        except:
            price = 'None'
        try:
            td_element = right_menu.find('td', text='SKU')
            sku = td_element.find_next('td').text.strip()
        except:
            sku = 'None'
        return title, price, sku


def telegram_msg(bot_status):
    bot = telebot.TeleBot(telegram_token)
    if bot_status == "start":
        bot.send_message(telegram_user_id, "🟢 The bot started scraping websites. 🟢\nDo not open the excel spreadsheet that the bot is currently working with, otherwise some information may be lost.")
    elif bot_status == "end":
        bot.send_message(telegram_user_id, '♦️ The bot has finished parsing websites. ♦️\nThe bot has stopped working️, now you can work with excel spreadsheets. (folder "out")')


def main(browser, link, current_excel_name):
    wait = WebDriverWait(browser, 10)
    elem = html_elements
    objects = Objects(wait)
    try:
        if objects.check_for_general_info(elem['general_info_css']) == True:
            item_info_list = list()
            item_title, item_price, item_sku = objects.parse_info_from_page(browser, elem['right_menu_class'], elem['title_class'], elem['price_class'], elem['alt_price_class'])
            if objects.instock_status(elem['buy_btn']) == True:
                item_stock = 'IN STOCK'
            else:
                item_stock = 'OUT OF STOCK'
            item_info_list.append(item_sku)
            item_info_list.append(item_price)
            item_info_list.append(item_stock)
            item_info_list.append(item_title)
            item_info_list.append(link)

            if item_exists(item_sku) == False:
                add_item_to_db(item_sku, item_price, item_stock)
                add_new_data([item_info_list], current_excel_name)
            elif item_exists(item_sku) == True:
                prev_item_price, prev_item_stock = get_item_info(item_sku)
                if prev_item_price == item_price and prev_item_stock == item_stock:
                    add_new_data([item_info_list], current_excel_name)
                elif prev_item_stock != item_stock and prev_item_price == item_price:
                    add_data_with_stock_change([item_info_list], current_excel_name, red_color)
                    update_item_info(item_sku, item_stock, item_price)
                elif prev_item_stock == item_stock and prev_item_price != item_price:
                    add_data_with_price_change([item_info_list], current_excel_name, yellow_color)
                    update_item_info(item_sku, item_stock, item_price)
                elif prev_item_stock != item_stock and prev_item_price != item_price:
                    add_data_with_stock_and_price_change([item_info_list], current_excel_name, yellow_color, red_color)
                    update_item_info(item_sku, item_stock, item_price)
            time.sleep(10)
    except:
        time.sleep(10)


if __name__ == "__main__":
    options = Options()
    warnings.filterwarnings("ignore", category=DeprecationWarning)
    options.binary_location = ""
    try:
        browser = Chrome("chromedriver.exe", chrome_options=options)
        browser.maximize_window()

        current_excel_name = create_new_excel()

        df = pd.read_excel(aosom_link_excel_file)
        link_column = df['Supplier Product Link']

        telegram_msg("start")

        for i in range(len(link_column)):
            link = link_column.iloc[i]
            browser.get(link)
            main(browser, link, current_excel_name)

        telegram_msg("end")
    except:
        telegram_msg("end")