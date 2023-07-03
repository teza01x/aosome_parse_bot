import sqlite3
from config import *


def item_exists(sku):
    conn = sqlite3.connect('items.db')
    cursor = conn.cursor()

    result = cursor.execute("SELECT sku FROM item_info WHERE sku = ?", (sku,))
    exists = bool(len(result.fetchall()))

    conn.close()

    return exists


def add_item_to_db(sku, price, stock):
    conn = sqlite3.connect('items.db')
    cursor = conn.cursor()

    cursor.execute("INSERT INTO item_info (sku, price, stock) VALUES(?, ?, ?)", (sku, price, stock,))

    conn.commit()
    conn.close()


def get_item_info(sku):
    conn = sqlite3.connect('items.db')
    cursor = conn.cursor()

    result = cursor.execute("SELECT price, stock FROM item_info WHERE sku = ?", (sku,))
    result = result.fetchall()[0]
    conn.close()

    return result


def update_item_info(sku, stock_status, item_price):
    conn = sqlite3.connect('items.db')
    cursor = conn.cursor()

    cursor.execute("UPDATE item_info SET price = ?, stock = ? WHERE sku = ?", (item_price, stock_status, sku,))

    conn.commit()
    conn.close()
