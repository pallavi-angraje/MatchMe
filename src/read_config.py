from bs4 import BeautifulSoup
import os
from os import path

sourcedir = path.dirname(__file__)


def get_handler():
    conf_file = open(path.join(sourcedir, '../config/conf_invoices.xml'),"r")
    contents = conf_file.read()
    soup = BeautifulSoup(contents, 'xml')
    return soup


def get_requred_fields(handler):
    return handler.find_all('fields')[0].get_text().split(',')


def get_source_file_list(handler):
    source_path = handler.find_all('filepath')[0].get_text()
    file_list = [path.join(source_path, file_name) for file_name in os.listdir(source_path) if path.isfile(os.path.join(source_path, file_name))]
    return file_list


def get_db_config(handler):
    host = handler.find_all('host')[0].get_text()
    user = handler.find_all('user')[0].get_text()
    db = handler.find_all('db')[0].get_text()
    passwd = handler.find_all('passwd')[0].get_text()
    return host, user, db, passwd


def get_invoice_tables(handler):
    invoice_data_storage_table = handler.find_all('invoice-data-table')[0].get_text()
    invoice_master = handler.find_all('invoice-master-table')[0].get_text()
    return invoice_data_storage_table, invoice_master


def get_payment_tables(handler):
    payment_master = handler.find_all('payment-master-table')[0].get_text()
    payment_data_storage_table = handler.find_all('payment-data-table')[0].get_text()
    return payment_data_storage_table, payment_master

def get_matched_data_table(handler):
    matched_data_table = handler.find_all('matched-data-table')[0].get_text()
    return matched_data_table


