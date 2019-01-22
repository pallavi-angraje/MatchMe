import re
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import StringIO
import numpy as np
from tabula import read_pdf
from read_config import *
import pymysql as pm
import docx2txt
import datetime
import logging
import datetime
import os
import pandas as pd
import comtypes.client
import time
from dateutil.parser import parse



def extract_required_data_from_excel(file, field_list, doc_num):
    df = pd.read_excel(file)
    df = df.loc[:, df.columns.isin(field_list)]
    cols = df.loc[:, df.dtypes == 'datetime64[ns]'].columns.tolist()
    for col in cols:
        df[col] = pd.to_datetime(df[col], format("%Y-%m-%d"), "%Y-%m-%d").astype(str)
    res = df.T.to_dict().values()
    res_list = list()
    for a in res:
        a.update({'source_file_name': file.split('/')[-1], 'DocNum_Inv': doc_num})
        res_list.append(a)
    return res_list


def convert_doc_to_pdf(in_file, file):
    wdFormatPDF = 17
    out_file = file.split('.')[0]+'.pdf'
    # print out filenames
    print(in_file)
    print(out_file)

    # create COM object
    word = comtypes.client.CreateObject('Word.Application')
    # key point 1: make word visible before open a new document
    word.Visible = True
    # key point 2: wait for the COM Server to prepare well.
    time.sleep(3)

    # convert docx file 1 to pdf file 1
    doc = word.Documents.Open(in_file)  # open docx file 1
    doc.SaveAs(out_file, FileFormat=wdFormatPDF)  # conversion
    doc.Close()  # close docx file 1
    word.Visible = False
    word.Quit()  # close Word Application
    return out_file

# Synonyms for field names
def update_field_list(field_list):
    if 'order number' in field_list:
        field_list.append('Invoice Number')
    if 'Account No' in field_list:
        field_list.append('Acoount Number')

    return field_list

def insert_data_to_table_single(item, connection, master_table, logger, column_mapping):
    insert_query = 'insert into {0} ({1}) values ({2})'
    cursor = connection.cursor()
    item = alter_result_dictionary(item, column_mapping)
    values = tuple(item.values())
    column_names = ','.join(item.keys())
    subtitute_str = '%s,' * len(item)
    subtitute_str = subtitute_str[:-1]
    print(master_table, column_names, subtitute_str)
    try:
        cursor.execute(insert_query.format(master_table, column_names, subtitute_str), values)
    except Exception as e2:
        print("failed: ", item)
        print(e2)
        logger.error(str(e2)+':'+str(item.values()))
        pass
    else:
        logger.info("Data extracted for :"+column_names)
        logger.info("Extracted values: "+str(list(values)))
        print("Successfully inserted")
        logger.info("Processing completed")
    connection.commit()
    cursor.close()



def insert_invoice_data_batch(result_list, connection, logger, master_table, column_mapping):
    insert_query = 'insert into {0} ({1}) values ({2})'
    values_list = list()
    cursor = connection.cursor()
    for result_dictionary in result_list:
        result_dictionary = alter_result_dictionary(result_dictionary, column_mapping)
        values_list.append(tuple(result_dictionary.values()))
    column_names = ','.join(result_list[0].keys())
    subtitute_str = '%s,'*len(result_list[0].keys())
    subtitute_str = subtitute_str[:-1]
    logger.info("Number of rows :" + str(len(result_list)))
    logger.info("Data extracted for :" + column_names)
    logger.info("Extracted values: "+str(values_list))
    try:
        cursor.executemany(
            insert_query.format(master_table, column_names, subtitute_str), values_list)
        print("batch")
    except Exception as e1:
        print("single")
        print(e1)
        for item in values_list:
            subtitute_str = '%s,' * len(result_list[0].keys())
            subtitute_str = subtitute_str[:-1]
            try:
                cursor.execute(insert_query.format(master_table, column_names, subtitute_str), item)
            except Exception as e2:
                print("single failed")
                print(e2)
                pass
        pass
    connection.commit()
    cursor.close()
    logger.info("Processing completed")


def alter_result_dictionary(result_dictionary, column_mapping):
    #support for date column in YYYY-MM-DD or DD/MM/YYYY
    for key in list(result_dictionary.keys()):
        if key not in column_mapping:
            print(key, " not found in could mapping")
            del result_dictionary[key]
        else:
            val = result_dictionary[key]
            del result_dictionary[key]
            result_dictionary.update({column_mapping[key]: val})
    return result_dictionary



def is_date(string):
    try:
        parse(string)
        return True
    except ValueError:
        return False


def get_date_string(content_list):
    try:
        for element in content_list:
            ele_list = element.split()
            for ele in ele_list:
                print("****")
                print(ele)
                # check if the date is string of single  word
                if len(ele)==1 and '-' in ele:
                    if is_date(ele):
                        return ele
                elif len(ele)==2:
                    if is_date(ele) or is_date(ele[0]) or is_date(ele[1]):
                        return ele
                elif len(ele) == 3:
                    # 10 jan 2018
                    if is_date(ele):
                        return ele
                    # combination of 2
                    elif is_date(element.rsplit(' ', 1)[0]) or is_date(element.split(' ', 1)[0]):
                        return ele
                else:
                    for i in range(len(ele_list)-2):
                        # check if date is of 3 words
                        a = ele_list[i]+' '+ele_list[i+1]+' '+ele_list[i+2]
                        if is_date(ele_list[i]+' '+ele_list[i+1]+' '+ele_list[i+2]):
                            return ele_list[i]+' '+ele_list[i+1]+' '+ele_list[i+2]
                        # check if date is of 2 words
                        elif is_date(ele_list[i]+' '+ele_list[i+1]):
                            str1=ele_list[i]+' '+ele_list[i+1]
                            if len(str1.replace('-',' ').split())==3:
                                return ele_list[i]+' '+ele_list[i+1]
                        elif is_date(ele_list[i+1]+' '+ele_list[i+2]):
                            str2 = ele_list[i+1]+' '+ele_list[i+2]
                            if len(str2.replace('-', ' ').split()) == 3:
                                return ele_list[i+1]+' '+ele_list[i+2]
                        # check if date is of one word
                        elif ele_list[i].count('-') == 2 and is_date(ele_list[i]):
                            return ele
                        elif ele_list[i].count('-') == 2 and is_date(ele_list[i+1]):
                            return ele
                        elif ele_list[i].count('-') == 2 and is_date(ele_list[i+2]):
                            return ele
    except Exception as e:
        print(e)




def extract_required_data_from_text(content, field_list, logger):
    content_list = content.split('\n')
    # check for first occurance of date
    content_list = list(filter(None, content_list))
    content_list = [ele for ele in content_list if not re.match('^[\W*]$', ele)]
    date_str = get_date_string(content_list)
    date_str = impose_standard_date_format(date_str, logger)
    print(date_str)
    regex = re.compile('[,!:-]')
    data_dict = dict()
    for field in field_list:
        index = 0
        data_dict.update({field: None})
        for ele in content_list:
            value = None
            if field.lower() in ele.lower():
                print(ele)
                value = regex.sub('', ele).strip()
                value = value.replace(field, '')
                value = value.strip()
                if value == '':
                    value = content_list[index + 1]
                    value = regex.sub('', value).strip()
            if value:
                data_dict.update({field: value})
                break
            index = index + 1
    data_dict.update({'statement date':date_str})
    return data_dict


def convert_pdf_data_to_text(path):
    rsrmgr = PDFResourceManager()
    retstr = StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    device = TextConverter(rsrmgr, retstr, codec = codec, laparams=laparams)
    fp = open(path, 'rb')
    interpreter = PDFPageInterpreter(rsrmgr, device)
    password = ""
    maxpages = 0
    caching = True
    pagenos = set()
    for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password, caching=caching
                                  ,check_extractable=True):
        interpreter.process_page(page)
    text = retstr.getvalue()
    fp.close()
    device.close()
    retstr.close()
    return text


def extract_required_data_from_df(df_list, field_list, logger):
    data_dict = dict()
    for field in field_list:
        field_original = field
        for df in df_list:
            df_field_list = [x for x in df.columns.tolist() if x is not np.nan]
            required_cols = [col for col in df_field_list if field in col]
            if len(required_cols)> 1:
                logger.error("Found ambiguity between "+str(required_cols)+" while searching for :"+field)
                break
            elif len(required_cols) == 1:
                field = required_cols[0]
                if len(df[df[field].notnull()][field]) == 0:
                    value = None
                else:
                    value = df[df[field].notnull()][field][-1]
                if value is not None and field in value:
                    value = value[len(field_original):]
                    value = re.sub(r'[\x00-\x7F]+|^[\W]*|[\W]*$', '', value)
                if value:
                    data_dict.update({field_original: value})
            else:
                pass
    return data_dict


def convert_pdf_to_df(path):
    df_list = read_pdf(path, pandas_options={'header': 0}, multiple_tables=True)
    df_list[0].to_csv("sdsasd.csv")
    df_list = list(filter(None.__ne__, df_list))
    df_transpose_list = list()
    index = 0
    for df in df_list:
        df_transpose = df.T
        df_transpose.drop(columns=[np.nan], inplace=True, errors='ignore')
        new_header = df_transpose.iloc[0]
        df_transpose = df_transpose[1:]
        df_transpose.columns = new_header
        df_transpose_list.append(df_transpose)
        index = index+1
    return df_transpose_list


def connect_to_db(handler):
    host, user, db, password = get_db_config(handler)
    return pm.connect(host, user, password, db, use_unicode=True, charset="utf8mb4")



def impose_standard_date_format(date_str, logger):
    # 20 dec 2018, 20-dec-2018, 2018 dec 20, 2018-dec-20, 20-12-2018 DD/MM/YYYY to YYYYMMDD
    try:
        datetime_object = datetime.datetime.strptime(date_str, "%d/%m/%Y")
    except Exception as e:
        if '/' in date_str:
            try:
                datetime_object = datetime.datetime.strptime(date_str, "%d/%m/%Y")
            except Exception as b1:
                datetime_object = datetime.datetime.strptime(date_str, "%Y/%m/%d")
        else:
            try:
                datetime_object = datetime.datetime.strptime(date_str, "%d %b %Y")
            except Exception as e1:
                try:
                    datetime_object = datetime.datetime.strptime(date_str, "%d-%b-%Y")
                except Exception as e2:
                    try:
                        datetime_object = datetime.datetime.strptime(date_str, "%Y %b %d")
                    except Exception as e3:
                        try:
                            datetime_object = datetime.datetime.strptime(date_str, "%Y-%b-%d")
                        except Exception as e4:
                            try:
                                datetime_object = datetime.datetime.strptime(date_str, "%d/%m/%Y")
                            except Exception as e5:
                                logger.error("Unable to bring " + date_str + " to yyyy-mm-dd format")
        date_str = datetime_object.strftime("%Y-%m-%d")
    return date_str
