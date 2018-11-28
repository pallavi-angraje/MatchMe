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


logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

column_mapping = {
    'Invoice Date': 'Invoice_Date',
    'Address': 'Address',
    'Party Name': 'Name',
    'Total': 'Total_Amount',
    'Delivery Address': 'Address',
    'source_file_name': 'File_Name',
    'Invoice Number': 'Invoice_Number',
    'Order Total': 'Total_Amount',
    'Amazon.in order number': 'Invoice_Number',
    'DocNum_Inv': 'DocNum_Inv'

}


def extract_required_data_from_text(content, field_list):
    content_list = content.split('\n')
    content_list = list(filter(None, content_list))
    content_list = [ele for ele in content_list if not re.match('^[\W*]$', ele)]
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


def extract_required_data_from_df(df_list, field_list):
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


def alter_result_dictionary(result_dictionary):
    #support for date column in YYYY-MM-DD or DD/MM/YYYY
    try:
        if 'Invoice Date' in result_dictionary and result_dictionary['Invoice Date'] is not None and isinstance(result_dictionary['Invoice Date'], str):
            datetime.datetime.strptime(result_dictionary['Invoice Date'], '%Y-%m-%d')
    except Exception as e:
        format_str = '%d/%m/%Y'
        datetime_obj = datetime.datetime.strptime(result_dictionary['Invoice Date'], format_str)
        result_dictionary.update({'Invoice Date': datetime_obj.date()})
        pass
    for key in list(result_dictionary.keys()):
        if key not in column_mapping:
            print(key, " not found in could mapping")
            del result_dictionary[key]
        else:
            val = result_dictionary[key]
            del result_dictionary[key]
            result_dictionary.update({column_mapping[key]: val})
    return result_dictionary


def insert_invoice_data_batch(result_list, connection):
    insert_query = 'insert into {0} ({1}) values ({2})'
    values_list = list()
    cursor = connection.cursor()
    for result_dictionary in result_list:
        result_dictionary = alter_result_dictionary(result_dictionary)
        values_list.append(tuple(result_dictionary.values()))
    column_names = ','.join(result_list[0].keys())
    subtitute_str = '%s,'*len(result_list[0].keys())
    subtitute_str = subtitute_str[:-1]
    logger.info("Number of rows :" + str(len(result_list)))
    logger.info("Data extracted for :" + column_names)
    logger.info("Extracted values: "+str(values_list))
    try:
        cursor.executemany(
            insert_query.format(data_table, column_names, subtitute_str), values_list)
        print("batch")
    except Exception as e1:
        print("single")
        print(e1)
        for item in values_list:
            subtitute_str = '%s,' * len(result_list[0].keys())
            subtitute_str = subtitute_str[:-1]
            try:
                cursor.execute(insert_query.format(data_table, column_names, subtitute_str), item)
            except Exception as e2:
                print("single failed")
                print(e2)
                pass
        pass
    connection.commit()
    cursor.close()
    logger.info("Processing completed")


def insert_invoice_data_single(item, connection):
    insert_query = 'insert into {0} ({1}) values ({2})'
    cursor = connection.cursor()
    item = alter_result_dictionary(item)
    values = tuple(item.values())
    column_names = ','.join(item.keys())
    subtitute_str = '%s,' * len(item)
    subtitute_str = subtitute_str[:-1]
    try:
        cursor.execute(insert_query.format(data_table, column_names, subtitute_str), values)
    except Exception as e2:
        print("failed: ", item)
        print(e2)
        logger.error(e2+':'+str(item.values()))
        pass
    else:
        logger.info("Data extracted for :"+column_names)
        logger.info("Extracted values: "+str(list(values)))
        print("Successfully inserted")
        logger.info("Processing completed")
    connection.commit()
    cursor.close()


def get_log_handler(file_path):
    current_time = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%d__%H_%M_%S_%f%Z")
    # helpful when files are with same name but different extensions
    log_path = os.path.join(os.path.dirname(__file__), '../logs/'+file_path.split('/')[-1].split('.')[0]+'_'+file_path.split('/')[-1].split('.')[-1]+'_Inv_'+current_time+'.log')
    print(log_path)
    handler = logging.FileHandler(log_path, encoding="UTF-8")
    formatter = logging.Formatter('%(levelname)s :  %(message)s AT - %(asctime)s')
    handler.setFormatter(formatter)
    return handler


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


def master_data_insert(file, connection):
    query = 'insert into {0} (Extract_Location) Values("{1}")'
    cursor = connection.cursor()
    num_rows = cursor.execute(query.format(master_table, '/'.join(file.split('/')[:-1])))
    connection.commit()
    if num_rows > 0:
        cursor.execute('select max(DocNum_Inv) from {0}'.format(master_table))
        doc_num = cursor.fetchone()[0]
        cursor.close()
        return doc_num
    else:
        logger.error("Unable to insert master data for "+file)
        return None


def master_data_update(doc_num, connection):
    query = 'update {0} set Processed = "Y" where DocNum_Inv = {1}'
    cursor = connection.cursor()
    cursor.execute(query.format(master_table, doc_num))
    connection.commit()
    cursor.close()


def convert_doc_to_pdf(in_file):
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


if __name__ == '__main__':
    conf_handler = get_handler()
    file_list = get_source_file_list(conf_handler)
    field_list = get_requred_fields(conf_handler)
    data_table, master_table = get_tables(conf_handler)
    connection = connect_to_db(conf_handler)
    print("File List:", file_list)
    for file in file_list:
        handler = get_log_handler(file)
        logger.addHandler(handler)
        print("File ", file)
        logger.info('File: '+file)
        doc_num = master_data_insert(file, connection)
        print(doc_num)
        result = dict()
        if doc_num is not None:
            if file.endswith('.pdf'):
                df_list = convert_pdf_to_df(file)
                if len(df_list) > 0:
                    result = extract_required_data_from_df(df_list, field_list)
                    result.update({'source_file_name': file.split('/')[-1], 'DocNum_Inv': doc_num})
                    print(result)
                    insert_invoice_data_single(result, connection)
                else:
                    content = convert_pdf_data_to_text(file)
                    result = extract_required_data_from_text(content, field_list)
                    result.update({'source_file_name': file.split('/')[-1], 'DocNum_Inv': doc_num})
                    print(result)
                    insert_invoice_data_single(result, connection)
            elif file.endswith('.docx'):
                content = docx2txt.process(file)
                result = extract_required_data_from_text(content, field_list)
                result.update({'source_file_name': file.split('/')[-1], 'DocNum_Inv': doc_num})
                print(result)
                insert_invoice_data_single(result, connection)
            elif file.endswith('.xlsx') or file.endswith('.xls'):
                results = extract_required_data_from_excel(file, field_list, doc_num)
                insert_invoice_data_batch(results, connection)
            master_data_update(doc_num, connection)
        logger.removeHandler(handler)
    connection.close()
