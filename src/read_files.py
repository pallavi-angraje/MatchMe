from common_functions import *



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
    'order number': 'Invoice_Number',
    'DocNum_Inv': 'DocNum_Inv',

}


def get_log_handler(file_path):
    current_time = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%d__%H_%M_%S_%f%Z")
    # helpful when files are with same name but different extensions
    log_path = os.path.join(os.path.dirname(__file__), '../logs/'+file_path.split('/')[-1].split('.')[0]+'_'+file_path.split('/')[-1].split('.')[-1]+'_Inv_'+current_time+'.log')
    print(log_path)
    handler = logging.FileHandler(log_path, encoding="UTF-8")
    formatter = logging.Formatter('%(levelname)s :  %(message)s AT - %(asctime)s')
    handler.setFormatter(formatter)
    return handler




def inv_master_data_insert(file, connection):
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


# def master_data_update(doc_num, connection):
#     query = 'update {0} set Processed = "Y" where DocNum_Inv = {1}'
#     cursor = connection.cursor()
#     cursor.execute(query.format(master_table, doc_num))
#     connection.commit()
#     cursor.close()



if __name__ == '__main__':
    conf_handler = get_handler()
    file_list = get_source_file_list(conf_handler)
    field_list = get_requred_fields(conf_handler)
    field_list = update_field_list(field_list)
    data_table, master_table = get_invoice_tables(conf_handler)
    connection = connect_to_db(conf_handler)
    print("File List:", file_list)
    for file in file_list:
        handler = get_log_handler(file)
        logger.addHandler(handler)
        print("File ", file)
        logger.info('File: '+file)
        doc_num = inv_master_data_insert(file, connection)
        print(doc_num)
        result = dict()
        if doc_num is not None:
            if file.endswith('.pdf'):
                df_list = convert_pdf_to_df(file)
                if len(df_list) > 0:
                    print("1111")
                    result = extract_required_data_from_df(df_list, field_list, logger)
                    result.update({'source_file_name': file.split('/')[-1], 'DocNum_Inv': doc_num})
                    print(result)
                    insert_data_to_table_single(result, connection, master_table, logger, column_mapping)
                # else:
                content = convert_pdf_data_to_text(file)
                print("****")
                print(content)
                result = extract_required_data_from_text(content, field_list, logger)
                result.update({'source_file_name': file.split('/')[-1], 'DocNum_Inv': doc_num})
                print(result)
                insert_data_to_table_single(result, connection, master_table, logger, column_mapping)
            elif file.endswith('.docx'):
                content = docx2txt.process(file)
                result = extract_required_data_from_text(content, field_list, logger)
                result.update({'source_file_name': file.split('/')[-1], 'DocNum_Inv': doc_num})
                print(result)
                insert_data_to_table_single(result, connection, master_table, logger, column_mapping)
            elif file.endswith('.xlsx') or file.endswith('.xls'):
                results = extract_required_data_from_excel(file, field_list, doc_num)
                insert_invoice_data_batch(results, connection, logger, master_table)
            # master_data_update(doc_num, connection)
        logger.removeHandler(handler)
    connection.close()
