from common_functions import *
import itertools


def get_non_processed_records(connection, payment_data_table, payment_master_table, invoice_data_table, invoice_master_table):
    query_invoice = "select distinct Invoice_Number from {0},{1} where {1}.DocNum_Inv={0}.DocNum_Inv and {1}.Processed='N' and Invoice_Number iS not NULL"
    cursor = connection.cursor()
    cursor.execute(query_invoice.format(invoice_data_table, invoice_master_table))
    docnum_inv = list(itertools.chain.from_iterable(cursor))
    print(docnum_inv)

    query_payment = "select distinct description, DocNum_Bank from {0},{1} where {1}.DocNum_Bank={0}.docnum_pymt and {1}.Processed='N' and description iS not NULL"
    cursor = connection.cursor()
    cursor.execute(query_payment.format(payment_data_table, payment_master_table))
    desc_pymt = list(itertools.chain.from_iterable(cursor))
    print(desc_pymt)
    cursor.close()
    return docnum_inv, desc_pymt


def get_matching_invoice_num(docnum_inv, desc_pymt):
    # desc_pymt = ['Ref No. 45EF- INV/35464', 5, 'Ref No. 29-WEE INV/50487578', 9]
    matching_inv = []
    matching_pymt = []
    for ele in docnum_inv:
        idx = 0
        for ele1 in desc_pymt:
            idx = idx + 1
            if ele in str(ele1):
                matching_inv.append(ele)
                matching_pymt.append(desc_pymt[idx])
    return matching_inv, matching_pymt


def store_matching_recs(matching_inv, matching_pymt, payment_data_table, invoice_data_table,invoice_master_table, payment_master_table, connection, matched_data_table):
    for i in range(len(matching_inv)):
        docnum_pymt = matching_pymt[i]
        docnum_inv = matching_inv[i]
        print("**********")
        print(docnum_inv, docnum_pymt)
        query_inv =  "select {0}.DocNum_Inv, Invoice_Number	, Total_Amount, Invoice_Date from {0},{2} where Invoice_Number = '{1}' and {2}.Processed='N' and {0}.DocNum_Inv={2}.DocNum_Inv "
        cursor = connection.cursor()

        # , docnum_pymt, date_of_transaction, amount, description
        # from
        # and docnum_pymt = {3}
        cursor.execute(query_inv.format(invoice_data_table, docnum_inv, invoice_master_table))
        print(query_inv.format(invoice_data_table, docnum_inv, invoice_master_table))
        inv_data_list = []
        for matching_data in cursor.fetchall():
            inv_data_list.append(matching_data)

        query_inv = "select docnum_pymt, date_of_transaction, amount, description from {0},{2} where docnum_pymt = {1} and {2}.Processed='N' and {0}.docnum_pymt={2}.DocNum_Bank and description is not null"
        cursor = connection.cursor()
        cursor.execute(query_inv.format(payment_data_table, docnum_pymt, payment_master_table))
        print(query_inv.format(payment_data_table, docnum_pymt, payment_master_table))
        pymt_data_list = []
        for matching_data in cursor.fetchall():
            pymt_data_list.append(matching_data)
        inv_data_list[0] = list(inv_data_list[0])
        inv_data_list[0][3] = inv_data_list[0][3].strftime('%Y-%m-%d')
        inv_data_list[0] = tuple(inv_data_list[0])
        print(inv_data_list)
        print(pymt_data_list)
        matched_data = inv_data_list[0]+pymt_data_list[0]
        query_insert = """insert into {0}(DocNum_Inv,Invoice_Number,Invoice_Amount,Invoice_Date,	DocNum_Bank	,Bank_Payment_Date,Bank_Payment_Amount,Bank_Payment_Description)
values{1}
"""
        print(query_insert.format(matched_data_table, matched_data))
        cursor.execute(query_insert.format(matched_data_table, matched_data))
    cursor.close()

if __name__=='__main__':
    conf_handler = get_handler()
    payment_data_table, payment_master_table = get_payment_tables(conf_handler)
    invoice_data_table, invoice_master_table = get_invoice_tables(conf_handler)
    connection = connect_to_db(conf_handler)
    docnum_inv, desc_pymt = get_non_processed_records(connection, payment_data_table, payment_master_table, invoice_data_table,
                              invoice_master_table)
    matching_inv, matching_pymt = get_matching_invoice_num(docnum_inv, desc_pymt)
    print(matching_inv, matching_pymt)
    matched_data_table = get_matched_data_table(conf_handler)
    store_matching_recs(matching_inv, matching_pymt, payment_data_table, invoice_data_table, invoice_master_table, payment_master_table, connection, matched_data_table)

