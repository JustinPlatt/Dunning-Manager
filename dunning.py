# -*- coding: utf-8 -*-
"""
Created on Thu Feb 11 16:54:40 2021

@author: Justin Platt

package versions used:

python - 3.8.3
    glob
    os
    re
    shutil
    sys
pandas - 1.0.5
pyPDF2 - 1.26.0

"""

import glob
import os
import pandas as pd
import PyPDF2
import re
from re import sub
import shutil
import sys

# globals
LAST_UPDATED = '3/22/2021'
ERT_PATH = ''  # updated in main()
NETWORK_PDF_PATH = r'//fs1.bgeltd.com/Proc/ERT_Reports/PR/ECS/PDF/'
LOCAL_PDF_PATH = './ert/'
DATA_PATH = './data/'
PDF_PATH = './pdfs/'
VOD_PATH = './vods/'
DATA_FILE = DATA_PATH + 'data.csv'


# regex patterns
invoice_regex = re.compile(r"""
    (\d{2}-\d{5}-\d{3})-        #group 0 - PROD10
    (\d{3})                     #group 1 - DUNNING_NUM
    .+"([^"]+)"                 #group 2 = ITEM_NAME
    .+Past[ ]Due                #the [ ] represents a blank space
    \$([\d,]+.\d{2})            #group 3 - TOTAL_DUE
    (\d{9})                     #group 4 - CUST_ID
    \s{2}\d\s(\d{12})           #group 5 - ORDER_ID (order #)
    \s{3}(\d{3})                #group 6 - ORDER_ID (line #)
    """, re.VERBOSE)

get_fname = re.compile(r'\d[\d_]+')  # pulls unique part of filename

order_regex = re.compile(r'^\d{15}$|^[bB]$')  # find 15 digit , 'b' or 'B'

menu_regex = re.compile(r'^[ioqIOQ]$')  # find i/o/q in upper/lower case


def get_ert_path():
    if os.path.isdir(NETWORK_PDF_PATH):
        print('Network PDF folder found...')
        return NETWORK_PDF_PATH
    elif os.path.isdir(LOCAL_PDF_PATH):
        print('Local PDF folder found...')
        return LOCAL_PDF_PATH
    else:
        print("Can't find local or network PDF folder.  Quitting.")
        print("Network folder path: " + NETWORK_PDF_PATH)
        print("Local folder path: " + LOCAL_PDF_PATH)
        sys.exit()


def get_file_list():
    """Create a list of all unique files in data.csv

    Returns
    -------
    file_list : list
        Contains all unique values of the FILE column.
        Returns [] if data.csv doesn't exist
    """
    if os.path.isfile(DATA_FILE):
        cur_data = pd.read_csv(DATA_FILE, sep=',', dtype={'ORDER_ID': str})
        file_list = cur_data['FILE'].unique()
    else:
        file_list = []
    return file_list


def get_menu_choice():
    """Get main menu choice from user

    Returns
    -------
    menu_input : str
        Returns 'i', 'o', or 'q'
        Loops until one of these is picked
    """
    menu_prompt = 'Choose one: [I]mport pdf / [O]rder search / [Q]uit:  '
    menu_input = input(menu_prompt).lower()
    while not menu_regex.search(menu_input):
        print('\nInvalid choice!')
        menu_input = input(menu_prompt).lower()
    return menu_input


def get_order():
    """Get an order number or return to main menu

    Returns
    -------
    order_input : str
        Returns 15 digit order id, 'b' or 'B'
        Loops until one of these is picked
    """
    order_prompt = '\nEnter a 15 digit order id, or [B] to go back:  '
    order_input = input(order_prompt)
    while not order_regex.search(order_input):
        print('\nInvalid choice!')
        order_input = input(order_prompt)
    return str(order_input)


def find_order():
    if os.path.isfile(DATA_FILE):
        ord_num = get_order()
        if ord_num in ('b', 'B'):
            print('Returning to main menu...')
        else:
            matches = pd.read_csv(
                DATA_FILE, sep=',',
                usecols=['ORDER_ID', 'DUNNING_NUM', 'FILE', 'PAGE'],
                dtype={'ORDER_ID': str})

            matches = matches[matches['ORDER_ID'] == ord_num]
            match_ct = len(matches)
            if match_ct == 0:
                print('No matches found.')
            elif match_ct == 1:
                file = matches['FILE'].values[0]
                start_page = matches['PAGE'].values[0]
                print('1 match found - printing invoice')
                print_order(file, start_page, ord_num)
            else:
                print(str(match_ct) + ' matches found - printing newest invoice')
                matches['FILE_DATE'] = matches['FILE'].str.extract(r'_(\d{8})_').astype(int)
                newest_match_data = list(matches.sort_values(by=['FILE_DATE', 'DUNNING_NUM'], ascending=False).iloc[0])
                file = newest_match_data[2]
                start_page = newest_match_data[3]
                ord_num = newest_match_data[0]
                # items = [order, dunning, file, page, file_date]
                print_order(file, start_page, ord_num)
    else:
        print('No data.csv file found.  Import some data.')


def print_order(file_name, start_page, order_id):
    """Saves an order invoice as a pdf

    Parameters
    ----------
    file_name : str
        The name of the pdf containing the invoice as stored in /data/data.csv
    start_page : int
        The first page of the invoice in the pdf
    order_id : str
        The 15 digit order+line number

    """
    full_file = PDF_PATH + 'BGE_DUNNING_' + file_name + '.DP.pdf'
    d_pdf = open(full_file, 'rb')
    pdf_reader = PyPDF2.PdfFileReader(d_pdf, strict=False, warndest=None)
    pdf_writer = PyPDF2.PdfFileWriter()
    for page in range(start_page-1, start_page+1):
        pdf_writer.addPage(pdf_reader.getPage(page))
    output_filename = 'DUNNING-ORDER_{}.pdf'.format(order_id)
    with open(VOD_PATH + output_filename, 'wb') as out:
        pdf_writer.write(out)
    d_pdf.close()
    print('SAVED AS ' + os.path.abspath(VOD_PATH + output_filename))


def import_pdf(pdf_to_open):
    fname = get_fname.search(pdf_to_open)
    if get_fname.search(pdf_to_open):
        fname = get_fname.search(pdf_to_open).group()
    else:
        fname = pdf_to_open.rsplit('.', maxsplit=1)[0]
    d_pdf = open(ERT_PATH + pdf_to_open, 'rb')  # open the file to be read by PyPDF2
    pdfReader = PyPDF2.PdfFileReader(d_pdf)  # use this to navigate the pdf
    page_ct = pdfReader.getNumPages()  # get number of pages
    print('Processing ' + str(page_ct) + ' pages...')
    data_tmp = []
    for page_num in range(2, page_ct, 2):  # loop through page 3, 5, 7, ... pg_ct
        if page_num % 1000 == 0:
            print('Processing page ' + str(page_num) + ' of ' + str(page_ct))
        page_txt = pdfReader.getPage(page_num).extractText()
        try:
            gps = invoice_regex.search(page_txt).groups()
            new_row = [gps[5]+gps[6], gps[0], gps[4], gps[1], gps[2],
                       sub(r'[^\d.]', '', gps[3]), fname, page_num+1]
            data_tmp.append(new_row)
        except (AttributeError):
            print('check page ' + str(page_num))
    d_pdf.close()
    shutil.copy(ERT_PATH + pdf_to_open, PDF_PATH)
    data_cols = ['ORDER_ID', 'PROD10', 'CUST_ID', 'DUNNING_NUM', 'ITEM_NAME', 'TOTAL_DUE', 'FILE', 'PAGE']
    data_inv = pd.DataFrame(data=data_tmp, columns=data_cols)
    data_inv['TOTAL_DUE'] = data_inv['TOTAL_DUE'].astype(float)
    if os.path.isfile(DATA_FILE):  # look for data.csv - the . is for current dir
        print('Appending to data.csv')
        prior_data = pd.read_csv(DATA_FILE, sep=',', dtype={'ORDER_ID': str})
        data_inv = prior_data.append(data_inv)
    else:
        print('No data.csv file found.  Creating now.')
    data_inv.to_csv(DATA_FILE, index=False)
    print('Done.  File copied to ' + os.path.abspath(PDF_PATH))


def import_check():
    # count number of dunning PDFs in folder
    # check which ones are already in data.csv
    # ask whether to pull in the new files if there are any
    print(('Searching for dunning invoice pdfs in {}').format(ERT_PATH))
    dunning_files = glob.glob(ERT_PATH + '*BGE_DUNNING*.pdf')
    dunning_count = len(dunning_files)
    print(('Found {!s} files').format(dunning_count))
    print('Comparing to data.csv ...')
    data_files = get_file_list()
    to_import = []
    for file in dunning_files:
        file_name = get_fname.search(file).group(0)  # 'data.csv'-styled name
        if file_name not in data_files:
            to_import.append(os.path.basename(file))  # add file name to list
    if to_import == []:  # if nothing to import
        print('No new files to import.  Returning to main menu.')
    else:
        print(('Found {!s} file(s) to import').format(len(to_import)))
        import_prompt = 'Import now? [Y]es / [N]o:  '
        import_choice = input(import_prompt).lower()
        while import_choice not in ('y', 'n'):
            print('\nInvalid choice!')
            import_choice = input(import_prompt).lower()
        if import_choice == 'n':
            print('Returning to main menu...')
        else:
            for file_name in to_import:
                print('Importing '+file_name)
                import_pdf(file_name)
            print('Finished.')


def main():
    print('DUNNING INVOICE MANAGER - LAST UPDATED ' + LAST_UPDATED)
    global ERT_PATH   # Needed to modify global copy of ERT_PATH
    ERT_PATH = get_ert_path()
    while True:
        choice = get_menu_choice()
        if choice == 'i':               # import pdfs
            import_check()
        elif choice == 'o':             # find an order
            find_order()
        elif choice == 'q':             # import pdfs
            print('Exiting.')
            break
        else:
            print('Unknown error occurred.  Exiting.')
            break


if __name__ == "__main__":
    main()
