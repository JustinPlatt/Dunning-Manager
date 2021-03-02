# -*- coding: utf-8 -*-
"""
Created on Thu Feb 11 16:54:40 2021

@author: Justin Platt
"""

import glob
import os
from pathlib import Path
from pdfreader import PDFDocument, SimplePDFViewer
import pandas as pd
import PyPDF2
import re
from re import sub
import sys
import warnings

warnings.filterwarnings("ignore", category=FutureWarning)

#  "globals"
ord_regex = re.compile(r"""
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

ON_NETWORK = False

if ON_NETWORK:
    ERT_PATH = r'\\fs1.bgeltd.com\Proc\ERT_Reports\PR\ECS\PDF\\'
else:
    ERT_PATH = ''
DATA_PATH = os.getcwd()


def get_file_list():
    """Create a list of all unique files in data.csv

    Returns
    -------
    file_list : list
        Contains all unique values of the FILE column.
        Returns [] if data.csv doesn't exist
    """
    if os.path.isfile('./data.csv'):  # leading dot indicates relative import
        cur_data = pd.read_csv('data.csv', sep='|', dtype={'ORDER_ID': str})
        file_list = cur_data['FILE'].unique()
    else:
        file_list = []
    return file_list


def get_file():
    """Create a numbered list of all pdf files in the folder,
    then prompt user to pick one.

    This should probably find all the pdfs that match some regex
    and aren't already in file_list.  Or choose to import one/all

    Returns
    -------
    pdf_to_open : string
        Returns the file name including the .pdf extension

    """
    print('Showing PDF files in ' + os.getcwd())
    files = []
    f_count = 0
    for file in os.listdir(ERT_PATH):
        if file.endswith(".pdf") or file.endswith(".PDF"):
            f_name = file.rsplit('.', maxsplit=1)[0]
            print('(' + str(f_count) + ')  ' + file)
            files.append(f_name)  # take the file name w/o the .pdf
            f_count += 1
    prompt = '\nEnter the number corresponding to the target pdf, or q to quit: '
    choice = input(prompt)
    while not is_valid(choice, f_count):
        if choice == 'q':
            print('Quitting.')
            sys.exit()
        else:
            print('Invalid choice - try again.')
            choice = input(prompt)
    if choice != 'q':
        pdf_to_open = str(files[int(choice)]) + '.pdf'
        return pdf_to_open


def is_valid(choice, file_ct):
    """Check if the report choice is a valid one"""
    x = False
    if choice.isnumeric():
        if int(choice) in range(file_ct):
            x = True
    return x


def get_choice():
    menu_prompt = 'Choose one: [I]mport pdf / [O]rder search / [Q]uit:  '
    menu_choice = input(menu_prompt).lower()
    while menu_choice not in ('i', 'o', 'q'):
        print('\nInvalid choice!')
        menu_choice = input(menu_prompt).lower()
    return menu_choice


def get_order():
    ord_regex = re.compile(r'\d{15}|[bB]')  # find a 15 digit number, 'b' or 'B'
    ord_prompt = '\nEnter a 15 digit order id, or [B] to go back:  '
    ord_choice = input(ord_prompt)
    while not ord_regex.search(ord_choice):
        print('\nInvalid choice!')
        ord_choice = input(ord_prompt)
        if ord_choice in ('b', 'B'):
            break
    return str(ord_choice)


def find_order():
    if os.path.isfile('./data.csv'):  # placeholder for now
        ord_num = get_order()
        if ord_num in ('b', 'B'):
            print('Returning to main menu...')
        else:
            df = pd.read_csv('data.csv', sep='|', dtype={'ORDER_ID': str})
            matches = df[df['ORDER_ID'] == ord_num]
            match_ct = matches['DUNNING_NUM'].count()
            if match_ct == 0:
                print('No matches found.')
            elif match_ct == 1:
                file = matches['FILE'].values[0]
                start_page = matches['PAGE'].values[0]
                print('1 match found:')
                print(matches.to_string(index=False))
                print('Printing dunning invoice.')
                print_order(file, start_page, ord_num,1)
            else:
                print(str(match_ct) + ' matches found:')
                print(matches.to_string(index=False))
                for a in range(match_ct):
                    print('Printing match ' + str(a+1) + ' of ' + str(match_ct))
                    file = matches['FILE'].values[a]
                    start_page = matches['PAGE'].values[a]
                    print_order(file, start_page, ord_num, a+1)
    else:
        print('No data.csv file found.  Import some data.')


def print_order(file, start_page, order_id, num):
    full_file = ERT_PATH + 'BGE_DUNNING_' + file +'.DP.pdf'
    d_pdf = open(full_file, 'rb')
    pdf_reader = PyPDF2.PdfFileReader(d_pdf, strict=False, warndest=None)
    pdf_writer = PyPDF2.PdfFileWriter()
    for page in range(start_page-1, start_page+1):
        pdf_writer.addPage(pdf_reader.getPage(page))
    output_filename = 'DUNNING-{}-INV_{}.pdf'.format(order_id, str(num))
    with open(output_filename, 'wb') as out:
        pdf_writer.write(out)
    d_pdf.close()
    print('SAVED AS ' + output_filename)


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
        if ord_regex.search(page_txt):
            gps = ord_regex.search(page_txt).groups()
            new_row = [gps[5]+gps[6], gps[0], gps[4], gps[1], gps[2],
                       sub(r'[^\d.]', '', gps[3]), fname, page_num+1]
            data_tmp.append(new_row)
        else:
            print('check page ' + str(page_num))
    d_pdf.close()
    data_cols = ['ORDER_ID', 'PROD10', 'CUST_ID', 'DUNNING_NUM', 'ITEM_NAME', 'TOTAL_DUE', 'FILE', 'PAGE']
    data_inv = pd.DataFrame(data=data_tmp, columns=data_cols)
    data_inv['TOTAL_DUE'] = data_inv['TOTAL_DUE'].astype(float)
    if os.path.isfile('./data.csv'):  #look for data.csv - the . is for current dir
        print('Appending to data.csv')
        prior_data = pd.read_csv('data.csv', sep='|', dtype={'ORDER_ID': str})
        data_inv = prior_data.append(data_inv)
    else:
        print('No data.csv file found.  Creating now.')
    data_inv.to_csv('data.csv', index=False, sep='|')
    print('Done.')


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
            print('This is where we will import everything')
            for file_name in to_import:
                print('Importing '+file_name)
                import_pdf(file_name)
            print('Finished.')


def main():
    print('DUNNING INVOICE MANAGER - LAST UPDATED 2/28/2021')
    while True:
        choice = get_choice()
        if choice == 'i':               # import pdfs
            import_check()
            #import_pdfs(file_list)
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
