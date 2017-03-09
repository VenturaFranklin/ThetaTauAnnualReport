'''
Created on Mar 1, 2017

@author: venturf2
from AnnualReport import combine_files
combine_files.run()
'''
import os
from shutil import copy, move
import glob
import csv
import datetime

MAIN_FOLDER = r"E:\ThetaTau Drive\Theta Tau\Nationals\AnnualReport\TOPROCESS"
INCOMING_FOLDER = r"E:\ThetaTau Drive\Theta Tau\Nationals\AnnualReport\SUBMISSIONS_ANNUAL_REPORT"

FILE_TYPES = ['INIT', 'DEPL', 'MSCR', 'COOP', 'OER']
DATE = datetime.date.isoformat(datetime.date.today()).replace('-', '')


def find_all_new_csvs():
    all_files = list(os.walk(MAIN_FOLDER))
    _, _, files_to_process = all_files[0]
    _, _, files_processed = all_files[1]
    old_filenames = files_to_process + files_processed
    for dirpath, _, filenames in os.walk(INCOMING_FOLDER):
        for filename in [f for f in filenames if f.endswith(".csv") and
                         f not in old_filenames]:
            new_file = os.path.join(dirpath, filename)
            new_file_loc = os.path.join(MAIN_FOLDER, filename)
            copy(new_file, new_file_loc)
            print("Copied file: ", filename)


def run():
    for file_type in FILE_TYPES:
        first_row_complete = False
        files = glob.glob(MAIN_FOLDER + '/*'+file_type+'*.csv')
        file_out_name = os.path.join(MAIN_FOLDER,
                                     'PROCESSED',
                                     DATE+'_'+file_type+'.csv')
        print("FILE OUT: ", file_out_name)
        with open(file_out_name, 'a+', newline='') as file_out:
            file_out_writer = csv.writer(file_out)
            for file_name in files:
                print("FILE IN: ", file_name)
                with open(file_name) as file_in:
                    mscr_row = 0
                    reader = csv.reader(file_in)
                    for i, row in enumerate(reader):
                        if i == 0 and first_row_complete:
                            continue
                        if file_type == 'MSCR':
                            if row[7] == 'Graduated from school':
                                if mscr_row == 0:
                                    phones = row[5].split(',')
                                    emails = row[6].split(',')
                                    degree = row[8].split(',')
                                row[5] = phones[mscr_row]
                                row[6] = emails[mscr_row]
                                row[8] = degree[mscr_row]
                                mscr_row += 1
                        row = [val if val.upper() not in
                               ['UNKNOWN', 'NONE', 'N/A']
                               else '' for val in row]
                        print(row)
                        file_out_writer.writerow(row)
                new_file_loc = os.path.join(MAIN_FOLDER,
                                            'PROCESSED',
                                            os.path.split(file_name)[1])
                print("MOVE FILE: ", new_file_loc)
                first_row_complete = True
                move(file_name, new_file_loc)

if __name__ == '__main__':
    find_all_new_csvs()
