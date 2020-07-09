"""step through whole filetree, list every file in csv"""
#TODO: Size changes between bytes, KB, MB. write into separate column

import os
import csv
import win32com.client # to read meta data
from pathlib import Path # to get usable path

def read_all_files(dir_src, categories, custom_columns): #pylint: disable=too-many-locals
    """walk entire filetree of given path, for every file create a dictionary
    with meta data and return list of created dictionaries """

    file_list = [] # every entry is one dictionary holding the meta data of one file

    if custom_columns:
        category_dict = categories
    else:
        category_dict = {}

    counter = 0 # counts files

    for root, _, files in os.walk(dir_src):
        for filename in files:

            counter += 1

            filepath = os.path.join(root, filename)
            print("{0}: {1}".format(str(counter), filepath))
            filepath = Path(filepath)

            ns = sh.NameSpace(str(filepath.parent))
            item = ns.ParseName(str(filepath.name))

            file_dict = {} # holds meta data of a single file

            if custom_columns:
                for key, field in categories.items():
                    colname = field
                    attr_val = str(ns.GetDetailsOf(item, int(key))) # the meta data of the specific file, e.g. "test.txt"
                    new_file_entry = {str(key): [str(colname), str(attr_val)]}
                    file_dict.update(new_file_entry)
                file_list.append(file_dict)

            else:
                colnum = 0
                while colnum <= 350: # value arbitrary, last meta column should be 320
                    colname = str(ns.GetDetailsOf(None, colnum)) # meta data category, e.g. "Name"
                    attr_val = str(ns.GetDetailsOf(item, colnum)) # the meta data of the specific file, e.g. "test.txt"
                    if attr_val != '' and colname != '':
                        new_file_entry = {str(colnum): [str(colname), str(attr_val)]}
                        file_dict.update(new_file_entry)

                        if str(colnum) not in category_dict:
                            new_category_entry = {str(colnum) : str(colname)} 
                            category_dict.update(new_category_entry)
                    colnum += 1
                file_list.append(file_dict) 

    return file_list, category_dict

def create_file_name(source, dest):
    """directory that is crawled is used as output name"""
    source = source.replace(' ', '') # remove whitespace
    source = source.replace('-', '') # remove existing -
    source = source.replace(':/', '-') 
    source = source.replace(':', '-')
    source = source.replace('/', '-')
    source = source.replace('\\', '-') # transform G\\ into G--
    source = source.replace('--', '-') # transform G-- into G-
    return dest + source + '.csv'

def read_config(configfile):
    """from config file read:
    1) directory to crawl
    2) where to save csv output
    3) if pre-configured categories should be used
    returns crawl directory, output directory, bool value"""
    source = ''
    with open(configfile, 'r', encoding="utf-8") as config:
        for line in config.readlines():
            if line[0] == '#':
                continue
            elif "crawl directory" in line:
                # remove first part of the line and '\n' at the end
                source = line.replace("crawl directory = ", '')
                if source[-1] == '\n':
                    source = source[:-1]
                if source == "here":
                    source = os.getcwd()

                # delete trailing '/'
                if source[-1] == '/':
                    source = source[:-1]

            elif "output directory" in line:
                dest = line.split()[-1]

                if dest == "here":
                    dest = os.getcwd()

                # add trailing '/'
                if dest[-1] != '/':
                    dest += '/'

            elif "pre-configured categories" in line:
                if line.split()[-1].lower() in ['false', 'no', 'n']:
                    custom_columns = False
                else:
                    # remove first part of the line and '\n' at the end
                    custom_columns = line.replace("pre-configured categories = ", '')
                    if custom_columns[-1] == '\n':
                        custom_columns = custom_columns[:-1]

    return source, dest, custom_columns

def save_to_csv(output_name, file_list, category_dict):
    """writes found files into csv file, 1 row per file, every metadata category in separate column"""

    category_codes = [] # this list determines the order in which the category columns will be printed
    for key, field in category_dict.items():
        if int(key) not in category_codes:
            category_codes.append(int(key))

    category_codes.sort() # start with lowest code

    # create first line of csv file with category names
    first_line = []
    for i in category_codes:
        if str(i) in category_dict:
            first_line.append(category_dict[str(i)])

#    encoding = "utf-8"
#    encoding = "cp1252"
    encoding = "utf-16"

    with open(output_name, "w", encoding=encoding) as csv_file:
        filewriter = csv.writer(csv_file, delimiter=';', quotechar='|', quoting=csv.QUOTE_MINIMAL, lineterminator = '\n')
        filewriter.writerow(first_line)

        for file in file_list:
            writeout = []

            for i in category_codes:
                if str(i) in file:
                    writeout.append(file[str(i)][1])
                else:
                    writeout.append('')

            filewriter.writerow(writeout)

def read_needed_categories(categoryfile):
    print("categoryfile = {0}".format(repr(categoryfile)))
    output_dict ={} 
    with open(categoryfile, "r") as category:
        for line in category.readlines():
            if line[0] != '#':
                split = line.split()
                new_entry = {split[0]: split[2]}
                output_dict.update(new_entry)

    return output_dict


if __name__ == '__main__':
    sh=win32com.client.gencache.EnsureDispatch('Shell.Application',0) # set up meta data reader

    # get directory to crawl through and directory to save output to
    configfile = "config.txt"

    if not os.path.exists(configfile):
        source = os.getcwd()
        dest = os.getcwd()
        custom_columns = False
    else:
        source, dest, custom_columns = read_config(configfile)

    output_name = create_file_name(source, dest)
    print("path to crawl: {0}".format(repr(source)))
    print("output file: {0}".format(repr(output_name)))

    if custom_columns:
        needed_categories = read_needed_categories(custom_columns)

        print("metadata categories to store in csv: ")
        for key, field in needed_categories.items():
            print("{0} - {1}".format(str(key), str(field)))
    else: 
        needed_categories = None

    file_list, category_dict = read_all_files(source, needed_categories, custom_columns)

    print("files found: {0}".format(str(len(file_list))))

    if custom_columns:
        save_to_csv(output_name, file_list, needed_categories)
    else:
        save_to_csv(output_name, file_list, category_dict)