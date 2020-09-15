""" tool to look up metadata of every file in specified directory and store data
in csv file. if config file provides path to file with user specified categories
then program will only store those metadata categories, otherwise a new column
is created for every category entry found """

import csv
import os
from pathlib import Path  # to transform filepath into something usable
import win32com.client  # to read meta data
import logging  # for error logging to file


class DiscScraperWin:
    def __init__(self, config_path):
        logging.basicConfig(filename='error.log',
                            level=logging.DEBUG,
                            format='%(asctime)s %(levelname)s %(name)s %(message)s')
        self._logger = logging.getLogger(__name__)

        self._sh = None
        self._configs = self._read_config(config_path)

        self._source_dir = self._configs['crawl directory']
        self._output_dir = self._configs['output directory']
        self._output_file = self._create_file_name()
        if self._configs['pre-configured categories']:
            self._categories = self._read_categories_file()
            self._setup_csv_file()
        else:
            self._categories = None

        print("self._configs: ")
        print(repr(self._configs))

        print("self._source_dir: ")
        print(repr(self._source_dir))

        print("self._output_dir: ")
        print(repr(self._output_dir))

        print("self._output_file: ")
        print(repr(self._output_file))

        print("self._categories: ")
        print(repr(self._categories))

    @staticmethod
    def _read_config(config_path):
        """
        read input, output directories and specified categories from file. if no
        config file exists, use the current working directory.
        input: config_path:str, filepath to config file
        output: out_dict:dict, holds all config file entries
        """
        out_dict = {}
        try:
            with open(config_path, "r", encoding="utf-8") as config_file:
                for line in config_file.readlines():
                    if line.startswith('#') or line == '\n':
                        continue
                    split = line.split(': ')
                    split[-1] = split[-1].rstrip()

                    if split[0] == "crawl directory":
                        # use current working directory
                        if split[-1] == "here":
                            split[-1] = os.getcwd()

                        # remove trailing '/'
                        if split[-1][-1] == '/':
                            split[-1] = split[-1][:-1]
                        out_dict.update({split[0]: split[-1]})

                    if split[0] == "output directory":
                        # use current working directory
                        if split[-1] == "here":
                            split[-1] = os.getcwd()

                        # add trailing '/'
                        if split[-1][-1] != '/':
                            split[-1] += '/'

                        out_dict.update({split[0]: split[-1]})
                    elif split[0] == "pre-configured categories":
                        if split[-1].lower() == "false":
                            out_dict.update({split[0]: None})
                        else:
                            out_dict.update({split[0]: split[-1]})
        except FileNotFoundError:
            cwd = os.getcwd()
            out_dict.update({"crawl directory": cwd})
            out_dict.update({"output directory": cwd})
            out_dict.update({"pre-configured categories": None})

        return out_dict

    def _read_categories_file(self):
        """ read file with categories, ignore lines starting with '#' and add
        all other lines as dictionary entries
        output: output_dict:dict, holds metadata category number and name,
        e.g. {0: 'name'} """
        output_list = list()
        try:
            with open(self._configs['pre-configured categories'], 'r') as categories_file:
                for line in categories_file.readlines():
                    if not line.startswith('#'):
                        split = line.split(':')
                        split[0] = split[0].strip()
                        split[1] = split[1].strip()
                        output_list.append(split[1])
            return output_list

        except FileNotFoundError:
            print("specified categories file not found")
            raise SystemError

    def _read_categories_file_old(self):
        """ read file with categories, ignore lines starting with '#' and add
        all other lines as dictionary entries
        output: output_dict:dict, holds metadata category number and name,
        e.g. {0: 'name'} """
        output_dict = dict()
        try:
            with open(self._configs['pre-configured categories'], 'r') as categories_file:
                for line in categories_file.readlines():
                    if not line.startswith('#'):
                        split = line.split('-')
                        split[0] = split[0].strip()
                        split[1] = split[1].strip()
                        output_dict.update({split[0]: split[1]})
            return output_dict

        except FileNotFoundError:
            print("specified categories file not found")
            raise SystemError

    def _create_file_name(self):
        """ directory that is crawled through is used as output name
        output: path:str, full path of output csv file"""
        source = self._source_dir
        dest = self._output_dir
        source = source.replace(' ', '')  # remove whitespace
        source = source.replace('-', '')  # remove existing -
        source = source.replace(':/', '-')
        source = source.replace(':', '-')
        source = source.replace('/', '-')
        source = source.replace('\\', '-')  # transform G\\ into G--
        source = source.replace('--', '-')  # transform G-- into G-
        path = os.path.join(dest, source) + '.csv'
        return path

    def _setup_csv_file(self):
        """ create the first row with all column names """

        if not os.path.isfile(self._output_file):
            with open(self._output_file, 'w', encoding='utf-16', newline='') as csv_file:
                file_writer = csv.writer(csv_file, delimiter=';', quotechar='|',
                                         quoting=csv.QUOTE_MINIMAL)
                file_writer.writerow(self._categories)

    def read_files(self):
        """ reads every file in self._source_dir and calls self._read_meta_data on it """

        # set up meta data reader
        self._sh = win32com.client.gencache.EnsureDispatch('Shell.Application', 0)

        counter = 0

        for root, _, files in os.walk(self._source_dir):
            for file_name in files:
                counter += 1

                file_path = os.path.join(root, file_name)
                print("{}: {}".format(str(counter), file_path))

                self._read_meta_data(Path(file_path))

    def _read_meta_data(self, win_path):
        """ read the meta data of the given file, either every meta data
        category or just the ones specified in the config file. after every read
        file, save the meta data as a row in the output csv file """

        try:
            ns = self._sh.NameSpace(str(win_path.parent))
            item = ns.ParseName(str(win_path.name))

            meta_data_dict = dict()

            for category_num in range(350):
                category_name = str(ns.GetDetailsOf(None, category_num))
                category_value = str(ns.GetDetailsOf(item, category_num))
                if category_name != '' and category_value != '':
                    # temp_dict = {str(category_num): [category_name, category_value]}
                    temp_dict = {category_name: category_value}
                    meta_data_dict.update(temp_dict)

            self._save_single_entry(meta_data_dict)

        except AttributeError as err:
            self._logger.error('{}: {}'.format(err, win_path))

    def _save_single_entry(self, meta_data_dict):
        with open(self._output_file, 'a', encoding='utf-16', newline='') as csv_file:
            file_writer = csv.writer(csv_file, delimiter=';', quotechar='|',
                                     quoting=csv.QUOTE_MINIMAL)
            row = self._categories.copy()
            for index, value in enumerate(row):
                try:
                    row[index] = meta_data_dict[value]
                except KeyError:
                    row[index] = ''
            file_writer.writerow(row)


if __name__ == '__main__':
    scraper = DiscScraperWin("config.txt")
    scraper.read_files()
