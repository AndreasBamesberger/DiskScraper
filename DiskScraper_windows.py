""" Tool to look up metadata of every file in the specified directory and its
sub-directories and store data in a csv file. If the config file provides a
path to a file with user specified categories then the program will only read
those metadata categories. """

import csv  # To create output file
import logging  # For error logging to file
import os  # To walk through directories
from datetime import datetime  # For timestamps
from pathlib import Path  # To transform filepath into something usable

import win32com.client  # To read metadata


class DiscScraperWin:
    """
        A class to extract metadata from files in a given directory.

        ...

        Attributes
        ----------
        _error_file: str
            name of the error log file
        _logger: Logger __main__ (DEBUG)
            the logging functionality
        _existing_files: list
            list of paths of all files read in previously (extracted from
            previous output csv file)
        _sh: win32com. ... .IShellDispatch6

        _configs: dict
            dictionary holding all user input from config file (source and
            output directory and path of file listing metadata categories)
        _source_dir: str
            path of directory that should be walked through
        _output_dir: str
            path of directory where output csv file should be created
        _output_file: str
            path name of output csv file
        _category_indices: list
            list holding the indices of the user specified metadata categories
        _categories: list
            list holding metadata category names (either every category or
            only user specified categories)

        Methods
        -------
        _read_config(config_path):
            Read user input from config file.
        _read_categories_file():
            Read list of user defined metadata categories.
        _create_file_name():
            Use path of source directory to create the name of the output csv
            file.
        _setup_csv_file():
            Create the header row of the csv file.
        read_files():
            Read every file in _source_dir and call _read_meta_data on it.
        _read_meta_data(win_path):
            Read in metadata categories of file win_path and save contents
            using _save_single_entry.
        _save_single_entry(win_path, meta_data_dict):
            Write metadata of the file win_path into the output csv file.
        _save_failed(win_path, error):
            If reading the file win_path fails, write the error message as a
            row in the output file.
        _read_existing():
            Read an existing output csv file and store the previously read in
            entries in _existing_files.
        _get_all_categories():
            Use the just created "error.log" file to read all metadata
            categories.
    """
    def __init__(self, config_path):
        self._error_file = 'error.log'
        logging.basicConfig(filename=self._error_file,
                            level=logging.DEBUG,
                            format='%(asctime)s %(levelname)s %(name)s %('
                                   'message)s')
        self._logger = logging.getLogger(__name__)

        self._existing_files = list()

        # Set up metadata reader
        # TODO: maybe also close shell again
        self._sh = win32com.client.gencache.EnsureDispatch('Shell.Application',
                                                           0)
        self._configs = self._read_config(config_path)
        self._source_dir = self._configs['crawl directory']
        self._output_dir = self._configs['output directory']
        self._output_file = self._create_file_name()

        # Check if output file already exists. If it does, read all entries
        # from this file
        if os.path.isfile(self._output_file):
            print("self._output_file already exists")
            self._read_existing()

        # If there are user specified metadata categories, read them from
        # the file and store them in self._categories
        if self._configs['pre-configured categories']:
            self._category_indices = list()
            self._categories = self._read_categories_file()

        # If not, use the just created 'error.log' file to get all metadata
        # category names
        else:
            self._categories = list()
            self._get_all_categories()
            self._category_indices = None

        # Add these 2 custom columns to the output
        self._categories.insert(0, "Timestamp of reading")
        self._categories.insert(0, "Filepath read by program")
        # Write the header of the output csv file
        self._setup_csv_file()

        print("Configuration read in from config file: ")
        print(repr(self._configs))

        print("Source directory: ")
        print(repr(self._source_dir))

        print("Output directory: ")
        print(repr(self._output_dir))

        print("Name of the output file: ")
        print(repr(self._output_file))

        print("Metadata categories to read in: ")
        print(repr(self._categories))

        print("Number of read in entries from existing output file: ")
        print(repr(len(self._existing_files)))

        print("---")

    @staticmethod
    def _read_config(config_path):
        """
        Search config file for source and output directory paths as well as
        the path to the file holding the user specified metadata categories.
        If no config file exists, use the current working directory.

        Parameters:
            config_path (str): File path to config file

        Returns:
            out_dict (dict): Holds all config file entries
        """
        out_dict = dict()
        try:
            with open(config_path, "r", encoding="utf-8") as config_file:
                for line in config_file.readlines():
                    if line.startswith('#') or line == '\n':
                        continue
                    split = line.split(' = ')
                    split[-1] = split[-1].strip()

                    if split[0] == "crawl directory":
                        if split[-1] == "here":
                            split[-1] = os.getcwd()

                        # remove trailing '/'
                        if split[-1][-1] == '/':
                            split[-1] = split[-1][:-1]
                        out_dict.update({split[0]: split[-1]})

                    elif split[0] == "output directory":
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
        """
        Read file containing user defined metadata categories and store them
        in list. Lines starting with '#' are ignored.

        Returns:
            output_list (list): holds all read in metadata categories
        """

        output_list = list()
        try:
            with open(self._configs['pre-configured categories'],
                      'r') as categories_file:
                for line in categories_file.readlines():
                    if not line.startswith('#'):
                        split = line.split(':')
                        split[0] = split[0].strip()
                        split[1] = split[1].strip()
                        output_list.append(split[1])
                        self._category_indices.append(split[0])

            return output_list

        except FileNotFoundError:
            print("Specified categories file not found")
            raise SystemError

    def _create_file_name(self):
        """
        Edit the name of the source directory to use it as the name of
        the output file.

        Returns:
            path (str): Full path and filename of the output file
        """
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
        """
        Create the header row of the csv output file. The first 2 columns
        are custom, all others are metadata category names.
        """

        if not os.path.isfile(self._output_file):
            with open(self._output_file, 'w', encoding='utf-16',
                      newline='') as csv_file:
                file_writer = csv.writer(csv_file, delimiter=';', quotechar='|',
                                         quoting=csv.QUOTE_MINIMAL)
                file_writer.writerow(self._categories)

    def read_files(self):
        """
        Reads every file in self._source_dir and calls self._read_meta_data
        on it.
        """

        counter = 0

        for root, _, files in os.walk(self._source_dir):
            for file_name in files:
                counter += 1

                file_path = os.path.join(root, file_name)
                print("{}: {}".format(str(counter), file_path))

                # Skip if the file has been read before
                if self._existing_files:
                    if file_path in self._existing_files:
                        print("File already read")
                        continue

                self._read_meta_data(Path(file_path))

    def _read_meta_data(self, win_path):
        """
        Read the metadata of the given file, either every metadata category
        or just the ones specified in the config file. After every read file,
        save the metadata as a row in the output csv file.

        Parameters:
            win_path (class 'pathlib.WindowsPath'): Path to single file
        """
        if self._configs['pre-configured categories']:
            iter_range = self._category_indices
        else:
            # 330 is arbitrary, metadata normally goes up to 320
            iter_range = range(330)

        try:
            ns = self._sh.NameSpace(str(win_path.parent))
            item = ns.ParseName(str(win_path.name))

            meta_data_dict = dict()

            for category_num in iter_range:
                category_name = str(ns.GetDetailsOf(None, category_num))
                category_value = str(ns.GetDetailsOf(item, category_num))

                if category_name != '' and category_value != '':
                    temp_dict = {category_name: category_value}
                    meta_data_dict.update(temp_dict)

        # Have to use broad Exception because I can't find how to handle
        # "pywintypes.com_error"
        except Exception as err:
            print('{}: {}'.format(err, win_path))
            self._logger.error('{}, filepath={}'.format(err, win_path))
        else:
            self._save_single_entry(win_path, meta_data_dict)

    def _save_single_entry(self, win_path, meta_data_dict):
        """
        Write metadata of single file into output csv file.

        Parameters:
            win_path (class 'pathlib.WindowsPath'): Path to single file
            meta_data_dict (dict): Holds all metadata of single file
        """

        with open(self._output_file, 'a', encoding='utf-16',
                  newline='') as csv_file:
            file_writer = csv.writer(csv_file, delimiter=';', quotechar='|',
                                     quoting=csv.QUOTE_MINIMAL)
            row = self._categories.copy()

            for index, value in enumerate(row):
                try:
                    row[index] = meta_data_dict[value]
                except KeyError:
                    row[index] = ''

            # add timestamp as second column
            row[1] = str(datetime.now())
            # add read in file path as first column
            row[0] = win_path

            file_writer.writerow(row)

    def _save_failed(self, win_path, error):
        """
        This method is not used currently.
        If reading a file fails, write error as a new row in output csv file.

        Parameters
            win_path (class 'pathlib.WindowsPath'): Path to single file
            error (str): Error message
        """
        output = list()
        output.append(win_path)
        output.append(error)
        with open(self._output_file, 'a', encoding='utf-16',
                  newline='') as csv_file:
            file_writer = csv.writer(csv_file, delimiter=';', quotechar='|',
                                     quoting=csv.QUOTE_MINIMAL)
            file_writer.writerow(output)

    def _read_existing(self):
        """
        Read the entries of a previously generated csv file. The first column
        of every entry holds the path to the entry which is saved in
        self._existing_files.
        """

        print("Reading file entries of previously created csv file...")
        with open(self._output_file, 'r', encoding='utf-16') as previous_file:
            # Dismiss the csv header row
            next(previous_file)

            for line in previous_file:
                line_split = line.split(';')
                # First column of csv holds the file path
                self._existing_files.append(line_split[0])

    def _get_all_categories(self):
        """
        Read all metadata categories from the just created 'error.log' file.
        """
        test_path = os.path.join(os.getcwd(), self._error_file)
        win_path = Path(test_path)
        ns = self._sh.NameSpace(str(win_path.parent))
        item = ns.ParseName(str(win_path.name))

        for category_num in range(350):
            category_name = str(ns.GetDetailsOf(None, category_num))
            category_value = str(ns.GetDetailsOf(item, category_num))

            if category_name != '' or category_value != '':
                self._categories.append(category_name)


if __name__ == '__main__':
    scraper = DiscScraperWin("config.txt")
    scraper.read_files()
