import argparse
import getopt
import logging
import sys
from pathlib import Path

import pandas as pd
import numpy as np


# pip install odfpy
# pd.read_excel("the_document.ods", engine="odf")

logger = logging.getLogger()
logger.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s',
                              '%d-%m-%Y %H:%M:%S')

stdout_handler = logging.StreamHandler(sys.stdout)
stdout_handler.setLevel(logging.DEBUG)
stdout_handler.setFormatter(formatter)

file_handler = logging.FileHandler('logs.log')
file_handler.setLevel(logging.DEBUG)
file_handler.setFormatter(formatter)

logger.addHandler(file_handler)
logger.addHandler(stdout_handler)

def get_sheets(excel_file_1, excel_file_2):
    excel_1_sheets = pd.ExcelFile(excel_file_1, engine="odf").sheet_names.sort()
    excel_2_sheets = pd.ExcelFile(excel_file_2, engine="odf").sheet_names.sort()
    logging.debug('test3')
    return excel_1_sheets == excel_2_sheets, excel_1_sheets


def compare_excel_files(excel_1, excel_2, out_dir):
    diff_annot = f'COMPARISON OF\n\t{excel_1}\n\tWITH\n\t{excel_2}' \
                 f'\n\nDifferences:\n'
    equal_sheets, sheets = get_sheets(excel_1, excel_2)
    # print(exl_1.compare.exl_2)
    if equal_sheets:
        for i in range(len(sheets)):
            exl_1 = pd.read_excel(excel_1, sheets[i])
            exl_2 = pd.read_excel(excel_2, sheets[i])
            exl_1.equals(exl_2)
            match_map = exl_1.values == exl_2.values
            print(match_map)
            diff_rows, dif_cols = np.where(not match_map)
            for item in zip(diff_rows, dif_cols):
                col_position = exl_1.columns.values.tolist()[dif_cols]
                if not (pd.isnull(exl_1.iloc[item[0], item[1]]) and
                        pd.isnull(exl_1.iloc[item[0], item[1]])):
                    exl_1_val = exl_1.iloc[item[0], item[1]]
                    exl_2_val = exl_2.iloc[item[0], item[1]]
                    diff_annot += f'\t\tIn {sheets[i]} sheet ' \
                                  f'[row: {diff_rows + 1}, ' \
                                  f'col: ' \
                                  f'{col_position}]:' \
                                  f' {exl_1_val} >>>' \
                                  f' {exl_2_val}\n'
                    exl_1.iloc[item[0], item[1]] = f'{exl_1_val} ' \
                                                   f'>>> {exl_2_val}'
            out_path = f'{out_dir}/excel_diff/'
            res_exl_file = f'ExcelDiff_sheet_{sheets[i]}.xlsx'
            annot_file = 'ExcelDiff_annotations.txt'
            exl_1.to_excel(out_path + res_exl_file, index=False,
                           header=True, encoding='utf-8')
            with open(out_path + annot_file, 'w') as f:
                f.write(diff_annot)

def main():
    parser = argparse.ArgumentParser(description="A tool to compare two excel files "
                                                 "with annotation of the differences.")
    parser.add_argument("-i", "--input-files",
                        type=list,
                        nargs=2,
                        help="Two paths to the Excel files to "
                             "be compared with each other.",
                        required=True)
    parser.add_argument("-o", "--out-dir", type=str,
                        nargs=1,
                        default='',
                        help="Path to the output directory.", required=True)

    parser.add_argument("-v", "--verbose", action='store_true',
                        help="Increase output verbosity", required=False)
    parser.parse_args()
    argument_list = sys.argv[1:]
    input_files, out_dir = None, None

    if "-h" in sys.argv[1:] or "--help" in sys.argv[1:]:
        parser.print_help()
        exit(0)
    else:
        # Options
        options = "i:o:v"

        # Long options
        long_options = [
            "input-files=",
            "out-dir=",
            "verbose"
        ]

        try:
            # Parsing argument
            arguments, values = getopt.getopt(argument_list, options, long_options)

            # checking each argument
            for currentArgument, currentValue in arguments:
                if currentArgument in ("-i", "--input-files"):
                    if len(currentValue) == 2:
                        path_1 = Path(currentValue[0])
                        path_2 = Path(currentValue[1])
                        if path_1.is_file() and path_2.is_file():
                            input_files = [path_1, path_2]

                elif currentArgument in ("-o", "--out-dir"):
                    if Path(currentValue).is_dir():
                        out_dir = currentValue

                elif currentArgument in ("-v", "--verbose"):
                    # logging.basicConfig(level=logging.DEBUG)
                    logging.getLogger().setLevel(logging.DEBUG)

        except getopt.error as err:
            # output error, and return with an error code
            logging.exception(err)

        logging.debug("hello world3!")
        logger.debug('test')

        if input_files and out_dir:
            compare_excel_files(input_files[0], input_files[1], out_dir)


if __name__ == "__main__":
    logging.getLogger().setLevel(logging.WARNING)
    try:
        main()
    except Exception as e:
        logging.exception(e)
