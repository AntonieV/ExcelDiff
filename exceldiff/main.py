import argparse
import logging
import os
import sys
from logging.handlers import RotatingFileHandler
from pathlib import Path
import pandas as pd
import numpy as np

logger = logging.getLogger()
logger.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s',
                              '%d-%m-%Y %H:%M:%S')

stdout_handler = logging.StreamHandler(sys.stdout)
stdout_handler.setLevel(logging.DEBUG)
stdout_handler.setFormatter(formatter)

file_handler = RotatingFileHandler('logs.log', maxBytes=500 * 1024, backupCount=1)
file_handler.setLevel(logging.DEBUG)
file_handler.setFormatter(formatter)

logger.addHandler(file_handler)
logger.addHandler(stdout_handler)


def compare_excel_files(excel_1, excel_2, out_dir):
    logger.info('Starting ExcelDiff analysis...')
    diff_annot = f'COMPARISON OF\n\t{excel_1}\n\tWITH\n\t{excel_2}' \
                 f'\n\nDifferences:\n'
    exl_1 = pd.read_excel(excel_1, sheet_name=None)
    exl_2 = pd.read_excel(excel_2, sheet_name=None)
    out_path = f'{out_dir}/'
    if exl_1.keys() == exl_2.keys():
        sheets = list(exl_1.keys())
        res_exl_file = f'{out_path}ExcelDiff_{os.path.basename(excel_1)}_vs' \
                       f'_{os.path.basename(excel_2)}.xlsx'
        with pd.ExcelWriter(res_exl_file) as diff_writer:
            for idx in range(len(sheets)):
                logger.info(f"Analysing sheet '{sheets[idx]}'")
                diff_annot += f"\tSheet '{sheets[idx]}':\n"
                if not exl_1[sheets[idx]].equals(exl_2[sheets[idx]]):
                    match_map = exl_1[sheets[idx]] == exl_2[sheets[idx]]
                    diff_rows, dif_cols = np.where(match_map == False)
                    for cell in zip(diff_rows, dif_cols):
                        col_position = exl_1[sheets[idx]].columns.values[cell[1]]
                        row_idx_shift = 1 if len(col_position) == 0 else 2
                        if not (pd.isnull(exl_1[sheets[idx]].iloc[cell[0], cell[1]]) and
                                pd.isnull(exl_1[sheets[idx]].iloc[cell[0], cell[1]])):
                            exl_1_val = exl_1[sheets[idx]].iloc[cell[0], cell[1]]
                            exl_2_val = exl_2[sheets[idx]].iloc[cell[0], cell[1]]
                            diff_msg = f"In sheet '{sheets[idx]}' " \
                                       f"[row: {cell[0] + row_idx_shift}, " \
                                       f"col: " \
                                       f"{col_position}]: " \
                                       f"{exl_1_val} >>> " \
                                       f"{exl_2_val}"
                            logger.info(diff_msg)
                            diff_annot += f"\t\t{diff_msg}\n"
                            exl_1[sheets[idx]].iloc[cell[0], cell[1]] = f'{exl_1_val} ' \
                                                                        f'>>> {exl_2_val}'
                else:
                    diff_annot += f"\t\tSheet '{sheets[idx]}' " \
                                  f"does not show any differences.\n"

                exl_1[sheets[idx]].to_excel(diff_writer, index=False,
                                            header=True, encoding='utf-8',
                                            sheet_name=sheets[idx])

    else:
        solution_msg = 'Please adjust the sheets of both ' \
                       'files before.'
        logger.error(f'For a differentiated comparison of the excel files, '
                     f'the sheets of both files must match in name and '
                     f'number. {solution_msg}')
        diff_annot += f'The two files to be compared have a different ' \
                      f'number of sheets or have different sheet names. ' \
                      f'An analysis for differences in their sheets is ' \
                      f'therefore not possible. {solution_msg}'
    annot_file = 'ExcelDiff_annotations.txt'
    with open(out_path + annot_file, 'w') as f:
        f.write(diff_annot)

    logger.info('ExcelDiff analysis finished!')


def main():
    logging.getLogger().setLevel(logging.WARNING)
    parser = argparse.ArgumentParser(description="A tool to compare two excel files "
                                                 "with annotation of the differences.")
    parser.add_argument("-i", "--input-files",
                        nargs=2,
                        help="Two paths to the Excel files (.xlsx or .ods format) "
                             "to be compared with each other.",
                        required=True)
    parser.add_argument("-o", "--out-dir",
                        nargs=1,
                        help="Path to the output directory.",
                        required=True)

    parser.add_argument("-v", "--verbose", action='store_true',
                        help="Increase output verbosity", required=False)
    args = parser.parse_args()
    input_files = args.input_files
    out_dir = ''.join(args.out_dir)
    verbose = args.verbose

    if "-h" in sys.argv[1:] or "--help" in sys.argv[1:]:
        parser.print_help()
        exit(0)
    else:
        if input_files and len(input_files) == 2:
            path_1 = Path(input_files[0])
            path_2 = Path(input_files[1])
            if not (path_1.is_file() and path_2.is_file()):
                input_files = None
        if out_dir and not os.path.exists(out_dir):
            os.makedirs(out_dir)
        if verbose:
            logging.getLogger().setLevel(logging.DEBUG)

        if input_files and os.path.exists(out_dir):
            compare_excel_files(input_files[0], input_files[1], out_dir)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        logging.exception(e)
