import logging
import os
from pathlib import Path
import pandas as pd
import numpy as np
import xlsxwriter

from spreadsheetdiff import command_line_options
from spreadsheetdiff import log_handler

logger = log_handler.init_logger()


def style_diffs(res_exl_file: str, diffs: dict, style: list):
    """
    Displays in resulting spreadsheet bold style and/or background color to cells with
    differences between the two input files.
    """
    res_exl_sheets = pd.read_excel(res_exl_file, sheet_name=None)
    workbook = xlsxwriter.Workbook(res_exl_file)
    style_format = {}
    for entry in style:
        style_format[entry[0]] = entry[1]
    diff_format = workbook.add_format(style_format)
    for sheet in res_exl_sheets:
        sheet_data = res_exl_sheets[sheet]
        worksheet = workbook.add_worksheet(sheet)
        col_names = list(sheet_data.columns.values)
        for col in sheet_data:
            col_idx = col_names.index(col)
            for row_idx, _ in sheet_data.iterrows():
                if (row_idx, col_idx) in diffs[sheet]:
                    worksheet.write(row_idx, col_idx,
                                    f"{sheet_data.iloc[row_idx][col]}",
                                    diff_format)
                else:
                    worksheet.write(row_idx,
                                    col_idx,
                                    f"{sheet_data.iloc[row_idx][col]}")
    workbook.close()


def compare_sheets(exl_1, exl_2, sheet, diff_writer, diff_annot):
    """Compares corresponding sheets of input files"""
    logger.debug(f"Analysing sheet '{sheet}'")
    diff_annot += f"\tSheet '{sheet}':\n"
    diffs = []

    if not exl_1[sheet].equals(exl_2[sheet]):
        match_map = exl_1[sheet] == exl_2[sheet]
        diff_rows, diff_cols = np.where(match_map == False)
        for cell in zip(diff_rows, diff_cols):
            col_position = exl_1[sheet].columns.values[cell[1]]
            row_idx_shift = 1 if len(col_position) == 0 else 2
            exl_1_val = exl_1[sheet].iloc[cell[0], cell[1]]
            exl_2_val = exl_2[sheet].iloc[cell[0], cell[1]]
            if not (pd.isnull(exl_1_val) and pd.isnull(exl_2_val)):
                diff_msg = f"In sheet '{sheet}' " \
                           f"[row: {cell[0] + row_idx_shift}, " \
                           f"col: " \
                           f"{col_position}]: " \
                           f"{exl_1_val} >>> " \
                           f"{exl_2_val}"
                logger.debug(diff_msg)
                diff_annot += f"\t\t{diff_msg}\n"
                diffs.append((cell[0], cell[1]))
                exl_1[sheet].iloc[cell[0], cell[1]] = f'{exl_1_val} >>> {exl_2_val}'
    else:
        diff_annot += f"\t\tSheet '{sheet}' does not show any differences.\n"

    exl_1[sheet].to_excel(diff_writer, index=False,
                          header=True, encoding='utf-8',
                          sheet_name=sheet)
    return diff_annot, diffs


def compare_excel_files(excel_1, excel_2, out_dir, style=None):
    """Compares two input files in .xlsx or .ods format"""
    logger.info('Starting SpreadSheetDiff analysis...')
    diff_annot = f'## COMPARISON OF\n##\t{excel_1}\n##\tWITH\n##\t{excel_2}' \
                 f'\n##\n## Differences:\n'
    exl_1 = pd.read_excel(excel_1, sheet_name=None)
    exl_2 = pd.read_excel(excel_2, sheet_name=None)
    out_path = f'{out_dir}/'
    diffs = {}
    if exl_1.keys() == exl_2.keys():
        sheets = list(exl_1.keys())
        res_exl_file = f'{out_path}SpreadSheetDiff_' \
                       f'{Path(os.path.basename(excel_1)).stem}_vs' \
                       f'_{Path(os.path.basename(excel_2)).stem}.xlsx'
        with pd.ExcelWriter(res_exl_file) as diff_writer:
            for idx in range(len(sheets)):
                diff_annot, diffs[sheets[idx]] = compare_sheets(exl_1, exl_2,
                                                                sheets[idx],
                                                                diff_writer, diff_annot)
        if style:
            style_diffs(res_exl_file, diffs, style)
        annot_file = f'{res_exl_file}_annotations.txt'
        with open(out_path + annot_file, 'w') as f:
            f.write(diff_annot)

        logger.info('SpreadSheetDiff analysis finished!')
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


def main():
    input_files, out_dir, style = command_line_options.parse_command_line_opts()
    if input_files and os.path.exists(out_dir):
        compare_excel_files(input_files[0], input_files[1], out_dir, style)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        logging.exception(e)
